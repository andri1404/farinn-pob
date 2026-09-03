"""
LAPORAN HARIAN KEGIATAN PEKERJAAN POB / JURU JARINGAN / PPA
DAERAH IRIGASI RIAM KANAN

Vercel-compatible Flask app:
- Foto di-handle in-memory (BytesIO) — gak butuh disk persistent
- Templates & static relatif ke module dir (PROJECT_DIR = parent of api/)
- Entry point: `app` (WSGI) — Vercel auto-detect
- Single-day generate (PDF) + bulk generate satu bulan (ZIP)

Source layout: LAPORAN HARIAN POB-JURU-PPA.xls
- Tabel 1 — Pemeriksaan Pagi: Waktu, TMA, Status, Cuaca, TMA Pagi, Selfi
- Tabel 2 — Kegiatan Pekerjaan: Jam Mulai/Akhir, Alat, Cuaca, Dokumentasi (rating)
- TTD: Pengamat DI Riam Kanan (Akhmad Muhazir) — Petugas
"""

import os
import sys
# Ensure reportlab/Pillow in user-local site-packages are importable on hosts
# where PYTHONPATH might not be set (e.g. some sandboxed shells).
_USER_SITE = '/home/ubuntu/.local/lib/python3.12/site-packages'
if os.path.isdir(_USER_SITE) and _USER_SITE not in sys.path:
    sys.path.insert(0, _USER_SITE)

import io
import re
import base64
from datetime import datetime, date
from io import BytesIO

from flask import Flask, render_template, request, send_file, jsonify
from reportlab.lib.pagesizes import A4, landscape, letter
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image,
    PageBreak, KeepTogether, Flowable,
)
from reportlab.lib.utils import ImageReader


class ImageStack(Flowable):
    """Flowable that renders N images side-by-side (or stacked) inside a cell.

    Each image is drawn as a ReportLab Image with given size. We override
    wrap() so the cell reserves proper height for ALL images. draw() draws
    each image scaled to fit the cell width.
    """
    def __init__(self, images, max_w=4.5 * cm, max_h=3.0 * cm, gap=4, layout='horizontal'):
        Flowable.__init__(self)
        self.images = images or []
        self.max_w = max_w
        self.max_h = max_h
        self.gap = gap
        self.layout = layout  # 'horizontal' or 'vertical'
        # Force each image's drawWidth/drawHeight for consistent sizing
        for img in self.images:
            try:
                img.drawWidth = self.max_w
                img.drawHeight = self.max_h
            except Exception:
                pass

    def wrap(self, availWidth, availHeight):
        n = max(len(self.images), 1)
        if self.layout == 'vertical':
            per_h = self.max_h
            total_h = n * per_h + (n - 1) * self.gap
        else:  # horizontal: all images in single row, max_h tall
            total_h = self.max_h
        if availHeight and total_h > availHeight:
            total_h = availHeight
        return (self.max_w, total_h)

    def split(self, availWidth, availHeight):
        return []

    def draw(self):
        if not self.images:
            return
        n = len(self.images)
        canvas = self.canv

        if self.layout == 'horizontal':
            # All images side-by-side, scaled to fit total width
            # Total available width = self.max_w, minus gaps
            avail_w = self.width if self.width else self.max_w
            per_w = (avail_w - (n - 1) * self.gap) / n
            # height = self.height
            x = 0
            for img in self.images:
                try:
                    iw, ih = img.imageWidth, img.imageHeight
                except Exception:
                    iw, ih = per_w, self.height
                # Scale to fit per_w x self.height
                scale_w = per_w / iw
                scale_h = self.height / ih
                scale = min(scale_w, scale_h)
                dw = iw * scale
                dh = ih * scale
                img_x = x + (per_w - dw) / 2
                img_y = (self.height - dh) / 2
                try:
                    img.drawOn(canvas, img_x, img_y)
                except Exception as e:
                    import sys
                    print(f'ImageStack.draw error: {e}', file=sys.stderr)
                x += per_w + self.gap
        else:
            # vertical layout
            per_h = (self.height - (n - 1) * self.gap) / n
            y = self.height
            for img in self.images:
                y -= per_h
                try:
                    iw, ih = img.imageWidth, img.imageHeight
                except Exception:
                    iw, ih = self.max_w, per_h
                scale_w = self.max_w / iw
                scale_h = per_h / ih
                scale = min(scale_w, scale_h)
                dw = iw * scale
                dh = ih * scale
                x = (self.max_w - dw) / 2
                try:
                    img.drawOn(canvas, x, y + (per_h - dh) / 2)
                except Exception as e:
                    import sys
                    print(f'ImageStack.draw error: {e}', file=sys.stderr)
                y -= self.gap
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ----------------- App setup -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)  # parent of api/

app = Flask(
    __name__,
    template_folder=os.path.join(PROJECT_DIR, 'templates'),
    static_folder=os.path.join(PROJECT_DIR, 'static'),
    static_url_path='/static',
)
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25 MB


# ----------------- Helpers -----------------

HARI_ID = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
            'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']


def hari_id(date_str):
    """date_str 'YYYY-MM-DD' -> 'Senin/01/09/2026'"""
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
        return f"{HARI_ID[d.weekday()]}/{d.strftime('%d/%m/%Y')}"
    except Exception:
        return date_str


def tgl_indonesia(date_str):
    """date_str 'YYYY-MM-DD' -> '1 September 2026'"""
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
        return f"{d.day} {BULAN_ID[d.month - 1]} {d.year}"
    except Exception:
        return date_str


def decode_images(file_list_files, max_px=900):
    """Decode list of uploaded files -> list of ReportLab Images.
    Returns [] if empty. Each image is resized & compressed in-memory.
    Each image gets a unique filename so ReportLab doesn't de-dup them.
    """
    images = []
    if not file_list_files:
        return images
    for idx, fs in enumerate(file_list_files):
        if not fs or not fs.filename:
            continue
        raw = fs.read()
        if not raw:
            continue
        try:
            from PIL import Image as PILImage
            pil = PILImage.open(BytesIO(raw))
            pil.thumbnail((max_px, max_px))
            if pil.mode in ('RGBA', 'LA', 'P'):
                pil = pil.convert('RGB')
            buf = BytesIO()
            pil.save(buf, format='JPEG', quality=80, optimize=True)
            buf.seek(0)
            # Unique filename per image to prevent ReportLab from de-duplicating
            unique_name = f'img_{idx}_{fs.filename}'
            img = Image(buf, width=3.0 * cm, height=2.2 * cm)
            img.filename = unique_name
            images.append(img)
        except Exception:
            try:
                buf = BytesIO(raw)
                img = Image(buf, width=3.0 * cm, height=2.2 * cm)
                img.filename = f'img_{idx}.jpg'
                images.append(img)
            except Exception:
                continue
    return images


def decode_image(file_storage, max_px=900):
    """Single-file convenience wrapper."""
    result = decode_images(file_storage if isinstance(file_storage, list) else [file_storage], max_px)
    return result[0] if result else None


def decode_b64(b64str, max_px=900):
    """Decode base64 dataURL -> Image."""
    if not b64str:
        return None
    try:
        m = re.match(r'^data:image/[^;]+;base64,(.*)$', b64str)
        payload = m.group(1) if m else b64str
        raw = base64.b64decode(payload)
        from PIL import Image as PILImage
        pil = PILImage.open(BytesIO(raw))
        pil.thumbnail((max_px, max_px))
        if pil.mode in ('RGBA', 'LA', 'P'):
            pil = pil.convert('RGB')
        buf = BytesIO()
        pil.save(buf, format='JPEG', quality=80, optimize=True)
        buf.seek(0)
        return Image(buf, width=3.0 * cm, height=2.2 * cm)
    except Exception:
        return None


# ----------------- PDF builder -----------------

# Layout Letter landscape (792 x 612 pts = 28.0 x 21.6 cm) — matches sample from Excel
PAGE_W, PAGE_H = landscape(letter)
MARGIN = 0.4 * cm
USABLE_W = PAGE_W - 2 * MARGIN  # ~27.2 cm


def build_pdf(meta, pagi_rows, kerja_rows, signature_pengamat, signature_petugas):
    """Build single Laporan PDF in memory -> BytesIO."""
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(letter),
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN,
        title="Laporan Harian POB/JURU/PPA",
        author="Pengamat DI Riam Kanan",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleX', parent=styles['Heading1'],
        fontName='Helvetica-Bold', fontSize=11,
        alignment=TA_CENTER, spaceAfter=2, leading=13,
    )
    sub_style = ParagraphStyle(
        'SubX', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=10,
        alignment=TA_CENTER, spaceAfter=4, leading=12,
    )
    meta_style = ParagraphStyle(
        'Meta', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.5, leading=11, spaceAfter=2,
    )

    story = []

    # ---------- Header ----------
    story.append(Paragraph("LAPORAN HARIAN KEGIATAN PEKERJAAN "
                          "PETUGAS OPERASI BENDUNG / JURU JARINGAN / PPA",
                          title_style))
    story.append(Paragraph("DAERAH IRIGASI RIAM KANAN", sub_style))

    # Metadata block (Nama, Petugas, Hari/Tanggal)
    # Width: 2.0 + 9.5 + 2.5 + 12.8 = 26.8 cm (fits Letter 27.2 cm)
    meta_table = Table(
        [
            [Paragraph("<b>Nama</b>", meta_style),
             Paragraph(": " + str(meta.get('nama', '-')), meta_style),
             Paragraph("<b>Hari / Tanggal</b>", meta_style),
             Paragraph(": " + str(meta.get('hari_tanggal', '-')), meta_style)],
            [Paragraph("<b>Petugas</b>", meta_style),
             Paragraph(": " + str(meta.get('petugas', '-')), meta_style),
             Paragraph("<b>Tanggal Cetak</b>", meta_style),
             Paragraph(": " + str(meta.get('tanggal_cetak', '-')), meta_style)],
        ],
        colWidths=[1.8 * cm, 9.5 * cm, 2.5 * cm, 13.0 * cm],
    )
    meta_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 0.15 * cm))

    # ---------- TABEL 1 - PEMERIKSAAN PAGI ----------
    # 11 cols total. Cols 8=TMA Pagi (text), 9=Selfi (image). Col 10 = filler merged into Selfi for extra width.
    # Total ~27.0 cm to fit Letter landscape (27.2 cm usable)
    col_widths_t1_cm = [
        0.8,  # No
        2.5,  # Hari/Tanggal
        4.0,  # Lokasi
        4.0,  # Jenis
        1.5,  # Waktu
        1.3,  # TMA
        1.6,  # Status
        1.4,  # Cuaca
        2.2,  # TMA Pagi (text + optional photo)
        2.5,  # Selfi (IMAGE)
        5.0,  # filler (merged into Selfi body cells)
    ]
    col_widths_t1 = [w * cm for w in col_widths_t1_cm]

    story.append(Paragraph("<b>Tabel 1. Pemeriksaan Pagi</b>", meta_style))

    # Header baris 1 - 11 cols
    t1_header1 = [
        'No.', 'Hari / Tanggal', 'Titik Lokasi Pekerjaan', 'Jenis Pekerjaan',
        'Pagi', '', '', '',
        'Dokumentasi', '', '',
    ]
    t1_header2 = [
        '', '', '', '',
        'Waktu', 'TMA', 'Status', 'Cuaca',
        'TMA Pagi', 'Selfi', '',
    ]

    t1_data = [t1_header1, t1_header2]
    for i, row in enumerate(pagi_rows, start=1):
        selfi_imgs = row.get('selfi_imgs') or []
        # Each image gets max_h, stack vertically — up to 4 images fit in cell
        selfi_cell = ImageStack(selfi_imgs[:4], max_w=8.5 * cm, max_h=2.2 * cm) if selfi_imgs else ''
        # TMA Pagi: combine text + optional photo
        tma_pagi_text = str(row.get('tma_pagi', ''))
        tma_pagi_imgs = row.get('tma_pagi_imgs') or []
        if tma_pagi_imgs:
            tma_pagi_cell = [
                Paragraph(tma_pagi_text, meta_style) if tma_pagi_text else '',
                ImageStack(tma_pagi_imgs[:2], max_w=2.0 * cm, max_h=1.5 * cm, layout='vertical')
            ]
            tma_pagi_rendered = tma_pagi_cell
        else:
            tma_pagi_rendered = tma_pagi_text
        t1_data.append([
            str(i),
            str(row.get('hari_tanggal', '')),
            Paragraph(str(row.get('lokasi', '')), meta_style),
            Paragraph(str(row.get('jenis', '')), meta_style),
            str(row.get('waktu', '')),
            str(row.get('tma', '')),
            str(row.get('status', '')),
            str(row.get('cuaca', '')),
            tma_pagi_rendered,              # TMA Pagi: text + optional photo
            selfi_cell,                      # Selfi = 1+ images stacked
            '',
        ])

    # Pad rows to minimum 3 visible rows even if empty
    while len(t1_data) < 5:
        t1_data.append([''] * 11)

    t1 = Table(t1_data, colWidths=col_widths_t1, repeatRows=2)
    t1.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 1), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 1), 7.5),
        ('ALIGN', (0, 0), (-1, 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 1), 'MIDDLE'),
        ('FONTSIZE', (0, 2), (-1, -1), 8),
        ('ALIGN', (0, 2), (0, -1), 'CENTER'),
        ('ALIGN', (4, 2), (8, -1), 'CENTER'),
        ('VALIGN', (0, 2), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('MINROWHEIGHT', (0, 2), (-1, -1), 1.0 * cm),
        ('MINROWHEIGHT', (9, 2), (10, -1), 9.0 * cm),  # Selfi col grows with images
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (1, 1)),
        ('SPAN', (2, 0), (2, 1)),
        ('SPAN', (3, 0), (3, 1)),
        ('SPAN', (4, 0), (7, 0)),    # Pagi (4 cols)
        ('SPAN', (8, 0), (10, 0)),   # Dokumentasi (TMA Pagi + Selfi + filler)
        # Body: merge filler (col 10) into Selfi (col 9) so image gets full width
        ('SPAN', (9, 2), (10, -1)),
    ]))
    story.append(t1)
    story.append(Spacer(1, 0.25 * cm))

    # ---------- TABEL 2 - KEGIATAN PEKERJAAN ----------
    # 11 cols: No | Hari/Tgl | Lokasi | Jenis | Jam Mulai | Jam Akhir | Cuaca | Alat | Foto0 | Foto0.5 | Foto1
    # Total ~27.0 cm to fit Letter landscape (27.2 cm usable)
    col_widths_t2_cm = [
        0.7,  # No
        2.4,  # Hari/Tanggal
        3.5,  # Lokasi
        3.5,  # Jenis
        1.4,  # Jam Mulai
        1.4,  # Jam Akhir
        1.3,  # Cuaca
        2.2,  # Alat
        3.4,  # Foto 0
        3.4,  # Foto 1
        3.5,  # Foto 2 (was Foto 1)
    ]
    col_widths_t2 = [w * cm for w in col_widths_t2_cm]

    story.append(Paragraph("<b>Tabel 2. Kegiatan Pekerjaan</b>", meta_style))

    t2_header1 = [
        'No.', 'Hari / Tanggal', 'Titik Lokasi Pekerjaan', 'Jenis Pekerjaan',
        'Jam', '', 'Cuaca', 'Alat yang Digunakan',
        'Dokumentasi', '', '',
    ]
    t2_header2 = [
        '', '', '', '',
        'Mulai', 'Akhir', '', '',
        'Foto 0', 'Foto 0.5', 'Foto 1',
    ]

    t2_data = [t2_header1, t2_header2]
    for i, row in enumerate(kerja_rows, start=1):
        f0 = row.get('foto0_imgs') or []
        f05 = row.get('foto05_imgs') or []
        f1 = row.get('foto1_imgs') or []
        t2_data.append([
            str(i),
            str(row.get('hari_tanggal', '')),
            Paragraph(str(row.get('lokasi', '')), meta_style),
            Paragraph(str(row.get('jenis', '')), meta_style),
            str(row.get('jam_mulai', '')),
            str(row.get('jam_akhir', '')),
            str(row.get('cuaca', '')),
            Paragraph(str(row.get('alat', '')), meta_style),
            ImageStack(f0, max_w=3.2 * cm, max_h=3.0 * cm) if f0 else '',
            ImageStack(f05, max_w=3.2 * cm, max_h=3.0 * cm) if f05 else '',
            ImageStack(f1, max_w=3.6 * cm, max_h=3.0 * cm) if f1 else '',
        ])

    while len(t2_data) < 5:
        t2_data.append([''] * 11)

    t2 = Table(t2_data, colWidths=col_widths_t2, repeatRows=2)
    t2.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 1), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 1), 7.5),
        ('ALIGN', (0, 0), (-1, 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 1), 'MIDDLE'),
        ('FONTSIZE', (0, 2), (-1, -1), 8),
        ('ALIGN', (0, 2), (0, -1), 'CENTER'),
        ('ALIGN', (4, 2), (6, -1), 'CENTER'),
        ('VALIGN', (0, 2), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('MINROWHEIGHT', (0, 2), (-1, -1), 1.6 * cm),
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (1, 1)),
        ('SPAN', (2, 0), (2, 1)),
        ('SPAN', (3, 0), (3, 1)),
        ('SPAN', (4, 0), (5, 0)),    # Jam (Mulai+Akhir)
        ('SPAN', (6, 0), (6, 1)),    # Cuaca
        ('SPAN', (7, 0), (7, 1)),    # Alat
        ('SPAN', (8, 0), (10, 0)),   # Dokumentasi (3 foto cols)
    ]))
    story.append(t2)
    story.append(Spacer(1, 0.35 * cm))

    # ---------- TTD ----------
    ttd_data = [[
        Paragraph("<b>Mengetahui :</b><br/>Pengamat DI. Riam Kanan", meta_style),
        Paragraph("<b>Dibuat oleh :</b><br/>" + str(meta.get('petugas_label', 'Petugas')), meta_style),
    ], [
        signature_pengamat or Paragraph('<br/><br/><br/><br/>', meta_style),
        signature_petugas or Paragraph('<br/><br/><br/><br/>', meta_style),
    ], [
        Paragraph("<b><u>" + str(meta.get('pengamat', 'AKHMAD MUHAZIR')) + "</u></b>", meta_style),
        Paragraph("<b><u>" + str(meta.get('nama', '.........................')) + "</u></b>", meta_style),
    ], [
        Paragraph("NIP. " + str(meta.get('pengamat_nip', '.................................')), meta_style),
        Paragraph("NIP. " + str(meta.get('petugas_nip', '.................................')), meta_style),
    ]]

    ttd = Table(ttd_data, colWidths=[14 * cm, 14 * cm])
    ttd.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    story.append(ttd)

    # ---------- Catatan kecil ----------
    story.append(Spacer(1, 0.2 * cm))
    note = Paragraph(
        "<i>Dokumen dicetak otomatis oleh sistem - Martapura, "
        + tgl_indonesia(datetime.now().strftime('%Y-%m-%d')) + "</i>",
        ParagraphStyle('Note', parent=meta_style, fontSize=7,
                       textColor=colors.grey, alignment=TA_CENTER),
    )
    story.append(note)

    doc.build(story)
    buf.seek(0)
    return buf


# ----------------- Form parsing -----------------

def parse_form_pagi():
    """Parse repeated 'pagi-*' fields from request.form -> list of dicts.
    Image fields use request.files.getlist() to support multiple uploads.
    TMA Pagi supports BOTH text value AND photo upload (optional).
    """
    rows = []
    idx_set = set()
    for k in request.form.keys():
        m = re.match(r'^pagi-(\d+)-', k)
        if m:
            idx_set.add(int(m.group(1)))
    for i in sorted(idx_set):
        rows.append({
            'hari_tanggal': request.form.get(f'pagi-{i}-hari_tanggal', '').strip(),
            'lokasi': request.form.get(f'pagi-{i}-lokasi', '').strip(),
            'jenis': request.form.get(f'pagi-{i}-jenis', '').strip(),
            'waktu': request.form.get(f'pagi-{i}-waktu', '').strip(),
            'tma': request.form.get(f'pagi-{i}-tma', '').strip(),
            'status': request.form.get(f'pagi-{i}-status', '').strip(),
            'cuaca': request.form.get(f'pagi-{i}-cuaca', '').strip(),
            'tma_pagi': request.form.get(f'pagi-{i}-tma_pagi', '').strip(),
            'tma_pagi_imgs': decode_images(request.files.getlist(f'pagi-{i}-tma_pagi_foto')),
            'selfi_imgs': decode_images(request.files.getlist(f'pagi-{i}-selfi')),
        })
    return rows


def parse_form_kerja():
    """Parse repeated 'kerja-*' fields. Image fields support multiple uploads."""
    rows = []
    idx_set = set()
    for k in request.form.keys():
        m = re.match(r'^kerja-(\d+)-', k)
        if m:
            idx_set.add(int(m.group(1)))
    for i in sorted(idx_set):
        rows.append({
            'hari_tanggal': request.form.get(f'kerja-{i}-hari_tanggal', '').strip(),
            'lokasi': request.form.get(f'kerja-{i}-lokasi', '').strip(),
            'jenis': request.form.get(f'kerja-{i}-jenis', '').strip(),
            'jam_mulai': request.form.get(f'kerja-{i}-jam_mulai', '').strip(),
            'jam_akhir': request.form.get(f'kerja-{i}-jam_akhir', '').strip(),
            'alat': request.form.get(f'kerja-{i}-alat', '').strip(),
            'cuaca': request.form.get(f'kerja-{i}-cuaca', '').strip(),
            'foto0_imgs': decode_images(request.files.getlist(f'kerja-{i}-foto0')),
            'foto05_imgs': decode_images(request.files.getlist(f'kerja-{i}-foto05')),
            'foto1_imgs': decode_images(request.files.getlist(f'kerja-{i}-foto1')),
        })
    return rows


def auto_fill_dates(pagi_rows, kerja_rows, base_date):
    """If user left 'Hari/Tanggal' empty, auto-fill from base_date."""
    formatted = hari_id(base_date)
    for r in pagi_rows:
        if not r.get('hari_tanggal'):
            r['hari_tanggal'] = formatted
    for r in kerja_rows:
        if not r.get('hari_tanggal'):
            r['hari_tanggal'] = formatted


# ----------------- Routes -----------------

@app.route('/')
def index():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('form.html', default_date=today, bulan_id=BULAN_ID)


@app.route('/preview', methods=['POST'])
def preview():
    """Show preview of what the PDF will look like (HTML mockup)."""
    meta = {
        'nama': request.form.get('nama', 'Muhammad Yasir'),
        'petugas': request.form.get('petugas', ''),
        'petugas_label': request.form.get('petugas_label', 'Petugas'),
        'pengamat': request.form.get('pengamat', 'AKHMAD MUHAZIR'),
        'pengamat_nip': request.form.get('pengamat_nip', ''),
        'petugas_nip': request.form.get('petugas_nip', ''),
        'hari_tanggal': hari_id(request.form.get('tanggal', datetime.now().strftime('%Y-%m-%d'))),
        'tanggal_cetak': tgl_indonesia(datetime.now().strftime('%Y-%m-%d')),
    }
    pagi = parse_form_pagi()
    kerja = parse_form_kerja()
    auto_fill_dates(pagi, kerja, request.form.get('tanggal', datetime.now().strftime('%Y-%m-%d')))
    return render_template('preview.html', meta=meta, pagi=pagi, kerja=kerja)


@app.route('/generate', methods=['POST'])
def generate():
    meta = {
        'nama': request.form.get('nama', 'Muhammad Yasir'),
        'petugas': request.form.get('petugas', ''),
        'petugas_label': request.form.get('petugas_label', 'Petugas'),
        'pengamat': request.form.get('pengamat', 'AKHMAD MUHAZIR'),
        'pengamat_nip': request.form.get('pengamat_nip', ''),
        'petugas_nip': request.form.get('petugas_nip', ''),
        'hari_tanggal': hari_id(request.form.get('tanggal', datetime.now().strftime('%Y-%m-%d'))),
        'tanggal_cetak': tgl_indonesia(datetime.now().strftime('%Y-%m-%d')),
    }
    pagi = parse_form_pagi()
    kerja = parse_form_kerja()
    auto_fill_dates(pagi, kerja, request.form.get('tanggal', datetime.now().strftime('%Y-%m-%d')))

    sig_pengamat = decode_image(request.files.get('signature_pengamat'))
    sig_petugas = decode_image(request.files.get('signature_petugas'))

    pdf = build_pdf(meta, pagi, kerja, sig_pengamat, sig_petugas)
    fname = "Laporan-POB-" + request.form.get('tanggal', datetime.now().strftime('%Y-%m-%d')) + ".pdf"
    return send_file(pdf, mimetype='application/pdf', as_attachment=True,
                     download_name=fname)


@app.route('/healthz')
def healthz():
    return jsonify({'status': 'ok', 'service': 'farinn-pob'})


# ----------------- WSGI -----------------
# Vercel auto-detects `app`; locally we run dev server
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
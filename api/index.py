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
# Load user-local site-packages when available (for Vercel build compat)
for sp in ('/home/ubuntu/.local/lib/python3.12/site-packages',
          '/home/ubuntu/.local/lib/python3.11/site-packages'):
    if os.path.isdir(sp) and sp not in sys.path:
        sys.path.append(sp)

import io
import re
import base64
from datetime import datetime, date
from io import BytesIO

from flask import Flask, render_template, request, send_file, jsonify, redirect
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

    - Always renders within availWidth (cell width) for proper centering.
    - Images preserve aspect ratio (uniform scale based on PIXEL dimensions).
    - All images in same row get SAME height (h = max_h) for visual consistency.
    """
    def __init__(self, images, max_w=4.4 * cm, max_h=3.0 * cm, gap=4, layout='horizontal'):
        Flowable.__init__(self)
        self.images = images or []
        self.max_w = max_w
        self.max_h = max_h
        self.gap = gap
        self.layout = layout  # 'horizontal' or 'vertical'

    def wrap(self, availWidth, availHeight):
        self._avail_w = availWidth
        self._avail_h = availHeight
        n = max(len(self.images), 1)
        if self.layout == 'vertical':
            total_h = n * self.max_h + (n - 1) * self.gap
        else:
            total_h = self.max_h
        if availHeight and total_h > availHeight:
            total_h = availHeight
        self._h = total_h
        return (availWidth, total_h)

    def split(self, availWidth, availHeight):
        return []

    def draw(self):
        if not self.images:
            return
        n = len(self.images)
        canvas = self.canv
        cell_w = self._avail_w
        cell_h = getattr(self, '_h', None) or self.max_h
        if self.layout == 'horizontal':
            slot_w = (cell_w - (n - 1) * self.gap) / n
            slot_h = cell_h
        else:
            slot_w = cell_w
            slot_h = (cell_h - (n - 1) * self.gap) / n
        # FILL mode: stretch setiap image persis mengisi slot-nya
        # (edge-to-edge, sesuai garis cell) — seperti formulir asli.
        for i, img in enumerate(self.images):
            if self.layout == 'horizontal':
                img_x = i * (slot_w + self.gap)
                img_y = 0
            else:
                img_x = 0
                img_y = cell_h - (i + 1) * slot_h - i * self.gap
            img.drawWidth = slot_w
            img.drawHeight = slot_h
            try:
                img.drawOn(canvas, img_x, img_y)
            except Exception as e:
                import sys
                print(f'ImageStack.draw error: {e}', file=sys.stderr)
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
# Limit Vercel untuk request body = 4.5 MB. Disamakan di sini supaya Flask balas
# 413 dengan halaman ramah, bukan error mentah. Foto dikompres di browser dulu.
app.config['MAX_CONTENT_LENGTH'] = 4500000  # 4.5 MB (batas function Vercel)


# ----------------- Helpers -----------------

HARI_ID = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
            'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']


def hari_id(date_str):
    """date_str 'YYYY-MM-DD' -> 'Senin/01/09/2026'"""
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
        return d.strftime('%d/%m/%Y')
    except Exception:
        return date_str


def tgl_indonesia(date_str):
    """date_str 'YYYY-MM-DD' -> '1 September 2026'"""
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
        return f"{d.day} {BULAN_ID[d.month - 1]} {d.year}"
    except Exception:
        return date_str


def _process_image(raw_bytes, max_px=900):
    """Decode + auto-rotate (EXIF) + resize image. Returns (buf, orig_w, orig_h)."""
    from PIL import Image as PILImage, ImageOps
    pil = PILImage.open(BytesIO(raw_bytes))
    # Auto-rotate based on EXIF orientation tag (HP/smartphone photos)
    pil = ImageOps.exif_transpose(pil)
    # Resize (preserves aspect ratio)
    pil.thumbnail((max_px, max_px))
    # Convert mode
    if pil.mode in ('RGBA', 'LA', 'P'):
        pil = pil.convert('RGB')
    orig_w, orig_h = pil.size
    buf = BytesIO()
    pil.save(buf, format='JPEG', quality=80, optimize=True)
    buf.seek(0)
    return buf, orig_w, orig_h


# Default natural render size used by decode_images when no target given.
# Matches sample PDF aspect (~1.45:1) and gives consistent visuals in PDF.
DEFAULT_RW = 4.2 * cm
DEFAULT_RH = 2.9 * cm


def decode_images(file_list_files, max_px=900, target_w=None, target_h=None):
    """Decode list of uploaded files -> list of ReportLab Images.

    - Auto-rotates via EXIF
    - Resizes to max_px preserving aspect ratio
    - target_w/target_h: explicit render size. If None, uses DEFAULT_RW/DEFAULT_RH
      with aspect ratio adjustment.
    Each image gets a unique filename so ReportLab doesn't de-dup.
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
            buf, iw, ih = _process_image(raw, max_px)
            # Determine render size
            if target_w is not None and target_h is not None:
                rw, rh = target_w, target_h
            else:
                # Use DEFAULT_RW as max width, scale height by aspect
                aspect = iw / ih if ih > 0 else 1.4
                rw = DEFAULT_RW
                rh = rw / aspect
                # Cap height too
                if rh > DEFAULT_RH:
                    rh = DEFAULT_RH
                    rw = rh * aspect
            unique_name = f'img_{idx}_{fs.filename}'
            img = Image(buf, width=rw, height=rh)
            img.filename = unique_name
            images.append(img)
        except Exception:
            try:
                buf = BytesIO(raw)
                img = Image(buf, width=DEFAULT_RW, height=DEFAULT_RH)
                img.filename = f'img_{idx}.jpg'
                images.append(img)
            except Exception:
                continue
    return images


def decode_image(file_storage, max_px=900, target_w=None, target_h=None):
    """Single-file convenience wrapper."""
    result = decode_images(
        file_storage if isinstance(file_storage, list) else [file_storage],
        max_px, target_w, target_h
    )
    return result[0] if result else None


def decode_b64(b64str, max_px=900, target_w=None, target_h=None):
    """Decode base64 dataURL -> Image."""
    if not b64str:
        return None
    try:
        m = re.match(r'^data:image/[^;]+;base64,(.*)$', b64str)
        payload = m.group(1) if m else b64str
        raw = base64.b64decode(payload)
        buf, iw, ih = _process_image(raw, max_px)
        if target_w is not None and target_h is not None:
            rw, rh = target_w, target_h
        else:
            aspect = iw / ih
            if aspect > 1:
                rw = 3.0 * cm
                rh = rw / aspect
            else:
                rh = 3.0 * cm
                rw = rh * aspect
        return Image(buf, width=rw, height=rh)
    except Exception:
        return None


# ----------------- PDF builder -----------------

# Layout A4 landscape (841.89 x 595.28 pts = 29.7 x 21.0 cm)
# Sumber: file .xls asli -> BIFF SETUP record paper_size=9 (A4), flag landscape.
PAGE_W, PAGE_H = landscape(A4)
# Margin disamakan dengan .xls: kolom "No." di .xls mulai di x=34pt dari tepi kiri
# (0.55cm terlalu mepet). 1.1cm bikin inset kiri/kanan mirip cetakan .xls.
MARGIN = 1.1 * cm
USABLE_W = PAGE_W - 2 * MARGIN  # ~27.5 cm


def build_pdf(meta, pagi_rows, kerja_rows, signature_pengamat, signature_petugas):
    """Build single Laporan PDF in memory -> BytesIO."""
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
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
        alignment=TA_CENTER,
    )
    cell_style = ParagraphStyle(
        'Cell', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8, leading=10,
        alignment=TA_CENTER,
    )

    story = []

    # ---------- Header Kementerian PU ----------
    # Logo PU di kiri + 5 baris alamat di kanan + garis separator di bawah
    pu_logo_path = os.path.join(PROJECT_DIR, 'static', 'assets', 'logo-pu.png')
    pu_logo = None
    if os.path.exists(pu_logo_path):
        try:
            pu_logo = Image(pu_logo_path, width=2.2 * cm, height=2.2 * cm)
        except Exception:
            pu_logo = None

    pu_addr_style = ParagraphStyle(
        'PuAddr', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=9.5, leading=12.5, alignment=TA_LEFT,
    )
    pu_addr_top = ParagraphStyle(
        'PuAddrTop', parent=pu_addr_style, fontSize=10.5, leading=13.5,
    )

    pu_lines = [
        Paragraph('KEMENTERIAN PEKERJAAN UMUM', pu_addr_top),
        Paragraph('DIREKTORAT JENDERAL SUMBER DAYA AIR', pu_addr_style),
        Paragraph('BALAI WILAYAH SUNGAI KALIMANTAN III BANJARMASIN', pu_addr_style),
        Paragraph('SATUAN KERJA OPERASI DAN PEMELIHARAAN SUMBER DAYA AIR KALIMANTAN III', pu_addr_style),
        Paragraph('Jl. Pemajatan KM. 1 Gambut 70652 Kab. Banjar Prov. Kalimantan Selatan, '
                  'Telepon/Faksimili (0511) 6775967', ParagraphStyle(
                      'PuAddrSm', parent=pu_addr_style, fontName='Helvetica', fontSize=7)),
    ]

    header_left = pu_logo if pu_logo else Paragraph('', pu_addr_style)
    header_right_cells = [[line] for line in pu_lines]
    header_right = Table(header_right_cells, colWidths=[USABLE_W - 2.5 * cm])
    header_right.setStyle(TableStyle([
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))

    header_row = Table(
        [[header_left, header_right]],
        colWidths=[2.5 * cm, USABLE_W - 2.5 * cm],
    )
    header_row.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_row)

    # Garis separator hitam di bawah header PU
    sep_table = Table([['']], colWidths=[USABLE_W], rowHeights=[0.05 * cm])
    sep_table.setStyle(TableStyle([
        ('LINEBELOW', (0, 0), (-1, -1), 1.4, colors.black),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(sep_table)
    story.append(Spacer(1, 0.18 * cm))

    # ---------- Judul Laporan ----------
    story.append(Paragraph("LAPORAN HARIAN KEGIATAN PEKERJAAN "
                          "PETUGAS OPERASI BENDUNG / JURU JARINGAN / PPA",
                          title_style))
    story.append(Paragraph("DAERAH IRIGASI RIAM KANAN", sub_style))

    # Metadata block — persis .xls: label rata KIRI di tepi margin (x~31.5pt) dan
    # titik dua di x~122.6pt (jarak 91pt = 3.21cm dari label).
    meta_left = ParagraphStyle('MetaLeft', parent=meta_style, alignment=TA_LEFT)
    meta_table = Table(
        [
            [Paragraph("Nama", meta_left),
             Paragraph(": " + str(meta.get('nama', '-')), meta_left)],
            [Paragraph("Petugas", meta_left),
             Paragraph(": " + str(meta.get('petugas', '-')), meta_left)],
        ],
        colWidths=[3.21 * cm, USABLE_W - 3.21 * cm],
        hAlign='LEFT',
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
    # 10 cols total. Header group: Pagi (4 sub-cols), Dokumentasi (TMA Pagi + Selfi).
    # Proporsi lebar kolom diambil dari lebar kolom .xls asli (col width /256 char),
    # di-scale ke USABLE_W A4 landscape = 28.6 cm.
    #   xls: No=1137, Hari=2929, Lokasi=5774, Jenis=5774, Waktu=2190, TMA=2218,
    #        Status=2218, Cuaca=2190, TMA Pagi=3*2759, Selfi=3*2759
    # Status dikasih sedikit ekstra (2.4cm) biar "Tidak Normal" gak ke-split jelek.
    col_widths_t1_frac = [
        0.0262,  # No
        0.0682,  # Hari/Tanggal
        0.1311,  # Titik Lokasi Pekerjaan
        0.1311,  # Jenis Pekerjaan
        0.0507,  # Waktu
        0.0524,  # TMA
        0.0839,  # Status  (dikasih ekstra biar "Tidak Normal" gak ke-split jelek)
        0.0524,  # Cuaca
        0.2021,  # Foto TMA Pagi
        0.2018,  # Selfi (foto)
    ]  # sum = 1.0
    col_widths_t1 = [w * USABLE_W for w in col_widths_t1_frac]

    # Header baris 1 - 10 cols (Pagi + Dokumentasi merged)
    # Header pakai Paragraph biar auto-wrap + center, gak nabrak cell sebelah
    header_style = ParagraphStyle(
        'Header', parent=cell_style,
        fontName='Helvetica-Bold', fontSize=7.5, leading=9,
        alignment=TA_CENTER,
    )
    t1_header1 = [
        Paragraph('No.', header_style),
        Paragraph('Hari /<br/>Tanggal', header_style),
        Paragraph('Titik Lokasi Pekerjaan', header_style),
        Paragraph('Jenis Pekerjaan', header_style),
        Paragraph('Pagi', header_style),
        '', '', '',
        Paragraph('Dokumentasi', header_style),
        '',
    ]
    t1_header2 = [
        '', '', '', '',
        Paragraph('Waktu', header_style),
        Paragraph('TMA', header_style),
        Paragraph('Status', header_style),
        Paragraph('Cuaca', header_style),
        Paragraph('TMA Pagi', header_style),
        Paragraph('Selfi', header_style),
    ]

    # Helper: ensure Status text doesn't overflow cell by manual word wrap.
    # ReportLab Paragraph uses XML for formatting; <br/> forces line break.
    def wrap_status(text, max_len=10):
        if not text or len(text) <= max_len:
            return text
        words = text.split()
        lines, cur = [], ''
        for w in words:
            if len(cur) + len(w) + 1 > max_len and cur:
                lines.append(cur)
                cur = w
            else:
                cur = (cur + ' ' + w).strip()
        if cur:
            lines.append(cur)
        # <br/> = explicit line break in ReportLab Paragraph XML
        return '<br/>'.join(lines)

    t1_data = [t1_header1, t1_header2]
    for i, row in enumerate(pagi_rows, start=1):
        selfi_imgs = row.get('selfi_imgs') or []
        # Selfi cell: 1-3 foto, side-by-side kalau <=2, stacked kalau 3+
        # Selfi col = 4.4cm. Padding 2 left + 2 right = 4.0cm effective.
        # 1 foto: max_w 4.0cm. 2 foto: per-image ~2.0cm. 3 foto stacked: 1.5cm each.
        if selfi_imgs:
            if len(selfi_imgs) == 1:
                selfi_cell = ImageStack(selfi_imgs, max_w=4.0 * cm, max_h=3.2 * cm, layout='horizontal', gap=0)
            elif len(selfi_imgs) == 2:
                selfi_cell = ImageStack(selfi_imgs, max_w=2.0 * cm, max_h=3.2 * cm, layout='horizontal', gap=2)
            else:
                selfi_cell = ImageStack(selfi_imgs[:3], max_w=4.0 * cm, max_h=0.9 * cm, layout='vertical', gap=2)
        else:
            selfi_cell = ''
        # TMA Pagi cell: FOTO SAJA (input teks dihapus — permintaan user 03/10).
        # Cell 5.78cm x 3.2cm, padding 0 -> foto stretch penuh mengisi cell.
        tma_pagi_imgs = row.get('tma_pagi_imgs') or []
        if tma_pagi_imgs:
            tma_pagi_rendered = ImageStack(
                tma_pagi_imgs[:1], max_w=5.78 * cm, max_h=3.2 * cm,
                layout='vertical', gap=0,
            )
        else:
            tma_pagi_rendered = ''
        status_text = wrap_status(str(row.get('status', '')))
        t1_data.append([
            str(i),
            str(row.get('hari_tanggal', '')),
            Paragraph(str(row.get('lokasi', '')), cell_style),
            Paragraph(str(row.get('jenis', '')), cell_style),
            str(row.get('waktu', '')),
            str(row.get('tma', '')),
            Paragraph(status_text, cell_style),  # Wrapped & centered
            str(row.get('cuaca', '')),
            tma_pagi_rendered,
            selfi_cell,
        ])

    # .xls cuma punya 1 baris body per tabel -> jangan tambahin baris kosong banyak.
    while len(t1_data) < 3:
        t1_data.append([''] * 10)

    # rowHeights: baris berisi data/foto = 3.2cm (ruang foto, sama seperti .xls),
    # baris kosong tambahan = 0.85cm. Baris body pertama selalu tinggi.
    t1_row_heights = [0.8 * cm, 0.8 * cm]
    for r in range(2, len(t1_data)):
        has = False
        if r - 2 < len(pagi_rows):
            _r = pagi_rows[r - 2]
            has = bool(_r.get('selfi_imgs') or _r.get('tma_pagi_imgs') or any(
                str(_r.get(k) or '').strip()
                for k in ('lokasi', 'jenis', 'waktu', 'tma', 'status', 'cuaca')))
        t1_row_heights.append(3.2 * cm if has else 0.85 * cm)
    if t1_row_heights[2] < 3.0 * cm:
        t1_row_heights[2] = 3.2 * cm      # baris pertama selalu tinggi (ruang foto)

    t1 = Table(t1_data, colWidths=col_widths_t1, rowHeights=t1_row_heights, repeatRows=2)
    t1.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 1), colors.yellow),
        ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 1), 7.5),
        ('ALIGN', (0, 0), (-1, 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 1), 'MIDDLE'),
        ('FONTSIZE', (0, 2), (-1, -1), 8),
        ('ALIGN', (0, 2), (-1, -1), 'CENTER'),     # All body cells centered
        ('VALIGN', (0, 2), (-1, -1), 'MIDDLE'),  # Text cols vertical center
        ('LEFTPADDING', (0, 0), (-1, 4), 3),     # Text cols padding 3
        ('RIGHTPADDING', (0, 0), (-1, 4), 3),
        ('TOPPADDING', (0, 0), (-1, 4), 3),
        ('BOTTOMPADDING', (0, 0), (-1, 4), 3),
        # Foto cols (8=TMA Pagi, 9=Selfi): padding 5 top/bottom = 0.18cm
        # supaya image centered (row=3.2cm, image max=2.9cm, sisa 0.3cm = 0.15cm each)
        # Foto cols (8=TMA Pagi, 9=Selfi): padding 0 biar image nempel ke garis cell
        ('LEFTPADDING', (8, 0), (9, -1), 0),
        ('RIGHTPADDING', (8, 0), (9, -1), 0),
        ('TOPPADDING', (8, 0), (9, -1), 0),
        ('BOTTOMPADDING', (8, 0), (9, -1), 0),
        ('MINROWHEIGHT', (0, 2), (-1, -1), 0.85 * cm),  # baris kosong tetap punya tinggi minimum
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (1, 1)),
        ('SPAN', (2, 0), (2, 1)),
        ('SPAN', (3, 0), (3, 1)),
        ('SPAN', (4, 0), (7, 0)),    # Pagi (4 cols: Waktu/TMA/Status/Cuaca)
        ('SPAN', (8, 0), (9, 0)),    # Dokumentasi (TMA Pagi + Selfi)
    ]))
    story.append(t1)
    story.append(Spacer(1, 0.25 * cm))

    # ---------- TABEL 2 - KEGIATAN PEKERJAAN ----------
    # 11 cols, URUTAN PERSIS .xls:
    #   No | Hari/Tgl | Titik Lokasi | Jenis | Jam Mulai | Jam Akhir | Alat | Cuaca |
    #   Dokumentasi{0% , 50% , 100%}
    # (Catatan: versi lama salah urut — Cuaca sebelum Alat; .xls = Alat dulu baru Cuaca.)
    # Lebar diambil dari proporsi kolom .xls, di-scale ke USABLE_W = 28.6 cm.
    col_widths_t2_frac = [
        0.0262,  # No
        0.0682,  # Hari/Tanggal
        0.1329,  # Titik Lokasi Pekerjaan
        0.1329,  # Jenis Pekerjaan
        0.0507,  # Jam Mulai
        0.0514,  # Jam Akhir
        0.1014,  # Alat yang Digunakan
        0.0524,  # Cuaca
        0.1280,  # Dokumentasi 0%
        0.1280,  # Dokumentasi 50%
        0.1279,  # Dokumentasi 100%
    ]  # sum = 1.0
    col_widths_t2 = [w * USABLE_W for w in col_widths_t2_frac]

    t2_header1 = [
        Paragraph('No.', header_style),
        Paragraph('Hari /<br/>Tanggal', header_style),
        Paragraph('Titik Lokasi Pekerjaan', header_style),
        Paragraph('Jenis Pekerjaan', header_style),
        Paragraph('Jam', header_style),
        '',
        Paragraph('Alat yang<br/>Digunakan', header_style),
        Paragraph('Cuaca', header_style),
        Paragraph('Dokumentasi', header_style),
        '', '',
    ]
    t2_header2 = [
        '', '', '', '',
        Paragraph('Mulai', header_style),
        Paragraph('Akhir', header_style),
        '', '',
        Paragraph('0%', header_style),
        Paragraph('50%', header_style),
        Paragraph('100%', header_style),
    ]

    t2_data = [t2_header1, t2_header2]
    for i, row in enumerate(kerja_rows, start=1):
        f0 = row.get('foto0_imgs') or []
        f05 = row.get('foto05_imgs') or []
        f1 = row.get('foto1_imgs') or []
        t2_data.append([
            str(i),
            str(row.get('hari_tanggal', '')),
            Paragraph(str(row.get('lokasi', '')), cell_style),
            Paragraph(str(row.get('jenis', '')), cell_style),
            str(row.get('jam_mulai', '')),
            str(row.get('jam_akhir', '')),
            Paragraph(str(row.get('alat', '')), cell_style),   # Alat (sebelum Cuaca, sesuai .xls)
            str(row.get('cuaca', '')),
            # Dokumentasi 0% / 50% / 100%. Cell ~3.66cm, padding 0 -> image full cell.
            ImageStack(f0, max_w=3.6 * cm, max_h=3.2 * cm) if f0 else '',
            ImageStack(f05, max_w=3.6 * cm, max_h=3.2 * cm) if f05 else '',
            ImageStack(f1, max_w=3.6 * cm, max_h=3.2 * cm) if f1 else '',
        ])

    # .xls cuma punya 1 baris body per tabel
    while len(t2_data) < 3:
        t2_data.append([''] * 11)

    # rowHeights: baris berisi data/foto = 3.2cm, baris kosong tambahan = 0.85cm.
    t2_row_heights = [0.8 * cm, 0.8 * cm]
    for r in range(2, len(t2_data)):
        has = False
        if r - 2 < len(kerja_rows):
            _k = kerja_rows[r - 2]
            has = bool(_k.get('foto0_imgs') or _k.get('foto05_imgs') or _k.get('foto1_imgs') or any(
                str(_k.get(k) or '').strip()
                for k in ('lokasi', 'jenis', 'jam_mulai', 'jam_akhir', 'alat', 'cuaca')))
        t2_row_heights.append(3.2 * cm if has else 0.85 * cm)
    if t2_row_heights[2] < 3.0 * cm:
        t2_row_heights[2] = 3.2 * cm      # baris pertama selalu tinggi (ruang foto)

    t2 = Table(t2_data, colWidths=col_widths_t2, rowHeights=t2_row_heights, repeatRows=2)
    t2.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 1), colors.yellow),
        ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 1), 7.5),
        ('ALIGN', (0, 0), (-1, 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 1), 'MIDDLE'),
        ('FONTSIZE', (0, 2), (-1, -1), 8),
        ('ALIGN', (0, 2), (-1, -1), 'CENTER'),     # All body cells centered
        ('VALIGN', (0, 2), (-1, 7), 'MIDDLE'),  # Text cols (0-7, sudah termasuk Alat+Cuaca)
        ('LEFTPADDING', (0, 0), (-1, 7), 3),
        ('RIGHTPADDING', (0, 0), (-1, 7), 3),
        ('TOPPADDING', (0, 0), (-1, 7), 3),
        ('BOTTOMPADDING', (0, 0), (-1, 7), 3),
        # Foto cols (8=0%, 9=50%, 10=100%): padding 0 biar image nempel ke garis cell
        ('LEFTPADDING', (8, 0), (10, -1), 0),
        ('RIGHTPADDING', (8, 0), (10, -1), 0),
        ('TOPPADDING', (8, 0), (10, -1), 0),
        ('BOTTOMPADDING', (8, 0), (10, -1), 0),
        ('MINROWHEIGHT', (0, 2), (-1, -1), 0.85 * cm),  # baris kosong tetap punya tinggi minimum
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (1, 1)),
        ('SPAN', (2, 0), (2, 1)),
        ('SPAN', (3, 0), (3, 1)),
        ('SPAN', (4, 0), (5, 0)),    # Jam (Mulai+Akhir)
        ('SPAN', (6, 0), (6, 1)),    # Alat yang Digunakan
        ('SPAN', (7, 0), (7, 1)),    # Cuaca
        ('SPAN', (8, 0), (10, 0)),   # Dokumentasi (0% / 50% / 100%)
    ]))
    story.append(t2)
    story.append(Spacer(1, 0.15 * cm))

    # ---------- TTD ----------
    # Batasi ukuran gambar TTD biar blok tanda tangan gak mendorong baris nama ke
    # halaman berikutnya (row 3.2cm x 2 tabel + TTD udah mepet di A4 landscape).
    def _fit_sig(img, max_w=6.0 * cm, max_h=1.5 * cm):
        if img is None:
            return None
        w, h = img.drawWidth, img.drawHeight
        s = min(max_w / w, max_h / h)
        img.drawWidth, img.drawHeight = w * s, h * s
        return img

    # Blok tanda tangan persis .xls: "Mengetahui : / Pengamat DI. Riam Kanan" (kiri),
    # "Dibuat oleh : / <label petugas>" (kanan), lalu nama (bold+underline).
    # Baris NIP hanya muncul kalau memang diisi (di .xls asli tidak ada baris NIP).
    ttd_rows = [[
        Paragraph("<b>Mengetahui :</b><br/>Pengamat DI. Riam Kanan", meta_style),
        Paragraph("<b>Dibuat oleh :</b><br/>" + str(meta.get('petugas_label', 'Petugas')), meta_style),
    ], [
        _fit_sig(signature_pengamat) or Paragraph('<br/><br/><br/>', meta_style),
        _fit_sig(signature_petugas) or Paragraph('<br/><br/><br/>', meta_style),
    ], [
        Paragraph("<b><u>" + str(meta.get('pengamat', 'AKHMAD MUHAZIR')).upper() + "</u></b>", meta_style),
        Paragraph("<b><u>" + str(meta.get('nama', '.........................')).upper() + "</u></b>", meta_style),
    ]]
    _nip_pengamat = str(meta.get('pengamat_nip', '') or '').strip()
    _nip_petugas = str(meta.get('petugas_nip', '') or '').strip()
    if _nip_pengamat or _nip_petugas:
        ttd_rows.append([
            Paragraph("NIP. " + _nip_pengamat, meta_style) if _nip_pengamat else Paragraph('', meta_style),
            Paragraph("NIP. " + _nip_petugas, meta_style) if _nip_petugas else Paragraph('', meta_style),
        ])
    ttd_data = ttd_rows

    ttd = Table(ttd_data, colWidths=[USABLE_W / 2, USABLE_W / 2])
    ttd.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    story.append(ttd)

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

def _too_large_page():
    """Halaman ramah buat HTTP 413 (payload kelebihan limit 4,5 MB Vercel)."""
    return (
        "<!doctype html><html lang='id'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<title>Foto terlalu besar</title></head>"
        "<body style=\"font-family:system-ui,-apple-system,Segoe UI,sans-serif;"
        "max-width:560px;margin:0 auto;padding:28px;line-height:1.6\">"
        "<h2>📸 Kiriman terlalu besar</h2>"
        "<p>Server menolak karena body request melebihi <b>4,5 MB</b> "
        "(batas function Vercel). Biasanya karena foto kamera belum dikompres.</p>"
        "<p><b>Solusi:</b></p><ol>"
        "<li>Kurangi jumlah foto (mis. 4–6 foto saja per laporan).</li>"
        "<li>Di HP, set kamera ke resolusi lebih kecil sebelum foto.</li>"
        "<li>Buka ulang form lalu tekan tombol lagi — form versi baru "
        "<b>otomatis mengompres foto</b> di browser sebelum dikirim.</li>"
        "</ol><p><a href='/'>← Balik ke form laporan</a></p>"
        "</body></html>"
    ), 413


@app.errorhandler(413)
def handle_413(e):
    return _too_large_page()


@app.route('/')
def index():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('form.html', default_date=today, bulan_id=BULAN_ID)


@app.route('/preview', methods=['POST', 'GET'])
def preview():
    """Show preview of what the PDF will look like (HTML mockup)."""
    # GET ke /preview cuma bisa kejadian kalau user reload/back setelah submit
    # (dulu: 'Method Not Allowed' 405). Sekarang dibalikin ke form.
    if request.method == 'GET':
        return redirect('/')
    meta = {
        'nama': request.form.get('nama', 'Muhammad Andri'),
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


@app.route('/generate', methods=['POST', 'GET'])
def generate():
    # GET ke /generate (reload/back setelah submit) -> balikin ke form, bukan 405
    if request.method == 'GET':
        return redirect('/')
    meta = {
        'nama': request.form.get('nama', 'Muhammad Andri'),
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
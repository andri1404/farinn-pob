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
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image,
    PageBreak, KeepTogether,
)
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


def decode_image(file_storage, max_px=900):
    """Decode upload -> ReportLab Image (downscaled). Returns None if empty."""
    if not file_storage or not file_storage.filename:
        return None
    raw = file_storage.read()
    if not raw:
        return None
    try:
        from PIL import Image as PILImage
        pil = PILImage.open(BytesIO(raw))
        pil.thumbnail((max_px, max_px))
        if pil.mode in ('RGBA', 'LA', 'P'):
            pil = pil.convert('RGB')
        buf = BytesIO()
        pil.save(buf, format='JPEG', quality=80, optimize=True)
        buf.seek(0)
        img = Image(buf, width=3.0 * cm, height=2.2 * cm)
        return img
    except Exception:
        try:
            buf = BytesIO(raw)
            return Image(buf, width=3.0 * cm, height=2.2 * cm)
        except Exception:
            return None


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

# Layout A4 landscape (29.7 x 21.0 cm)
PAGE_W, PAGE_H = landscape(A4)
MARGIN = 0.4 * cm
USABLE_W = PAGE_W - 2 * MARGIN  # ~28.9 cm


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
    )

    story = []

    # ---------- Header ----------
    story.append(Paragraph("LAPORAN HARIAN KEGIATAN PEKERJAAN "
                          "PETUGAS OPERASI BENDUNG / JURU JARINGAN / PPA",
                          title_style))
    story.append(Paragraph("DAERAH IRIGASI RIAM KANAN", sub_style))

    # Metadata block (Nama, Petugas, Hari/Tanggal)
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
        colWidths=[2.0 * cm, 10.5 * cm, 2.5 * cm, 13.9 * cm],
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
    # 11 cols total. Sum must = USABLE_W/cm (~28.9)
    col_widths_t1_cm = [
        0.8,  # No
        2.6,  # Hari/Tanggal
        4.4,  # Lokasi
        4.4,  # Jenis
        1.6,  # Waktu
        1.4,  # TMA
        1.6,  # Status
        1.5,  # Cuaca
        1.9,  # TMA Pagi
        3.3,  # Selfi (img)
        5.4,  # stretchable filler (merged header span)
    ]
    # Re-balance: filler should not be in final merged area; we'll handle
    # via SPAN + layout. Total ~28.9
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
        t1_data.append([
            str(i),
            str(row.get('hari_tanggal', '')),
            Paragraph(str(row.get('lokasi', '')), meta_style),
            Paragraph(str(row.get('jenis', '')), meta_style),
            str(row.get('waktu', '')),
            str(row.get('tma', '')),
            str(row.get('status', '')),
            str(row.get('cuaca', '')),
            str(row.get('tma_pagi', '')) if not row.get('tma_pagi_img') else row['tma_pagi_img'],
            row.get('selfi_img') or '',
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
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (1, 1)),
        ('SPAN', (2, 0), (2, 1)),
        ('SPAN', (3, 0), (3, 1)),
        ('SPAN', (4, 0), (7, 0)),    # Pagi (4 cols)
        ('SPAN', (8, 0), (10, 0)),   # Dokumentasi (3 cols incl. filler)
        # Body cell merges: span last filler cell into prior empty
        ('SPAN', (10, 2), (10, -1)),  # merge filler column on body rows so it acts as extra width for TMA Pagi / Selfi
    ]))
    story.append(t1)
    story.append(Spacer(1, 0.25 * cm))

    # ---------- TABEL 2 - KEGIATAN PEKERJAAN ----------
    # 11 cols: No | Hari/Tgl | Lokasi | Jenis | Jam Mulai | Jam Akhir | Cuaca | Alat | Foto0 | Foto0.5 | Foto1
    col_widths_t2_cm = [
        0.8,  # No
        2.6,  # Hari/Tanggal
        4.0,  # Lokasi
        4.0,  # Jenis
        1.5,  # Jam Mulai
        1.5,  # Jam Akhir
        1.4,  # Cuaca
        2.5,  # Alat
        3.4,  # Foto 0
        3.4,  # Foto 0.5
        3.8,  # Foto 1
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
        t2_data.append([
            str(i),
            str(row.get('hari_tanggal', '')),
            Paragraph(str(row.get('lokasi', '')), meta_style),
            Paragraph(str(row.get('jenis', '')), meta_style),
            str(row.get('jam_mulai', '')),
            str(row.get('jam_akhir', '')),
            str(row.get('cuaca', '')),
            Paragraph(str(row.get('alat', '')), meta_style),
            row.get('foto0_img') or '',
            row.get('foto05_img') or '',
            row.get('foto1_img') or '',
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
    """Parse repeated 'pagi-*' fields from request.form -> list of dicts."""
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
            'selfi_img': decode_image(request.files.get(f'pagi-{i}-selfi')),
        })
    return rows


def parse_form_kerja():
    """Parse repeated 'kerja-*' fields."""
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
            'foto0_img': decode_image(request.files.get(f'kerja-{i}-foto0')),
            'foto05_img': decode_image(request.files.get(f'kerja-{i}-foto05')),
            'foto1_img': decode_image(request.files.get(f'kerja-{i}-foto1')),
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
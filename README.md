# Laporan Harian POB / JURU JARINGAN / PPA — DI Riam Kanan

Web app generator laporan harian Petugas Operasi Bendung dalam format PDF.
Dibangun ulang dari source Excel `LAPORAN HARIAN POB-JURU-PPA.xls` dengan layout
2 tabel:

- **Tabel 1 — Pemeriksaan Pagi**: Waktu, TMA, Status, Cuaca, TMA Pagi, Selfi
- **Tabel 2 — Kegiatan Pekerjaan**: Jam Mulai/Akhir, Alat, Cuaca, Dokumentasi (Foto 0/0.5/1)
- **TTD**: Mengetahui (Pengamat DI Riam Kanan — Akhmad Muhazir) / Dibuat oleh (Petugas)

## Stack

- **Flask** — web framework
- **ReportLab** — PDF generator
- **Pillow** — image resize/compress (in-memory)
- **Vercel** — deployment (@vercel/python adapter)

## Struktur

```
api/index.py          # Flask app — semua logic (form, PDF, ZIP)
templates/form.html   # Form input
templates/preview.html# Preview sebelum generate
static/style.css      # Styling
vercel.json           # Vercel routing
requirements.txt      # Deps
```

## Endpoints

| Method | Path        | Fungsi                                |
| ------ | ----------- | ------------------------------------- |
| GET    | `/        | Form input                            |
| POST   | `/preview`  | Preview layout (HTML)                 |
| POST   | `/generate` | Generate PDF single-day (download)     |
| GET    | `/healthz`  | Health check                          |

## Deploy

```bash
vercel login
vercel --prod
```

Tanpa token → CLI flow device-code (butuh browser auth). Lihat README
farinn-vercel lama untuk workaround paste token manual.

## Development lokal

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python api/index.py
# buka http://localhost:5000
```
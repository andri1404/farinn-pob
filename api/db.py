"""
STORAGE - Multi-backend: Vercel KV (primary) or GitHub Gist (fallback)

Menyimpan & membaca data laporan harian PPA.
Primary: Vercel KV via REST API (KV_REST_API_URL + KV_REST_API_TOKEN)
Fallback: GitHub API (gh token, buat gist 'laporanppa-data')

Env vars:
  KV_REST_API_URL     - Vercel KV REST API URL
  KV_REST_API_TOKEN   - Vercel KV REST API token
  GH_TOKEN            - GitHub token (fallback)

Data disimpan sebagai JSON array di Redis list (KV) atau file JSON di Gist (GH).
Tiap entry: tanggal, nama, petugas, petugas_label, pengamat, lokasi,
             jenis_pekerjaan, tma, cuaca, status, created_at
"""

import os
import json
import time
from urllib.request import Request, urlopen, HTTPError
from urllib.parse import urlencode


# --- Backend: Vercel KV ---

KV_URL = os.environ.get('KV_REST_API_URL', '').rstrip('/')
KV_TOKEN = os.environ.get('KV_REST_API_TOKEN', '')
KV_KEY = 'laporan:reports'


def _kv_ok():
    return bool(KV_URL and KV_TOKEN)


def _kv_headers():
    return {'Authorization': f'Bearer {KV_TOKEN}', 'Content-Type': 'application/json'}


def _kv_post(path, body=None):
    if not _kv_ok():
        return None
    try:
        url = f'{KV_URL}/{path.lstrip("/")}'
        data = json.dumps(body).encode() if body else None
        req = Request(url, data=data, headers=_kv_headers(), method='POST')
        with urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode())
    except (HTTPError, OSError, json.JSONDecodeError) as e:
        print(f'[KV POST] {path}: {e}')
        return None


def _kv_get(path):
    if not _kv_ok():
        return None
    try:
        url = f'{KV_URL}/{path.lstrip("/")}'
        req = Request(url, headers=_kv_headers(), method='GET')
        with urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode())
    except (HTTPError, OSError, json.JSONDecodeError) as e:
        print(f'[KV GET] {path}: {e}')
        return None


# --- Backend: GitHub Gist (fallback) ---

GH_TOKEN = os.environ.get('GH_TOKEN', '')
GIST_ID_KEY = 'laporanppa_gist_id'
GIST_FILE = 'laporan_data.json'
GIST_DESC = 'Laporan Harian PPA - Data Storage'


def _gh_ok():
    return bool(GH_TOKEN)


def _gh_headers():
    return {
        'Authorization': f'Bearer {GH_TOKEN}',
        'Accept': 'application/vnd.github+json',
        'User-Agent': 'laporanppa/1.0',
    }


def _gh_get_gist_id():
    """Dapatkan Gist ID dari environment atau buat baru."""
    gist_id = os.environ.get(GIST_ID_KEY, '')
    if gist_id:
        return gist_id
    # Cari gist yang ada
    try:
        req = Request('https://api.github.com/gists', headers=_gh_headers())
        with urlopen(req, timeout=10) as resp:
            gists = json.loads(resp.read().decode())
        for g in gists:
            if g.get('description') == GIST_DESC:
                os.environ[GIST_ID_KEY] = g['id']
                return g['id']
    except Exception as e:
        print(f'[GH] list gists: {e}')
    return ''


def _gh_create_gist():
    """Buat gist baru untuk menyimpan data."""
    if not _gh_ok():
        return ''
    payload = json.dumps({
        'description': GIST_DESC,
        'public': False,
        'files': {GIST_FILE: {'content': json.dumps([], ensure_ascii=False)}},
    }).encode()
    try:
        req = Request('https://api.github.com/gists', data=payload,
                      headers=_gh_headers(), method='POST')
        with urlopen(req, timeout=15) as resp:
            gist = json.loads(resp.read().decode())
            os.environ[GIST_ID_KEY] = gist['id']
            return gist['id']
    except Exception as e:
        print(f'[GH] create gist: {e}')
        return ''


def _gh_read_reports():
    """Baca data laporan dari Gist."""
    gist_id = _gh_get_gist_id()
    if not gist_id:
        gist_id = _gh_create_gist()
    if not gist_id:
        return []
    try:
        req = Request(f'https://api.github.com/gists/{gist_id}', headers=_gh_headers())
        with urlopen(req, timeout=10) as resp:
            gist = json.loads(resp.read().decode())
        content = gist.get('files', {}).get(GIST_FILE, {}).get('content', '[]')
        data = json.loads(content)
        return data if isinstance(data, list) else []
    except Exception as e:
        print(f'[GH] read: {e}')
        return []


def _gh_write_reports(reports):
    """Tulis data laporan ke Gist."""
    gist_id = _gh_get_gist_id()
    if not gist_id:
        gist_id = _gh_create_gist()
    if not gist_id:
        return False
    payload = json.dumps({
        'files': {GIST_FILE: {'content': json.dumps(reports, ensure_ascii=False, indent=2)}},
    }).encode()
    try:
        req = Request(f'https://api.github.com/gists/{gist_id}', data=payload,
                      headers=_gh_headers(), method='PATCH')
        with urlopen(req, timeout=15) as resp:
            return True
    except Exception as e:
        print(f'[GH] write: {e}')
        return False


# --- Public API ---

def available():
    """Cek apakah ada backend yang aktif."""
    return _kv_ok() or _gh_ok()


def backend_name() -> str:
    if _kv_ok():
        return 'Vercel KV'
    if _gh_ok():
        return 'GitHub Gist'
    return '-'


def save_report(data: dict) -> bool:
    """Simpan 1 laporan ke storage. LPUSH ke KV atau append ke GH Gist."""
    entry = dict(data)
    entry.setdefault('created_at', int(time.time() * 1000))
    # Primary: KV
    if _kv_ok():
        result = _kv_post(f'lpush/{KV_KEY}', {'value': json.dumps(entry, ensure_ascii=False)})
        if result is not None:
            return True
        print('[KV] lpush failed, fallback to GH')
    # Fallback: GitHub Gist
    if _gh_ok():
        reports = _gh_read_reports()
        reports.insert(0, entry)
        return _gh_write_reports(reports[:1000])
    return False


def get_reports(limit: int = 100) -> list:
    """Ambil daftar laporan terbaru (terbaru dulu)."""
    # Primary: KV
    if _kv_ok():
        result = _kv_post(f'lrange/{KV_KEY}/0/{limit - 1}')
        if result and 'result' in result:
            reports = []
            for raw in result['result']:
                try:
                    reports.append(json.loads(raw))
                except (json.JSONDecodeError, TypeError):
                    continue
            return reports
        print('[KV] lrange failed, fallback to GH')
    # Fallback: GitHub Gist
    if _gh_ok():
        return _gh_read_reports()[:limit]
    return []


def get_reports_filtered(tahun: str = '', bulan: str = '', lokasi: str = '') -> list:
    """Ambil laporan dengan filter."""
    all_reports = get_reports(500)
    result = []
    for r in all_reports:
        tgl = r.get('tanggal', '')
        if tahun and not tgl.startswith(tahun):
            continue
        if bulan:
            parts = tgl.split('-')
            if len(parts) >= 2 and parts[1] != bulan.zfill(2):
                continue
        if lokasi:
            lok = (r.get('lokasi', '') or '').lower()
            if lokasi.lower() not in lok:
                continue
        result.append(r)
    return result


def count_reports() -> int:
    """Hitung total laporan tersimpan."""
    return len(get_reports(1000))
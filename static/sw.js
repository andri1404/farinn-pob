/* Service worker minimal untuk Laporan Harian Riam Kanan.
 *
 * Fungsi: bikin web ini memenuhi syarat "Install app" di Android Chrome,
 * jadi bisa dipasang di layar utama dan dibuka tanpa address bar.
 *
 * SENGAJA TANPA CACHE. Kita cuma lewatkan request ke jaringan apa adanya.
 * (Dulu sempat ada masalah user kejebak versi lama karena cache — jangan diulang.)
 */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', (e) => {
  // Cuma GET yang dilewati. POST /generate & /preview jangan disentuh sama sekali.
  if (e.request.method !== 'GET') return;
  e.respondWith(fetch(e.request));
});

/* オフラインでも開けるように、アプリ一式を端末に保存する */
const CACHE = 'jishiyaku-067abd98';
const FILES = ['./', './index.html', './tesseract.min.js', './worker.min.js', './tesseract-core-lstm.wasm.js', './jpn-traineddata-gz.wasm', './manifest.webmanifest', './icon.png', './zxing-reader.js', './zxing_reader.wasm'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// ページ本体は新しい版を優先（通信できなければ保存版）、その他は保存版を優先
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put('./index.html', c)); return r; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(r => r || fetch(req)));
});

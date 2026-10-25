/* オフラインでも開けるように、アプリ一式を端末に保存する */
const NDL_CACHE = 'ndl-v1';
const CACHE = 'jishiyaku-16d43dc4';
const FILES = ['./', './index.html', './tesseract.min.js', './worker.min.js', './tesseract-core-lstm.wasm.js', './jpn-traineddata-gz.wasm', './manifest.webmanifest', './icon.png', './zxing-reader.js', './zxing_reader.wasm'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  // OCRモデル（ndl-…）は大きいので、アプリの更新では消さない
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE && !k.startsWith('ndl-')).map(k => caches.delete(k)))).then(() => self.clients.claim()));
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
  // OCRエンジン（ort/）とモデル（ndl/）は、初めて使うときに取得して端末に保存する
  const path = new URL(req.url).pathname;
  if (/\/(ort|ndl)\/[^/]+$/.test(path)) {
    e.respondWith(caches.open(NDL_CACHE).then(c => c.match(req, { ignoreSearch: true }).then(r => r || fetch(req).then(res => { if (res.ok) c.put(req, res.clone()); return res; }))));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(r => r || fetch(req)));
});

// Minimal service worker, present only so the browser recognizes this page as an installable
// app. It doesn't cache anything or serve offline -- every request still hits the network,
// so the dashboard always shows live data. A fetch listener is required for Chrome's install
// criteria even if it does nothing but pass the request through.
self.addEventListener('fetch', () => {});

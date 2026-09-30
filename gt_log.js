// Login log, backed by the Google Apps Script web app in login_log/Code.gs.
// Leave GT_LOG_URL empty to disable it.
var GT_LOG_URL = 'https://script.google.com/macros/s/AKfycbxxjEWilEaIGvWP_Pz9BfE9RJH8zjpFZNmB_fyZoNN9mAr7wzY0cvjkhpq7Am75oet_9A/exec';

function gtLogLogin(name, id, access) {
  if (!GT_LOG_URL) return;
  fetch(GT_LOG_URL, {
    method: 'POST',
    mode: 'no-cors',
    keepalive: true,
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: 'log', name: name, id: id, access: access, page: location.pathname, device: navigator.userAgent })
  }).catch(function () {});
}

function gtLogApi(params) {
  if (!GT_LOG_URL) return Promise.reject(new Error('Login log is not set up (GT_LOG_URL is empty).'));
  var qs = Object.keys(params).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  return fetch(GT_LOG_URL + '?' + qs, { cache: 'no-store' })
    .then(function (r) { return r.json(); })
    .then(function (d) {
      if (d.error) throw new Error(d.error);
      return d;
    });
}

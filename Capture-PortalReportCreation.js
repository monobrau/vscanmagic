/**
 * ConnectSecure Portal Report Creation Capturer
 *
 * INSTRUCTIONS:
 * 1. Log into ConnectSecure in your browser
 * 2. Open DevTools (F12) → Console tab
 * 3. Paste this entire script and press Enter
 * 4. Create a standard report in the UI (Reports → Create → pick report, company, etc.)
 * 5. When done, run: copy(JSON.stringify(window.__capturedReportCalls, null, 2))
 *    to copy captured calls to clipboard, or inspect window.__capturedReportCalls
 */
(function () {
  window.__capturedReportCalls = [];

  function matchReportUrl(url) {
    const u = (url || '').toString();
    return (
      /report|create_report|get_report_link|report_jobs|standard_reports/i.test(u)
    );
  }

  function capture(type, url, reqBody, respBody, status, method) {
    if (!matchReportUrl(url)) return;
    const entry = {
      time: new Date().toISOString(),
      type,
      method: method || 'GET',
      url,
      requestBody: reqBody,
      responseStatus: status,
      responseBody: respBody
    };
    window.__capturedReportCalls.push(entry);
    console.log(
      '%c[Report Capture]',
      'background:#333;color:#0f0',
      method || type,
      url,
      entry
    );
  }

  // Intercept fetch
  const origFetch = window.fetch;
  window.fetch = async function (resource, init) {
    const url = typeof resource === 'string' ? resource : resource?.url || '';
    const method = init?.method || 'GET';
    let reqBody = null;
    try {
      if (init?.body) {
        reqBody = typeof init.body === 'string' ? JSON.parse(init.body) : init.body;
      }
    } catch (_) {
      reqBody = init?.body;
    }

    const resp = await origFetch.apply(this, arguments);
    let respBody = null;
    try {
      const clone = resp.clone();
      respBody = await clone.json();
    } catch (_) {
      try {
        const c = resp.clone();
        respBody = await c.text();
      } catch (__) {}
    }
    capture('fetch', url, reqBody, respBody, resp.status, method);
    return resp;
  };

  // Intercept XMLHttpRequest
  const origOpen = XMLHttpRequest.prototype.open;
  const origSend = XMLHttpRequest.prototype.send;

  XMLHttpRequest.prototype.open = function (method, url) {
    this._csUrl = url;
    this._csMethod = method;
    return origOpen.apply(this, arguments);
  };

  XMLHttpRequest.prototype.send = function (body) {
    const xhr = this;
    const url = xhr._csUrl || '';
    const method = xhr._csMethod || 'GET';

    const onLoad = function () {
      let respBody = null;
      try {
        respBody = JSON.parse(xhr.responseText);
      } catch (_) {
        respBody = xhr.responseText;
      }
      let reqBody = null;
      try {
        reqBody = body ? JSON.parse(body) : null;
      } catch (_) {
        reqBody = body;
      }
      capture('xhr', url, reqBody, respBody, xhr.status, method);
    };

    xhr.addEventListener('load', onLoad);
    return origSend.apply(this, arguments);
  };

  console.log(
    '%c[Report Capturer Active]',
    'background:#0a0;color:#fff;padding:4px 8px',
    'Create a report in the UI. Capture stored in window.__capturedReportCalls'
  );
  console.log(
    'To export: copy(JSON.stringify(window.__capturedReportCalls, null, 2))'
  );
})();

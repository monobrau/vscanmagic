/**
 * Capture-PortalReportCreation.js
 * Paste this into the ConnectSecure portal browser console (F12 > Console) to capture
 * API calls when creating reports. Captures create_report_job, get_report_link,
 * standard_reports, and related report-builder requests.
 *
 * After creating a report in the UI, run: copy(JSON.stringify(window.__capturedReportCalls, null, 2))
 * Then paste into portal-report-capture.json for analysis.
 */

(function() {
  window.__capturedReportCalls = [];

  function shouldCapture(url) {
    var u = (url || '').toLowerCase();
    return u.indexOf('report') !== -1 ||
           u.indexOf('standard_reports') !== -1 ||
           u.indexOf('create_report') !== -1 ||
           u.indexOf('get_report_link') !== -1 ||
           u.indexOf('report_builder') !== -1 ||
           u.indexOf('report_job') !== -1;
  }

  function safeJson(obj) {
    try {
      if (typeof obj === 'string') return obj;
      return JSON.stringify(obj);
    } catch (e) { return String(obj); }
  }

  function captureEntry(url, method, requestBody, responseData, status) {
    if (!shouldCapture(url)) return;
    window.__capturedReportCalls.push({
      timestamp: new Date().toISOString(),
      url: url,
      method: method || 'GET',
      requestBody: requestBody,
      responseData: responseData,
      status: status
    });
    console.log('[Report Capturer]', method, url);
  }

  // Hook fetch
  var origFetch = window.fetch;
  window.fetch = function(url, opts) {
    opts = opts || {};
    var method = (opts.method || 'GET').toUpperCase();
    var body = opts.body;
    return origFetch.apply(this, arguments).then(function(r) {
      var clone = r.clone();
      clone.text().then(function(text) {
        var respData = null;
        try { respData = JSON.parse(text); } catch (e) { respData = text; }
        captureEntry(
          typeof url === 'string' ? url : (url && url.url) || '',
          method,
          body ? (typeof body === 'string' ? (function(){ try { return JSON.parse(body); } catch(e){ return body; } })() : body) : null,
          respData,
          r.status
        );
      });
      return r;
    });
  };

  // Hook XMLHttpRequest
  var origOpen = XMLHttpRequest.prototype.open;
  var origSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(method, url) {
    this._captureUrl = url;
    this._captureMethod = method;
    return origOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function(body) {
    var xhr = this;
    var url = xhr._captureUrl || '';
    var method = xhr._captureMethod || 'GET';
    var reqBody = null;
    if (body && typeof body === 'string') {
      try { reqBody = JSON.parse(body); } catch (e) { reqBody = body; }
    }
    xhr.addEventListener('load', function() {
      var respData = null;
      try { respData = JSON.parse(xhr.responseText); } catch (e) { respData = xhr.responseText; }
      captureEntry(url, method, reqBody, respData, xhr.status);
    });
    return origSend.apply(this, arguments);
  };

  console.log('[Report Capturer Active] Create a report in the UI (company or global), then run: copy(JSON.stringify(window.__capturedReportCalls, null, 2))');
})();

/**
 * Capture-PortalCompanyReview.js
 * Paste this into the ConnectSecure portal browser console (F12 > Console) to capture
 * API calls when viewing company data (agents, assets, external scans, etc.).
 * Use to discover server-side filtering params (company_id, condition, etc.).
 *
 * After navigating to a company view and loading its data, run:
 *   copy(JSON.stringify(window.__capturedCompanyCalls, null, 2))
 * Then paste into portal-company-review-capture.json for analysis.
 */

(function() {
  window.__capturedCompanyCalls = [];

  function shouldCapture(url) {
    var u = (url || '').toLowerCase();
    return u.indexOf('external_asset_externalscan') !== -1 ||
           u.indexOf('lightweight_assets') !== -1 ||
           u.indexOf('company_stats') !== -1 ||
           u.indexOf('discovery_settings') !== -1 ||
           u.indexOf('/r/company/agents') !== -1 ||
           u.indexOf('/r/company/company_stats') !== -1 ||
           u.indexOf('/r/company/credentials') !== -1 ||
           u.indexOf('agent_credentials_mapping') !== -1 ||
           u.indexOf('agent_discoverysettings_mapping') !== -1 ||
           u.indexOf('asset_firewall_policy') !== -1 ||
           (u.indexOf('/r/asset/assets') !== -1 && u.indexOf('report_queries') === -1) ||
           u.indexOf('report_queries') !== -1 && (u.indexOf('asset') !== -1 || u.indexOf('agent') !== -1);
  }

  function parseUrlParams(url) {
    var params = {};
    try {
      var idx = url.indexOf('?');
      if (idx === -1) return params;
      var qs = url.substring(idx + 1);
      qs.split('&').forEach(function(pair) {
        var eq = pair.indexOf('=');
        if (eq !== -1) {
          var k = decodeURIComponent(pair.substring(0, eq));
          var v = decodeURIComponent(pair.substring(eq + 1));
          params[k] = v;
        }
      });
    } catch (e) {}
    return params;
  }

  function captureEntry(url, method, requestBody, responseData, status) {
    if (!shouldCapture(url)) return;
    var fullUrl = typeof url === 'string' ? url : (url && url.url) || '';
    var path = fullUrl.split('?')[0];
    var queryParams = parseUrlParams(fullUrl);
    window.__capturedCompanyCalls.push({
      timestamp: new Date().toISOString(),
      url: fullUrl,
      path: path,
      method: method || 'GET',
      queryParams: queryParams,
      requestBody: requestBody,
      responseRecordCount: (responseData && responseData.data && Array.isArray(responseData.data)) ? responseData.data.length : null,
      responseStatus: (responseData && responseData.status) ? responseData.status : null,
      status: status
    });
    console.log('[Company Capturer]', method, path, 'params:', JSON.stringify(queryParams));
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

  console.log('[Company Review Capturer Active] Navigate to a company view (e.g. Company > Assets, External Scans, Agents) and load its data. Then run: copy(JSON.stringify(window.__capturedCompanyCalls, null, 2))');
})();

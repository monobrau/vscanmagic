/**
 * Capture-PortalCompanyReview-DEBUG.js
 * DEBUG version - logs ALL API requests so you can see what the portal actually calls.
 * Paste into console, then navigate. You should see [API] logs for every request.
 *
 * If you see NO [API] logs when clicking External Assets:
 *   - The portal may use WebSockets or a different mechanism
 *   - Try opening DevTools Network tab, filter by Fetch/XHR, then click External Assets
 */
(function() {
  window.__capturedCompanyCalls = [];
  window.__capturedAllApiCalls = [];  // DEBUG: all API requests

  function isApiUrl(url) {
    var u = (url || '').toLowerCase();
    return u.indexOf('myconnectsecure.com') !== -1 ||
           u.indexOf('connectsecure.com') !== -1 ||
           u.indexOf('/r/') !== -1 ||
           u.indexOf('/w/') !== -1;
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

  function sampleDateFields(obj) {
    if (!obj || typeof obj !== 'object') return null;
    var keys = Object.keys(obj);
    var sample = {};
    ['updated', 'created', 'last_scan', 'external_last_scan', 'ad_last_scan', 'date', 'company_id'].forEach(function(k) {
      keys.filter(function(x) { return x.toLowerCase().indexOf(k) !== -1; }).forEach(function(x) { sample[x] = obj[x]; });
    });
    return Object.keys(sample).length ? sample : null;
  }

  function captureEntry(url, method, requestBody, responseData, status) {
    var fullUrl = typeof url === 'string' ? url : (url && url.url) || '';
    if (!isApiUrl(fullUrl)) return;

    var path = fullUrl.split('?')[0];
    var queryParams = parseUrlParams(fullUrl);
    var data = responseData && responseData.data !== undefined ? responseData.data : (Array.isArray(responseData) ? responseData : responseData);
    var recordCount = (data && Array.isArray(data)) ? data.length : (data && typeof data === 'object' ? 1 : null);
    var firstRecord = (data && Array.isArray(data) && data[0]) ? data[0] : (data && typeof data === 'object' && !Array.isArray(data) ? data : null);
    var firstSample = sampleDateFields(firstRecord) || sampleDateFields(responseData);

    var entry = {
      timestamp: new Date().toISOString(),
      url: fullUrl,
      path: path,
      method: method || 'GET',
      queryParams: queryParams,
      responseRecordCount: recordCount,
      firstRecordSample: firstSample,
      status: status
    };
    window.__capturedAllApiCalls.push(entry);

    if (path.indexOf('external_asset') !== -1 || path.indexOf('lightweight_assets') !== -1 ||
        path.indexOf('discovery_settings') !== -1 || path.indexOf('company_stats') !== -1 ||
        path.indexOf('jobs_view') !== -1 || path.indexOf('agents') !== -1 || path.indexOf('assets') !== -1) {
      window.__capturedCompanyCalls.push(entry);
      console.log('[Company Capturer]', method, path, 'params:', JSON.stringify(queryParams), 'records:', recordCount, firstSample ? '(has date fields)' : '');
    } else {
      console.log('[API]', method, path, 'records:', recordCount);
    }
  }

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
        captureEntry(typeof url === 'string' ? url : (url && url.url) || '', method,
          body ? (typeof body === 'string' ? (function(){ try { return JSON.parse(body); } catch(e){ return body; } })() : body) : null,
          respData, r.status);
      });
      return r;
    });
  };

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

  console.log('[DEBUG Capturer Active] Logs [API] and [Company Capturer] for every request. firstRecordSample shows date fields (updated, last_scan, etc.).');
  console.log('For LAST EXTERNAL SCAN DATE: Find where the portal displays it, refresh, then navigate there. Export: copy(JSON.stringify(window.__capturedAllApiCalls, null, 2))');
})();

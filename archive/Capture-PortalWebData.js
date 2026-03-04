/**
 * Capture-PortalWebData.js
 * Captures API calls from the ConnectSecure portal for hostname, username, vulnerability,
 * asset, and agent data. Use to discover what endpoints and fields the portal exposes.
 *
 * Steps:
 * 1. Log into ConnectSecure portal
 * 2. Open DevTools (F12) > Console
 * 3. Paste this script and press Enter
 * 4. Navigate to: Vulnerability reports, All Vulnerabilities, Assets, Agents, Host lists, etc.
 * 5. Let data load; expand tables/details that show hostname, username, IP
 * 6. Export: copy(JSON.stringify(window.__capturedWebDataCalls, null, 2))
 * 7. Save to portal-web-data-capture.json
 */
(function() {
  window.__capturedWebDataCalls = [];

  var relevantPaths = [
    'vulnerabilit',
    'asset',
    'agent',
    'host',
    'report',
    'lightweight',
    'discovery',
    'credentials',
    'external_asset',
    'company/agents',
    'company_stats',
    'report_queries',
    'all_vulnerabilities',
    'pending',
    'epss',
    'software',
    'product'
  ];

  function isRelevantUrl(url) {
    var u = (url || '').toLowerCase();
    if (u.indexOf('myconnectsecure.com') === -1 && u.indexOf('connectsecure.com') === -1 && u.indexOf('/r/') === -1) return false;
    return relevantPaths.some(function(p) { return u.indexOf(p) !== -1; });
  }

  function parseUrlParams(url) {
    var params = {};
    try {
      var idx = (url || '').indexOf('?');
      if (idx === -1) return params;
      var qs = url.substring(idx + 1);
      qs.split('&').forEach(function(pair) {
        var eq = pair.indexOf('=');
        if (eq !== -1) {
          params[decodeURIComponent(pair.substring(0, eq))] = decodeURIComponent(pair.substring(eq + 1));
        }
      });
    } catch (e) {}
    return params;
  }

  function sampleRecord(obj, keys) {
    if (!obj || typeof obj !== 'object') return null;
    var sample = {};
    (keys || []).forEach(function(k) {
      if (obj.hasOwnProperty(k) && obj[k] !== undefined && obj[k] !== null) {
        sample[k] = obj[k];
      }
    });
    return Object.keys(sample).length ? sample : null;
  }

  var usernameKeys = ['username', 'user_name', 'user', 'last_user', 'last_logged_in_user', 'logged_in_user', 'owner', 'primary_user', 'account', 'login'];
  var hostKeys = ['host_name', 'hostname', 'hostName', 'computer_name', 'device_name', 'asset_name', 'name', 'fqdn', 'fqdn_name'];
  var allSampleKeys = hostKeys.concat(usernameKeys).concat(['ip', 'ip_address', 'software_name', 'product', 'severity', 'cve', 'epss_score', 'vulnerability_count']);

  function captureEntry(url, method, responseData) {
    var fullUrl = typeof url === 'string' ? url : (url && url.url) || '';
    if (!isRelevantUrl(fullUrl)) return;

    var path = fullUrl.split('?')[0];
    var data = responseData && responseData.data !== undefined ? responseData.data : (Array.isArray(responseData) ? responseData : responseData);
    var records = Array.isArray(data) ? data : (data && typeof data === 'object' ? [data] : []);
    var firstRecord = records[0] || (data && typeof data === 'object' && !Array.isArray(data) ? data : null);

    var firstSample = sampleRecord(firstRecord, allSampleKeys);
    var hasUsername = firstRecord && usernameKeys.some(function(k) { return firstRecord[k] !== undefined && firstRecord[k] !== ''; });
    var hasHostname = firstRecord && hostKeys.some(function(k) { return firstRecord[k] !== undefined && firstRecord[k] !== ''; });

    var entry = {
      timestamp: new Date().toISOString(),
      path: path,
      queryParams: parseUrlParams(fullUrl),
      method: method || 'GET',
      responseRecordCount: records.length,
      firstRecordSample: firstSample,
      hasUsernameField: !!hasUsername,
      hasHostnameField: !!hasHostname
    };

    if (firstRecord && hasUsername) {
      entry.usernameFieldSample = usernameKeys.reduce(function(s, k) {
        if (firstRecord[k] !== undefined) s[k] = firstRecord[k];
        return s;
      }, {});
    }

    window.__capturedWebDataCalls.push(entry);
    console.log('[WebData Capturer]', path, 'records:', records.length,
      hasUsername ? 'HAS USERNAME' : '',
      hasHostname ? 'hostname' : '',
      firstSample ? 'sample: ' + JSON.stringify(firstSample).substring(0, 120) + '...' : '');
  }

  var origFetch = window.fetch;
  window.fetch = function(url, opts) {
    return origFetch.apply(this, arguments).then(function(res) {
      var urlStr = typeof url === 'string' ? url : (url && url.url) || '';
      if (isRelevantUrl(urlStr)) {
        res.clone().json().then(function(data) { captureEntry(urlStr, (opts && opts.method) || 'GET', data); }).catch(function() {});
      }
      return res;
    });
  };

  var origOpen = XMLHttpRequest.prototype.open;
  var origSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(method, url) {
    this._captureUrl = url;
    this._captureMethod = method;
    return origOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function() {
    var xhr = this;
    xhr.addEventListener('load', function() {
      var data = null;
      try { data = JSON.parse(xhr.responseText); } catch (e) {}
      if (data) captureEntry(xhr._captureUrl || '', xhr._captureMethod || 'GET', data);
    });
    return origSend.apply(this, arguments);
  };

  console.log('[WebData Capturer Active] Navigate to: Vulnerability reports, All Vulnerabilities, Assets, Agents, Host lists. Look for hostname/username data. Export: copy(JSON.stringify(window.__capturedWebDataCalls, null, 2))');
})();

/**
 * Capture-PortalOffline.js
 * Captures offline-agent/offline-asset related API calls from the ConnectSecure web portal.
 * Use to discover: which endpoints the portal uses for offline counts (7d/14d/30+d), params, response structure.
 *
 * Paste into DevTools Console (F12), then navigate to views that show offline agents/assets for a company.
 * Export: copy(JSON.stringify(window.__capturedOfflineCalls, null, 2))
 */
(function() {
  window.__capturedOfflineCalls = [];

  var offlinePaths = [
    'offline',
    'last_ping',
    'lastPing',
    'lightweight_assets',
    'lightweight-assets',
    'company/agents',
    'agents',
    'assets',
    'offline_assets',
    'offline_agents',
    'agent_offline',
    'asset_status'
  ];

  function isOfflineUrl(url) {
    var u = (url || '').toLowerCase();
    return offlinePaths.some(function(p) { return u.indexOf(p.toLowerCase()) !== -1; });
  }

  function isApiUrl(url) {
    var u = (url || '').toLowerCase();
    return u.indexOf('myconnectsecure.com') !== -1 ||
           u.indexOf('connectsecure.com') !== -1 ||
           u.indexOf('/r/') !== -1;
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

  function captureOfflineResponse(url, data) {
    if (!isOfflineUrl(url) || !isApiUrl(url)) return;
    var path = (typeof url === 'string' ? url : '').split('?')[0];
    var respData = data && data.data !== undefined ? data.data : (Array.isArray(data) ? data : data);
    var recordCount = (respData && Array.isArray(respData)) ? respData.length : (respData && typeof respData === 'object' ? 1 : null);
    var firstRecord = (respData && Array.isArray(respData) && respData[0]) ? respData[0] : (respData && typeof respData === 'object' && !Array.isArray(respData) ? respData : null);
    var sampleKeys = ['last_ping_time', 'lastPingTime', 'host_name', 'agent_type', 'is_deprecated', 'company_id', 'id', 'name', 'ip', 'offline_assets_count', 'offline_assets', 'online_assets_count'];
    var firstSample = firstRecord ? sampleKeys.reduce(function(s, k) {
      if (firstRecord[k] !== undefined) s[k] = firstRecord[k];
      return s;
    }, {}) : null;
    var entry = {
      timestamp: new Date().toISOString(),
      path: path,
      queryParams: parseUrlParams(url),
      responseRecordCount: recordCount,
      firstRecordSample: firstSample && Object.keys(firstSample).length ? firstSample : null,
      hasOfflineField: firstRecord && (firstRecord.offline_assets_count !== undefined || firstRecord.last_ping_time !== undefined || (firstRecord.offline_assets && firstRecord.offline_assets.length !== undefined))
    };
    if (typeof data === 'object' && data.offline_assets_count !== undefined) {
      entry.topLevelOfflineCount = data.offline_assets_count;
    }
    if (typeof data === 'object' && data.data && Array.isArray(data.data) && data.data[0]) {
      var d0 = data.data[0];
      if (d0.offline_assets_count !== undefined) entry.inlineOfflineCount = d0.offline_assets_count;
      if (d0.offline_assets && Array.isArray(d0.offline_assets)) entry.inlineOfflineAssetsLength = d0.offline_assets.length;
    }
    window.__capturedOfflineCalls.push(entry);
    console.log('[Offline Capturer]', path, 'params:', JSON.stringify(entry.queryParams), 'records:', recordCount, 
      entry.topLevelOfflineCount ? 'offline_count=' + entry.topLevelOfflineCount : '',
      firstSample ? 'sample: ' + JSON.stringify(firstSample) : '');
  }

  var origFetch = window.fetch;
  window.fetch = function(url, opts) {
    return origFetch.apply(this, arguments).then(function(res) {
      var urlStr = typeof url === 'string' ? url : (url && url.url) || '';
      if (isOfflineUrl(urlStr) && isApiUrl(urlStr)) {
        res.clone().json().then(function(data) { captureOfflineResponse(urlStr, data); }).catch(function() {});
      }
      return res;
    });
  };

  var origOpen = XMLHttpRequest.prototype.open;
  var origSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(method, url) {
    this._captureUrl = url;
    return origOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function() {
    var xhr = this;
    var url = xhr._captureUrl || '';
    xhr.addEventListener('load', function() {
      var data = null;
      try { data = JSON.parse(xhr.responseText); } catch (e) {}
      if (data) captureOfflineResponse(url, data);
    });
    return origSend.apply(this, arguments);
  };

  console.log('[Offline Capturer Active] Navigate to views showing offline agents/assets for a company. Then run: copy(JSON.stringify(window.__capturedOfflineCalls, null, 2))');
})();

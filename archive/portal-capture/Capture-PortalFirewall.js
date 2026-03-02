/**
 * Capture-PortalFirewall.js
 * Captures firewall-related API calls from the ConnectSecure web portal.
 * Use to discover: which endpoints show Fortigate/Sonicwall counts, what params the portal uses.
 *
 * Paste into DevTools Console (F12), then navigate to the Firewall section for a company.
 * Export: copy(JSON.stringify(window.__capturedFirewallCalls, null, 2))
 */
(function() {
  window.__capturedFirewallCalls = [];

  var firewallPaths = [
    'firewall',
    'firewall_asset_view',
    'asset_firewall',
    'asset_firewall_policy',
    'firewall_groups',
    'firewall_interfaces',
    'firewall_rules',
    'firewall_zones',
    'firewall_users',
    'firewall_license'
  ];

  function isFirewallUrl(url) {
    var u = (url || '').toLowerCase();
    return firewallPaths.some(function(p) { return u.indexOf(p) !== -1; });
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

  function captureFirewallResponse(url, data) {
    if (!isFirewallUrl(url) || !isApiUrl(url)) return;
    var path = (typeof url === 'string' ? url : '').split('?')[0];
    var respData = data && data.data !== undefined ? data.data : (Array.isArray(data) ? data : data);
    var recordCount = (respData && Array.isArray(respData)) ? respData.length : (respData && typeof respData === 'object' ? 1 : null);
    var firstRecord = (respData && Array.isArray(respData) && respData[0]) ? respData[0] : (respData && typeof respData === 'object' && !Array.isArray(respData) ? respData : null);
    var sampleKeys = ['policy_type', 'asset_id', 'company_id', 'manufacturer', 'is_firewall', 'asset_type', 'name', 'interface_type', 'group_type'];
    var firstSample = firstRecord ? sampleKeys.reduce(function(s, k) {
      if (firstRecord[k] !== undefined) s[k] = firstRecord[k];
      return s;
    }, {}) : null;
    var entry = {
      timestamp: new Date().toISOString(),
      path: path,
      queryParams: parseUrlParams(url),
      responseRecordCount: recordCount,
      firstRecordSample: firstSample && Object.keys(firstSample).length ? firstSample : null
    };
    window.__capturedFirewallCalls.push(entry);
    console.log('[Firewall Capturer]', path, 'params:', JSON.stringify(entry.queryParams), 'records:', recordCount, firstSample ? 'sample: ' + JSON.stringify(firstSample) : '');
  }

  var origFetch = window.fetch;
  window.fetch = function(url, opts) {
    return origFetch.apply(this, arguments).then(function(res) {
      if (isFirewallUrl(url) && isApiUrl(url)) {
        res.clone().json().then(function(data) { captureFirewallResponse(typeof url === 'string' ? url : (url && url.url) || '', data); }).catch(function() {});
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
      if (data) captureFirewallResponse(url, data);
    });
    return origSend.apply(this, arguments);
  };

  console.log('[Firewall Capturer Active] Navigate to the Firewall section for a company in the portal. Then run: copy(JSON.stringify(window.__capturedFirewallCalls, null, 2))');
})();

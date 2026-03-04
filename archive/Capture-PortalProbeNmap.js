/**
 * Capture-PortalProbeNmap.js
 * Captures API calls when viewing probe agents / nmap interface in the ConnectSecure portal.
 * Use to discover how the portal filters probe vs lightweight agents, and what agent_type/nmap_interface values look like.
 *
 * Steps:
 * 1. Paste into ConnectSecure portal console (F12 > Console)
 * 2. Navigate to: Company > Probes (or Agents / Probe Agents / wherever probe/nmap is shown)
 * 3. Let data load; expand agent details if the portal shows nmap_interface per agent
 * 4. Export: copy(JSON.stringify(window.__capturedProbeNmapCalls, null, 2))
 * 5. Save to portal-probe-nmap-capture.json
 */
(function() {
  window.__capturedProbeNmapCalls = [];
  window.__capturedAllApiCalls = [];

  function isRelevantForProbes(url) {
    var u = (url || '').toLowerCase();
    return u.indexOf('/r/company/agents') !== -1 ||
           u.indexOf('agent_discovery_credentials') !== -1 ||
           u.indexOf('agent_credentials_mapping') !== -1 ||
           u.indexOf('agent_discoverysettings_mapping') !== -1 ||
           u.indexOf('agent_discoverysettings') !== -1 ||
           u.indexOf('lightweight_assets') !== -1 ||
           u.indexOf('probe') !== -1 ||
           u.indexOf('discovery_settings') !== -1;
  }

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

  function sampleAgentFields(obj) {
    if (!obj || typeof obj !== 'object') return null;
    var sample = {};
    var keyList = ['id', 'agent_id', 'agent_type', 'agentType', 'host_name', 'hostname', 'name', 'ip',
                   'nmap_interface', 'nmapInterface', 'probe_setting', 'probeSetting',
                   'company_id', 'discovery_settings_id', 'credentials_id'];
    keyList.forEach(function(k) {
      if (obj.hasOwnProperty(k) && obj[k] !== undefined && obj[k] !== null) {
        sample[k] = obj[k];
      }
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
    var agentSample = sampleAgentFields(firstRecord);

    var entry = {
      timestamp: new Date().toISOString(),
      url: fullUrl,
      path: path,
      method: method || 'GET',
      queryParams: queryParams,
      responseRecordCount: recordCount,
      firstRecordSample: agentSample,
      status: status
    };
    window.__capturedAllApiCalls.push(entry);

    if (isRelevantForProbes(fullUrl)) {
      window.__capturedProbeNmapCalls.push(entry);
      console.log('[Probe/Nmap Capturer]', method, path, 'params:', JSON.stringify(queryParams), 'records:', recordCount,
        agentSample && (agentSample.agent_type || agentSample.agentType || agentSample.nmap_interface) ? '(agent_type/nmap in sample)' : '');
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

  console.log('[Probe/Nmap Capturer Active] Navigate to Probes or Agents view for a company. Then run:');
  console.log('  copy(JSON.stringify(window.__capturedProbeNmapCalls, null, 2))');
  console.log('Or for all API calls: copy(JSON.stringify(window.__capturedAllApiCalls, null, 2))');
})();

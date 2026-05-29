window.vscanmagic = window.vscanmagic || {};

window.vscanmagic.copyText = function (text) {
    if (text == null) {
        return Promise.resolve(false);
    }

    var value = String(text);
    var copyViaTextarea = function () {
        var textarea = document.createElement('textarea');
        textarea.value = value;
        textarea.setAttribute('readonly', '');
        textarea.style.position = 'fixed';
        textarea.style.left = '0';
        textarea.style.top = '0';
        textarea.style.opacity = '0';
        textarea.style.pointerEvents = 'none';
        document.body.appendChild(textarea);
        textarea.focus();
        textarea.select();
        textarea.setSelectionRange(0, value.length);

        var copied = false;
        try {
            copied = document.execCommand('copy');
        } finally {
            document.body.removeChild(textarea);
        }

        return copied ? Promise.resolve(true) : Promise.reject(new Error('copy failed'));
    };

    if (navigator.clipboard && window.isSecureContext) {
        return navigator.clipboard.writeText(value)
            .then(function () { return true; })
            .catch(function () { return copyViaTextarea(); });
    }

    return copyViaTextarea();
};

window.vscanmagic.copyFromTarget = function (targetId, htmlTargetId) {
    var source = targetId ? document.getElementById(targetId) : null;
    if (!source) {
        return Promise.reject(new Error('missing copy source'));
    }

    var plainText = source.textContent || source.innerText || '';
    if (htmlTargetId) {
        var htmlSource = document.getElementById(htmlTargetId);
        var html = htmlSource ? htmlSource.innerHTML : '';
        return window.vscanmagic.copyRichHtml(plainText, html);
    }

    return window.vscanmagic.copyText(plainText);
};

window.vscanmagic.wrapHtmlFragment = function (html) {
    var fragment = String(html || '');
    var startMarker = '<!--StartFragment-->';
    var endMarker = '<!--EndFragment-->';
    var body = '<html><body>' + startMarker + fragment + endMarker + '</body></html>';
    var startHtml = body.indexOf(startMarker);
    var endHtml = body.indexOf(endMarker) + endMarker.length;
    var pad = function (value) {
        return ('0000000000' + value).slice(-10);
    };

    return [
        'Version:0.9',
        'StartHTML:' + pad(body.indexOf('<html>')),
        'EndHTML:' + pad(body.length),
        'StartFragment:' + pad(startHtml),
        'EndFragment:' + pad(endHtml),
        body
    ].join('\r\n');
};

window.vscanmagic.copyViaSelection = function (html) {
    var rich = String(html || '');
    if (!rich.trim()) {
        return Promise.reject(new Error('missing html'));
    }

    var container = document.createElement('div');
    container.contentEditable = 'true';
    container.innerHTML = rich;
    container.style.position = 'fixed';
    container.style.left = '0';
    container.style.top = '0';
    container.style.opacity = '0';
    container.style.pointerEvents = 'none';
    container.setAttribute('aria-hidden', 'true');
    document.body.appendChild(container);
    container.focus();

    var range = document.createRange();
    range.selectNodeContents(container);
    var selection = window.getSelection();
    if (!selection) {
        document.body.removeChild(container);
        return Promise.reject(new Error('copy failed'));
    }

    selection.removeAllRanges();
    selection.addRange(range);

    var copied = false;
    try {
        copied = document.execCommand('copy');
    } finally {
        selection.removeAllRanges();
        document.body.removeChild(container);
    }

    return copied ? Promise.resolve(true) : Promise.reject(new Error('copy failed'));
};

window.vscanmagic.copyRichHtml = function (plainText, html) {
    var plain = String(plainText || '');
    var rich = String(html || '');

    if (!rich.trim()) {
        return window.vscanmagic.copyText(plain);
    }

    var copyViaClipboardApi = function () {
        if (!(navigator.clipboard && window.ClipboardItem && window.isSecureContext)) {
            return Promise.reject(new Error('clipboard api unavailable'));
        }

        var wrappedHtml = window.vscanmagic.wrapHtmlFragment(rich);
        var item = new ClipboardItem({
            'text/html': Promise.resolve(new Blob([wrappedHtml], { type: 'text/html' })),
            'text/plain': Promise.resolve(new Blob([plain], { type: 'text/plain' }))
        });
        return navigator.clipboard.write([item]);
    };

    return copyViaClipboardApi()
        .catch(function () { return window.vscanmagic.copyViaSelection(rich); })
        .catch(function () { return window.vscanmagic.copyText(plain); });
};

window.vscanmagic.showCopyToast = function (message, isError) {
    var toast = document.getElementById('vscan-copy-toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'vscan-copy-toast';
        toast.style.cssText = 'position:fixed;top:1rem;right:1rem;z-index:1080;max-width:24rem;';
        document.body.appendChild(toast);
    }

    toast.className = 'alert py-2 mb-0 shadow-sm ' + (isError ? 'alert-warning' : 'alert-success');
    toast.textContent = message;
    toast.hidden = false;

    window.clearTimeout(window.vscanmagic._copyToastTimer);
    window.vscanmagic._copyToastTimer = window.setTimeout(function () {
        toast.hidden = true;
    }, 2500);
};

window.vscanmagic.initCopyButtons = function () {
    if (window.vscanmagic._copyButtonsReady) {
        return;
    }

    window.vscanmagic._copyButtonsReady = true;
    document.addEventListener('click', function (event) {
        var button = event.target.closest('[data-vscan-copy-target]');
        if (!button) {
            return;
        }

        event.preventDefault();
        event.stopPropagation();
        var targetId = button.getAttribute('data-vscan-copy-target');
        var htmlTargetId = button.getAttribute('data-vscan-copy-html-target');
        window.vscanmagic.copyFromTarget(targetId, htmlTargetId)
            .then(function () {
                window.vscanmagic.showCopyToast(
                    htmlTargetId ? 'Copied body with links to clipboard.' : 'Copied to clipboard.',
                    false);
            })
            .catch(function () {
                window.vscanmagic.showCopyToast('Copy failed — select the text manually.', true);
            });
    }, true);
};

window.vscanmagic.initCopyButtons();

window.vscanmagic.scrollToHelpSection = function (sectionId) {
    if (!sectionId) {
        return;
    }

    var target = document.getElementById(sectionId);
    if (!target) {
        return;
    }

    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    if (history.replaceState) {
        history.replaceState(null, '', '/help#' + sectionId);
    }
};

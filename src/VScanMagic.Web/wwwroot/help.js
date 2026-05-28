window.vscanmagic = window.vscanmagic || {};

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

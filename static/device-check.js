/**
 * Block phones and tablets; allow desktop/laptop browsers.
 * Used by participant-facing pages (not /manage).
 */
(function (global) {
  'use strict';

  function isMobileOrTablet() {
    var ua = navigator.userAgent || '';
    if (/iPad|Tablet|PlayBook|Silk|Kindle|KFAPWI|Tablet PC/i.test(ua)) return true;
    if (/Android/i.test(ua) && !/Mobile/i.test(ua)) return true;
    if (/Android.*Mobile|iPhone|iPod|webOS|BlackBerry|IEMobile|Opera Mini|Windows Phone/i.test(ua)) {
      return true;
    }
    if (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1) return true;
    return false;
  }

  function showBlock() {
    global.__DEVICE_BLOCKED__ = true;
    document.documentElement.classList.add('device-blocked');
    var overlay = document.getElementById('device-block-overlay');
    if (overlay) overlay.classList.remove('hidden');
    var app = document.getElementById('app');
    if (app) app.style.display = 'none';
    var header = document.querySelector('body > .header');
    if (header) header.style.display = 'none';
  }

  function isBlocked() {
    return !!global.__DEVICE_BLOCKED__;
  }

  global.DeviceCheck = {
    isMobileOrTablet: isMobileOrTablet,
    isBlocked: isBlocked,
    showBlock: showBlock,
  };

  if (isMobileOrTablet()) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', showBlock);
    } else {
      showBlock();
    }
  }
})(window);

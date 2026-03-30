/**
 * Utility functions for AscDesktopEditor stub
 * Extracted for testability - used by both desktop-stub.js and unit tests
 */

// Font sprite constants
const FONT_SPRITE_BASE_WIDTH = 300;
const FONT_SPRITE_ROW_HEIGHT = 28;

/**
 * Parse scale argument for font sprite generation
 * @param {string|number|boolean} arg - Scale argument
 * @returns {number} Parsed scale value (minimum 0.5)
 */
function parseSpriteScale(arg) {
  if (typeof arg === 'string') {
    var match = arg.match(/^@([0-9]+(?:\.[0-9]+)?)x$/);
    if (match && match[1]) {
      var parsed = parseFloat(match[1]);
      return isNaN(parsed) ? 1 : Math.max(parsed, 0.5);
    }
  } else if (typeof arg === 'number') {
    return Math.max(arg, 0.5);
  } else if (arg === true) {
    return 2;
  }
  return 1;
}

/**
 * Detect language code from URL or window.location
 * @param {object} location - Location object with search property
 * @returns {string} Language code or empty string
 */
function detectLanguageCode(location) {
  try {
    var search = location && location.search ? location.search : '';
    var match = search.match(/[?&]lang=([^&#]+)/i);
    if (match && match[1]) {
      return decodeURIComponent(match[1]).toLowerCase();
    }
  } catch (err) {
    // Silently fail
  }
  return '';
}

/**
 * Normalize incoming theme values to the two supported modes.
 * Accepts either host-friendly names ("dark"/"light") or ONLYOFFICE theme ids.
 * Any unknown value falls back to light.
 * @param {string} theme - Theme query parameter or theme id
 * @returns {"dark"|"light"} Normalized theme choice
 */
function normalizeThemeChoice(theme) {
  var normalized = typeof theme === 'string' ? theme.toLowerCase() : '';
  if (
    normalized === 'dark' ||
    normalized === 'theme-night' ||
    normalized === 'theme-dark'
  ) {
    return 'dark';
  }
  return 'light';
}

/**
 * Resolve the ONLYOFFICE uiTheme id for an incoming theme value.
 * @param {string} theme - Theme query parameter or theme id
 * @returns {string} ONLYOFFICE theme id
 */
function resolveUiThemeId(theme) {
  return normalizeThemeChoice(theme) === 'dark'
    ? 'theme-night'
    : 'theme-classic-light';
}

/**
 * Build the desktop stub theme object expected by the web apps.
 * @param {string} theme - Theme query parameter or theme id
 * @returns {{id: string, type: string, system: string}} Theme config
 */
function buildThemeConfig(theme) {
  var normalized = normalizeThemeChoice(theme);
  if (normalized === 'dark') {
    return {
      id: 'theme-night',
      type: 'dark',
      system: 'dark'
    };
  }

  return {
    id: 'theme-classic-light',
    type: 'light',
    system: 'light'
  };
}

/**
 * Check if East Asian font variant should be used
 * @param {string} lang - Language code
 * @returns {boolean} True if EA variant should be used
 */
function shouldUseEastAsiaVariant(lang) {
  if (!lang) return false;
  var baseLang = lang.replace(/[_-].*$/, '');
  return baseLang === 'zh' || baseLang === 'ja' || baseLang === 'ko';
}

/**
 * Get sprite generation options from arguments
 * @param {Array} args - Arguments array
 * @param {string} detectedLang - Detected language code
 * @returns {object} Options with scale and useEA properties
 */
function getSpriteOptions(args, detectedLang) {
  var scale = 1;
  var forceEA = false;

  for (var i = 0; i < args.length; i++) {
    var arg = args[i];
    if (typeof arg === 'string' && /_ea/.test(arg)) {
      forceEA = true;
    }
    var parsedScale = parseSpriteScale(arg);
    if (parsedScale !== 1) {
      scale = parsedScale;
    }
  }

  return {
    scale: scale,
    useEA: forceEA || shouldUseEastAsiaVariant(detectedLang)
  };
}

/**
 * Collect font names from __fonts_infos array
 * @param {Array} fontsInfos - The window.__fonts_infos array
 * @param {boolean} useEA - Whether to use East Asian variant names
 * @returns {Array<string>} Array of font names
 */
function collectFontNames(fontsInfos, useEA) {
  if (!Array.isArray(fontsInfos) || fontsInfos.length === 0) {
    return [];
  }

  var names = [];
  for (var i = 0; i < fontsInfos.length; i++) {
    var entry = fontsInfos[i];
    if (!Array.isArray(entry) || !entry[0]) {
      continue;
    }
    var name = useEA && entry.length > 9 && entry[9] ? entry[9] : entry[0];
    names.push(name);
  }

  return names;
}

/**
 * Check if a file path is an image based on extension
 * @param {string} path - File path or name
 * @returns {boolean} True if file is an image
 */
function isImageFile(path) {
  if (!path) return false;
  var ext = path.split('.').pop().toLowerCase();
  return ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].indexOf(ext) !== -1;
}

/**
 * Extract filename from a media path
 * Handles various path formats: media/img.png, ./media/img.png, Editor.bin/media/img.png
 * @param {string} path - The media path
 * @returns {string} Extracted filename
 */
function extractMediaFilename(path) {
  if (!path) return '';
  
  var filename = path;
  if (path.startsWith('media/')) {
    filename = path.substring(6);
  } else if (path.startsWith('./media/')) {
    filename = path.substring(8);
  } else if (path.indexOf('/media/') !== -1) {
    filename = path.substring(path.indexOf('/media/') + 7);
  }
  return filename;
}

/**
 * Build server URL for media file
 * @param {string} baseUrl - Server base URL
 * @param {string} fileHash - Document file hash
 * @param {string} filename - Media filename
 * @returns {string|null} Full URL to media file, or null if fileHash is missing
 */
function buildMediaUrl(baseUrl, fileHash, filename) {
  if (!fileHash) return null;
  return baseUrl + '/api/media/' + fileHash + '/' + encodeURIComponent(filename);
}

/**
 * Extract blob URL from SDK-wrapped path
 * @param {string} path - Potentially wrapped path
 * @returns {string|null} Extracted blob URL or null
 */
function extractBlobUrl(path) {
  if (!path) return null;
  var blobMatch = path.match(/blob:https?:\/\/[^\/]+\/[a-f0-9-]+/);
  return blobMatch ? blobMatch[0] : null;
}

// UMD export - works in browser (global) and Node.js/Bun (CommonJS)
(function(root, factory) {
  var exports = {
    FONT_SPRITE_BASE_WIDTH: FONT_SPRITE_BASE_WIDTH,
    FONT_SPRITE_ROW_HEIGHT: FONT_SPRITE_ROW_HEIGHT,
    parseSpriteScale: parseSpriteScale,
    detectLanguageCode: detectLanguageCode,
    normalizeThemeChoice: normalizeThemeChoice,
    resolveUiThemeId: resolveUiThemeId,
    buildThemeConfig: buildThemeConfig,
    shouldUseEastAsiaVariant: shouldUseEastAsiaVariant,
    getSpriteOptions: getSpriteOptions,
    collectFontNames: collectFontNames,
    isImageFile: isImageFile,
    extractMediaFilename: extractMediaFilename,
    buildMediaUrl: buildMediaUrl,
    extractBlobUrl: extractBlobUrl
  };
  
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = exports;
  } else if (typeof root !== 'undefined') {
    root.DesktopStubUtils = exports;
  }
})(typeof window !== 'undefined' ? window : this);

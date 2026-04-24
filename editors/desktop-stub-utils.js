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
 * Extract theme query parameter from a search string.
 * @param {string} search - URL search string
 * @returns {string} Theme query value or empty string
 */
function extractThemeQueryValue(search) {
  try {
    var searchValue = typeof search === 'string' ? search : '';
    if (!searchValue) {
      return '';
    }
    return new URLSearchParams(searchValue).get('theme') || '';
  } catch (err) {
    return '';
  }
}

/**
 * Find the first theme value from a list of search strings.
 * @param {Array<string>} searches - Ordered search strings to inspect
 * @returns {string} First non-empty theme query value or empty string
 */
function findThemeChoice(searches) {
  if (!Array.isArray(searches)) {
    return '';
  }

  for (var i = 0; i < searches.length; i++) {
    var theme = extractThemeQueryValue(searches[i]);
    if (theme) {
      return theme;
    }
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

var SELECTION_CHANGED_MESSAGE_TYPE = 'ONLYOFFICE_SELECTION_CHANGED';

function isPrimitiveSelectionValue(value) {
  return value === null ||
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean';
}

function toSelectionString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value);
}

function readSelectionValue(source, names) {
  if (!source) {
    return undefined;
  }

  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    try {
      var value = source[name];
      if (typeof value === 'function') {
        value = value.call(source);
      }
      if (value !== undefined && value !== null) {
        return value;
      }
    } catch (err) {}
  }

  return undefined;
}

function readSelectionFunction(source, name, args) {
  if (!source || typeof source[name] !== 'function') {
    return undefined;
  }

  try {
    return source[name].apply(source, args || []);
  } catch (err) {
    return undefined;
  }
}

function extractSelectedText(api) {
  var selectedText = readSelectionFunction(api, 'asc_GetSelectedText', [false]);
  return toSelectionString(selectedText);
}

function getCellText(cellInfo) {
  var value = readSelectionValue(cellInfo, [
    'text',
    'Text',
    'value',
    'Value',
    'asc_getObjectValue',
    'asc_getText',
    'asc_getValue',
    'getText',
    'getValue'
  ]);
  return value === undefined || value === null ? null : toSelectionString(value);
}

function getCellName(cellInfo) {
  var value = readSelectionValue(cellInfo, [
    'name',
    'Name',
    'cell',
    'cellRef',
    'asc_getName',
    'getName'
  ]);
  return value === undefined || value === null ? null : toSelectionString(value);
}

function columnIndexToName(index) {
  var col = Number(index);
  if (!isFinite(col) || col < 0) {
    return '';
  }

  var colStr = '';
  do {
    colStr = String.fromCharCode(65 + (col % 26)) + colStr;
    col = Math.floor(col / 26) - 1;
  } while (col >= 0);
  return colStr;
}

function getSelectionRangeCellName(api, iframeWindow) {
  if (!api || !iframeWindow || !iframeWindow.AscCommonExcel) {
    return null;
  }

  var wbModel = api.wbModel || readSelectionFunction(api, 'asc_getModel');
  if (!wbModel) {
    return null;
  }

  var activeWs = readSelectionFunction(wbModel, 'getActiveWs');
  if (!activeWs || !activeWs.selectionRange) {
    return null;
  }

  var range = readSelectionFunction(activeWs.selectionRange, 'getLast');
  if (!range || range.c1 === undefined || range.r1 === undefined) {
    return null;
  }

  var columnName = columnIndexToName(range.c1);
  if (!columnName) {
    return null;
  }
  return columnName + (Number(range.r1) + 1);
}

function getCellReference(api, iframeWindow, eventCellInfo) {
  var info = readSelectionFunction(api, 'asc_getCellInfo') || eventCellInfo;
  var cellRef = getCellName(info) || getCellName(eventCellInfo);

  if (!cellRef) {
    cellRef = readSelectionFunction(api, 'asc_getActiveRangeStr');
  }

  if (!cellRef && api && api.wb) {
    var ws = readSelectionFunction(api.wb, 'getWorksheet');
    if (ws) {
      cellRef = readSelectionFunction(ws, 'getSelectionRangeStr');
      if (!cellRef && ws.model && ws.model.selectionRange) {
        cellRef = readSelectionFunction(ws.model.selectionRange, 'getName');
      }
    }
  }

  if (!cellRef) {
    cellRef = getSelectionRangeCellName(api, iframeWindow);
  }

  return cellRef === undefined || cellRef === null ? null : toSelectionString(cellRef);
}

function buildCellSelectionPayload(api, iframeWindow, eventCellInfo) {
  var info = readSelectionFunction(api, 'asc_getCellInfo') || eventCellInfo || null;
  var cellRef = getCellReference(api, iframeWindow, eventCellInfo);
  var range = readSelectionFunction(api, 'asc_getActiveRangeStr', [undefined, false, true]);
  var activeCell = readSelectionFunction(api, 'asc_getActiveRangeStr', [undefined, true, true]);

  if (range === undefined || range === null || range === '') {
    range = cellRef;
  }
  if (activeCell === undefined || activeCell === null || activeCell === '') {
    activeCell = cellRef;
  }

  var sheetIndex = readSelectionFunction(api, 'asc_getActiveWorksheetIndex');
  if (sheetIndex === undefined || sheetIndex === null) {
    sheetIndex = readSelectionFunction(api, 'getActiveWorksheetIndex');
  }
  if (sheetIndex === undefined || sheetIndex === null) {
    sheetIndex = 0;
  }

  var text = getCellText(info);
  if (text === null) {
    text = getCellText(eventCellInfo);
  }

  return {
    kind: 'cell',
    cell: cellRef,
    range: range === undefined || range === null ? null : toSelectionString(range),
    activeCell: activeCell === undefined || activeCell === null ? null : toSelectionString(activeCell),
    sheetIndex: sheetIndex,
    text: text
  };
}

function getSelectedElements(api, focusedElements) {
  if (Array.isArray(focusedElements)) {
    return focusedElements;
  }

  var selectedElements = readSelectionFunction(api, 'getSelectedElements', [true]);
  return Array.isArray(selectedElements) ? selectedElements : [];
}

function getImageUrlFromElement(element, seenElements) {
  seenElements = seenElements || [];
  if (!element || seenElements.indexOf(element) !== -1) {
    return null;
  }
  seenElements.push(element);

  var value = readSelectionValue(element, [
    'ImageUrl',
    'imageUrl',
    'Url',
    'url',
    'src',
    'Src',
    'Source',
    'source',
    'getImageUrl',
    'get_ImageUrl',
    'asc_getImageUrl',
    'getUrl'
  ]);
  if (value !== undefined && value !== null) {
    return toSelectionString(value);
  }

  var nestedValue = readSelectionValue(element, ['Value', 'value', 'asc_getObjectValue', 'getValue']);
  if (nestedValue && typeof nestedValue === 'object') {
    return getImageUrlFromElement(nestedValue, seenElements);
  }

  return null;
}

function stripImageUrlNoise(imageUrl) {
  return toSelectionString(imageUrl).split('?')[0].split('#')[0];
}

function getMediaFilenameFromImageUrl(imageUrl) {
  var cleanedUrl = stripImageUrlNoise(imageUrl);
  if (!cleanedUrl) {
    return '';
  }

  if (cleanedUrl.indexOf('/media/') !== -1 || cleanedUrl.indexOf('media/') === 0 || cleanedUrl.indexOf('./media/') === 0) {
    return extractMediaFilename(cleanedUrl);
  }

  if (cleanedUrl.indexOf('/') === -1 && isImageFile(cleanedUrl)) {
    return cleanedUrl;
  }

  return '';
}

function normalizeSelectedElementImageUrl(imageUrl, mediaContext) {
  if (!imageUrl) {
    return null;
  }

  var cleanedUrl = stripImageUrlNoise(imageUrl);
  if (/^https?:\/\//i.test(cleanedUrl) || /^data:/i.test(cleanedUrl) || /^blob:/i.test(cleanedUrl)) {
    return cleanedUrl;
  }

  var filename = getMediaFilenameFromImageUrl(cleanedUrl);
  if (!filename) {
    return cleanedUrl;
  }

  mediaContext = mediaContext || {};
  if (mediaContext.baseUrl && mediaContext.fileHash) {
    return buildMediaUrl(mediaContext.baseUrl, mediaContext.fileHash, filename);
  }

  if (mediaContext.docBaseUrl) {
    return mediaContext.docBaseUrl.replace(/\/$/, '') + '/media/' + encodeURIComponent(filename);
  }

  return cleanedUrl;
}

function serializeSelectedElements(elements, mediaContext) {
  if (!Array.isArray(elements)) {
    return [];
  }

  var serialized = [];
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    if (!element) {
      continue;
    }

    var type = readSelectionValue(element, [
      'Type',
      'type',
      'ObjectType',
      'objectType',
      'asc_getObjectType',
      'getType',
      'get_Type',
      'asc_getType'
    ]);
    var value = readSelectionValue(element, ['Value', 'value', 'asc_getObjectValue', 'getValue']);
    var rawImageUrl = getImageUrlFromElement(element);
    var imageUrl = normalizeSelectedElementImageUrl(rawImageUrl, mediaContext);
    var imageName = getMediaFilenameFromImageUrl(rawImageUrl);
    var objectId = readSelectionValue(element, [
      'Id',
      'id',
      'ObjectId',
      'objectId',
      'getId',
      'get_Id'
    ]);

    var item = {};
    if (isPrimitiveSelectionValue(type)) {
      item.type = type;
    }
    if (isPrimitiveSelectionValue(value)) {
      item.value = value;
    }
    if (isPrimitiveSelectionValue(objectId)) {
      item.id = objectId;
    }
    if (imageUrl) {
      item.imageUrl = imageUrl;
    }
    if (imageName) {
      item.imageName = imageName;
    }

    var typeText = item.type === undefined || item.type === null ? '' : String(item.type).toLowerCase();
    item.hasImage = !!imageUrl || typeText.indexOf('image') !== -1 || typeText.indexOf('picture') !== -1;

    serialized.push(item);
  }

  return serialized;
}

function buildDocumentSelectionPayload(api, focusedElements, mediaContext) {
  var text = extractSelectedText(api);
  var objects = serializeSelectedElements(getSelectedElements(api, focusedElements), mediaContext);

  if (text) {
    var textPayload = {
      kind: 'text',
      text: text
    };
    if (objects.length > 0) {
      textPayload.objects = objects;
    }
    return textPayload;
  }

  if (objects.length > 0) {
    var hasImage = objects.some(function(object) { return object.hasImage; });
    return {
      kind: hasImage ? 'image' : 'object',
      objects: objects
    };
  }

  return {
    kind: 'empty',
    text: ''
  };
}

function buildSelectionMessage(meta, selection, timestamp) {
  meta = meta || {};
  return {
    type: SELECTION_CHANGED_MESSAGE_TYPE,
    filePath: meta.filePath || meta.filepath || '',
    filename: meta.filename || '',
    doctype: meta.doctype || '',
    selection: selection || { kind: 'empty', text: '' },
    timestamp: timestamp === undefined ? Date.now() : timestamp
  };
}

function createSelectionStream(options) {
  options = options || {};
  var parentWindow = options.parentWindow;
  var targetOrigin = options.targetOrigin || '*';
  var now = typeof options.now === 'function' ? options.now : function() { return Date.now(); };
  var api = options.api || null;
  var iframeWindow = options.iframeWindow || null;
  var meta = {
    doctype: options.doctype || '',
    filePath: options.filePath || options.filepath || '',
    filename: options.filename || ''
  };
  var mediaContext = {
    baseUrl: options.baseUrl || '',
    fileHash: options.fileHash || '',
    docBaseUrl: options.docBaseUrl || ''
  };
  var lastPayloadKey = null;

  function emit(selection) {
    var payloadKey = JSON.stringify(selection || null);
    if (payloadKey === lastPayloadKey) {
      return null;
    }
    lastPayloadKey = payloadKey;

    var message = buildSelectionMessage(meta, selection, now());
    if (parentWindow && typeof parentWindow.postMessage === 'function') {
      parentWindow.postMessage(message, targetOrigin);
    }
    return message;
  }

  return {
    emit: emit,
    emitCell: function(cellInfo) {
      return emit(buildCellSelectionPayload(api, iframeWindow, cellInfo));
    },
    emitDocument: function(focusedElements) {
      return emit(buildDocumentSelectionPayload(api, focusedElements, mediaContext));
    }
  };
}

// UMD export - works in browser (global) and Node.js/Bun (CommonJS)
(function(root, factory) {
  var exports = {
    FONT_SPRITE_BASE_WIDTH: FONT_SPRITE_BASE_WIDTH,
    FONT_SPRITE_ROW_HEIGHT: FONT_SPRITE_ROW_HEIGHT,
    parseSpriteScale: parseSpriteScale,
    detectLanguageCode: detectLanguageCode,
    extractThemeQueryValue: extractThemeQueryValue,
    findThemeChoice: findThemeChoice,
    normalizeThemeChoice: normalizeThemeChoice,
    resolveUiThemeId: resolveUiThemeId,
    buildThemeConfig: buildThemeConfig,
    shouldUseEastAsiaVariant: shouldUseEastAsiaVariant,
    getSpriteOptions: getSpriteOptions,
    collectFontNames: collectFontNames,
    isImageFile: isImageFile,
    extractMediaFilename: extractMediaFilename,
    buildMediaUrl: buildMediaUrl,
    extractBlobUrl: extractBlobUrl,
    SELECTION_CHANGED_MESSAGE_TYPE: SELECTION_CHANGED_MESSAGE_TYPE,
    extractSelectedText: extractSelectedText,
    serializeSelectedElements: serializeSelectedElements,
    normalizeSelectedElementImageUrl: normalizeSelectedElementImageUrl,
    buildCellSelectionPayload: buildCellSelectionPayload,
    buildDocumentSelectionPayload: buildDocumentSelectionPayload,
    buildSelectionMessage: buildSelectionMessage,
    createSelectionStream: createSelectionStream
  };
  
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = exports;
  } else if (typeof root !== 'undefined') {
    root.DesktopStubUtils = exports;
  }
})(typeof window !== 'undefined' ? window : this);

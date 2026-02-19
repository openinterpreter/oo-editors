/**
 * AscDesktopEditor Stub
 * Browser implementation of ONLYOFFICE's native desktop API (AscDesktopEditor).
 *
 * NOTE(victor): ARCHITECTURE OVERVIEW
 * ====================================
 * In native ONLYOFFICE Desktop, AscDesktopEditor is a C++ object injected via CEF
 * (Chromium Embedded Framework) that provides file I/O, font loading, clipboard, etc.
 *
 * This stub replaces that native API with HTTP calls to our Express server.
 *
 * KEY FUNCTION: LocalStartOpen()
 * ------------------------------
 * The SDK calls LocalStartOpen() when it's ready to receive document data. This is
 * the SDK's signal that initialization is complete. Our implementation:
 * 1. Waits for window._commonJsReady (common.js overrides complete)
 * 2. Retrieves document binary from window._currentDocumentBinary
 * 3. Calls DesktopOfflineAppDocumentEndLoad() to load into SDK
 *
 * CRITICAL: Do NOT call DesktopOfflineAppDocumentEndLoad() from anywhere else.
 * The SDK orchestrates loading via LocalStartOpen(). Calling it directly causes
 * race conditions where SDK objects don't exist yet.
 *
 * LOADING SEQUENCE:
 * offline-loader-proper.html -> DocsAPI -> SDK init -> LocalStartOpen() -> load doc
 *
 * See common.js header for full architecture documentation.
 */

(function() {
    'use strict';

    var APP_VERSION = '1.0.18';

    // NOTE(victor): sdk-all-min.js and sdk-all.js are code-split chunks from Closure
    // Compiler, NOT minified vs unminified. Both must load. The "-min" suffix means
    // "minimum/core" not "minified". Setting this false ensures loadSdk() in
    // editorscommon.js always loads the second chunk (sdk-all.js).
    window['AscNotLoadAllScript'] = false;

    // Threshold for detecting infinite loop calls (calls per second)
    var INFINITE_LOOP_THRESHOLD = 30;

    // Store dropped files for GetDropFiles()
    // Upload images to server for persistence
    window._droppedFiles = [];
    window._uploadedFileMap = {}; // blob URL -> server filename

    document.addEventListener('drop', function(e) {
        window._droppedFiles = [];
        window._uploadedFileMap = {};

        if (!e.dataTransfer || !e.dataTransfer.files) return;

        var fileHash = window._ONLYOFFICE_FILE_HASH;
        if (!fileHash && window.parent && window.parent !== window) {
            fileHash = window.parent._ONLYOFFICE_FILE_HASH;
        }

        for (var i = 0; i < e.dataTransfer.files.length; i++) {
            var file = e.dataTransfer.files[i];
            var ext = file.name.split('.').pop().toLowerCase();
            var isImage = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].indexOf(ext) !== -1;

            if (isImage && fileHash) {
                // Synchronously upload image to server for persistence
                var xhr = new XMLHttpRequest();
                var serverFilename = 'dropped_' + Date.now() + '_' + i + '.' + ext;
                xhr.open('POST', 'http://localhost:8080/api/media/' + fileHash + '?filename=' + encodeURIComponent(serverFilename), false);
                xhr.setRequestHeader('Content-Type', 'application/octet-stream');

                // Send File directly (File extends Blob, XHR handles it)
                xhr.send(file);

                if (xhr.status === 200) {
                    var response = JSON.parse(xhr.responseText);
                    var blobUrl = URL.createObjectURL(file);
                    window._uploadedFileMap[blobUrl] = response.filename;
                    window._droppedFiles.push(file);
                    console.log('[BROWSER] Uploaded dropped image:', response.filename);
                } else {
                    console.error('[BROWSER] Failed to upload dropped image:', xhr.status);
                }
            } else {
                window._droppedFiles.push(file);
            }
        }
    }, true);

    // Only create if it doesn't exist
    if (window.AscDesktopEditor) {
        console.log('AscDesktopEditor already exists, skipping stub creation');
        return;
    }

    console.log('Creating AscDesktopEditor stub for browser environment');

    // Server configuration
    var SERVER_PORT = 8080;
    var SERVER_BASE_URL = 'http://localhost:' + SERVER_PORT;

    // Import utils from global (loaded before this script)
    var utils = window.DesktopStubUtils || {};
    var FONT_SPRITE_BASE_WIDTH = utils.FONT_SPRITE_BASE_WIDTH || 300;
    var FONT_SPRITE_ROW_HEIGHT = utils.FONT_SPRITE_ROW_HEIGHT || 28;
    var fontSpriteCache = {};

    // Use utils functions with fallbacks
    var parseSpriteScale = utils.parseSpriteScale || function(arg) { return 1; };
    var extractMediaFilename = utils.extractMediaFilename || function(p) { return p; };
    var buildMediaUrl = utils.buildMediaUrl || function(base, hash, name) { return hash ? base + '/api/media/' + hash + '/' + name : null; };
    var extractBlobUrl = utils.extractBlobUrl || function(p) { return null; };

    function detectLanguageCode() {
        if (utils.detectLanguageCode) {
            return utils.detectLanguageCode(window.location);
        }
        return '';
    }

    function shouldUseEastAsiaVariant() {
        var lang = detectLanguageCode();
        if (!lang && window.Common && window.Common.Locale && typeof window.Common.Locale.getCurrentLanguage === 'function') {
            try {
                lang = (window.Common.Locale.getCurrentLanguage() || '').toLowerCase();
            } catch (err) {
                console.warn('[BROWSER] Failed to read Common.Locale language:', err);
            }
        }
        return utils.shouldUseEastAsiaVariant ? utils.shouldUseEastAsiaVariant(lang) : false;
    }

    function getSpriteOptions(args) {
        var lang = detectLanguageCode();
        if (!lang && window.Common && window.Common.Locale && typeof window.Common.Locale.getCurrentLanguage === 'function') {
            try {
                lang = (window.Common.Locale.getCurrentLanguage() || '').toLowerCase();
            } catch (err) {}
        }
        if (utils.getSpriteOptions) {
            return utils.getSpriteOptions(args, lang);
        }
        return { scale: 1, useEA: false };
    }

    function collectFontNames(useEA) {
        if (utils.collectFontNames) {
            return utils.collectFontNames(window["__fonts_infos"], useEA);
        }
        return [];
    }

    function createTransparentPixel() {
        return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4XmNk+M8ABAwDASMWrt0AAAAASUVORK5CYII=';
    }

    function renderFontSprite(scale, useEA) {
        try {
            var fontNames = collectFontNames(useEA);
            if (fontNames.length === 0) {
                return createTransparentPixel();
            }

            var canvas = document.createElement('canvas');
            var widthPx = Math.max(1, Math.round(FONT_SPRITE_BASE_WIDTH * scale));
            var rowHeightPx = Math.max(12, Math.round(FONT_SPRITE_ROW_HEIGHT * scale));
            var heightPx = Math.max(rowHeightPx, fontNames.length * rowHeightPx);

            canvas.width = widthPx;
            canvas.height = heightPx;

            var ctx = canvas.getContext('2d');
            if (!ctx) {
                return createTransparentPixel();
            }

            ctx.clearRect(0, 0, widthPx, heightPx);
            ctx.textBaseline = 'middle';
            ctx.imageSmoothingEnabled = true;

            var fontSize = Math.max(12, Math.round(13 * scale));
            var labelX = Math.round(8 * scale);

            for (var row = 0; row < fontNames.length; row++) {
                var y = row * rowHeightPx;
                if (row % 2 === 1) {
                    ctx.fillStyle = 'rgba(0, 0, 0, 0.04)';
                    ctx.fillRect(0, y, widthPx, rowHeightPx);
                }

                var centerY = y + rowHeightPx / 2;
                var fontName = fontNames[row];

                ctx.fillStyle = '#222';
                ctx.font = fontSize + 'px "' + fontName.replace(/"/g, '') + '", sans-serif';
                ctx.fillText(fontName, labelX, centerY);
            }

            return canvas.toDataURL('image/png');
        } catch (err) {
            console.error('[BROWSER] Failed to render font sprite:', err);
            return createTransparentPixel();
        }
    }

    // Initialize global state for tracking document changes
    window._documentChanges = [];
    window._hasUnsavedChanges = false;
    window._documentLoaded = false;

    // Initialize image cache for base64 encoded images
    window._imageCache = {};

    // Create the main AscDesktopEditor object
    window.AscDesktopEditor = {
        // Features object that tracks desktop capabilities
        features: {
            singlewindow: false
        },

        // Main command execution interface
        execCommand: function(command, params) {
            console.log('AscDesktopEditor.execCommand:', command, params);

            // Handle specific commands that might be expected
            switch(command) {
                case 'webapps:entry':
                    console.log('Desktop editor entry point called with features:', params);
                    break;
                case 'portal:login':
                    console.log('Portal login requested');
                    break;
                case 'portal:logout':
                    console.log('Portal logout requested');
                    break;
                case 'editor:event':
                    console.log('Editor event:', params);
                    break;
                default:
                    console.log('Unhandled desktop command:', command);
            }

            return true;
        },

        // Call function for synchronous operations
        Call: function(method, ...args) {
            console.log('AscDesktopEditor.Call:', method, args);
            return null;
        },

        // Check if running in desktop mode
        isDesktopApp: function() {
            return false;
        },

        // Get application version
        GetVersion: function() {
            return {
                major: 0,
                minor: 0,
                build: 0,
                revision: 0
            };
        },

        // Theme support
        theme: {
            id: 'theme-classic-light',
            type: 'light',
            system: 'light'
        },

        // Crypto plugin support (stub)
        CryptoMode: 0,

        // Local file operations (stub)
        LocalFileOpen: function() {
            console.log('LocalFileOpen called - not available in browser');
            return false;
        },

        LocalFileSave: function() {
            console.log('LocalFileSave called - not available in browser');
            return false;
        },

        LocalFileRecents: function() {
            console.log('LocalFileRecents called - not available in browser');
            return [];
        },

        // Print support (stub)
        Print: function() {
            console.log('Print called - using browser print instead');
            window.print();
        },

        // Download support (stub)
        DownloadAs: function() {
            console.log('DownloadAs called - not available in browser stub');
        },

        // Clipboard operations - called by SDK for both context menu and keyboard shortcuts
        Copy: function() {
            console.log('[BROWSER] AscDesktopEditor.Copy() called');
            try {
                var result = document.execCommand('copy');
                console.log('[BROWSER] Copy execCommand result:', result);
                return result;
            } catch (e) {
                console.error('[BROWSER] Copy failed:', e);
                return false;
            }
        },

        Cut: function() {
            console.log('[BROWSER] AscDesktopEditor.Cut() called');
            try {
                var result = document.execCommand('cut');
                console.log('[BROWSER] Cut execCommand result:', result);
                return result;
            } catch (e) {
                console.error('[BROWSER] Cut failed:', e);
                return false;
            }
        },

        Paste: function() {
            console.log('[BROWSER] AscDesktopEditor.Paste() called');
            try {
                var result = document.execCommand('paste');
                console.log('[BROWSER] Paste execCommand result:', result);
                return result;
            } catch (e) {
                console.error('[BROWSER] Paste failed:', e);
                return false;
            }
        },

        // Native dialog support (stub)
        OpenFilenameDialog: function() {
            console.log('OpenFilenameDialog called - not available in browser');
            return null;
        },

        // Window management
        SetWindowSize: function(width, height) {
            console.log('SetWindowSize called:', width, height);
        },

        SetWindowPosition: function(x, y) {
            console.log('SetWindowPosition called:', x, y);
        },

        // System info
        GetSystemTheme: function() {
            return 'light';
        },

        // Create Editor API - required for desktop mode
        // This is THE CRITICAL function that creates the editor
        CreateEditorApi: function(elementId, config) {
            console.log('[DESKTOP] CreateEditorApi called');
            console.log('[DESKTOP] elementId:', elementId);
            console.log('[DESKTOP] config:', config);

            // IMPORTANT: Return the element itself, not create the API
            // The SDK's internal code will call the constructor itself
            // We just need to return the DOM element it should use

            var element = null;

            if (typeof elementId === 'string') {
                // Try common container element IDs
                element = document.getElementById(elementId) ||
                         document.getElementById('editor_sdk') ||
                         document.getElementById('editor');
                console.log('[DESKTOP] Found element by ID:', element?.id || 'not found');
            } else if (elementId && elementId.nodeType) {
                element = elementId;
                console.log('[DESKTOP] Using provided DOM element:', element.id);
            }

            if (!element) {
                // Fallback: use editor_sdk which should exist from app.js
                element = document.getElementById('editor_sdk');
                console.log('[DESKTOP] Fallback to editor_sdk:', !!element);
            }

            if (!element) {
                console.error('[DESKTOP] FATAL: No container element found!');
                return null;
            }

            console.log('[DESKTOP] Returning element for SDK:', element.id);

            // Return the element - the SDK will handle the rest
            return element;
        },

        // Set document name
        SetDocumentName: function(name) {
            console.log('SetDocumentName called:', name);
        },

        // Get installed plugins - must return JSON string with proper structure
        // The SDK's UpdateInstallPlugins expects array with 2 elements (system + user plugins)
        // Each element must have 'url' and 'pluginsData' properties
        GetInstallPlugins: function() {
            console.log('[BROWSER] GetInstallPlugins called');
            return JSON.stringify([
                {
                    "url": "file:///system/plugins/",
                    "pluginsData": []
                },
                {
                    "url": "file:///user/plugins/",
                    "pluginsData": []
                }
            ]);
        },

        // Get opened file binary data - used by offline document loading
        GetOpenedFile: function(ref) {
            console.log('[BROWSER] GetOpenedFile called with:', ref);
            // ref will be "binary_content://filename.xlsx"
            // Return the ArrayBuffer stored earlier
            if (window._currentDocumentBinary) {
                console.log('[BROWSER] Returning stored document binary, size:', window._currentDocumentBinary.byteLength);
                return window._currentDocumentBinary;  // Must return ArrayBuffer
            }
            console.log('[BROWSER] ERROR: No document binary stored!');
            return null;
        },

        // Check user ID for offline editing
        CheckUserId: function() {
            var userId = "browser-user-" + Date.now();
            console.log('[BROWSER] CheckUserId returning:', userId);
            return userId;
        },

        // Get source path for local files
        LocalFileGetSourcePath: function(path) {
            if (path === undefined) {
                var urlParams = new URLSearchParams(window.location.search);
                var parentFilename = null;
                try { parentFilename = window.parent._ONLYOFFICE_FILENAME || window.parent._ONLYOFFICE_FILEPATH; } catch(e) {}
                var filename = urlParams.get('title') || urlParams.get('filename') || parentFilename || 'document.xlsx';
                console.log('[BROWSER] LocalFileGetSourcePath (no args), returning:', filename);
                return filename;
            }
            console.log('[BROWSER] LocalFileGetSourcePath called with:', path);
            return path;
        },

        // LocalStartOpen - called by SDK when ready for document loading
        // This is the SDK's signal that all initialization is complete
        LocalStartOpen: function() {
            console.log('[BROWSER] LocalStartOpen called - SDK is ready for document');

            // Extract and store filename from URL for later save operations
            var urlParams = new URLSearchParams(window.location.search);
            var filenameFromTitle = urlParams.get('title');
            var filenameFromParam = urlParams.get('filename');

            // Also check the URL path (e.g., /open/OLIVER.xlsx)
            var pathMatch = window.location.pathname.match(/\/open\/([^\/]+)$/);
            var filenameFromPath = pathMatch ? pathMatch[1] : null;

            var filename = filenameFromTitle || filenameFromParam || filenameFromPath;

            if (filename) {
                window._ONLYOFFICE_FILENAME = filename;
                console.log('[BROWSER] Stored filename for save:', filename);
            }

            // CRITICAL: Trigger document loading now that SDK is ready
            // Wait for common.js overrides to complete, then load the document
            var ready = window._commonJsReady || Promise.resolve();
            ready.then(function() {
                console.log('[BROWSER] common.js ready, triggering document load');

                // Get document info from parent window or current window
                var docBinary = window._currentDocumentBinary;
                var docFilename = window._ONLYOFFICE_FILENAME;
                var docUrl = window._ONLYOFFICE_DOC_BASE_URL;

                // Try parent window if not found in current
                if (!docBinary && window.parent && window.parent !== window) {
                    docBinary = window.parent._currentDocumentBinary;
                    window._currentDocumentBinary = docBinary;
                }
                if (!docFilename && window.parent && window.parent !== window) {
                    docFilename = window.parent._ONLYOFFICE_FILENAME;
                    window._ONLYOFFICE_FILENAME = docFilename;
                }
                if (!docUrl && window.parent && window.parent !== window) {
                    docUrl = window.parent._ONLYOFFICE_DOC_BASE_URL;
                    window._ONLYOFFICE_DOC_BASE_URL = docUrl;
                }
                // Also copy file hash if needed
                if (!window._ONLYOFFICE_FILE_HASH && window.parent && window.parent !== window) {
                    window._ONLYOFFICE_FILE_HASH = window.parent._ONLYOFFICE_FILE_HASH;
                }
                if (!window._ONLYOFFICE_FILEPATH && window.parent && window.parent !== window) {
                    window._ONLYOFFICE_FILEPATH = window.parent._ONLYOFFICE_FILEPATH;
                }

                if (!docBinary) {
                    console.error('[BROWSER] No document binary available for loading!');
                    return;
                }

                docFilename = docFilename || 'document';

                // Call DesktopOfflineAppDocumentEndLoad to load the document
                if (window.DesktopOfflineAppDocumentEndLoad) {
                    var binaryRef = 'binary_content://' + docFilename;
                    var effectiveUrl = docUrl || docFilename;
                    console.log('[BROWSER] Calling DesktopOfflineAppDocumentEndLoad');
                    console.log('[BROWSER]   URL:', effectiveUrl);
                    console.log('[BROWSER]   Binary ref:', binaryRef);
                    console.log('[BROWSER]   Binary size:', docBinary.byteLength);
                    window.DesktopOfflineAppDocumentEndLoad(effectiveUrl, binaryRef, 0);
                } else {
                    console.error('[BROWSER] DesktopOfflineAppDocumentEndLoad not available!');
                }
            }).catch(function(err) {
                console.error('[BROWSER] Error in LocalStartOpen:', err);
            });
        },

        // Document title/caption
        setDocumentCaption: function(caption) {
            console.log('[BROWSER] setDocumentCaption:', caption);
            document.title = caption || 'ONLYOFFICE';
        },

        // Editor configuration - CRITICAL for fixing customization errors
        GetEditorConfig: function() {
            console.log('[BROWSER] GetEditorConfig called');
            return JSON.stringify({
                customization: {
                    autosave: true,
                    chat: false,
                    comments: false,
                    help: false,
                    hideRightMenu: false,
                    compactHeader: true
                },
                mode: 'edit',
                canCoAuthoring: false,
                canBackToFolder: false,
                canPlugins: true,
                isDesktopApp: false
            });
        },

        // Font sprite generation - creates canvas preview of fonts
        GetFontsSprite: function() {
            console.log('[BROWSER] GetFontsSprite called - delegating to getFontsSprite');
            return window.AscDesktopEditor.getFontsSprite.apply(window.AscDesktopEditor, arguments);
        },

        getFontsSprite: function() {
            var options = getSpriteOptions(arguments);
            var cacheKey = options.scale + '::' + (options.useEA ? 'ea' : 'default');

            if (!fontSpriteCache[cacheKey]) {
                console.log('[BROWSER] Generating font sprite with scale', options.scale, 'EA variant:', options.useEA);
                fontSpriteCache[cacheKey] = renderFontSprite(options.scale, options.useEA);
            } else {
                console.log('[BROWSER] Font sprite cache hit for', cacheKey);
            }

            return fontSpriteCache[cacheKey];
        },

        GetSpellCheckLanguages: function() {
            console.log('[BROWSER] GetSpellCheckLanguages called');
            return [];
        },

        // Event handlers - just log and return
        on: function(event, callback) {
            console.log('[BROWSER] AscDesktopEditor.on:', event);
        },

        sendSystemMessage: function(msg) {
            console.log('[BROWSER] sendSystemMessage:', msg);
        },

        SetEventToParent: function(event) {
            console.log('[BROWSER] SetEventToParent:', event);
        },

        // Cloud/sync features (disabled)
        isOffline: function() {
            return true;
        },

        isViewMode: function() {
            return false;
        },

        IsLocalFile: function() {
            console.log('[BROWSER] IsLocalFile called - returning true');
            return true;
        },

        // Plugins
        InstallPlugin: function(guid, data) {
            console.log('[BROWSER] InstallPlugin:', guid);
        },

        RemovePlugin: function(guid) {
            console.log('[BROWSER] RemovePlugin:', guid);
        },

        PluginInstall: function(data) {
            console.log('[BROWSER] PluginInstall:', data);
            return null;
        },

        PluginUninstall: function(guid) {
            console.log('[BROWSER] PluginUninstall:', guid);
        },

        // File operations stubs
        LocalFileSaveChanges: function(changes, deleteIndex, count) {
            console.log('[BROWSER] LocalFileSaveChanges:', { deleteIndex, count, changesLength: changes ? changes.length : 0 });
            // Store changes in global state for tracking
            if (!window._documentChanges) {
                window._documentChanges = [];
            }
            if (changes && changes.length > 0) {
                window._documentChanges.push({
                    changes: changes,
                    deleteIndex: deleteIndex,
                    count: count,
                    timestamp: Date.now()
                });
                // Mark as having unsaved changes AND that user has edited
                window._hasUnsavedChanges = true;
                window._userHasEdited = true;
                window._documentModified = true;
                console.log('[BROWSER] LocalFileSaveChanges: User has made changes, _userHasEdited=true');
            }
        },

        LocalFileGetSaved: function() {
            // If we've already detected infinite loop, always return undefined
            if (window._infiniteLoopGetSavedDetected) {
                return undefined;
            }

            // Track call frequency to detect infinite loops
            if (!window._localFileGetSavedCallTimes) {
                window._localFileGetSavedCallTimes = [];
            }

            const now = Date.now();
            window._localFileGetSavedCallTimes.push(now);

            // Keep only calls from the last second
            window._localFileGetSavedCallTimes = window._localFileGetSavedCallTimes.filter(
                time => now - time < 1000
            );

            // If being called more than INFINITE_LOOP_THRESHOLD times per second, we have an infinite loop
            if (window._localFileGetSavedCallTimes.length > INFINITE_LOOP_THRESHOLD && !window._infiniteLoopGetSavedWarningShown) {
                console.error('[BROWSER] WARNING: LocalFileGetSaved being called too frequently (' +
                    window._localFileGetSavedCallTimes.length + ' times/sec). This indicates an infinite loop in the SDK.');
                console.error('[BROWSER] This usually means the SDK is waiting for a save operation to complete.');
                console.error('[BROWSER] Returning undefined permanently to signal that local save tracking is not available.');
                window._infiniteLoopGetSavedWarningShown = true;
                window._infiniteLoopGetSavedDetected = true; // Mark as permanently detected

                // After detecting the loop, return undefined to signal feature not available
                return undefined;
            }

            // Initialize state if needed
            if (window._hasUnsavedChanges === undefined) {
                window._hasUnsavedChanges = false;
            }

            // Throttle logging to prevent console spam from infinite loops
            if (!window._localFileGetSavedLastLog || now - window._localFileGetSavedLastLog > 1000) {
                console.log('[BROWSER] LocalFileGetSaved - returning:', !window._hasUnsavedChanges);
                window._localFileGetSavedLastLog = now;
            }

            // Return true if file is saved (no unsaved changes), false otherwise
            // For initial state, return true to indicate the file starts in a saved state
            return !window._hasUnsavedChanges;
        },

        // Get current document binary data for saving
        GetFileData: function(options) {
            console.log('[BROWSER] GetFileData called with options:', options);

            // In desktop mode, there's no iframe - the editor is embedded directly
            var iframeCount = document.querySelectorAll('iframe').length;
            console.log('[BROWSER] GetFileData: Detected iframes:', iframeCount);

            var editorWindow = null;

            if (iframeCount === 0) {
                // Desktop mode - editor is in the current window
                console.log('[BROWSER] GetFileData: Using desktop mode (no iframe)');
                editorWindow = window;
            } else {
                // Iframe mode - find the editor iframe
                console.log('[BROWSER] GetFileData: Using iframe mode');
                var iframe = document.querySelector('iframe[id*="placeholder"]') ||
                            document.querySelector('iframe[id*="frameEditor"]') ||
                            document.querySelector('iframe');

                console.log('[BROWSER] GetFileData: Selected iframe:', iframe ? ('ID: ' + (iframe.id || '(no id)')) : 'null');

                if (!iframe || !iframe.contentWindow) {
                    console.error('[BROWSER] GetFileData: No editor iframe found');
                    return null;
                }

                editorWindow = iframe.contentWindow;
            }

            if (!editorWindow) {
                console.error('[BROWSER] GetFileData: No editor window found');
                return null;
            }

            var iframeWindow = editorWindow;

            // Find the editor instance
            var editor = iframeWindow.editor || (iframeWindow.Asc && iframeWindow.Asc.editor) || null;

            // Try to get the binary data from the editor
            if (editor && editor.asc_nativeGetFile) {
                console.log('[BROWSER] GetFileData: Using editor.asc_nativeGetFile()');
                try {
                    // Get binary in ONLYOFFICE format (format 8192)
                    var binaryData = editor.asc_nativeGetFile();
                    console.log('[BROWSER] GetFileData: Got binary data, size:', binaryData ? binaryData.byteLength : 'null');
                    return binaryData;
                } catch(e) {
                    console.error('[BROWSER] GetFileData: Error getting file data:', e);
                    return null;
                }
            } else {
                console.error('[BROWSER] GetFileData: editor or asc_nativeGetFile not available');
                return null;
            }
        },

        // Save file - called by keyboard shortcut or File > Save
        // Parameters from SDK: (options, format, param3, param4, jsonParams)
        LocalFileSave: function(options, format, param3, param4, jsonParams) {
            console.log('[BROWSER] ===== LocalFileSave CALLED =====');
            console.log('[BROWSER] Parameters received:');
            console.log('[BROWSER]   options:', options);
            console.log('[BROWSER]   format:', format);
            console.log('[BROWSER]   param3:', param3);
            console.log('[BROWSER]   param4:', param4);
            console.log('[BROWSER]   jsonParams:', jsonParams);
            console.log('[BROWSER] Current state:');
            console.log('[BROWSER]   window._documentModified:', window._documentModified);
            console.log('[BROWSER]   window._hasUnsavedChanges:', window._hasUnsavedChanges);
            console.log('[BROWSER]   window._documentLoaded:', window._documentLoaded);

            // FIX FOR BLOCKING DIALOG: If document not modified, skip save
            // This prevents the "Saving..." dialog from appearing during document load
            // when the SDK calls LocalFileSave but the user hasn't made any changes yet
            if (!window._documentModified) {
                console.log('[BROWSER] *** EARLY RETURN: Document not modified, skipping save ***');
                return true;  // Return success without actually saving
            }
            console.log('[BROWSER] Document has been modified, proceeding with save');

            // Get filepath - ONLY use absolute filepath
            var filepathFromWindow = window._ONLYOFFICE_FILEPATH;
            var filepathFromParent = null;
            try {
                if (window.parent && window.parent !== window) {
                    if (window.parent._ONLYOFFICE_FILEPATH) {
                        filepathFromParent = window.parent._ONLYOFFICE_FILEPATH;
                    }
                }
            } catch(e) {}

            var filepath = filepathFromWindow || filepathFromParent;

            if (!filepath) {
                console.error('[BROWSER] ERROR: No absolute filepath available for save!');
                console.error('[BROWSER] window._ONLYOFFICE_FILEPATH:', filepathFromWindow);
                console.error('[BROWSER] parent._ONLYOFFICE_FILEPATH:', filepathFromParent);
                return false;
            }

            var filename = filepath.split('/').pop();
            console.log('[BROWSER] Saving to:', filename);
            console.log('[BROWSER] Full filepath:', filepath);

            // Get editor - we're already running IN the iframe with the editor
            var editor = window.editor || (window.Asc && window.Asc.editor) || null;
            if (!editor) {
                console.error('[BROWSER] Editor not available');
                return false;
            }

            // Commit pending changes
            if (editor.asc_Save) {
                console.log('[BROWSER] Calling editor.asc_Save() to commit pending changes...');
                editor.asc_Save();
            }

            // Wait 2000ms for commit to fully process, then export
            setTimeout(function() {
                console.log('[SAVE] Exporting document binary...');
                console.log('[SAVE] Available export methods:', {
                    asc_nativeGetFile: typeof editor.asc_nativeGetFile,
                    asc_nativeGetFileData: typeof editor.asc_nativeGetFileData,
                    asc_nativeGetFile2: typeof editor.asc_nativeGetFile2
                });

                // Get document as ONLYOFFICE binary format
                // Try asc_nativeGetFileData first (user says this works), then fall back to asc_nativeGetFile
                var binaryData = null;
                if (editor.asc_nativeGetFileData) {
                    console.log('[SAVE] Using editor.asc_nativeGetFileData()');
                    binaryData = editor.asc_nativeGetFileData();
                } else if (editor.asc_nativeGetFile) {
                    console.log('[SAVE] Using editor.asc_nativeGetFile()');
                    binaryData = editor.asc_nativeGetFile();
                } else {
                    console.error('[SAVE] No binary export method available!');
                    console.log('[SAVE] All editor methods:', Object.keys(editor).filter(function(k) {
                        return k.indexOf('native') !== -1 || k.indexOf('File') !== -1 || k.indexOf('Data') !== -1;
                    }).join(', '));
                    return;
                }

                console.log('[SAVE] Got binary data, size:', binaryData ? binaryData.byteLength || binaryData.length : 0);
                if (binaryData && binaryData.byteLength) {
                    var firstBytes = new Uint8Array(binaryData.slice(0, 20));
                    var firstBytesStr = String.fromCharCode.apply(null, firstBytes);
                    console.log('[SAVE] First 20 bytes:', firstBytesStr);
                }

                if (!binaryData) {
                    console.error('[SAVE] Failed to get binary data!');
                    return;
                }

                // POST to server - ONLY use absolute filepath in query parameter
                // Also pass file hash so server can locate media files
                var fileHash = window._ONLYOFFICE_FILE_HASH;
                if (!fileHash && window.parent && window.parent !== window) {
                    fileHash = window.parent._ONLYOFFICE_FILE_HASH;
                }
                var saveUrl = SERVER_BASE_URL + '/api/save?filepath=' + encodeURIComponent(filepath);
                if (fileHash) {
                    saveUrl += '&filehash=' + encodeURIComponent(fileHash);
                }
                console.log('[SAVE] Saving to URL:', saveUrl);
                console.log('[SAVE] File hash:', fileHash);

                fetch(saveUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/octet-stream' },
                    body: binaryData
                })
                .then(function(response) {
                    console.log('[SAVE] Server response:', response.status);
                    return response.json();
                })
                .then(function(data) {
                    console.log('[SAVE] Save completed successfully!', data);

                    // Tell SDK save is done
                    if (window._originalDesktopOfflineAppDocumentEndSave) {
                        window._originalDesktopOfflineAppDocumentEndSave(true);
                    }

                    // Mark document as saved
                    window._hasUnsavedChanges = false;
                })
                .catch(function(error) {
                    console.error('[SAVE] Save failed:', error);
                });
            }, 2000);

            return true;
        },

        LocalFileGetImageUrl: function(path) {
            console.log('[BROWSER] LocalFileGetImageUrl:', path);

            // Handle blob URLs from drag-drop - return uploaded filename for persistence
            if (path && path.startsWith('blob:')) {
                var uploadedFilename = window._uploadedFileMap && window._uploadedFileMap[path];
                if (uploadedFilename) {
                    console.log('[BROWSER] Returning uploaded filename for blob:', uploadedFilename);
                    return uploadedFilename;
                }
                console.warn('[BROWSER] No upload mapping for blob URL, image will not persist');
                return path;
            }

            // Extract filename from path using util
            var filename = extractMediaFilename(path);

            // Get the file hash - check both current window and parent window
            var fileHash = window._ONLYOFFICE_FILE_HASH;
            if (!fileHash && window.parent && window.parent !== window) {
                fileHash = window.parent._ONLYOFFICE_FILE_HASH;
            }

            if (!fileHash) {
                console.error('[BROWSER] No file hash available for image loading!');
                return SERVER_BASE_URL + '/api/media/UNKNOWN/' + encodeURIComponent(filename);
            }

            var imageUrl = buildMediaUrl(SERVER_BASE_URL, fileHash, filename);
            console.log('[BROWSER] Mapped image path to URL:', imageUrl);
            return imageUrl;
        },

        LocalFileGetImageUrlCorrect: function(path) {
            console.log('[BROWSER] LocalFileGetImageUrlCorrect:', path);

            // Extract blob URL if wrapped by SDK's getImageUrl (legacy support)
            var blobUrl = extractBlobUrl(path);
            if (blobUrl) {
                var uploadedFilename = window._uploadedFileMap && window._uploadedFileMap[blobUrl];
                if (uploadedFilename) {
                    var fileHash = window._ONLYOFFICE_FILE_HASH || (window.parent && window.parent._ONLYOFFICE_FILE_HASH);
                    var serverUrl = buildMediaUrl(SERVER_BASE_URL, fileHash, uploadedFilename);
                    if (serverUrl) {
                        console.log('[BROWSER] Corrected blob to server URL:', serverUrl);
                        return serverUrl;
                    }
                }
                return blobUrl;
            }

            // Extract filename from wrapped path and convert to server URL
            if (path && path.indexOf('/media/') !== -1) {
                var filename = extractMediaFilename(path);
                filename = filename.split('?')[0].split('#')[0];
                if (filename && !filename.startsWith('http')) {
                    var fileHash = window._ONLYOFFICE_FILE_HASH || (window.parent && window.parent._ONLYOFFICE_FILE_HASH);
                    if (fileHash) {
                        var serverUrl = buildMediaUrl(SERVER_BASE_URL, fileHash, filename);
                        console.log('[BROWSER] Corrected media path to server URL:', serverUrl);
                        return serverUrl;
                    }
                }
            }

            return this.LocalFileGetImageUrl(path);
        },

        GetImageBase64: function(path) {
            console.log('[BROWSER] GetImageBase64:', path);

            // Check if we have this image in cache
            if (window._imageCache[path]) {
                console.log('[BROWSER] Returning cached base64 for:', path);
                return window._imageCache[path];
            }

            // If not cached, we need to return the HTTP URL
            // The image should have been pre-loaded, but if not, return the URL
            console.log('[BROWSER] Image not in cache, returning URL:', path);
            return this.LocalFileGetImageUrl(path);
        },

        IsImageFile: function(path) {
            var filename = path;
            // Check both maps for file info
            if (window._blobFileMap && window._blobFileMap[path]) {
                filename = window._blobFileMap[path].name;
            } else if (window._uploadedFileMap && window._uploadedFileMap[path]) {
                filename = window._uploadedFileMap[path];
            }
            console.log('[BROWSER] IsImageFile:', filename);
            return utils.isImageFile ? utils.isImageFile(filename) : false;
        },

        GetDropFiles: function() {
            // Return blob URLs that were created during drop event
            // The _uploadedFileMap already contains blob URL -> server filename mappings
            var blobUrls = window._uploadedFileMap ? Object.keys(window._uploadedFileMap) : [];
            console.log('[BROWSER] GetDropFiles, count:', blobUrls.length);

            // Also handle any non-uploaded files
            if (window._droppedFiles && window._droppedFiles.length > 0) {
                for (var i = 0; i < window._droppedFiles.length; i++) {
                    var file = window._droppedFiles[i];
                    // Check if this file was already uploaded (blob URL already created)
                    var alreadyHasBlobUrl = false;
                    for (var blobUrl in window._uploadedFileMap) {
                        if (window._uploadedFileMap.hasOwnProperty(blobUrl)) {
                            alreadyHasBlobUrl = true;
                            break;
                        }
                    }
                    if (!alreadyHasBlobUrl) {
                        var url = URL.createObjectURL(file);
                        window._blobFileMap = window._blobFileMap || {};
                        window._blobFileMap[url] = file;
                        blobUrls.push(url);
                    }
                }
                window._droppedFiles = [];
            }
            return blobUrls;
        },

        // Protection and signatures
        IsProtectionSupport: function() {
            console.log('[BROWSER] IsProtectionSupport');
            return false;
        },

        IsSignaturesSupport: function() {
            console.log('[BROWSER] IsSignaturesSupport');
            return false;
        },

        Sign: function() {
            console.log('[BROWSER] Sign - not supported');
        },

        RemoveSignature: function() {
            console.log('[BROWSER] RemoveSignature - not supported');
        },

        RemoveAllSignatures: function() {
            console.log('[BROWSER] RemoveAllSignatures - not supported');
        },

        ViewCertificate: function() {
            console.log('[BROWSER] ViewCertificate - not supported');
        },

        SelectCertificate: function() {
            console.log('[BROWSER] SelectCertificate - not supported');
        },

        GetDefaultCertificate: function() {
            console.log('[BROWSER] GetDefaultCertificate - not supported');
            return null;
        },

        // Media support
        IsSupportMedia: function() {
            console.log('[BROWSER] IsSupportMedia');
            return false;
        },

        AddVideo: function() {
            console.log('[BROWSER] AddVideo - not supported');
        },

        AddAudio: function() {
            console.log('[BROWSER] AddAudio - not supported');
        },

        // Advanced features
        isSupportPlugins: function() {
            console.log('[BROWSER] isSupportPlugins');
            return false;
        },

        isSupportMacroses: function() {
            console.log('[BROWSER] isSupportMacroses');
            return false;
        },

        isSupportNetworkFunctionality: function() {
            console.log('[BROWSER] isSupportNetworkFunctionality');
            return false;
        },

        SetFullscreen: function(enabled) {
            console.log('[BROWSER] SetFullscreen:', enabled);
        },

        CheckNeedWheel: function() {
            console.log('[BROWSER] CheckNeedWheel');
            return true;
        },

        SetAdvancedOptions: function(options) {
            console.log('[BROWSER] SetAdvancedOptions:', options);
        },

        // Spell checking
        SpellCheck: function(word, langId) {
            console.log('[BROWSER] SpellCheck:', word, langId);
            return [];
        },

        getDictionariesPath: function() {
            console.log('[BROWSER] getDictionariesPath');
            return '';
        },

        // Font loading
        LoadFontBase64: function(fontName) {
            console.log('[BROWSER] LoadFontBase64:', fontName);
        },

        GetFontThumbnailHeight: function() {
            console.log('[BROWSER] GetFontThumbnailHeight');
            return 0;
        },

        // Image operations
        GetImageFormat: function(path) {
            console.log('[BROWSER] GetImageFormat:', path);
            return '';
        },

        GetImageOriginalSize: function(path) {
            console.log('[BROWSER] GetImageOriginalSize:', path);
            return { width: 0, height: 0 };
        },

        // File existence check
        IsLocalFileExist: function(path) {
            console.log('[BROWSER] IsLocalFileExist:', path);
            return false;
        },

        // Document operations
        onDocumentContentReady: function() {
            console.log('[BROWSER] onDocumentContentReady - document loaded successfully');
            window._documentLoaded = true;
            window._hasUnsavedChanges = false;

            // NOTE(victor): Override Application display to show "officeExtension".
            // The value comes from docProps/app.xml in the OOXML file, read via
            // CApp.fromStream(). We patch the instance method so the info panel
            // (File > Info > Application) reflects our branding without modifying
            // the file's actual metadata.
            try {
                var _editor = Asc.editor || window.editor;
                if (_editor && _editor.asc_getAppProps) {
                    var _app = _editor.asc_getAppProps();
                    if (_app) {
                        _app.asc_getApplication = function() { return 'officeExtension'; };
                        _app.asc_getAppVersion = function() { return APP_VERSION; };
                    }
                }
            } catch (e) {
                console.log('[BROWSER] Could not override app branding:', e);
            }

            // Notify parent window that document is ready for interaction
            try {
                if (window && window.parent && window.parent.postMessage) {
                    window.parent.postMessage({
                        type: 'ONLYOFFICE_DOCUMENT_READY',
                        filePath: window._ONLYOFFICE_FILEPATH
                    }, '*');
                }
            } catch (e) {
                console.log('[BROWSER] Could not send ONLYOFFICE_DOCUMENT_READY message:', e);
            }
        },

        onDocumentModifiedChanged: function(modified) {
            // Throttle this callback to prevent infinite loops
            const now = Date.now();
            if (window._lastModifiedChangedTime && now - window._lastModifiedChangedTime < 100) {
                // Being called too frequently, skip this call
                return;
            }
            window._lastModifiedChangedTime = now;

            console.log('[BROWSER] onDocumentModifiedChanged:', modified, 'userHasEdited:', window._userHasEdited, 'documentChanges:', window._documentChanges ? window._documentChanges.length : 0);

            // Track when user actually modifies the document
            if (modified === true) {
                window._documentModified = true;
                window._hasUnsavedChanges = true;
                window._userHasEdited = true;
                console.log('[BROWSER] _documentModified set to: true (user made changes)');
            } else if (modified === false) {
                // Check if we have unsaved changes in the document
                var hasChanges = window._documentChanges && window._documentChanges.length > 0;

                // Only accept false if user hasn't edited AND there are no stored changes
                if (!window._userHasEdited && !hasChanges) {
                    window._documentModified = false;
                    window._hasUnsavedChanges = false;
                    console.log('[BROWSER] _documentModified set to: false (initialization, no changes)');
                } else {
                    console.log('[BROWSER] Ignoring false - user has edited or has pending changes, keeping flags true');
                    // Keep the modified flag true
                    window._documentModified = true;
                    window._hasUnsavedChanges = true;
                }
            }
        },

        OnSave: function() {
            console.log('[BROWSER] OnSave - save requested by SDK');

            // CRITICAL FIX: Check if document is modified BEFORE doing anything
            // This prevents the "Saving..." dialog from appearing during document initialization
            if (!window._documentModified) {
                console.log('[BROWSER] OnSave: Document not modified, returning false to prevent save dialog');
                return false;  // Return false to tell SDK not to show save UI
            }

            // Prevent infinite save loops - track last save time
            const now = Date.now();
            if (window._lastOnSaveTime && now - window._lastOnSaveTime < 2000) {
                console.log('[BROWSER] OnSave: Ignoring save request (too frequent, last save was', now - window._lastOnSaveTime, 'ms ago)');
                return false;
            }
            window._lastOnSaveTime = now;

            // OnSave is called by SDK for auto-save
            // We need to actually save the file here
            console.log('[BROWSER] OnSave: Triggering actual file save via LocalFileSave');

            // Call LocalFileSave to do the actual save
            if (window.AscDesktopEditor && window.AscDesktopEditor.LocalFileSave) {
                return window.AscDesktopEditor.LocalFileSave();
            } else {
                console.error('[BROWSER] OnSave: LocalFileSave not available!');
                return false;
            }
        },

        SaveQuestion: function() {
            console.log('[BROWSER] SaveQuestion');
            // Return false to indicate no save dialog needed
            return false;
        },

        // Crypto (all disabled)
        isBlockchainSupport: function() {
            console.log('[BROWSER] isBlockchainSupport');
            return false;
        },

        Crypto_GetLocalImageBase64: function(path) {
            console.log('[BROWSER] Crypto_GetLocalImageBase64:', path);
            return '';
        },

        GetEncryptedHeader: function() {
            console.log('[BROWSER] GetEncryptedHeader');
            return '';
        },

        CryptoAES_Init: function(password) {
            console.log('[BROWSER] CryptoAES_Init');
        },

        CryptoAES_Encrypt: function(data) {
            console.log('[BROWSER] CryptoAES_Encrypt');
            return data;
        },

        CryptoAES_Decrypt: function(data) {
            console.log('[BROWSER] CryptoAES_Decrypt');
            return data;
        },

        // Engine version
        getEngineVersion: function() {
            console.log('[BROWSER] getEngineVersion');
            return 120; // Chrome 120+
        },

        // Reporter (logging)
        startReporter: function() {
            console.log('[BROWSER] startReporter');
        },

        endReporter: function() {
            console.log('[BROWSER] endReporter');
        },

        sendToReporter: function(msg) {
            console.log('[BROWSER] sendToReporter:', msg);
        },

        sendFromReporter: function(msg) {
            console.log('[BROWSER] sendFromReporter:', msg);
        },

        // External operations
        openExternalReference: function(url) {
            console.log('[BROWSER] openExternalReference:', url);
        },

        // Printing
        IsFilePrinting: function() {
            console.log('[BROWSER] IsFilePrinting');
            return false;
        },

        Print_Start: function() {
            console.log('[BROWSER] Print_Start');
        },

        Print_Page: function() {
            console.log('[BROWSER] Print_Page');
        },

        Print_End: function() {
            console.log('[BROWSER] Print_End');
        },

        // File conversion
        convertFile: function(params) {
            console.log('[BROWSER] convertFile:', params);
        },

        startExternalConvertation: function(type, params) {
            console.log('[BROWSER] startExternalConvertation:', type, params);
        },

        // Get backup plugins
        GetBackupPlugins: function() {
            console.log('[BROWSER] GetBackupPlugins');
            return '[]';
        },

        // Relative path
        LocalFileGetRelativePath: function(path) {
            console.log('[BROWSER] LocalFileGetRelativePath:', path);
            return path;
        },

        // Local file operations count
        LocalFileGetOpenChangesCount: function() {
            // If we've already detected infinite loop, always return undefined
            if (window._infiniteLoopChangesCountDetected) {
                return undefined;
            }

            // Track call frequency to detect infinite loops
            if (!window._localFileGetOpenChangesCountCallTimes) {
                window._localFileGetOpenChangesCountCallTimes = [];
            }

            const now = Date.now();
            window._localFileGetOpenChangesCountCallTimes.push(now);

            // Keep only calls from the last second
            window._localFileGetOpenChangesCountCallTimes = window._localFileGetOpenChangesCountCallTimes.filter(
                time => now - time < 1000
            );

            // If being called more than INFINITE_LOOP_THRESHOLD times per second, we have an infinite loop
            if (window._localFileGetOpenChangesCountCallTimes.length > INFINITE_LOOP_THRESHOLD && !window._infiniteLoopWarningShown) {
                console.error('[BROWSER] WARNING: LocalFileGetOpenChangesCount being called too frequently (' +
                    window._localFileGetOpenChangesCountCallTimes.length + ' times/sec). This indicates an infinite loop in the SDK.');
                console.error('[BROWSER] Returning undefined permanently to signal that this feature is not available.');
                window._infiniteLoopWarningShown = true;
                window._infiniteLoopChangesCountDetected = true; // Mark as permanently detected

                // After detecting the loop, return undefined to signal feature not available
                return undefined;
            }

            // Throttle logging to prevent console spam
            if (!window._localFileGetOpenChangesCountLastLog || now - window._localFileGetOpenChangesCountLastLog > 1000) {
                console.log('[BROWSER] LocalFileGetOpenChangesCount - returning 0 (no pending changes)');
                window._localFileGetOpenChangesCountLastLog = now;
            }

            // Always return 0 - no pending changes
            // In a real desktop app, this would return the number of change operations waiting to be saved
            return 0;
        },

        // Call in all windows
        CallInAllWindows: function(method, args) {
            console.log('[BROWSER] CallInAllWindows:', method, args);
        },

        // Load JS
        LoadJS: function(path) {
            console.log('[BROWSER] LoadJS:', path);
        },

        // Local save restrictions
        SetLocalRestrictions: function(value) {
            console.log('[BROWSER] SetLocalRestrictions:', value);
        },

        // Native viewer
        NativeViewerOpen: function(password) {
            console.log('[BROWSER] NativeViewerOpen:', password);
        },

        // File remove
        RemoveFile: function(path) {
            console.log('[BROWSER] RemoveFile:', path);
        },

        // Resave file
        ResaveFile: function() {
            console.log('[BROWSER] ResaveFile');
        },

        // Download files
        DownloadFiles: function(files) {
            console.log('[BROWSER] DownloadFiles:', files);
        },

        // File locked
        onFileLockedClose: function() {
            console.log('[BROWSER] onFileLockedClose');
        },

        // Cloud crypto
        CryptoCloud_GetUserInfo: function() {
            console.log('[BROWSER] CryptoCloud_GetUserInfo');
            return null;
        },

        cloudCryptoCommandMainFrame: function(cmd, callback) {
            console.log('[BROWSER] cloudCryptoCommandMainFrame:', cmd);
            if (callback) callback({});
        },

        // RSA Crypto
        CryproRSA_EncryptPublic: function(key, data) {
            console.log('[BROWSER] CryproRSA_EncryptPublic');
            return data;
        },

        CryproRSA_DecryptPrivate: function(key, data) {
            console.log('[BROWSER] CryproRSA_DecryptPrivate');
            return data;
        },

        // Preload crypto image
        PreloadCryptoImage: function(url, path) {
            console.log('[BROWSER] PreloadCryptoImage:', url, path);
        },

        // Crypto download
        CryptoDownloadAs: function() {
            console.log('[BROWSER] CryptoDownloadAs');
        },

        // Build crypted
        buildCryptedStart: function() {
            console.log('[BROWSER] buildCryptedStart');
        },

        buildCryptedEnd: function() {
            console.log('[BROWSER] buildCryptedEnd');
        },

        // Open file crypt
        OpenFileCrypt: function() {
            console.log('[BROWSER] OpenFileCrypt');
        },

        // Comparison
        CompareDocumentFile: function(file) {
            console.log('[BROWSER] CompareDocumentFile:', file);
        },

        CompareDocumentUrl: function(url) {
            console.log('[BROWSER] CompareDocumentUrl:', url);
        },

        // Merge
        MergeDocumentFile: function(file) {
            console.log('[BROWSER] MergeDocumentFile:', file);
        },

        MergeDocumentUrl: function(url) {
            console.log('[BROWSER] MergeDocumentUrl:', url);
        },

        // Workbook
        OpenWorkbook: function() {
            console.log('[BROWSER] OpenWorkbook');
        },

        // Media player
        CallMediaPlayerCommand: function(cmd) {
            console.log('[BROWSER] CallMediaPlayerCommand:', cmd);
        },

        // PDF print cloud
        SetPdfCloudPrintFileInfo: function(info) {
            console.log('[BROWSER] SetPdfCloudPrintFileInfo:', info);
        },

        IsCachedPdfCloudPrintFileInfo: function() {
            console.log('[BROWSER] IsCachedPdfCloudPrintFileInfo');
            return false;
        },

        emulateCloudPrinting: function() {
            console.log('[BROWSER] emulateCloudPrinting');
        },

        // Local save to drawing format
        localSaveToDrawingFormat: function() {
            console.log('[BROWSER] localSaveToDrawingFormat');
        },

        localSaveToDrawingFormat2: function() {
            console.log('[BROWSER] localSaveToDrawingFormat2');
        },

        // Load local file
        loadLocalFile: function() {
            console.log('[BROWSER] loadLocalFile');
        },

        // Send by mail (plugin)
        SendByMail: function() {
            console.log('[BROWSER] SendByMail');
        },

        // Get supported scale values
        GetSupportedScaleValues: function() {
            console.log('[BROWSER] GetSupportedScaleValues');
            return [1, 1.25, 1.5, 1.75, 2];
        }
    };

    // Create window.desktop alias that editors expect
    window.desktop = window.AscDesktopEditor;

    // CRITICAL FIX: Install window.native functions AFTER SDK loads
    // The SDK needs to create window.native itself during initialization
    // We'll add our Save_End function later, after SDK is fully loaded

    console.log('[BROWSER] Deferring window.native setup - will install after SDK loads');

    // Function to install our native callbacks
    window._installNativeCallbacks = function() {
        console.log('[BROWSER] Installing window.native callbacks...');

        // NOTE(victor): Delay setting window.native until font engine loads.
        // The SDK's AscFonts.load() checks if window.native exists - if so, it skips
        // loading the WASM font engine (assumes native mode). On Windows, this function
        // can run before AscFonts.load(), causing font loading to be skipped entirely.
        function waitForFontEngineAndInstall(attempts) {
            attempts = attempts || 0;

            // Check if font engine has started loading (AscFonts.c7i becomes true when loaded)
            // OR if AscFonts.load has been called (AscFonts.oe gets set)
            var fontEngineStarted = window.AscFonts && (window.AscFonts.c7i || window.AscFonts.oe);

            if (fontEngineStarted || attempts >= 50) {
                if (attempts >= 50) {
                    console.log('[BROWSER] Font engine wait timeout reached, proceeding anyway');
                } else {
                    console.log('[BROWSER] Font engine has started loading, safe to set window.native');
                }
                doInstallNativeCallbacks();
            } else {
                setTimeout(function() { waitForFontEngineAndInstall(attempts + 1); }, 100);
            }
        }

        function doInstallNativeCallbacks() {
            console.log('[BROWSER] Creating window.native object');
            if (!window.native) {
                window.native = {};
            }

            // Add our callbacks to the existing object
            if (!window.native.Save_End) {
                window.native.Save_End = function(format, size) {
                    console.log('[BROWSER] window.native.Save_End called');
                    console.log('[BROWSER] Format:', format);
                    console.log('[BROWSER] Size:', size);
                };
                console.log('[BROWSER] Installed window.native.Save_End');
            }

            if (!window.native.GetOriginalImageSize) {
                window.native.GetOriginalImageSize = function(path) {
                    console.log('[BROWSER] window.native.GetOriginalImageSize:', path);
                    return [0, 0];
                };
            }

            if (!window.native.DD_GetOriginalImageSize) {
                window.native.DD_GetOriginalImageSize = function(path) {
                    console.log('[BROWSER] window.native.DD_GetOriginalImageSize:', path);
                    return [0, 0];
                };
            }

            if (!window.native.setUrlsCount) {
                window.native.setUrlsCount = function(count) {
                    console.log('[BROWSER] window.native.setUrlsCount:', count);
                };
            }

            if (!window.native.openFileCommand) {
                window.native.openFileCommand = function(path, binary, format) {
                    console.log('[BROWSER] window.native.openFileCommand:', path, binary, format);
                    return null;
                };
            }

            if (!window.native.AddImageInChanges) {
                window.native.AddImageInChanges = function(image) {
                    console.log('[BROWSER] window.native.AddImageInChanges:', image);
                };
            }

            if (!window.native.CheckNextChange) {
                window.native.CheckNextChange = function() {
                    return true;
                };
            }

            if (!window.native.BeginDrawStyle) {
                window.native.BeginDrawStyle = function(width, height) {
                    console.log('[BROWSER] window.native.BeginDrawStyle:', width, height);
                };
            }

            if (!window.native.EndDrawStyle) {
                window.native.EndDrawStyle = function() {
                    console.log('[BROWSER] window.native.EndDrawStyle');
                };
            }

            console.log('[BROWSER] window.native callbacks installed successfully');
        }

        waitForFontEngineAndInstall();
    };

    // Create RendererProcessVariable stub for theme and RTL support
    window.RendererProcessVariable = {
        theme: {
            id: 'theme-classic-light',
            type: 'light',
            system: 'light'
        },
        localthemes: {},
        rtl: false
    };

    // Create native message handling infrastructure
    window.native_message_cmd = [];
    window.on_native_message = function(cmd, param) {
        console.log('on_native_message:', cmd, param);

        if (/window:features/.test(cmd)) {
            try {
                var obj = JSON.parse(param);
                if (obj.singlewindow !== undefined) {
                    window.desktop.features.singlewindow = obj.singlewindow;
                }
            } catch(e) {
                console.error('Failed to parse window features:', e);
            }
        } else {
            window.native_message_cmd[cmd] = param;
        }
    };

    console.log('AscDesktopEditor stub initialized successfully');

    // Function to wrap DesktopOfflineAppDocumentEndSave
    // This wraps the function to intercept save completion
    // We use Object.defineProperty to prevent SDK from overwriting our wrapper
    window._installSaveCompletionCallback = function() {
        console.log('[BROWSER] Installing save completion callback wrapper...');

        // Check if already installed using defineProperty
        if (window._saveCompletionCallbackInstalled) {
            console.log('[BROWSER] Save completion callback already installed via defineProperty');
            return true;
        }

        // If the function already exists, store it
        var existingFunction = window.DesktopOfflineAppDocumentEndSave;
        if (typeof existingFunction === 'function') {
            console.log('[BROWSER] Found existing DesktopOfflineAppDocumentEndSave, wrapping it');
            window._originalDesktopOfflineAppDocumentEndSave = existingFunction;
        } else {
            console.log('[BROWSER] DesktopOfflineAppDocumentEndSave not yet defined, will intercept when defined');
        }

        window._saveCompletionCallbackInstalled = true;
        console.log('[BROWSER] Installing property interceptor for DesktopOfflineAppDocumentEndSave...');

        // Create the wrapper function that will handle saves
        var wrapperFunction = function(errorCode, param2, param3) {
        console.log('[BROWSER] ===== DesktopOfflineAppDocumentEndSave WRAPPER CALLED =====');
        console.log('[BROWSER] Error code:', errorCode, '(0=success, 2=error)');
        console.log('[BROWSER] Param2:', param2);
        console.log('[BROWSER] Param3:', param3);
        console.log('[BROWSER] Has pending save:', window._hasPendingSave);
        console.log('[BROWSER] Pending filename:', window._pendingSaveFilename);
        console.log('[BROWSER] Has editor ref:', !!window._editor);
        console.log('[BROWSER] Has editor window ref:', !!window._editorWindow);

        // If save completed successfully (errorCode=0) and we have a pending save
        if (errorCode === 0 && window._hasPendingSave) {
            console.log('[BROWSER] Save completed successfully! Now exporting XLSX...');

            var editor = window._editor;
            var filename = window._pendingSaveFilename;

            if (!editor) {
                console.error('[BROWSER] No editor reference available!');
                // Call original
                if (window._originalDesktopOfflineAppDocumentEndSave) {
                    return window._originalDesktopOfflineAppDocumentEndSave(errorCode, param2, param3);
                }
                return;
            }

            if (!filename) {
                console.error('[BROWSER] No filename available!');
                // Call original
                if (window._originalDesktopOfflineAppDocumentEndSave) {
                    return window._originalDesktopOfflineAppDocumentEndSave(errorCode, param2, param3);
                }
                return;
            }

            // Check if asc_nativeGetFileData is available
            if (typeof editor.asc_nativeGetFileData === 'function') {
                console.log('[BROWSER] Calling asc_nativeGetFileData(257) to export XLSX...');

                try {
                    // Format 257 = XLSX (Asc.c_oAscFileType.XLSX)
                    // Note: 65 = DOCX, which returns ONLYOFFICE binary format
                    var xlsxData = editor.asc_nativeGetFileData(257);

                    if (!xlsxData) {
                        console.error('[BROWSER] asc_nativeGetFileData returned null/undefined!');
                        // Call original
                        if (window._originalDesktopOfflineAppDocumentEndSave) {
                            return window._originalDesktopOfflineAppDocumentEndSave(errorCode, param2, param3);
                        }
                        return;
                    }

                    console.log('[BROWSER] Got XLSX data, type:', typeof xlsxData);
                    console.log('[BROWSER] XLSX data size:', xlsxData.byteLength || xlsxData.length || 'unknown');

                    // Convert to ArrayBuffer if needed
                    var xlsxArrayBuffer = null;
                    if (xlsxData instanceof ArrayBuffer) {
                        xlsxArrayBuffer = xlsxData;
                    } else if (xlsxData instanceof Uint8Array) {
                        xlsxArrayBuffer = xlsxData.buffer;
                    } else if (typeof xlsxData === 'object' && xlsxData.buffer instanceof ArrayBuffer) {
                        xlsxArrayBuffer = xlsxData.buffer;
                    } else {
                        console.error('[BROWSER] Unexpected data type from asc_nativeGetFileData:', typeof xlsxData);
                        // Call original
                        if (window._originalDesktopOfflineAppDocumentEndSave) {
                            return window._originalDesktopOfflineAppDocumentEndSave(errorCode, param2, param3);
                        }
                        return;
                    }

                    console.log('[BROWSER] Converted to ArrayBuffer, size:', xlsxArrayBuffer.byteLength, 'bytes');
                    console.log('[BROWSER] POSTing XLSX to server...');

                    // Get filepath from global state - MUST be absolute
                    var filepath = window._pendingSaveFilepath;

                    if (!filepath) {
                        console.error('[BROWSER] ERROR: No absolute filepath available for save!');
                        // Call original with error
                        if (window._originalDesktopOfflineAppDocumentEndSave) {
                            return window._originalDesktopOfflineAppDocumentEndSave(2, param2, param3);
                        }
                        return;
                    }

                    // Build save URL - ONLY use absolute filepath
                    var saveUrl = SERVER_BASE_URL + '/api/save?filepath=' + encodeURIComponent(filepath);
                    console.log('[BROWSER] Saving to absolute path:', filepath);
                    console.log('[BROWSER] Save URL:', saveUrl);

                    // POST the XLSX to the server
                    fetch(saveUrl, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/octet-stream'
                        },
                        body: xlsxArrayBuffer
                    })
                    .then(function(response) {
                        if (!response.ok) {
                            throw new Error('Save failed: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(result) {
                        console.log('[BROWSER] ✅ Save successful:', result);
                        console.log('[BROWSER] File saved to:', result.path);

                        // Clear pending save state
                        window._hasPendingSave = false;
                        window._pendingSaveFilename = null;
                        window._editor = null;
                        window._editorWindow = null;

                        // Clear unsaved changes state
                        window._hasUnsavedChanges = false;
                        window._documentChanges = [];
                        window._userHasEdited = false;  // Reset flag after successful save
                        window._documentModified = false;  // Reset modified flag

                        // Call original callback
                        if (window._originalDesktopOfflineAppDocumentEndSave) {
                            return window._originalDesktopOfflineAppDocumentEndSave(errorCode, param2, param3);
                        }
                    })
                    .catch(function(error) {
                        console.error('[BROWSER] ❌ Save failed:', error);

                        // Clear pending save state
                        window._hasPendingSave = false;

                        // Call original callback with error code
                        if (window._originalDesktopOfflineAppDocumentEndSave) {
                            return window._originalDesktopOfflineAppDocumentEndSave(2, param2, param3); // Error code
                        }
                    });

                } catch (e) {
                    console.error('[BROWSER] Error calling asc_nativeGetFileData:', e);

                    // Clear pending save state
                    window._hasPendingSave = false;

                    // Call original callback with error code
                    if (window._originalDesktopOfflineAppDocumentEndSave) {
                        return window._originalDesktopOfflineAppDocumentEndSave(2, param2, param3); // Error code
                    }
                }
            } else {
                console.error('[BROWSER] asc_nativeGetFileData not available on editor!');
                console.log('[BROWSER] Available methods:', Object.keys(editor).filter(function(k) {
                    return k.includes('native') || k.includes('get') || k.includes('File');
                }).join(', '));

                // Clear pending save state
                window._hasPendingSave = false;

                // Call original
                if (window._originalDesktopOfflineAppDocumentEndSave) {
                    return window._originalDesktopOfflineAppDocumentEndSave(errorCode, param2, param3);
                }
            }
        } else {
            // Not a successful save, or no pending save - just pass through
            console.log('[BROWSER] Passing through to original DesktopOfflineAppDocumentEndSave');
            if (window._originalDesktopOfflineAppDocumentEndSave) {
                return window._originalDesktopOfflineAppDocumentEndSave(errorCode, param2, param3);
            }
        }
        };

        // Use Object.defineProperty to create a property that:
        // 1. When SDK writes to it, we store their function as _originalDesktopOfflineAppDocumentEndSave
        // 2. When SDK reads/calls it, they get our wrapper that calls their original
        var internalFunction = existingFunction || null;

        Object.defineProperty(window, 'DesktopOfflineAppDocumentEndSave', {
            get: function() {
                return wrapperFunction;
            },
            set: function(newFunction) {
                console.log('[BROWSER] SDK is trying to set DesktopOfflineAppDocumentEndSave - intercepting!');
                if (typeof newFunction === 'function') {
                    console.log('[BROWSER] Storing SDK function as _originalDesktopOfflineAppDocumentEndSave');
                    window._originalDesktopOfflineAppDocumentEndSave = newFunction;
                    internalFunction = newFunction;
                } else {
                    console.log('[BROWSER] Warning: Non-function assigned to DesktopOfflineAppDocumentEndSave');
                }
            },
            configurable: true // Allow reconfiguration if needed
        });

        console.log('[BROWSER] ✅ Save completion callback installed successfully via defineProperty!');
        return true;
    };

    // Install the callback immediately using defineProperty
    // This will intercept when SDK tries to define the function
    console.log('[BROWSER] Installing save completion callback via defineProperty...');
    window._installSaveCompletionCallback();

    // Intercept script loading to patch UpdateSystemPlugins AFTER SDK loads
    // AND install window.native callbacks after SDK initializes
    // AND intercept blob downloads created by asc_DownloadAs
    const originalCreateElement = document.createElement;
    document.createElement = function(tagName) {
        const element = originalCreateElement.call(document, tagName);

        // Handle script loading
        if (tagName.toLowerCase() === 'script') {
            element.addEventListener('load', function() {
                // After any script loads, try to override UpdateSystemPlugins
                if (window.UpdateSystemPlugins && !window._UpdateSystemPluginsPatched) {
                    console.log('[BROWSER] Patching UpdateSystemPlugins after SDK load');
                    window._UpdateSystemPluginsPatched = true;
                    window.UpdateSystemPlugins = function() {
                        console.log('[BROWSER] UpdateSystemPlugins called - skipping');
                    };
                }

                // Install window.native callbacks after SDK loads
                // Wait a bit to ensure SDK has finished its initialization
                if (!window._nativeCallbacksInstalled && window._installNativeCallbacks) {
                    setTimeout(function() {
                        window._installNativeCallbacks();
                        window._nativeCallbacksInstalled = true;
                    }, 100);
                }

                // No need to reinstall callback - defineProperty handles it automatically
            });
        }

        // Handle download link creation (for asc_DownloadAs interception)
        if (tagName.toLowerCase() === 'a' && element instanceof HTMLAnchorElement) {
            // Intercept when download attribute is set
            const originalSetAttribute = element.setAttribute;
            element.setAttribute = function(name, value) {
                if (name === 'download') {
                    console.log('[DOWNLOAD] Intercepted download link creation, filename:', value);
                    // Mark this element as a download link
                    this._isDownloadLink = true;
                    this._downloadFilename = value;
                }
                return originalSetAttribute.call(this, name, value);
            };

            // Intercept click to capture the blob data
            const originalClick = element.click;
            element.click = function() {
                if (this._isDownloadLink && this.href && this.href.startsWith('blob:')) {
                    console.log('[DOWNLOAD] Intercepted blob download:', this.href);
                    console.log('[DOWNLOAD] Filename:', this._downloadFilename);

                    // Fetch the blob data
                    fetch(this.href)
                        .then(function(response) { return response.arrayBuffer(); })
                        .then(function(arrayBuffer) {
                            console.log('[DOWNLOAD] Captured blob data, size:', arrayBuffer.byteLength, 'bytes');

                            // Call the resolver to pass the data to LocalFileSave
                            if (window._saveDownloadResolver) {
                                console.log('[DOWNLOAD] Calling save download resolver');
                                window._saveDownloadResolver(arrayBuffer);
                            } else if (window._downloadResolver) {
                                console.log('[DOWNLOAD] Calling generic download resolver');
                                window._downloadResolver(arrayBuffer);
                                window._downloadResolver = null;
                                window._downloadRejector = null;
                            } else {
                                console.error('[DOWNLOAD] No download resolver found!');
                            }

                            // Revoke the blob URL to free memory
                            URL.revokeObjectURL(this.href);
                        }.bind(this))
                        .catch(function(error) {
                            console.error('[DOWNLOAD] Failed to capture blob:', error);
                            if (window._downloadRejector) {
                                window._downloadRejector(error);
                                window._downloadResolver = null;
                                window._downloadRejector = null;
                            }
                        });

                    // Prevent the actual download
                    return false;
                } else {
                    // Not our intercepted download, proceed normally
                    return originalClick.call(this);
                }
            };
        }

        return element;
    };

    // Intercept XMLHttpRequest to handle ascdesktop:// fonts protocol
    const originalXHROpen = XMLHttpRequest.prototype.open;
    const originalXHRSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function(method, url, ...args) {
        this._interceptedUrl = url;
        if (typeof url === 'string' && url.startsWith('ascdesktop://fonts/')) {
            const fontPath = url.replace('ascdesktop://fonts/', '');
            // NOTE(victor): encodeURIComponent handles Windows paths with colons (C:/Windows/Fonts/...)
            const newUrl = SERVER_BASE_URL + '/fonts/' + encodeURIComponent(fontPath);
            console.log('[FONT] Intercepting font request:', fontPath, '->', newUrl);
            return originalXHROpen.call(this, method, newUrl, ...args);
        }
        return originalXHROpen.call(this, method, url, ...args);
    };

    // ADDITIONAL: Intercept URL.createObjectURL to capture blobs before they're used
    const originalCreateObjectURL = URL.createObjectURL;
    URL.createObjectURL = function(blob) {
        console.log('[BLOB] URL.createObjectURL called with blob:', blob);
        console.log('[BLOB] Blob type:', blob.type);
        console.log('[BLOB] Blob size:', blob.size);

        // Check if this is an XLSX blob we should capture
        if (window._saveDownloadResolver &&
            (blob.type.includes('spreadsheet') ||
             blob.type.includes('xlsx') ||
             blob.type.includes('officedocument') ||
             blob.type.includes('application/vnd.openxmlformats'))) {
            console.log('[BLOB] This appears to be an XLSX blob for save operation!');
            console.log('[BLOB] Converting blob to ArrayBuffer...');

            // Convert blob to ArrayBuffer and resolve the promise
            blob.arrayBuffer().then(function(arrayBuffer) {
                console.log('[BLOB] Converted to ArrayBuffer, size:', arrayBuffer.byteLength);
                console.log('[BLOB] Resolving save promise with blob data');
                window._saveDownloadResolver(arrayBuffer);
                window._saveDownloadResolver = null;
                window._saveDownloadRejector = null;
            }).catch(function(error) {
                console.error('[BLOB] Failed to convert blob to ArrayBuffer:', error);
                if (window._saveDownloadRejector) {
                    window._saveDownloadRejector(error);
                    window._saveDownloadResolver = null;
                    window._saveDownloadRejector = null;
                }
            });
        }

        // Still create the object URL (for normal downloads or if we're not intercepting)
        return originalCreateObjectURL.call(URL, blob);
    };

    console.log('Font protocol interception initialized - redirecting to ' + SERVER_BASE_URL + '/fonts/');
    console.log('Download interception initialized - capturing blob downloads from asc_DownloadAs');
    console.log('Blob URL interception initialized - capturing XLSX blobs via URL.createObjectURL');

    // Position get/set APIs for Excel scroll preservation
    window.getEditorPosition = function() {
        try {
            var editor = window.editor || (window.Asc && window.Asc.editor);
            if (!editor || !editor.wb) return null;
            var ws = editor.wb.getWorksheet();
            if (!ws || !ws.model || !ws.model.selectionRange) return null;
            var selection = ws.model.selectionRange;
            var activeCell = selection.activeCell;
            return {
                row: activeCell.row,
                col: activeCell.col,
                sheetIndex: ws.model.getIndex()
            };
        } catch (e) {
            console.log('[BROWSER] getEditorPosition error:', e);
            return null;
        }
    };

    window.setEditorPosition = function(position) {
        try {
            var editor = window.editor || (window.Asc && window.Asc.editor);
            if (!editor || !editor.wb || !position) return;
            var api = editor.asc_getEditorApi ? editor.asc_getEditorApi() : editor;
            if (api && api.asc_setActiveCell) {
                api.asc_setActiveCell(position.col, position.row);
            }
        } catch (e) {
            console.log('[BROWSER] setEditorPosition error:', e);
        }
    };

    window.addEventListener('message', function(event) {
        try {
            if (!event.data || typeof event.data !== 'object') return;

            if (event.data.type === 'GET_EDITOR_POSITION') {
                window.parent.postMessage({
                    type: 'EDITOR_POSITION',
                    position: window.getEditorPosition()
                }, '*');
            }

            if (event.data.type === 'SET_EDITOR_POSITION' && event.data.position) {
                window.setEditorPosition(event.data.position);
            }
        } catch (e) {
            console.log('[BROWSER] message handler error:', e);
        }
    });

    console.log('Position get/set APIs initialized for Excel scroll preservation');
})();

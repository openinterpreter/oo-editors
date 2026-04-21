const express = require('express');
const compression = require('compression');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { version: OO_EDITORS_VERSION } = require('./package.json');
const {
  initOoEditorsSentry,
  addLifecycleBreadcrumb,
  captureLifecycleMessage,
  captureLifecycleException,
  flushSentry
} = require('./sentry');
const {
  getX2TFormatCode,
  getOutputFormatInfo,
  generateFileHash,
  getDocTypeFromFilename,
  isAbsolutePath,
  getContentType,
  isXLSXSignature,
  classifySendFileError,
  resolveX2TPath,
  describeChildExit,
  getBufferedStderrLog,
  runChildProcess
} = require('./server-utils');

const LOG_NAMESPACE = 'oo-editors';
const SHUTDOWN_FORCE_EXIT_TIMEOUT_MS = 5000;
let stdoutBrokenPipe = false;
let stderrBrokenPipe = false;
let brokenPipeTelemetrySent = false;

function normalizeLogScope(scope) {
  return Array.isArray(scope) ? scope : [scope];
}

function formatLogScope(scope) {
  return `[${LOG_NAMESPACE}:${normalizeLogScope(scope).join(':')}]`;
}

function isBrokenPipeError(error) {
  const message = error && typeof error.message === 'string' ? error.message : '';
  return Boolean(error && (
    error.code === 'EPIPE'
    || error.errno === 'EPIPE'
    || message.includes('EPIPE')
  ));
}

function getErrorOwnProperties(error) {
  if (!error || typeof error !== 'object') {
    return {};
  }

  const details = {};
  for (const key of Object.getOwnPropertyNames(error)) {
    details[key] = error[key];
  }

  return details;
}

function markBrokenPipeStream(streamName, error) {
  if (streamName === 'stdout') {
    stdoutBrokenPipe = true;
  } else {
    stderrBrokenPipe = true;
  }

  if (brokenPipeTelemetrySent) {
    return;
  }

  brokenPipeTelemetrySent = true;
  const details = {
    event: 'stdio broken pipe',
    stream: streamName,
    timestamp: new Date().toISOString(),
    pid: process.pid,
    ppid: process.ppid,
    runtime: typeof Bun !== 'undefined' && typeof Bun.version === 'string'
      ? `bun-${Bun.version}`
      : `node-${process.versions.node}`,
    brokenPipeState: {
      stdoutBrokenPipe,
      stderrBrokenPipe,
      brokenPipeTelemetrySent
    },
    errorType: error && error.constructor ? error.constructor.name : typeof error,
    error: {
      name: error && error.name ? error.name : 'Error',
      message: error && error.message ? error.message : 'write EPIPE',
      code: error && error.code ? error.code : 'EPIPE',
      errno: error && error.errno ? error.errno : 'EPIPE',
      syscall: error && error.syscall ? error.syscall : null,
      stack: error && error.stack ? error.stack : null,
      ownProperties: getErrorOwnProperties(error)
    }
  };

  addLifecycleBreadcrumb('stdio broken pipe', details, {
    category: 'oo-editors.process',
    level: 'warning'
  });
}

function writeLog(method, streamName, scope, message, args) {
  if (streamName === 'stdout' && stdoutBrokenPipe) {
    return;
  }

  if (streamName === 'stderr' && stderrBrokenPipe) {
    return;
  }

  const formatted = `${formatLogScope(scope)} ${message}`;

  try {
    console[method](formatted, ...args);
  } catch (error) {
    if (isBrokenPipeError(error)) {
      markBrokenPipeStream(streamName, error);
      return;
    }

    throw error;
  }
}

function logInfo(scope, message, ...args) {
  writeLog('log', 'stdout', scope, message, args);
}

function logWarn(scope, message, ...args) {
  writeLog('log', 'stdout', scope, message, args);
}

function logError(scope, message, ...args) {
  writeLog('error', 'stderr', scope, message, args);
}

initOoEditorsSentry();

if (!process.env.FONT_DATA_DIR) {
  logError('STARTUP', 'FONT_DATA_DIR environment variable is required');
  process.exit(1);
}

const FONT_DATA_DIR = isAbsolutePath(process.env.FONT_DATA_DIR)
  ? process.env.FONT_DATA_DIR
  : path.join(__dirname, process.env.FONT_DATA_DIR);
const THEME_DIR = path.join(__dirname, 'editors', 'sdkjs', 'slide', 'themes');
const { path: X2T_PATH, candidates: X2T_PATH_CANDIDATES } = resolveX2TPath(__dirname, {
  pathExists: fs.existsSync
});

if (!fs.existsSync(FONT_DATA_DIR)) {
  logError('STARTUP', `FONT_DATA_DIR does not exist: ${FONT_DATA_DIR}`);
  process.exit(1);
}

if (!fs.existsSync(path.join(FONT_DATA_DIR, 'AllFonts.js'))) {
  logError('STARTUP', `AllFonts.js not found under FONT_DATA_DIR: ${FONT_DATA_DIR}`);
  process.exit(1);
}

if (!fs.existsSync(THEME_DIR)) {
  logError('STARTUP', `theme directory not found: ${THEME_DIR}`);
  process.exit(1);
}

if (!fs.existsSync(X2T_PATH)) {
  logError('STARTUP', `x2t binary not found. Checked: ${X2T_PATH_CANDIDATES.join(', ')}`);
  process.exit(1);
}

logInfo('STARTUP', `version=${OO_EDITORS_VERSION}`);
logInfo('STARTUP', `FONT_DATA_DIR=${FONT_DATA_DIR}`);
logInfo('STARTUP', `THEME_DIR=${THEME_DIR}`);
logInfo('STARTUP', `x2t=${X2T_PATH}`);
addLifecycleBreadcrumb('startup checks passed', {
  version: OO_EDITORS_VERSION,
  fontDataDir: FONT_DATA_DIR,
  themeDir: THEME_DIR,
  x2tPath: X2T_PATH
}, {
  category: 'oo-editors.startup'
});

const app = express();
let server = null;
let shutdownStarted = false;
app.use(compression());
const PORT = Number.parseInt(process.env.PORT || '38123', 10);
const BASE_URL = `http://localhost:${PORT}`;

// Enable CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', '*');
  res.header('Access-Control-Expose-Headers', 'X-File-Hash');
  next();
});

// Parse binary data for POST requests
app.use(express.raw({ type: 'application/octet-stream', limit: '2gb' }));

// Parse JSON body
app.use(express.json());

// Log all requests
app.use((req, res, next) => {
  logInfo('HTTP', `${new Date().toISOString()} ${req.method} ${req.url}`);
  next();
});

function sendFileWithHttpErrors(res, filePath, options) {
  const {
    logScope,
    notFoundBody,
    notFoundLogLevel = 'warn',
    internalErrorBody,
    successMessage
  } = options;

  res.sendFile(filePath, (err) => {
    if (!err) {
      if (successMessage) {
        logInfo(logScope, successMessage);
      }
      return;
    }

    const failure = classifySendFileError(err, {
      notFoundBody,
      internalErrorBody
    });

    if (failure.action === 'ignore') {
      return;
    }

    if (failure.logLevel === 'warn' && notFoundLogLevel !== 'error') {
      logWarn(logScope, `sendFile returned 404: ${filePath}`);
    } else {
      logError(logScope, `Error serving ${filePath}:`, err);
    }

    if (!res.headersSent) {
      res.status(failure.statusCode).send(failure.body);
    }
  });
}

function cleanupFiles(filePaths, logScope) {
  for (const filePath of filePaths) {
    if (!filePath) {
      continue;
    }

    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    } catch (err) {
      logWarn(logScope, `Failed to delete ${filePath}: ${err.message}`);
    }
  }
}

function runX2T(paramsPath, options) {
  const { logScope } = options;
  const lifecycleScope = normalizeLogScope(logScope).join(':').toLowerCase();

  addLifecycleBreadcrumb('starting x2t', {
    logScope: lifecycleScope,
    paramsFile: path.basename(paramsPath)
  }, {
    category: 'oo-editors.x2t'
  });
  logInfo(logScope, `Starting x2t: ${X2T_PATH} ${paramsPath}`);

  return runChildProcess((command, args) => {
    const x2t = spawn(command, args);
    x2t.on('spawn', () => {
      addLifecycleBreadcrumb('x2t spawned', {
        logScope: lifecycleScope,
        pid: x2t.pid
      }, {
        category: 'oo-editors.x2t'
      });
      logInfo(logScope, `x2t pid=${x2t.pid}`);
    });
    return x2t;
  }, X2T_PATH, [paramsPath], {
    onStdout: (output) => {
      const trimmed = output.trim();
      if (trimmed) {
        logInfo([].concat(normalizeLogScope(logScope), ['x2t', 'stdout']), trimmed);
      }
    }
  }).then((result) => {
    const exitDescription = describeChildExit(result.code, result.signal);
    const stderrLog = getBufferedStderrLog(result);

    if (stderrLog) {
      const stderrScope = [].concat(normalizeLogScope(logScope), ['x2t', 'stderr']);
      if (stderrLog.level === 'info') {
        logInfo(stderrScope, stderrLog.output);
      } else {
        logError(stderrScope, stderrLog.output);
      }
    }

    addLifecycleBreadcrumb('x2t exited', {
      logScope: lifecycleScope,
      code: result.code,
      signal: result.signal,
      exitDescription
    }, {
      category: 'oo-editors.x2t',
      level: result.code === 0 ? 'info' : 'error'
    });
    logInfo(logScope, `x2t exited with ${exitDescription}`);
    return {
      ...result,
      exitDescription
    };
  }).catch((failure) => {
    const error = failure.error;
    error.stdout = failure.stdout;
    error.stderr = failure.stderr;
    addLifecycleBreadcrumb('x2t failed to start', {
      logScope: lifecycleScope,
      message: error.message
    }, {
      category: 'oo-editors.x2t',
      level: 'error'
    });
    captureLifecycleException(error, {
      level: 'error',
      tags: {
        phase: 'x2t-start',
        logScope: lifecycleScope
      }
    });
    logError(logScope, `Failed to start x2t at ${X2T_PATH}:`, error);
    throw error;
  });
}

function shutdownServer(reason, exitCode) {
  if (shutdownStarted) {
    return;
  }

  shutdownStarted = true;
  addLifecycleBreadcrumb('shutdown requested', {
    reason,
    exitCode,
    serverListening: Boolean(server && server.listening)
  }, {
    category: 'oo-editors.process',
    level: exitCode === 0 ? 'info' : 'error'
  });
  logInfo('PROCESS', `shutdown requested (${reason}), exitCode=${exitCode}, version=${OO_EDITORS_VERSION}`);

  const finalize = (finalCode) => {
    logInfo('PROCESS', `exiting with code ${finalCode}`);
    void flushSentry().finally(() => {
      process.exit(finalCode);
    });
  };

  const forceExitTimer = setTimeout(() => {
    captureLifecycleMessage('oo-editors shutdown timed out', {
      level: 'error',
      tags: {
        phase: 'shutdown-timeout'
      },
      data: {
        reason,
        exitCode,
        timeoutMs: SHUTDOWN_FORCE_EXIT_TIMEOUT_MS
      }
    });
    logError('PROCESS', 'shutdown timed out, forcing exit');
    process.exit(exitCode);
  }, SHUTDOWN_FORCE_EXIT_TIMEOUT_MS);
  forceExitTimer.unref();

  if (!server || !server.listening) {
    clearTimeout(forceExitTimer);
    finalize(exitCode);
    return;
  }

  server.close((err) => {
    clearTimeout(forceExitTimer);

    if (err) {
      captureLifecycleException(err, {
        level: 'error',
        tags: {
          phase: 'shutdown-close'
        },
        data: {
          reason
        }
      });
      logError('PROCESS', 'error while closing HTTP server:', err);
      finalize(1);
      return;
    }

    logInfo('PROCESS', 'HTTP server closed');
    finalize(exitCode);
  });
}

function handleStdIoStreamError(streamName, error) {
  if (isBrokenPipeError(error)) {
    markBrokenPipeStream(streamName, error);
    shutdownServer(`${streamName}-epipe`, 0);
    return;
  }

  addLifecycleBreadcrumb('stdio stream error', {
    stream: streamName,
    message: error && error.message ? error.message : String(error)
  }, {
    category: 'oo-editors.process',
    level: 'error'
  });
  captureLifecycleException(error, {
    level: 'error',
    tags: {
      phase: 'stdio',
      stream: streamName
    }
  });
  shutdownServer(`${streamName}-error`, 1);
}

process.stdout.on('error', (error) => {
  handleStdIoStreamError('stdout', error);
});
process.stderr.on('error', (error) => {
  handleStdIoStreamError('stderr', error);
});
process.on('SIGTERM', () => shutdownServer('SIGTERM', 0));
process.on('SIGINT', () => shutdownServer('SIGINT', 0));
process.on('disconnect', () => shutdownServer('disconnect', 0));
process.on('beforeExit', (code) => {
  logInfo('PROCESS', `beforeExit with code ${code}`);
});
process.on('exit', (code) => {
  logInfo('PROCESS', `exit with code ${code}`);
});
process.on('uncaughtException', (err) => {
  if (isBrokenPipeError(err)) {
    markBrokenPipeStream('stdout', err);
    shutdownServer('uncaughtException-epipe', 0);
    return;
  }

  addLifecycleBreadcrumb('uncaught exception', {
    message: err.message,
    name: err.name
  }, {
    category: 'oo-editors.process',
    level: 'error'
  });
  captureLifecycleException(err, {
    level: 'fatal',
    tags: {
      phase: 'uncaughtException'
    }
  });
  logError('PROCESS', 'uncaught exception:', err);
  shutdownServer('uncaughtException', 1);
});
process.on('unhandledRejection', (reason) => {
  if (isBrokenPipeError(reason)) {
    markBrokenPipeStream('stdout', reason);
    shutdownServer('unhandledRejection-epipe', 0);
    return;
  }

  addLifecycleBreadcrumb('unhandled rejection', {
    reason: reason instanceof Error ? reason.message : String(reason)
  }, {
    category: 'oo-editors.process',
    level: 'error'
  });
  captureLifecycleException(reason, {
    level: 'fatal',
    tags: {
      phase: 'unhandledRejection'
    }
  });
  logError('PROCESS', 'unhandled rejection:', reason);
  shutdownServer('unhandledRejection', 1);
});

// API Endpoint: Health check
app.get('/healthcheck', (req, res) => {
  res.setHeader('Content-Type', 'text/plain');
  res.send('true');
});

// Allow absolute ascdesktop:// font paths (e.g. ascdesktop://fonts//System/Library/Fonts/Arial.ttf)
app.get('/fonts/*', (req, res) => {
  const rawPath = req.params[0];
  if (!rawPath) {
    return res.status(404).send('Font not found');
  }

  const decoded = decodeURIComponent(rawPath);

  // Normalize: requests often arrive with a leading slash (absolute macOS path).
  const fontPath = isAbsolutePath(decoded) ? decoded : path.join(__dirname, decoded);
  logInfo('FONTS', `request for absolute font path: ${fontPath}`);

  if (!fs.existsSync(fontPath)) {
    logError('FONTS', `absolute font path not found: ${fontPath}`);
    return res.status(404).send('Font not found');
  }

  try {
    const ext = path.extname(fontPath).toLowerCase();
    const contentTypes = {
      '.ttf': 'font/ttf',
      '.otf': 'font/otf',
      '.ttc': 'font/collection',
      '.woff': 'font/woff',
      '.woff2': 'font/woff2'
    };
    const fontContentType = contentTypes[ext] || 'application/octet-stream';

    res.setHeader('Content-Type', fontContentType);
    res.setHeader('Access-Control-Allow-Origin', '*');
    sendFileWithHttpErrors(res, fontPath, {
      logScope: 'FONTS',
      notFoundBody: 'Font not found',
      internalErrorBody: 'Font error',
      successMessage: `served font: ${fontPath}`
    });
  } catch (err) {
    logError('FONTS', `error serving absolute font ${fontPath}:`, err);
    res.status(500).send('Font error');
  }
});

// Surface service worker expected at origin root (mirrors desktop packaging)
app.get('/document_editor_service_worker.js', (req, res) => {
  const workerPath = path.join(
    __dirname,
    'editors',
    'sdkjs',
    'common',
    'serviceworker',
    'document_editor_service_worker.js'
  );

  if (!fs.existsSync(workerPath)) {
    logError('SW', 'document_editor_service_worker.js not found at', workerPath);
    return res.status(404).send('Service worker not found');
  }

  res.setHeader('Content-Type', 'application/javascript');
  res.setHeader('Cache-Control', 'no-cache');
  logInfo('SW', 'serving document_editor_service_worker.js');
  sendFileWithHttpErrors(res, workerPath, {
    logScope: 'SW',
    notFoundBody: 'Service worker not found',
    notFoundLogLevel: 'error',
    internalErrorBody: 'Service worker error'
  });
});

// Serve the desktop AllFonts.js verbatim for metadata parity
app.get('/fonts-info.js', (req, res) => {
  const allFontsPath = path.join(FONT_DATA_DIR, 'AllFonts.js');
  logInfo('API', 'GET /fonts-info.js serving', allFontsPath);
  res.setHeader('Content-Type', 'application/javascript');
  res.setHeader('Cache-Control', 'no-cache');
  sendFileWithHttpErrors(res, allFontsPath, {
    logScope: 'API',
    notFoundBody: '// Failed to load font metadata',
    notFoundLogLevel: 'error',
    internalErrorBody: '// Failed to load font metadata'
  });
});

// Override the SDK's bundled AllFonts.js (Linux-oriented) with the desktop macOS version
app.get('/sdkjs/common/AllFonts.js', (req, res) => {
  const desktopAllFontsPath = path.join(FONT_DATA_DIR, 'AllFonts.js');
  logInfo('API', 'GET /sdkjs/common/AllFonts.js overriding with desktop AllFonts.js');
  res.setHeader('Content-Type', 'application/javascript');
  res.setHeader('Cache-Control', 'no-cache');
  sendFileWithHttpErrors(res, desktopAllFontsPath, {
    logScope: 'API',
    notFoundBody: '// Failed to load AllFonts override',
    notFoundLogLevel: 'error',
    internalErrorBody: '// Failed to load AllFonts override'
  });
});

// API Endpoint: Upload image to media directory (for drag-drop)
app.post('/api/media/:filehash', (req, res) => {
  const filehash = req.params.filehash;
  const filename = req.query.filename || `image_${Date.now()}.png`;

  logInfo('MEDIA-UPLOAD', `uploading ${filename} for hash ${filehash}`);

  const outputDir = path.join(__dirname, 'test', 'output', filehash);
  const mediaDir = path.join(outputDir, 'media');

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  if (!fs.existsSync(mediaDir)) {
    fs.mkdirSync(mediaDir, { recursive: true });
  }

  const imagePath = path.join(mediaDir, filename);
  fs.writeFileSync(imagePath, req.body);

  logInfo('MEDIA-UPLOAD', `saved ${req.body.length} bytes to ${imagePath}`);
  res.json({ filename: filename, path: `/api/media/${filehash}/${filename}` });
});

// API Endpoint: Serve images from converted documents
// Updated to support file-specific media directories
app.get('/api/media/:filehash/:imagefile', (req, res) => {
  const filehash = req.params.filehash;
  const imagefile = req.params.imagefile;
  logInfo('MEDIA', `request for image ${imagefile} (file hash: ${filehash})`);

  // Images are in hash-specific output directory
  const outputDir = path.join(__dirname, 'test', 'output', filehash);
  const mediaDir = path.join(outputDir, 'media');
  const imagePath = path.join(mediaDir, imagefile);

  logInfo('MEDIA', `looking for image at: ${imagePath}`);

  if (!fs.existsSync(imagePath)) {
    logError('MEDIA', `image not found: ${imagePath}`);

    // Check if media directory exists at all
    if (!fs.existsSync(mediaDir)) {
      logError('MEDIA', `media directory does not exist: ${mediaDir}`);
    } else {
      logInfo('MEDIA', 'media directory contents:', fs.readdirSync(mediaDir));
    }

    return res.status(404).send('Image not found');
  }

  // Detect content type based on file extension
  const ext = path.extname(imagefile).toLowerCase();
  const contentTypes = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.bmp': 'image/bmp',
    '.webp': 'image/webp',
    '.emf': 'image/x-emf',
    '.wmf': 'image/x-wmf'
  };

  const contentType = contentTypes[ext] || 'application/octet-stream';

  res.setHeader('Content-Type', contentType);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'public, max-age=31536000'); // Cache for 1 year

  logInfo('MEDIA', `serving image: ${imagePath} (${contentType})`);
  sendFileWithHttpErrors(res, imagePath, {
    logScope: 'MEDIA',
    notFoundBody: 'Image not found',
    notFoundLogLevel: 'error',
    internalErrorBody: 'Image error'
  });
});

// API Endpoint: Serve arbitrary files within the converted document directory (used for relative resource lookups)
app.get('/api/doc-base/:filehash/*', (req, res) => {
  const filehash = req.params.filehash;
  const relativePath = req.params[0] || '';
  logInfo('DOC-BASE', `request for path "${relativePath}" (file hash: ${filehash})`);

  if (!relativePath) {
    return res.status(400).send('Path is required');
  }

  const outputDir = path.join(__dirname, 'test', 'output', filehash);
  const normalizedPath = path
    .normalize(relativePath)
    .replace(/^(\.\.(\/|\\|$))+/, '');
  const requestedPath = path.join(outputDir, normalizedPath);

  if (!requestedPath.startsWith(outputDir)) {
    logWarn('DOC-BASE', 'attempted directory traversal:', requestedPath);
    return res.status(403).send('Forbidden');
  }

  if (!fs.existsSync(requestedPath) || !fs.statSync(requestedPath).isFile()) {
    logInfo('DOC-BASE', 'file not found:', requestedPath);
    return res.status(404).send('File not found');
  }

  sendFileWithHttpErrors(res, requestedPath, {
    logScope: 'DOC-BASE',
    notFoundBody: 'File not found',
    internalErrorBody: 'File error'
  });
});

// API Endpoint: List images in media directory for a file
app.get('/api/media-list/:filehash', (req, res) => {
  const filehash = req.params.filehash;
  logInfo('MEDIA-LIST', `request for file hash: ${filehash}`);

  const outputDir = path.join(__dirname, 'test', 'output', filehash);
  const mediaDir = path.join(outputDir, 'media');

  if (!fs.existsSync(mediaDir)) {
    logInfo('MEDIA-LIST', `no media directory found for hash: ${filehash}`);
    return res.json([]);
  }

  try {
    const files = fs.readdirSync(mediaDir);
    logInfo('MEDIA-LIST', `found ${files.length} files:`, files);
    res.json(files);
  } catch (err) {
    logError('MEDIA-LIST', 'error reading media directory:', err);
    res.status(500).json({ error: 'Error reading media directory' });
  }
});

// API Endpoint: Convert XLSX to ONLYOFFICE binary format
// ONLY supports absolute paths via query parameter
app.get('/api/convert', async (req, res) => {
  try {
    const timings = { start: performance.now() };
    const filepath = req.query.filepath;

    if (!filepath) {
      return res.status(400).json({ error: 'filepath query parameter is required' });
    }

    if (!isAbsolutePath(filepath)) {
      return res.status(400).json({ error: 'filepath must be an absolute path' });
    }

    logInfo('CONVERT', `requested file: ${filepath}`);
    timings.validated = performance.now();

    // Create unique output directory based on the source file path
    const fileHash = generateFileHash(filepath);
    const outputDir = path.join(__dirname, 'test', 'output', fileHash);
    const inputPath = filepath;

    logInfo('CONVERT', `file hash: ${fileHash}`);
    logInfo('CONVERT', `output directory: ${outputDir}`);

    if (!fs.existsSync(inputPath)) {
      logWarn('CONVERT', `file not found: ${inputPath}`);
      return res.status(404).json({ error: 'File not found at absolute path' });
    }

    const outputPath = path.join(outputDir, 'Editor.bin');
    const paramsPath = path.join(outputDir, 'params_temp.xml');

    // Extract filename from path for logging
    const filename = path.basename(filepath);

    const sourceMtime = fs.statSync(inputPath).mtimeMs;
    if (fs.existsSync(outputPath)) {
      const cacheMtime = fs.statSync(outputPath).mtimeMs;
      if (cacheMtime > sourceMtime) {
        logInfo('CONVERT', `cache hit using cached Editor.bin (source: ${new Date(sourceMtime).toISOString()}, cache: ${new Date(cacheMtime).toISOString()})`);
        timings.cacheHit = true;
        const binaryData = fs.readFileSync(outputPath);

        timings.end = performance.now();
        const breakdown = {
          total: (timings.end - timings.start).toFixed(1),
          cacheHit: true,
          validation: (timings.validated - timings.start).toFixed(1),
        };
        logInfo(['CONVERT', 'TIMING'], 'cache hit breakdown (ms):', JSON.stringify(breakdown));

        res.setHeader('Content-Type', 'application/octet-stream');
        res.setHeader('Content-Disposition', `attachment; filename="Editor.bin"`);
        res.setHeader('X-File-Hash', fileHash);
        res.setHeader('X-Timing', JSON.stringify(breakdown));
        res.setHeader('X-Cache', 'HIT');
        return res.send(binaryData);
      }
      logInfo('CONVERT', `cache stale, reconverting (source: ${new Date(sourceMtime).toISOString()}, cache: ${new Date(cacheMtime).toISOString()})`);
    }

    logInfo('CONVERT', `converting ${filename} to binary format`);
    logInfo('CONVERT', `input: ${inputPath}`);
    logInfo('CONVERT', `output: ${outputPath}`);

    timings.beforeMkdir = performance.now();
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    timings.afterMkdir = performance.now();

    // Create XML config for x2t
    // CRITICAL: Use the same fonts directory that contains AllFonts.js served to browser
    // This ensures x2t assigns the same font IDs that the browser expects
    const fontDir = FONT_DATA_DIR;

    const xmlConfig = `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<m_sKey>api_conversion</m_sKey>
<m_sFileFrom>${inputPath}</m_sFileFrom>
<m_sFileTo>${outputPath}</m_sFileTo>
<m_sTitle>${filename}</m_sTitle>
<m_nFormatTo>8192</m_nFormatTo>${path.extname(filename).toLowerCase() === '.csv' ? '\n<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding>\n<m_nCsvDelimiter>4</m_nCsvDelimiter>' : ''}
<m_bPaid xsi:nil="true" />
<m_bEmbeddedFonts xsi:nil="true" />
<m_bFromChanges>false</m_bFromChanges>
<m_sFontDir>${fontDir}</m_sFontDir>
<m_sThemeDir>${THEME_DIR}</m_sThemeDir>
<m_sJsonParams>{}</m_sJsonParams>
<m_nLcid xsi:nil="true" />
<m_oTimestamp>${new Date().toISOString()}</m_oTimestamp>
<m_bIsNoBase64 xsi:nil="true" />
<m_sConvertToOrigin xsi:nil="true" />
<m_oInputLimits>
</m_oInputLimits>
<options>
<allowNetworkRequest>false</allowNetworkRequest>
<allowPrivateIP>true</allowPrivateIP>
</options>
</TaskQueueDataConvert>
`;

    // Write XML config
    timings.beforeXmlWrite = performance.now();
    fs.writeFileSync(paramsPath, xmlConfig);
    timings.afterXmlWrite = performance.now();
    logInfo('CONVERT', `XML config written to ${paramsPath}`);

    // Run x2t converter
    timings.beforeX2t = performance.now();
    let x2tResult;
    try {
      x2tResult = await runX2T(paramsPath, { logScope: 'CONVERT' });
    } catch (error) {
      timings.afterX2t = performance.now();
      cleanupFiles([paramsPath], 'CONVERT');
      const details = error.stderr || error.stdout || error.message;
      logError('CONVERT', 'conversion failed before x2t completed:', error);
      return res.status(500).send(`Conversion failed: ${details}`);
    }

    timings.afterX2t = performance.now();
    cleanupFiles([paramsPath], 'CONVERT');

    if (x2tResult.code !== 0) {
      const details = x2tResult.stderr || x2tResult.stdout || 'x2t exited without output';
      logError('CONVERT', `conversion failed (${x2tResult.exitDescription})`);
      return res.status(500).send(`Conversion failed (${x2tResult.exitDescription}): ${details}`);
    }

    // Check if output file was created
    if (!fs.existsSync(outputPath)) {
      logError('CONVERT', 'output file not created');
      return res.status(500).send('Conversion failed: output file not created');
    }

    // Read and send the binary file
    timings.beforeReadOutput = performance.now();
    logInfo('CONVERT', `reading output file: ${outputPath}`);
    const binaryData = fs.readFileSync(outputPath);
    timings.afterReadOutput = performance.now();
    logInfo('CONVERT', `sending ${binaryData.length} bytes`);

    // Send the file hash in a custom header so the browser can use it for image URLs
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', `attachment; filename="Editor.bin"`);
    res.setHeader('X-File-Hash', fileHash);
    res.setHeader('X-Cache', 'MISS');

    timings.end = performance.now();
    const breakdown = {
      total: (timings.end - timings.start).toFixed(1),
      validation: (timings.validated - timings.start).toFixed(1),
      mkdir: (timings.afterMkdir - timings.beforeMkdir).toFixed(1),
      xmlWrite: (timings.afterXmlWrite - timings.beforeXmlWrite).toFixed(1),
      x2tConversion: (timings.afterX2t - timings.beforeX2t).toFixed(1),
      readOutput: (timings.afterReadOutput - timings.beforeReadOutput).toFixed(1),
    };
    logInfo(['CONVERT', 'TIMING'], 'breakdown (ms):', JSON.stringify(breakdown));
    res.setHeader('X-Timing', JSON.stringify(breakdown));

    logInfo('CONVERT', `sent file hash in header: ${fileHash}`);
    res.send(binaryData);
  } catch (error) {
    logError('CONVERT', 'error:', error);
    if (!res.headersSent) {
      res.status(500).send(`Conversion failed: ${error.message}`);
    }
  }
});

// POST /converter - OnlyOffice Document Server API compatibility endpoint
app.post('/converter', async (req, res) => {
  try {
    logInfo('CONVERTER', 'OnlyOffice Document Server API request received');

    // Parse request body - handle both JWT token and raw payload
    let payload;
    if (req.body.token) {
      // JWT token present - decode without verification (local trusted environment)
      const token = req.body.token;
      const parts = token.split('.');
      if (parts.length !== 3) {
        return res.status(400).json({ error: -1, message: 'Invalid JWT token format' });
      }
      const payloadBase64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      const payloadJson = Buffer.from(payloadBase64, 'base64').toString('utf8');
      payload = JSON.parse(payloadJson);
      logInfo('CONVERTER', 'decoded JWT payload:', { filetype: payload.filetype, outputtype: payload.outputtype });
    } else {
      payload = req.body;
      logInfo('CONVERTER', 'using raw payload:', { filetype: payload.filetype, outputtype: payload.outputtype });
    }

    const { filetype, key, outputtype, title, url: payloadUrl } = payload;

    if (!filetype || !key || !outputtype) {
      return res.status(400).json({
        error: -1,
        message: 'Missing required fields: filetype, key, outputtype'
      });
    }

    let inputPath = payloadUrl;
    if (!inputPath) {
      return res.status(400).json({
        error: -1,
        message: 'Missing required field: url'
      });
    }

    // Handle URLs by extracting the file path
    // Example: http://host:port/api/onlyoffice/files/absolute/path/to/file.xlsx
    if (inputPath.startsWith('http://') || inputPath.startsWith('https://')) {
      const urlObj = new URL(inputPath);
      const pathMatch = urlObj.pathname.match(/\/api\/onlyoffice\/files\/(.+)/);
      if (pathMatch) {
        inputPath = '/' + pathMatch[1].split('/').map(decodeURIComponent).join('/');
        logInfo('CONVERTER', 'extracted file path from URL:', inputPath);
      } else {
        return res.status(400).json({
          error: -1,
          message: 'Cannot extract file path from URL: ' + inputPath
        });
      }
    }

    logInfo('CONVERTER', `converting: ${inputPath}`);
    logInfo('CONVERTER', `format: ${filetype} -> ${outputtype}`);

    if (!fs.existsSync(inputPath)) {
      logWarn('CONVERTER', 'input file not found:', inputPath);
      return res.status(404).json({
        error: -1,
        message: 'Input file not found: ' + inputPath
      });
    }

    const formatToCode = getX2TFormatCode(outputtype);
    if (!formatToCode) {
      return res.status(400).json({
        error: -1,
        message: `Unsupported output format: ${outputtype}`
      });
    }

    const outputDir = path.join(__dirname, 'converted');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const fileHash = crypto.createHash('md5').update(key + Date.now()).digest('hex');
    const outputFilename = `${fileHash}.${outputtype}`;
    const outputPath = path.join(outputDir, outputFilename);
    const paramsPath = path.join(outputDir, `params_${fileHash}.xml`);

    logInfo('CONVERTER', 'output will be:', outputPath);

    // Create XML config for x2t converter
    const xmlConfig = `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<m_sKey>converter_${key}</m_sKey>
<m_sFileFrom>${inputPath}</m_sFileFrom>
<m_sFileTo>${outputPath}</m_sFileTo>
<m_sTitle>${title || path.basename(inputPath)}</m_sTitle>
<m_nFormatTo>${formatToCode}</m_nFormatTo>
<m_bPaid xsi:nil="true" />
<m_bEmbeddedFonts xsi:nil="true" />
<m_bFromChanges>false</m_bFromChanges>
<m_sFontDir xsi:nil="true" />
<m_sThemeDir>${THEME_DIR}</m_sThemeDir>
<m_sJsonParams>{}</m_sJsonParams>
<m_nLcid xsi:nil="true" />
<m_oTimestamp>${new Date().toISOString()}</m_oTimestamp>
<m_bIsNoBase64 xsi:nil="true" />
<m_sConvertToOrigin xsi:nil="true" />
<m_oInputLimits>
</m_oInputLimits>
<options>
<allowNetworkRequest>false</allowNetworkRequest>
<allowPrivateIP>true</allowPrivateIP>
</options>
</TaskQueueDataConvert>
`;

    fs.writeFileSync(paramsPath, xmlConfig);
    logInfo('CONVERTER', 'XML config written');

    let x2tResult;
    try {
      x2tResult = await runX2T(paramsPath, { logScope: 'CONVERTER' });
    } catch (error) {
      cleanupFiles([paramsPath], 'CONVERTER');
      return res.status(500).json({
        error: -1,
        message: 'Conversion failed to start',
        details: error.stderr || error.stdout || error.message
      });
    }

    cleanupFiles([paramsPath], 'CONVERTER');

    if (x2tResult.code !== 0) {
      const details = x2tResult.stderr || x2tResult.stdout || 'x2t exited without output';
      logError('CONVERTER', `conversion failed (${x2tResult.exitDescription}):`, details);
      return res.status(500).json({
        error: -1,
        message: `Conversion failed (${x2tResult.exitDescription})`,
        details
      });
    }

    if (!fs.existsSync(outputPath)) {
      logError('CONVERTER', 'output file not created');
      return res.status(500).json({
        error: -1,
        message: 'Output file not created'
      });
    }

    const resultUrl = `http://localhost:${PORT}/converted/${outputFilename}`;
    logInfo('CONVERTER', `conversion successful: ${resultUrl}`);

    res.json({
      url: resultUrl,
      fileType: outputtype,
      error: 0
    });

  } catch (error) {
    logError('CONVERTER', 'error:', error);
    res.status(500).json({
      error: -1,
      message: 'Internal server error',
      details: error.message
    });
  }
});

app.get('/converted/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'converted', filename);

  logInfo('CONVERTED', `request for file: ${filename}`);

  if (!fs.existsSync(filePath)) {
    logError('CONVERTED', `file not found: ${filePath}`);
    return res.status(404).send('File not found');
  }

  const ext = path.extname(filename).toLowerCase();
  const contentTypes = {
    '.pdf': 'application/pdf',
    '.txt': 'text/plain',
    '.html': 'text/html',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  };

  const contentType = contentTypes[ext] || 'application/octet-stream';
  res.setHeader('Content-Type', contentType);
  res.setHeader('Access-Control-Allow-Origin', '*');

  logInfo('CONVERTED', `serving file: ${filePath} (${contentType})`);
  sendFileWithHttpErrors(res, filePath, {
    logScope: 'CONVERTED',
    notFoundBody: 'File not found',
    notFoundLogLevel: 'error',
    internalErrorBody: 'File error'
  });
});

// API Endpoint: Save binary back to XLSX
// ONLY supports absolute paths via query parameter
app.post('/api/save', async (req, res) => {
  try {
    const filepath = req.query.filepath;
    const filehash = req.query.filehash;

    if (!filepath) {
      return res.status(400).json({ error: 'filepath query parameter is required' });
    }

    if (!isAbsolutePath(filepath)) {
      return res.status(400).json({ error: 'filepath must be an absolute path' });
    }

    logInfo('SAVE', `requested file: ${filepath}`);
    logInfo('SAVE', `file hash: ${filehash || 'not provided'}`);

    const outputDir = path.join(__dirname, 'test', 'output');
    const outputPath = filepath;

    // Use hash-specific directory so x2t can find media files
    // x2t looks for media/ relative to input binary location
    const hashDir = filehash ? path.join(outputDir, filehash) : outputDir;
    if (!fs.existsSync(hashDir)) {
      fs.mkdirSync(hashDir, { recursive: true });
    }

    // Ensure parent directory exists
    const parentDir = path.dirname(outputPath);
    if (!fs.existsSync(parentDir)) {
      logInfo('SAVE', `creating parent directory: ${parentDir}`);
      fs.mkdirSync(parentDir, { recursive: true });
    }

    const filename = path.basename(filepath);

    logInfo('SAVE', `saving file: ${filename}`);
    logInfo('SAVE', `output path: ${outputPath}`);
    logInfo('SAVE', `content-type: ${req.get('Content-Type')}`);
    logInfo('SAVE', `body size: ${req.body.length} bytes`);

    // Debug: Check first few bytes
    const firstBytes = req.body.slice(0, 20);
    logInfo('SAVE', `first 20 bytes (hex): ${Buffer.from(firstBytes).toString('hex')}`);
    logInfo('SAVE', `first 20 bytes (ASCII): ${Buffer.from(firstBytes).toString('ascii').replace(/[^\x20-\x7E]/g, '.')}`);

    // Check if this is an XLSX file (should start with PK - ZIP signature)
    const isXLSX = isXLSXSignature(req.body);
    logInfo('SAVE', `detected format: ${isXLSX ? 'XLSX (ZIP)' : 'Unknown/Binary'}`);

    if (isXLSX) {
      // This is already an XLSX file - just save it directly!
      logInfo('SAVE', 'file is already XLSX format, saving directly');
      try {
        fs.writeFileSync(outputPath, req.body);
        logInfo('SAVE', `successfully saved XLSX file to ${outputPath}`);

        const stats = fs.statSync(outputPath);
        logInfo('SAVE', `file size: ${stats.size} bytes`);

        res.json({ success: true, path: outputPath, size: stats.size });
      } catch (error) {
        logError('SAVE', 'failed to save XLSX file:', error);
        res.status(500).send('Save failed: ' + error.message);
      }
    } else {
      // This is ONLYOFFICE binary format - convert it to the appropriate output format
      logInfo('SAVE', 'file appears to be ONLYOFFICE binary format');

      // Determine output format based on file extension
      const ext = path.extname(filepath).toLowerCase();
      const formatInfo = getOutputFormatInfo(ext);

      if (!formatInfo) {
        logWarn('SAVE', `unsupported file extension: ${ext}`);
        return res.status(400).send('Unsupported file format');
      }

      const { code: formatTo, name: formatName } = formatInfo;
      logInfo('SAVE', `converting received data to ${formatName}`);

      const saveRequestId = crypto.randomUUID();
      const changesBinPath = path.join(hashDir, `temp_changes_${saveRequestId}.bin`);
      const paramsPath = path.join(hashDir, `params_save_${saveRequestId}.xml`);

      // Save the received binary data
      fs.writeFileSync(changesBinPath, req.body);
      logInfo('SAVE', `wrote binary data: ${changesBinPath}`);

      // Convert the received data (with changes) to the output format
      logInfo('SAVE', `converting received data (with changes) to ${formatName}`);
      logInfo('SAVE', `converting from: ${changesBinPath}`);
      logInfo('SAVE', `output format: ${formatTo} (${formatName})`);

      // CRITICAL: Use the same fonts directory for save operations
      const fontDir = FONT_DATA_DIR;

      const xmlConfig = `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<m_sKey>api_save</m_sKey>
<m_sFileFrom>${changesBinPath}</m_sFileFrom>
<m_nFormatFrom>8192</m_nFormatFrom>
<m_sFileTo>${outputPath}</m_sFileTo>
<m_sTitle>${filename}</m_sTitle>
<m_nFormatTo>${formatTo}</m_nFormatTo>${ext === '.csv' ? '\n<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding>\n<m_nCsvDelimiter>4</m_nCsvDelimiter>' : ''}
<m_bPaid xsi:nil="true" />
<m_bEmbeddedFonts xsi:nil="true" />
<m_bFromChanges>false</m_bFromChanges>
<m_sFontDir>${fontDir}</m_sFontDir>
<m_sThemeDir>${THEME_DIR}</m_sThemeDir>
<m_sJsonParams>{}</m_sJsonParams>
<m_nLcid xsi:nil="true" />
<m_oTimestamp>${new Date().toISOString()}</m_oTimestamp>
<m_bIsNoBase64 xsi:nil="true" />
<m_sConvertToOrigin xsi:nil="true" />
<m_oInputLimits>
</m_oInputLimits>
<options>
<allowNetworkRequest>false</allowNetworkRequest>
<allowPrivateIP>true</allowPrivateIP>
</options>
</TaskQueueDataConvert>
`;

      fs.writeFileSync(paramsPath, xmlConfig);

      let x2tResult;
      try {
        x2tResult = await runX2T(paramsPath, { logScope: 'SAVE' });
      } catch (error) {
        cleanupFiles([paramsPath, changesBinPath], 'SAVE');
        return res.status(500).send(`Save failed: ${error.stderr || error.stdout || error.message}`);
      }

      cleanupFiles([paramsPath, changesBinPath], 'SAVE');

      if (x2tResult.code !== 0) {
        const details = x2tResult.stderr || x2tResult.stdout || 'x2t exited without output';
        logError('SAVE', `conversion failed (${x2tResult.exitDescription})`);
        return res.status(500).send(`Save failed: conversion error (${x2tResult.exitDescription}): ${details}`);
      }

      if (!fs.existsSync(outputPath)) {
        logError('SAVE', 'output file not created');
        return res.status(500).send('Save failed: no output file');
      }

      const stats = fs.statSync(outputPath);
      logInfo('SAVE', `successfully saved to ${outputPath}`);
      logInfo('SAVE', `file size: ${stats.size} bytes`);
      res.json({ success: true, path: outputPath, size: stats.size });
    }
  } catch (error) {
    logError('SAVE', 'error:', error);
    if (!res.headersSent) {
      res.status(500).send(`Save failed: ${error.message}`);
    }
  }
});

// API Endpoint: Load document with simple loader (direct binary loading)
app.get('/load/:filename', (req, res) => {
  const filename = req.params.filename;

  logInfo('LOAD', `loading ${filename} with simple loader`);

  // Serve the simple-loader.html with filename as query parameter
  const loaderPath = path.join(__dirname, 'editors', 'simple-loader.html');

  fs.readFile(loaderPath, 'utf8', (err, html) => {
    if (err) {
      logError('LOAD', 'error reading simple-loader.html:', err);
      return res.status(500).send('Failed to load editor');
    }

    // Inject the filename into the HTML (replace URL to include filename)
    const modifiedHtml = html.replace(
      '</head>',
      `<script>window.ONLYOFFICE_FILENAME = '${filename}';</script></head>`
    );

    res.setHeader('Content-Type', 'text/html');
    res.send(modifiedHtml);
  });
});

// API Endpoint: Open document in desktop editor
// ONLY supports absolute paths via query parameter
app.get('/open', (req, res) => {
  const filepath = req.query.filepath;
  // NOTE(victor): OnlyOffice expects ISO 639-1 codes (e.g. "es", "fr"), not full names like "spanish"
  const lang = typeof req.query.lang === 'string' ? req.query.lang : '';
  const theme = typeof req.query.theme === 'string' ? req.query.theme : '';

  if (!filepath) {
    return res.status(400).json({ error: 'filepath query parameter is required' });
  }

  if (!isAbsolutePath(filepath)) {
    return res.status(400).json({ error: 'filepath must be an absolute path' });
  }

  logInfo('OPEN', `requested file: ${filepath}`);

  const filename = path.basename(filepath);
  const ext = filename.split('.').pop().toLowerCase();
  const docType = getDocTypeFromFilename(filename);

  // Build the document URL with proper encoding
  const documentUrl = `http://localhost:${PORT}/api/convert?filepath=${encodeURIComponent(filepath)}`;

  logInfo('OPEN', `opening ${filename} with offline loader`);
  logInfo('OPEN', `document URL: ${documentUrl}`);

  // Redirect to offline loader with parameters
  const redirectParams = new URLSearchParams({
    url: documentUrl,
    title: filename,
    filepath: String(filepath),
    filetype: ext,
    doctype: docType,
  });
  if (lang) {
    redirectParams.set('lang', lang);
  }
  if (theme) {
    redirectParams.set('theme', theme);
  }

  res.redirect(`/offline-loader-proper.html?${redirectParams.toString()}`);
});

// API Endpoint: Serve raw/original file (not converted)
app.get('/raw/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'test', filename);

  logInfo('RAW', `serving original file: ${filename}`);
  logInfo('RAW', `path: ${filePath}`);

  // Check if file exists
  if (!fs.existsSync(filePath)) {
    logWarn('RAW', `file not found: ${filePath}`);
    return res.status(404).send('File not found');
  }

  // Read and send the original file
  const fileData = fs.readFileSync(filePath);
  logInfo('RAW', `sending ${fileData.length} bytes`);

  res.setHeader('Content-Type', 'application/octet-stream');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(fileData);
});

// API Endpoint: Open document with offline loader (proper desktop offline mode)
app.get('/offline/:filename', (req, res) => {
  const filename = req.params.filename;
  const fileExt = path.extname(filename).substring(1);

  // Determine document type based on file extension
  const doctype = getDocTypeFromFilename(filename);

  // Generate a unique key for this document
  const key = Buffer.from(filename + Date.now()).toString('base64').replace(/[^a-zA-Z0-9]/g, '').substring(0, 20);

  // Build query string for offline loader
  const queryParams = new URLSearchParams({
    filename: filename,
    key: key,
    filetype: fileExt,
    doctype: doctype,
    type: 'desktop',
    mode: 'edit'
  });

  logInfo('OFFLINE', `opening ${filename} with offline loader`);
  logInfo('OFFLINE', `file type: ${fileExt}, document type: ${doctype}`);
  logInfo('OFFLINE', `query params: ${queryParams.toString()}`);

  // Redirect to the offline loader HTML with query parameters
  const redirectUrl = `/offline-loader-proper.html?${queryParams.toString()}`;
  logInfo('OFFLINE', `redirecting to: ${redirectUrl}`);

  res.redirect(redirectUrl);
});

// API Endpoint: Edit document with proper editor HTML
app.get('/edit/:filename', (req, res) => {
  const filename = req.params.filename;
  const ext = filename.split('.').pop().toLowerCase();
  const docType = getDocTypeFromFilename(filename);

  logInfo('EDIT', `serving editor for ${filename} (${ext} / ${docType})`);

  res.send(`<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Edit ${filename}</title>
  <style>
    body { margin: 0; padding: 0; overflow: hidden; }
    #placeholder { width: 100vw; height: 100vh; }
  </style>
</head>
<body>
  <div id="placeholder"></div>
  <script src="/web-apps/apps/api/documents/api.js"></script>
  <script>
    console.log('[oo-editors:EDITOR] Initializing ONLYOFFICE editor...');
    new DocsAPI.DocEditor("placeholder", {
      width: "100%",
      height: "100%",
      documentType: "${docType}",
      document: {
        fileType: "${ext}",
        key: "${filename}_" + Date.now(),
        title: "${filename}",
        url: "${BASE_URL}/file/${filename}",
        permissions: {
          edit: true,
          download: true,
          print: true
        }
      },
      editorConfig: {
        mode: "edit",
        callbackUrl: "${BASE_URL}/api/save/${filename}",
        customization: {
          autosave: false,
          chat: false,
          comments: false,
          compactHeader: true,
          help: false,
          logo: {
            visible: true,
            image: "${BASE_URL}/web-apps/apps/common/main/resources/img/header/header-logo_s.svg"
          }
        }
      }
    });
    console.log('[oo-editors:EDITOR] Configuration sent to DocsAPI');
  </script>
</body>
</html>`);
});

// API Endpoint: Serve file for editor
app.get('/file/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'test', filename);

  logInfo('FILE', `serving file: ${filename}`);
  logInfo('FILE', `path: ${filePath}`);

  // Check if file exists
  if (!fs.existsSync(filePath)) {
    logWarn('FILE', `file not found: ${filePath}`);
    return res.status(404).send('File not found');
  }

  // Read and send the file
  const fileData = fs.readFileSync(filePath);
  logInfo('FILE', `sending ${fileData.length} bytes`);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(fileData);
});

// API Endpoint: Serve test page for loading documents (legacy)
app.get('/api/document/:filename', (req, res) => {
  const filename = req.params.filename;
  const fileExt = path.extname(filename).substring(1);

  // Determine editor type based on file extension
  let editorType = 'cell';
  if (fileExt === 'docx' || fileExt === 'doc') {
    editorType = 'word';
  } else if (fileExt === 'pptx' || fileExt === 'ppt') {
    editorType = 'slide';
  }

  // Generate a simple key based on filename
  const key = Buffer.from(filename).toString('base64').replace(/[^a-zA-Z0-9]/g, '').substring(0, 20);

  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>ONLYOFFICE - ${filename}</title>
  <style>
    body { margin: 0; padding: 0; overflow: hidden; }
    #placeholder { width: 100%; height: 100vh; }
  </style>
</head>
<body>
  <div id="placeholder"></div>
  <script type="text/javascript" src="/fonts-info.js"></script>
  <script type="text/javascript" src="/web-apps/apps/api/documents/api.js"></script>
  <script type="text/javascript">
    window.docEditor = new DocsAPI.DocEditor("placeholder", {
      "document": {
        "fileType": "${fileExt}",
        "key": "${key}",
        "title": "${filename}",
        "url": "${BASE_URL}/api/convert/${filename}"
      },
      "documentType": "${editorType}",
      "editorConfig": {
        "mode": "edit",
        "callbackUrl": "${BASE_URL}/api/save/${filename}"
      },
      "width": "100%",
      "height": "100%"
    });

    console.log('[oo-editors:EDITOR] configuration:', {
      url: "${BASE_URL}/api/convert/${filename}",
      fileType: "${fileExt}",
      documentType: "${editorType}",
      key: "${key}"
    });
  </script>
</body>
</html>`;

  res.setHeader('Content-Type', 'text/html');
  res.send(html);
});

// Serve desktop-stub-utils.js from the editors directory
app.get('/desktop-stub-utils.js', (req, res) => {
  const utilsPath = path.join(__dirname, 'editors', 'desktop-stub-utils.js');
  res.setHeader('Content-Type', 'application/javascript');
  sendFileWithHttpErrors(res, utilsPath, {
    logScope: 'DESKTOP-STUB-UTILS',
    notFoundBody: 'Asset not found',
    notFoundLogLevel: 'error',
    internalErrorBody: 'Asset error'
  });
});

// Serve desktop-stub.js from the editors directory
app.get('/desktop-stub.js', (req, res) => {
  const stubPath = path.join(__dirname, 'editors', 'desktop-stub.js');
  res.setHeader('Content-Type', 'application/javascript');
  sendFileWithHttpErrors(res, stubPath, {
    logScope: 'DESKTOP-STUB',
    notFoundBody: 'Asset not found',
    notFoundLogLevel: 'error',
    internalErrorBody: 'Asset error'
  });
});

// Middleware to inject desktop-stub.js into HTML files
app.use((req, res, next) => {
  // Skip injection for static-test directory
  if (req.path.startsWith('/static-test/')) {
    return next();
  }

  // Only process HTML file requests (including .desktop files which are HTML)
  if (req.path.endsWith('.html') || req.path.endsWith('.desktop')) {
    const filePath = path.join(__dirname, 'editors', req.path);

    // Check if file exists
    if (fs.existsSync(filePath)) {
      fs.readFile(filePath, 'utf8', (err, html) => {
        if (err) {
          logError('DESKTOP-STUB', 'error reading HTML file:', err);
          return next();
        }

        // Inject fonts-info.js, desktop-stub-utils.js, and desktop-stub.js before the first <script> tag
        const stubScript = '<script src="/fonts-info.js"></script>\n    <script src="/desktop-stub-utils.js"></script>\n    <script src="/desktop-stub.js"></script>\n    ';

        // Find the first <script> tag and inject before it
        let modifiedHtml = html;
        const scriptTagMatch = html.match(/<script/i);

        if (scriptTagMatch) {
          const insertPosition = scriptTagMatch.index;
          modifiedHtml = html.slice(0, insertPosition) + stubScript + html.slice(insertPosition);
          logInfo('DESKTOP-STUB', `injected fonts-info.js + desktop-stub.js into ${req.path}`);
        } else {
          // If no script tag found, try to inject before </head>
          const headEndMatch = html.match(/<\/head>/i);
          if (headEndMatch) {
            const insertPosition = headEndMatch.index;
            modifiedHtml = html.slice(0, insertPosition) + '  ' + stubScript + html.slice(insertPosition);
            logInfo('DESKTOP-STUB', `injected fonts-info.js + desktop-stub.js into ${req.path} (before </head>)`);
          }
        }

        res.setHeader('Content-Type', 'text/html');
        res.send(modifiedHtml);
      });
    } else {
      next();
    }
  } else {
    next();
  }
});

// WASM files redirect - serve from correct locations regardless of requested path
// The SDK resolves WASM paths relative to the bundle location (e.g., /sdkjs/cell/fonts.wasm)
// but the actual files are in /sdkjs/common/*/engine/ directories
const wasmFiles = {
  'fonts.wasm': 'editors/sdkjs/common/libfont/engine/fonts.wasm',
  'engine.wasm': 'editors/sdkjs/common/hash/hash/engine.wasm',
  'spell.wasm': 'editors/sdkjs/common/spell/spell/spell.wasm',
  'zlib.wasm': 'editors/sdkjs/common/zlib/engine/zlib.wasm',
  'drawingfile.wasm': 'editors/sdkjs/pdf/src/engine/drawingfile.wasm'
};

app.get('*/*.wasm', (req, res, next) => {
  const filename = path.basename(req.path);
  if (wasmFiles[filename]) {
    const wasmPath = path.join(__dirname, wasmFiles[filename]);
    logInfo('WASM', `redirecting ${req.path} to ${wasmPath}`);
    sendFileWithHttpErrors(res, wasmPath, {
      logScope: 'WASM',
      notFoundBody: 'WASM not found',
      notFoundLogLevel: 'error',
      internalErrorBody: 'WASM error'
    });
  } else {
    next();
  }
});

// Serve editors directory as static files (must be first to serve SDK files)
app.use(express.static(path.join(__dirname, 'editors'), {
  setHeaders: (res, filepath) => {
    if (filepath.endsWith('.js')) {
      res.setHeader('Content-Type', 'application/javascript');
    }
  }
}));

// Serve static-test directory (will serve HTML files without desktop stub injection since middleware runs before)
// This works because the middleware only intercepts .html files, and by placing it after express.static,
// the static serving for non-.html files already happened
app.use('/static-test', express.static(path.join(__dirname, 'static-test'), {
  setHeaders: (res, filepath) => {
    if (filepath.endsWith('.js')) {
      res.setHeader('Content-Type', 'application/javascript');
    }
  }
}));

server = app.listen(PORT, () => {
  logInfo('STARTUP', `server running at ${BASE_URL}/ version=${OO_EDITORS_VERSION}`);
  logInfo('STARTUP', 'desktop stub injection enabled for all HTML files');
  logInfo('STARTUP', `static test directory at ${BASE_URL}/static-test/`);
  addLifecycleBreadcrumb('server listening', {
    baseUrl: BASE_URL,
    port: PORT,
    version: OO_EDITORS_VERSION
  }, {
    category: 'oo-editors.startup'
  });
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    logError('STARTUP', `port ${PORT} is already in use`);
  } else {
    logError('STARTUP', 'failed to start HTTP server:', err);
  }
  addLifecycleBreadcrumb('server listen failed', {
    code: err.code,
    message: err.message,
    port: PORT
  }, {
    category: 'oo-editors.startup',
    level: 'error'
  });
  captureLifecycleException(err, {
    level: 'error',
    tags: {
      phase: 'server-listen'
    },
    data: {
      port: PORT
    }
  });
  shutdownServer('server.listen error', 1);
});

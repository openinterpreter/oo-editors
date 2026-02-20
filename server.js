const express = require('express');
const compression = require('compression');
const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const {
  getX2TFormatCode,
  getOutputFormatInfo,
  generateFileHash,
  getDocTypeFromFilename,
  isAbsolutePath,
  getContentType,
  isXLSXSignature
} = require('./server-utils');

if (!process.env.FONT_DATA_DIR) {
  console.error('ERROR: FONT_DATA_DIR environment variable is required');
  process.exit(1);
}

const FONT_DATA_DIR = isAbsolutePath(process.env.FONT_DATA_DIR)
  ? process.env.FONT_DATA_DIR
  : path.join(__dirname, process.env.FONT_DATA_DIR);

const app = express();
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
  console.log(`${new Date().toISOString()} - ${req.method} ${req.url}`);
  next();
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
  console.log(`[FONTS] Request for absolute font path: ${fontPath}`);

  if (!fs.existsSync(fontPath)) {
    console.error(`[FONTS] Absolute font path not found: ${fontPath}`);
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
    res.sendFile(fontPath);
    console.log(`[FONTS] Served font: ${fontPath}`);
  } catch (err) {
    console.error(`[FONTS] Error serving absolute font ${fontPath}:`, err);
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
    console.error('[SW] document_editor_service_worker.js not found at', workerPath);
    return res.status(404).send('Service worker not found');
  }

  res.setHeader('Content-Type', 'application/javascript');
  res.setHeader('Cache-Control', 'no-cache');
  console.log('[SW] Serving document_editor_service_worker.js');
  res.sendFile(workerPath);
});

// Serve the desktop AllFonts.js verbatim for metadata parity
app.get('/fonts-info.js', (req, res) => {
  const allFontsPath = path.join(FONT_DATA_DIR, 'AllFonts.js');
  console.log('[API] GET /fonts-info.js - serving', allFontsPath);
  res.setHeader('Content-Type', 'application/javascript');
  res.setHeader('Cache-Control', 'no-cache');
  res.sendFile(allFontsPath, (err) => {
    if (err) {
      console.error('[API] Error serving AllFonts.js:', err);
      if (!res.headersSent) {
        res.status(err.statusCode || 500).send('// Failed to load font metadata');
      }
    }
  });
});

// Override the SDK's bundled AllFonts.js (Linux-oriented) with the desktop macOS version
app.get('/sdkjs/common/AllFonts.js', (req, res) => {
  const desktopAllFontsPath = path.join(FONT_DATA_DIR, 'AllFonts.js');
  console.log('[API] GET /sdkjs/common/AllFonts.js - overriding with desktop AllFonts.js');
  res.setHeader('Content-Type', 'application/javascript');
  res.setHeader('Cache-Control', 'no-cache');
  res.sendFile(desktopAllFontsPath, (err) => {
    if (err) {
      console.error('[API] Error overriding /sdkjs/common/AllFonts.js:', err);
      if (!res.headersSent) {
        res.status(err.statusCode || 500).send('// Failed to load AllFonts override');
      }
    }
  });
});

// API Endpoint: Upload image to media directory (for drag-drop)
app.post('/api/media/:filehash', (req, res) => {
  const filehash = req.params.filehash;
  const filename = req.query.filename || `image_${Date.now()}.png`;

  console.log(`[MEDIA-UPLOAD] Uploading ${filename} for hash ${filehash}`);

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

  console.log(`[MEDIA-UPLOAD] Saved ${req.body.length} bytes to ${imagePath}`);
  res.json({ filename: filename, path: `/api/media/${filehash}/${filename}` });
});

// API Endpoint: Serve images from converted documents
// Updated to support file-specific media directories
app.get('/api/media/:filehash/:imagefile', (req, res) => {
  const filehash = req.params.filehash;
  const imagefile = req.params.imagefile;
  console.log(`[MEDIA] Request for image: ${imagefile} (file hash: ${filehash})`);

  // Images are in hash-specific output directory
  const outputDir = path.join(__dirname, 'test', 'output', filehash);
  const mediaDir = path.join(outputDir, 'media');
  const imagePath = path.join(mediaDir, imagefile);

  console.log(`[MEDIA] Looking for image at: ${imagePath}`);

  if (!fs.existsSync(imagePath)) {
    console.error(`[MEDIA] Image not found: ${imagePath}`);

    // Check if media directory exists at all
    if (!fs.existsSync(mediaDir)) {
      console.error(`[MEDIA] Media directory does not exist: ${mediaDir}`);
    } else {
      console.log(`[MEDIA] Media directory contents:`, fs.readdirSync(mediaDir));
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

  console.log(`[MEDIA] Serving image: ${imagePath} (${contentType})`);
  res.sendFile(imagePath);
});

// API Endpoint: Serve arbitrary files within the converted document directory (used for relative resource lookups)
app.get('/api/doc-base/:filehash/*', (req, res) => {
  const filehash = req.params.filehash;
  const relativePath = req.params[0] || '';
  console.log(`[DOC-BASE] Request for path "${relativePath}" (file hash: ${filehash})`);

  if (!relativePath) {
    return res.status(400).send('Path is required');
  }

  const outputDir = path.join(__dirname, 'test', 'output', filehash);
  const normalizedPath = path
    .normalize(relativePath)
    .replace(/^(\.\.(\/|\\|$))+/, '');
  const requestedPath = path.join(outputDir, normalizedPath);

  if (!requestedPath.startsWith(outputDir)) {
    console.warn('[DOC-BASE] Attempted directory traversal:', requestedPath);
    return res.status(403).send('Forbidden');
  }

  if (!fs.existsSync(requestedPath) || !fs.statSync(requestedPath).isFile()) {
    console.log('[DOC-BASE] File not found:', requestedPath);
    return res.status(404).send('File not found');
  }

  res.sendFile(requestedPath);
});

// API Endpoint: List images in media directory for a file
app.get('/api/media-list/:filehash', (req, res) => {
  const filehash = req.params.filehash;
  console.log(`[MEDIA-LIST] Request for file hash: ${filehash}`);

  const outputDir = path.join(__dirname, 'test', 'output', filehash);
  const mediaDir = path.join(outputDir, 'media');

  if (!fs.existsSync(mediaDir)) {
    console.log(`[MEDIA-LIST] No media directory found for hash: ${filehash}`);
    return res.json([]);
  }

  try {
    const files = fs.readdirSync(mediaDir);
    console.log(`[MEDIA-LIST] Found ${files.length} files:`, files);
    res.json(files);
  } catch (err) {
    console.error(`[MEDIA-LIST] Error reading media directory:`, err);
    res.status(500).json({ error: 'Error reading media directory' });
  }
});

// API Endpoint: Convert XLSX to ONLYOFFICE binary format
// ONLY supports absolute paths via query parameter
app.get('/api/convert', async (req, res) => {
  const timings = { start: performance.now() };
  const filepath = req.query.filepath;

  if (!filepath) {
    return res.status(400).json({ error: 'filepath query parameter is required' });
  }

  if (!isAbsolutePath(filepath)) {
    return res.status(400).json({ error: 'filepath must be an absolute path' });
  }

  console.log(`[CONVERT] Requested file: ${filepath}`);
  timings.validated = performance.now();

  // Create unique output directory based on the source file path
  const fileHash = generateFileHash(filepath);
  const outputDir = path.join(__dirname, 'test', 'output', fileHash);
  const inputPath = filepath;

  console.log(`[CONVERT] File hash: ${fileHash}`);
  console.log(`[CONVERT] Output directory: ${outputDir}`);

  if (!fs.existsSync(inputPath)) {
    console.error(`[CONVERT] File not found: ${inputPath}`);
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
      console.log(`[CONVERT] Cache hit! Using cached Editor.bin (source: ${new Date(sourceMtime).toISOString()}, cache: ${new Date(cacheMtime).toISOString()})`);
      timings.cacheHit = true;
      const binaryData = fs.readFileSync(outputPath);

      timings.end = performance.now();
      const breakdown = {
        total: (timings.end - timings.start).toFixed(1),
        cacheHit: true,
        validation: (timings.validated - timings.start).toFixed(1),
      };
      console.log(`[CONVERT][TIMING] Cache hit breakdown (ms):`, JSON.stringify(breakdown));

      res.setHeader('Content-Type', 'application/octet-stream');
      res.setHeader('Content-Disposition', `attachment; filename="Editor.bin"`);
      res.setHeader('X-File-Hash', fileHash);
      res.setHeader('X-Timing', JSON.stringify(breakdown));
      res.setHeader('X-Cache', 'HIT');
      return res.send(binaryData);
    }
    console.log(`[CONVERT] Cache stale, reconverting (source: ${new Date(sourceMtime).toISOString()}, cache: ${new Date(cacheMtime).toISOString()})`);
  }

  console.log(`[CONVERT] Converting ${filename} to binary format...`);
  console.log(`[CONVERT] Input: ${inputPath}`);
  console.log(`[CONVERT] Output: ${outputPath}`);

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
<m_sThemeDir>${path.join(__dirname, 'editors', 'sdkjs', 'slide', 'themes')}</m_sThemeDir>
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
  console.log(`[CONVERT] XML config written to ${paramsPath}`);

  // Run x2t converter
  timings.beforeX2t = performance.now();
  const x2tPath = path.join(__dirname, 'converter', 'x2t');
  const x2t = spawn(x2tPath, [paramsPath]);

  let stdout = '';
  let stderr = '';

  x2t.stdout.on('data', (data) => {
    const output = data.toString();
    stdout += output;
    console.log(`[X2T STDOUT] ${output.trim()}`);
  });

  x2t.stderr.on('data', (data) => {
    const output = data.toString();
    stderr += output;
    console.error(`[X2T STDERR] ${output.trim()}`);
  });

  x2t.on('close', (code) => {
    timings.afterX2t = performance.now();
    console.log(`[CONVERT] x2t process exited with code ${code}`);

    // Clean up params file
    try {
      fs.unlinkSync(paramsPath);
    } catch (e) {
      console.warn('[CONVERT] Failed to delete params file:', e.message);
    }

    if (code !== 0) {
      console.error('[CONVERT] Conversion failed!');
      console.error('[CONVERT] stderr:', stderr);
      return res.status(500).send('Conversion failed: ' + stderr);
    }

    // Check if output file was created
    if (!fs.existsSync(outputPath)) {
      console.error('[CONVERT] Output file not created');
      return res.status(500).send('Conversion failed: output file not created');
    }

    // Read and send the binary file
    timings.beforeReadOutput = performance.now();
    console.log(`[CONVERT] Reading output file: ${outputPath}`);
    const binaryData = fs.readFileSync(outputPath);
    timings.afterReadOutput = performance.now();
    console.log(`[CONVERT] Sending ${binaryData.length} bytes`);

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
    console.log(`[CONVERT][TIMING] Breakdown (ms):`, JSON.stringify(breakdown));
    res.setHeader('X-Timing', JSON.stringify(breakdown));

    console.log(`[CONVERT] Sent file hash in header: ${fileHash}`);
    res.send(binaryData);
  });
});

// POST /converter - OnlyOffice Document Server API compatibility endpoint
app.post('/converter', async (req, res) => {
  try {
    console.log('[CONVERTER] OnlyOffice Document Server API request received');

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
      console.log('[CONVERTER] Decoded JWT payload:', { filetype: payload.filetype, outputtype: payload.outputtype });
    } else {
      payload = req.body;
      console.log('[CONVERTER] Using raw payload:', { filetype: payload.filetype, outputtype: payload.outputtype });
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
        console.log('[CONVERTER] Extracted file path from URL:', inputPath);
      } else {
        return res.status(400).json({
          error: -1,
          message: 'Cannot extract file path from URL: ' + inputPath
        });
      }
    }

    console.log(`[CONVERTER] Converting: ${inputPath}`);
    console.log(`[CONVERTER] Format: ${filetype} → ${outputtype}`);

    if (!fs.existsSync(inputPath)) {
      console.error('[CONVERTER] Input file not found:', inputPath);
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

    console.log('[CONVERTER] Output will be:', outputPath);

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
<m_sThemeDir>${path.join(__dirname, 'editors', 'sdkjs', 'slide', 'themes')}</m_sThemeDir>
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
    console.log('[CONVERTER] XML config written');

    const x2tPath = path.join(__dirname, 'converter', 'x2t');
    const x2t = spawn(x2tPath, [paramsPath]);

    let stdout = '';
    let stderr = '';

    x2t.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    x2t.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    x2t.on('close', (code) => {
      console.log(`[CONVERTER] x2t process exited with code ${code}`);

      try {
        fs.unlinkSync(paramsPath);
      } catch (e) {
        console.warn('[CONVERTER] Failed to delete params file:', e.message);
      }

      if (code !== 0) {
        console.error('[CONVERTER] Conversion failed:', stderr);
        return res.status(500).json({
          error: -1,
          message: 'Conversion failed',
          details: stderr
        });
      }

      if (!fs.existsSync(outputPath)) {
        console.error('[CONVERTER] Output file not created');
        return res.status(500).json({
          error: -1,
          message: 'Output file not created'
        });
      }

      const resultUrl = `http://localhost:${PORT}/converted/${outputFilename}`;
      console.log(`[CONVERTER] Conversion successful: ${resultUrl}`);

      res.json({
        url: resultUrl,
        fileType: outputtype,
        error: 0
      });
    });

  } catch (error) {
    console.error('[CONVERTER] Error:', error);
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

  console.log(`[CONVERTED] Request for file: ${filename}`);

  if (!fs.existsSync(filePath)) {
    console.error(`[CONVERTED] File not found: ${filePath}`);
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

  console.log(`[CONVERTED] Serving file: ${filePath} (${contentType})`);
  res.sendFile(filePath);
});

// API Endpoint: Save binary back to XLSX
// ONLY supports absolute paths via query parameter
app.post('/api/save', (req, res) => {
  const filepath = req.query.filepath;
  const filehash = req.query.filehash;

  if (!filepath) {
    return res.status(400).json({ error: 'filepath query parameter is required' });
  }

  if (!isAbsolutePath(filepath)) {
    return res.status(400).json({ error: 'filepath must be an absolute path' });
  }

  console.log(`[SAVE] Requested file: ${filepath}`);
  console.log(`[SAVE] File hash: ${filehash || 'not provided'}`);

  const outputDir = path.join(__dirname, 'test', 'output');
  const outputPath = filepath;

  // Use hash-specific directory so x2t can find media files
  // x2t looks for media/ relative to input binary location
  const hashDir = filehash ? path.join(outputDir, filehash) : outputDir;

  // Ensure parent directory exists
  const parentDir = path.dirname(outputPath);
  if (!fs.existsSync(parentDir)) {
    console.log(`[SAVE] Creating parent directory: ${parentDir}`);
    fs.mkdirSync(parentDir, { recursive: true });
  }

  const filename = path.basename(filepath);

  console.log(`[SAVE] Saving file: ${filename}`);
  console.log(`[SAVE] Output path: ${outputPath}`);
  console.log(`[SAVE] Content-Type: ${req.get('Content-Type')}`);
  console.log(`[SAVE] Body size: ${req.body.length} bytes`);

  // Debug: Check first few bytes
  const firstBytes = req.body.slice(0, 20);
  console.log(`[SAVE] First 20 bytes (hex): ${Buffer.from(firstBytes).toString('hex')}`);
  console.log(`[SAVE] First 20 bytes (ASCII): ${Buffer.from(firstBytes).toString('ascii').replace(/[^\x20-\x7E]/g, '.')}`);

  // Check if this is an XLSX file (should start with PK - ZIP signature)
  const isXLSX = isXLSXSignature(req.body);
  console.log(`[SAVE] Detected format: ${isXLSX ? 'XLSX (ZIP)' : 'Unknown/Binary'}`);

  if (isXLSX) {
    // This is already an XLSX file - just save it directly!
    console.log('[SAVE] File is already XLSX format, saving directly...');
    try {
      fs.writeFileSync(outputPath, req.body);
      console.log(`[SAVE] Successfully saved XLSX file to ${outputPath}`);

      const stats = fs.statSync(outputPath);
      console.log(`[SAVE] File size: ${stats.size} bytes`);

      res.json({ success: true, path: outputPath, size: stats.size });
    } catch (error) {
      console.error('[SAVE] Failed to save XLSX file:', error);
      res.status(500).send('Save failed: ' + error.message);
    }
  } else {
    // This is ONLYOFFICE binary format - convert it to the appropriate output format
    console.log('[SAVE] File appears to be ONLYOFFICE binary format');

    // Determine output format based on file extension
    const ext = path.extname(filepath).toLowerCase();
    const formatInfo = getOutputFormatInfo(ext);

    if (!formatInfo) {
      console.error(`[SAVE] Unsupported file extension: ${ext}`);
      return res.status(400).send('Unsupported file format');
    }

    const { code: formatTo, name: formatName } = formatInfo;
    console.log(`[SAVE] Converting received data to ${formatName}...`);

    const changesBinPath = path.join(hashDir, 'temp_changes.bin');
    const paramsPath = path.join(hashDir, 'params_save.xml');

    // Save the received binary data
    fs.writeFileSync(changesBinPath, req.body);
    console.log(`[SAVE] Wrote binary data: ${changesBinPath}`);

    // Convert the received data (with changes) to the output format
    console.log(`[SAVE] Converting received data (with changes) to ${formatName}...`);
    console.log(`[SAVE] Converting from: ${changesBinPath}`);
    console.log(`[SAVE] Output format: ${formatTo} (${formatName})`);

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
<m_sThemeDir>${path.join(__dirname, 'editors', 'sdkjs', 'slide', 'themes')}</m_sThemeDir>
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

    const x2tPath = path.join(__dirname, 'converter', 'x2t');
    const x2t = spawn(x2tPath, [paramsPath]);

    let stdout = '';
    let stderr = '';

    x2t.on('error', (err) => {
      console.error(`[SAVE] Failed to start x2t: ${err.message}`);
    });

    x2t.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    x2t.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    x2t.on('close', (code) => {
      console.log(`[SAVE] x2t exited with code ${code}`);

      // Clean up temp files
      try {
        if (fs.existsSync(paramsPath)) fs.unlinkSync(paramsPath);
        if (fs.existsSync(changesBinPath)) fs.unlinkSync(changesBinPath);
      } catch (e) {
        console.warn('[SAVE] Cleanup warning:', e.message);
      }

      if (code !== 0) {
        console.error('[SAVE] Conversion failed!');
        return res.status(500).send('Save failed: conversion error');
      }

      if (!fs.existsSync(outputPath)) {
        console.error('[SAVE] Output file not created');
        return res.status(500).send('Save failed: no output file');
      }

      const stats = fs.statSync(outputPath);
      console.log(`[SAVE] Successfully saved to ${outputPath}`);
      console.log(`[SAVE] File size: ${stats.size} bytes`);
      res.json({ success: true, path: outputPath, size: stats.size });
    });
  }
});

// API Endpoint: Load document with simple loader (direct binary loading)
app.get('/load/:filename', (req, res) => {
  const filename = req.params.filename;

  console.log(`[LOAD] Loading ${filename} with simple loader`);

  // Serve the simple-loader.html with filename as query parameter
  const loaderPath = path.join(__dirname, 'editors', 'simple-loader.html');

  fs.readFile(loaderPath, 'utf8', (err, html) => {
    if (err) {
      console.error('[LOAD] Error reading simple-loader.html:', err);
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

  if (!filepath) {
    return res.status(400).json({ error: 'filepath query parameter is required' });
  }

  if (!isAbsolutePath(filepath)) {
    return res.status(400).json({ error: 'filepath must be an absolute path' });
  }

  console.log(`[OPEN] Requested file: ${filepath}`);

  const filename = path.basename(filepath);
  const ext = filename.split('.').pop().toLowerCase();
  const docType = getDocTypeFromFilename(filename);

  // Build the document URL with proper encoding
  const documentUrl = `http://localhost:${PORT}/api/convert?filepath=${encodeURIComponent(filepath)}`;

  console.log(`[OPEN] Opening ${filename} with offline loader`);
  console.log(`[OPEN] Document URL: ${documentUrl}`);

  // Redirect to offline loader with parameters
  res.redirect(`/offline-loader-proper.html?url=${encodeURIComponent(documentUrl)}&title=${encodeURIComponent(filename)}&filepath=${encodeURIComponent(filepath)}&filetype=${ext}&doctype=${docType}`);
});

// API Endpoint: Serve raw/original file (not converted)
app.get('/raw/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'test', filename);

  console.log(`[RAW] Serving original file: ${filename}`);
  console.log(`[RAW] Path: ${filePath}`);

  // Check if file exists
  if (!fs.existsSync(filePath)) {
    console.error(`[RAW] File not found: ${filePath}`);
    return res.status(404).send('File not found');
  }

  // Read and send the original file
  const fileData = fs.readFileSync(filePath);
  console.log(`[RAW] Sending ${fileData.length} bytes`);

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

  console.log(`[OFFLINE] Opening ${filename} with offline loader`);
  console.log(`[OFFLINE] File type: ${fileExt}, Document type: ${doctype}`);
  console.log(`[OFFLINE] Query params: ${queryParams.toString()}`);

  // Redirect to the offline loader HTML with query parameters
  const redirectUrl = `/offline-loader-proper.html?${queryParams.toString()}`;
  console.log(`[OFFLINE] Redirecting to: ${redirectUrl}`);

  res.redirect(redirectUrl);
});

// API Endpoint: Edit document with proper editor HTML
app.get('/edit/:filename', (req, res) => {
  const filename = req.params.filename;
  const ext = filename.split('.').pop().toLowerCase();
  const docType = getDocTypeFromFilename(filename);

  console.log(`[EDIT] Serving editor for ${filename} (${ext} / ${docType})`);

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
    console.log('[EDITOR] Initializing ONLYOFFICE editor...');
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
    console.log('[EDITOR] Configuration sent to DocsAPI');
  </script>
</body>
</html>`);
});

// API Endpoint: Serve file for editor
app.get('/file/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'test', filename);

  console.log(`[FILE] Serving file: ${filename}`);
  console.log(`[FILE] Path: ${filePath}`);

  // Check if file exists
  if (!fs.existsSync(filePath)) {
    console.error(`[FILE] File not found: ${filePath}`);
    return res.status(404).send('File not found');
  }

  // Read and send the file
  const fileData = fs.readFileSync(filePath);
  console.log(`[FILE] Sending ${fileData.length} bytes`);

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

    console.log('Editor configuration:', {
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
  res.sendFile(utilsPath);
});

// Serve desktop-stub.js from the editors directory
app.get('/desktop-stub.js', (req, res) => {
  const stubPath = path.join(__dirname, 'editors', 'desktop-stub.js');
  res.setHeader('Content-Type', 'application/javascript');
  res.sendFile(stubPath);
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
          console.error('Error reading HTML file:', err);
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
          console.log(`Injected fonts-info.js + desktop-stub.js into ${req.path}`);
        } else {
          // If no script tag found, try to inject before </head>
          const headEndMatch = html.match(/<\/head>/i);
          if (headEndMatch) {
            const insertPosition = headEndMatch.index;
            modifiedHtml = html.slice(0, insertPosition) + '  ' + stubScript + html.slice(insertPosition);
            console.log(`Injected fonts-info.js + desktop-stub.js into ${req.path} (before </head>)`);
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
    console.log(`[WASM] Redirecting ${req.path} to ${wasmPath}`);
    res.sendFile(wasmPath);
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

app.listen(PORT, () => {
  console.log(`Server running at ${BASE_URL}/`);
  console.log('Desktop stub injection enabled for all HTML files');
  console.log(`Static test directory at ${BASE_URL}/static-test/`);
});

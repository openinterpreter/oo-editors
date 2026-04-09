/**
 * Server utility functions
 * Extracted for testability
 */

const path = require('path');
const crypto = require('crypto');

/**
 * Get x2t format code from file extension string
 * @param {string} formatString - File extension (xlsx, docx, pptx, pdf, etc.)
 * @returns {number|null} x2t format code or null if unsupported
 */
function getX2TFormatCode(formatString) {
  const formatMap = {
    'xlsx': 257, 'xls': 257, 'ods': 257, 'csv': 260,
    'docx': 65, 'doc': 65, 'odt': 65, 'txt': 65, 'rtf': 65, 'html': 65,
    'pptx': 129, 'ppt': 129, 'odp': 129,
    'pdf': 513
  };
  return formatMap[formatString.toLowerCase()] || null;
}

/**
 * Get x2t format code for output based on file extension
 * @param {string} ext - File extension with leading dot (e.g., '.xlsx')
 * @returns {object} Format info with code and name, or null if unsupported
 */
function getOutputFormatInfo(ext) {
  const normalized = ext.toLowerCase();
  
  if (normalized === '.xlsx' || normalized === '.xls' || normalized === '.ods') {
    return { code: 257, name: 'XLSX' };
  } else if (normalized === '.csv') {
    return { code: 260, name: 'CSV' };
  } else if (normalized === '.docx' || normalized === '.doc' || normalized === '.odt' || normalized === '.txt' || normalized === '.rtf' || normalized === '.html') {
    return { code: 65, name: 'DOCX' };
  } else if (normalized === '.pptx' || normalized === '.ppt' || normalized === '.odp') {
    return { code: 129, name: 'PPTX' };
  }
  return null;
}

/**
 * Generate MD5 hash for file path (used for cache directories)
 * @param {string} filepath - Absolute file path
 * @returns {string} MD5 hash
 */
function generateFileHash(filepath) {
  return crypto.createHash('md5').update(filepath).digest('hex');
}

/**
 * Determine document type from file extension
 * @param {string} filename - Filename with extension
 * @returns {string} Document type: 'cell', 'word', or 'slide'
 */
function getDocTypeFromFilename(filename) {
  const ext = filename.split('.').pop().toLowerCase();
  if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm' || ext === 'ods' || ext === 'csv') {
    return 'cell';
  } else if (ext === 'docx' || ext === 'doc' || ext === 'odt' || ext === 'txt' || ext === 'rtf' || ext === 'html') {
    return 'word';
  } else if (ext === 'pptx' || ext === 'ppt' || ext === 'odp') {
    return 'slide';
  }
  return 'slide';
}

/**
 * Check if a path is absolute (works for both Windows and POSIX)
 * @param {string} filepath - Path to check
 * @returns {boolean} True if absolute
 */
function isAbsolutePath(filepath) {
  return path.win32.isAbsolute(filepath) || path.posix.isAbsolute(filepath);
}

/**
 * Get content type for file extension
 * @param {string} ext - File extension with leading dot
 * @returns {string} MIME content type
 */
function getContentType(ext) {
  const contentTypes = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.bmp': 'image/bmp',
    '.webp': 'image/webp',
    '.emf': 'image/x-emf',
    '.wmf': 'image/x-wmf',
    '.ttf': 'font/ttf',
    '.otf': 'font/otf',
    '.ttc': 'font/collection',
    '.woff': 'font/woff',
    '.woff2': 'font/woff2',
    '.pdf': 'application/pdf',
    '.txt': 'text/plain',
    '.html': 'text/html',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  };
  return contentTypes[ext.toLowerCase()] || 'application/octet-stream';
}

/**
 * Generate XML config for x2t converter
 * @param {object} options - Configuration options
 * @param {string} options.inputPath - Input file path
 * @param {string} options.outputPath - Output file path
 * @param {string} options.filename - Display filename
 * @param {number} options.formatTo - Target format code
 * @param {string} options.fontDir - Font directory path
 * @param {string} options.themeDir - Theme directory path
 * @param {number} [options.formatFrom] - Source format code (optional)
 * @returns {string} XML configuration string
 */
function generateX2TConfig(options) {
  const {
    inputPath,
    outputPath,
    filename,
    formatTo,
    fontDir,
    themeDir,
    formatFrom
  } = options;

  let xml = `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<m_sKey>api_conversion</m_sKey>
<m_sFileFrom>${inputPath}</m_sFileFrom>
<m_sFileTo>${outputPath}</m_sFileTo>
<m_sTitle>${filename}</m_sTitle>
<m_nFormatTo>${formatTo}</m_nFormatTo>`;

  if (formatFrom !== undefined) {
    xml += `\n<m_nFormatFrom>${formatFrom}</m_nFormatFrom>`;
  }

  xml += `
<m_bPaid xsi:nil="true" />
<m_bEmbeddedFonts xsi:nil="true" />
<m_bFromChanges>false</m_bFromChanges>
<m_sFontDir>${fontDir}</m_sFontDir>
<m_sThemeDir>${themeDir}</m_sThemeDir>
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
  return xml;
}

/**
 * Extract file path from OnlyOffice Document Server URL format
 * Example: http://host:port/api/onlyoffice/files/absolute/path/to/file.xlsx
 * @param {string} url - URL to parse
 * @returns {string|null} Extracted file path or null if not matched
 */
function extractFilePathFromUrl(url) {
  if (!url) return null;
  
  if (url.startsWith('http://') || url.startsWith('https://')) {
    try {
      const urlObj = new URL(url);
      const pathMatch = urlObj.pathname.match(/\/api\/onlyoffice\/files\/(.+)/);
      if (pathMatch) {
        return '/' + pathMatch[1].split('/').map(decodeURIComponent).join('/');
      }
    } catch (e) {
      return null;
    }
  }
  
  return null;
}

/**
 * Check if file data starts with XLSX signature (PK - ZIP)
 * @param {Buffer} data - File data buffer
 * @returns {boolean} True if data is XLSX format
 */
function isXLSXSignature(data) {
  if (!data || data.length < 2) return false;
  return data[0] === 0x50 && data[1] === 0x4B;
}

/**
 * Classify an Express sendFile callback error into an HTTP response.
 * @param {Error & {statusCode?: number, status?: number, code?: string}} err
 * @param {object} options
 * @param {string} options.notFoundBody
 * @param {string} options.internalErrorBody
 * @returns {
 *   {action: 'ignore'} |
 *   {action: 'respond', statusCode: number, body: string, logLevel: 'warn' | 'error'}
 * }
 */
function classifySendFileError(err, options) {
  const { notFoundBody, internalErrorBody } = options;
  const statusCode = Number.isInteger(err?.statusCode)
    ? err.statusCode
    : Number.isInteger(err?.status)
      ? err.status
      : null;

  if (err?.code === 'ECONNABORTED') {
    return {
      action: 'ignore'
    };
  }

  if (statusCode === 404 || err?.code === 'ENOENT' || err?.code === 'EISDIR') {
    return {
      action: 'respond',
      statusCode: 404,
      body: notFoundBody,
      logLevel: 'warn'
    };
  }

  return {
    action: 'respond',
    statusCode: 500,
    body: internalErrorBody,
    logLevel: 'error'
  };
}

/**
 * Resolve the x2t binary path for the current platform.
 * On Windows we prefer x2t.exe but tolerate legacy packaging that omits the extension.
 * @param {string} rootDir
 * @param {object} [options]
 * @param {string} [options.platform]
 * @param {(filepath: string) => boolean} [options.pathExists]
 * @returns {{ path: string, candidates: string[] }}
 */
function resolveX2TPath(rootDir, options = {}) {
  const platform = options.platform || process.platform;
  const pathExists = options.pathExists || (() => true);
  const candidates = platform === 'win32'
    ? [
        path.join(rootDir, 'converter', 'x2t.exe'),
        path.join(rootDir, 'converter', 'x2t')
      ]
    : [path.join(rootDir, 'converter', 'x2t')];

  return {
    path: candidates.find((candidate) => pathExists(candidate)) || candidates[0],
    candidates
  };
}

/**
 * Convert a child-process close event into a short human-readable description.
 * @param {number|null|undefined} code
 * @param {NodeJS.Signals|string|null|undefined} signal
 * @returns {string}
 */
function describeChildExit(code, signal) {
  if (signal) {
    return `signal ${signal}`;
  }

  if (Number.isInteger(code)) {
    return `code ${code}`;
  }

  return 'unknown exit status';
}

/**
 * Decide how a completed child process stderr buffer should be logged.
 * We wait until exit so successful tools that write noisy diagnostics to stderr
 * do not get treated like failures, while failed exits still preserve the full
 * stderr buffer in one log entry for debugging/Sentry.
 * @param {{code: number|null, signal: NodeJS.Signals|string|null, stderr: string}} result
 * @returns {{level: 'info' | 'error', output: string} | null}
 */
function getBufferedStderrLog(result) {
  const output = result?.stderr?.trim();

  if (!output) {
    return null;
  }

  if (result.code === 0 && !result.signal) {
    return { level: 'info', output };
  }

  return { level: 'error', output };
}

/**
 * Run a spawned child process and capture its output without leaving unhandled
 * error events behind.
 * @param {(command: string, args: string[]) => import('child_process').ChildProcess} spawnFn
 * @param {string} command
 * @param {string[]} args
 * @param {object} [options]
 * @param {(chunk: string) => void} [options.onStdout]
 * @param {(chunk: string) => void} [options.onStderr]
 * @returns {Promise<{code: number|null, signal: NodeJS.Signals|null, stdout: string, stderr: string}>}
 */
function runChildProcess(spawnFn, command, args, options = {}) {
  const { onStdout, onStderr } = options;

  return new Promise((resolve, reject) => {
    let child;

    try {
      child = spawnFn(command, args);
    } catch (error) {
      reject({ error, stdout: '', stderr: '' });
      return;
    }

    let stdout = '';
    let stderr = '';
    let settled = false;

    const finishResolve = (code, signal) => {
      if (settled) return;
      settled = true;
      resolve({ code, signal, stdout, stderr });
    };

    const finishReject = (error) => {
      if (settled) return;
      settled = true;
      reject({ error, stdout, stderr });
    };

    if (child.stdout && typeof child.stdout.on === 'function') {
      child.stdout.on('data', (data) => {
        const output = data.toString();
        stdout += output;
        if (typeof onStdout === 'function') {
          onStdout(output);
        }
      });
    }

    if (child.stderr && typeof child.stderr.on === 'function') {
      child.stderr.on('data', (data) => {
        const output = data.toString();
        stderr += output;
        if (typeof onStderr === 'function') {
          onStderr(output);
        }
      });
    }

    child.once('error', finishReject);
    child.once('close', finishResolve);
  });
}

module.exports = {
  getX2TFormatCode,
  getOutputFormatInfo,
  generateFileHash,
  getDocTypeFromFilename,
  isAbsolutePath,
  getContentType,
  generateX2TConfig,
  extractFilePathFromUrl,
  isXLSXSignature,
  classifySendFileError,
  resolveX2TPath,
  describeChildExit,
  getBufferedStderrLog,
  runChildProcess
};

const { execFileSync } = require('child_process');

// Cap the port lookup so a wedged lsof/netstat can never freeze startup.
const PORT_LOOKUP_TIMEOUT_MS = 2000;

function isBrokenPipeError(error) {
  const message = error && typeof error.message === 'string' ? error.message : '';
  const mentionsBrokenPipe = /\bEPIPE\b/.test(message);
  return Boolean(error && (
    error.code === 'EPIPE'
    || error.errno === 'EPIPE'
    || mentionsBrokenPipe
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

function createStdIoState(options = {}) {
  const { onBrokenPipeTelemetry } = options;
  const state = {
    stdoutBrokenPipe: false,
    stderrBrokenPipe: false,
    brokenPipeTelemetrySent: false
  };

  function markBrokenPipeStream(streamName, error) {
    if (streamName === 'stdout') {
      state.stdoutBrokenPipe = true;
    } else {
      state.stderrBrokenPipe = true;
    }

    if (state.brokenPipeTelemetrySent) {
      return;
    }

    state.brokenPipeTelemetrySent = true;
    if (typeof onBrokenPipeTelemetry === 'function') {
      onBrokenPipeTelemetry({
        event: 'stdio broken pipe',
        stream: streamName,
        timestamp: new Date().toISOString(),
        pid: process.pid,
        ppid: process.ppid,
        runtime: typeof Bun !== 'undefined' && typeof Bun.version === 'string'
          ? `bun-${Bun.version}`
          : `node-${process.versions.node}`,
        brokenPipeState: { ...state },
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
      });
    }
  }

  function isStreamBroken(streamName) {
    return streamName === 'stdout' ? state.stdoutBrokenPipe : state.stderrBrokenPipe;
  }

  function writeLog(method, streamName, formattedMessage, args = [], consoleLike = console) {
    if (isStreamBroken(streamName)) {
      return false;
    }

    try {
      consoleLike[method](formattedMessage, ...args);
      return true;
    } catch (error) {
      if (isBrokenPipeError(error)) {
        markBrokenPipeStream(streamName, error);
        return false;
      }

      throw error;
    }
  }

  return {
    state,
    markBrokenPipeStream,
    writeLog
  };
}

function handleStdIoStreamError(streamName, error, handlers) {
  const {
    stdioState,
    addBreadcrumb,
    captureException,
    shutdown
  } = handlers;

  if (isBrokenPipeError(error)) {
    stdioState.markBrokenPipeStream(streamName, error);
    shutdown(`${streamName}-epipe`, 0);
    return;
  }

  addBreadcrumb('stdio stream error', {
    stream: streamName,
    message: error && error.message ? error.message : String(error)
  }, {
    category: 'oo-editors.process',
    level: 'error'
  });
  captureException(error, {
    level: 'error',
    tags: {
      phase: 'stdio',
      stream: streamName
    }
  });
  shutdown(`${streamName}-error`, 1);
}

function handleProcessFailure(kind, error, handlers) {
  const {
    addBreadcrumb,
    captureException,
    logError,
    shutdown
  } = handlers;

  if (kind === 'unhandledRejection') {
    addBreadcrumb('unhandled rejection', {
      reason: error instanceof Error ? error.message : String(error)
    }, {
      category: 'oo-editors.process',
      level: 'error'
    });
  } else {
    addBreadcrumb('uncaught exception', {
      message: error && error.message ? error.message : String(error),
      name: error && error.name ? error.name : 'Error'
    }, {
      category: 'oo-editors.process',
      level: 'error'
    });
  }

  captureException(error, {
    level: 'fatal',
    tags: {
      phase: kind
    }
  });
  logError('PROCESS', kind === 'unhandledRejection' ? 'unhandled rejection:' : 'uncaught exception:', error);
  shutdown(kind, 1);
}

function isEnvFlagEnabled(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const normalized = value.trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes';
}

function getOwnPortCleanupExclusions(processInfo = process) {
  return [processInfo.pid, processInfo.ppid].filter((pid) => Number.isInteger(pid) && pid > 0);
}

// Extract the PID(s) LISTENING on the port. macOS lsof (-sTCP:LISTEN) already
// prints one listener PID per line. Windows netstat -ano prints every
// connection, so keep only LISTENING rows whose local address ends with :port
// -- this is what stops us from ever targeting a client that merely has a
// connection open to the port (the desktop app keeps client sockets to us).
function parsePortHolderPids(output, platform, port) {
  const lines = String(output).split('\n').map((line) => line.trim()).filter(Boolean);

  if (platform === 'win32') {
    return lines
      .filter((line) => line.includes('LISTENING'))
      .map((line) => line.split(/\s+/))
      .filter((columns) => (columns[1] ?? '').endsWith(`:${port}`))
      .map((columns) => Number(columns[columns.length - 1]))
      .filter((pid) => Number.isInteger(pid) && pid > 0);
  }

  return lines
    .map((line) => Number(line))
    .filter((pid) => Number.isInteger(pid) && pid > 0);
}

function lookupListenerPids(port, platform, run) {
  const [file, args] = platform === 'win32'
    ? ['netstat', ['-ano']]
    : ['lsof', ['-t', `-iTCP:${port}`, '-sTCP:LISTEN']];

  const output = run(file, args, { encoding: 'utf8', timeout: PORT_LOOKUP_TIMEOUT_MS });
  return parsePortHolderPids(output, platform, port);
}

// Best-effort: free a port by killing whatever process is LISTENING on it.
// Hard contract: this NEVER throws under any input -- it returns the PIDs it
// killed (possibly empty), so a missing/wedged lookup tool, a vanished PID
// (ESRCH), a denied signal (EPERM), or even malformed options can never crash
// the host server. Listener-only targeting plus excluding this process and its
// parent keep it from killing the wrong thing.
function killPortProcess(port, options = {}) {
  try {
    const platform = options.platform ?? process.platform;
    const run = options.exec ?? execFileSync;
    const killPid = options.kill ?? ((pid) => process.kill(pid, 'SIGKILL'));
    const excluded = new Set(
      (options.excludePids ?? getOwnPortCleanupExclusions())
        .filter((pid) => Number.isInteger(pid) && pid > 0)
    );

    const pids = lookupListenerPids(port, platform, run);

    const killed = [];
    for (const pid of new Set(pids)) {
      if (excluded.has(pid)) {
        continue;
      }
      try {
        killPid(pid);
        killed.push(pid);
      } catch {
        // Already gone (ESRCH) or not permitted (EPERM) -- skip the pid but keep
        // killing the rest; the loop must not abandon other holders.
      }
    }
    return killed;
  } catch {
    // No listener (lsof/netstat exit non-zero), a missing tool, a timed-out
    // lookup, or bad options -- degrade to "freed nothing", never propagate.
    return [];
  }
}

// The desktop app frees port 38123 in the Electron process and then spawns this
// server as a separate Node process, so the "port is free" check and the real
// bind happen in different processes a spawn apart. A predecessor's listening
// socket may still be tearing down when this child binds, surfacing a transient
// EADDRINUSE. Retry the bind a bounded number of times before giving up so the
// editor backend rides out that window instead of dying on first conflict.
function startListeningWithRetry(options) {
  const {
    attemptListen,
    onListening,
    onRetry,
    onFatalError,
    scheduleRetry,
    maxRetries
  } = options;

  let listening = false;

  function tryListen(retriesLeft) {
    attemptListen({
      onListening() {
        listening = true;
        onListening();
      },
      onError(error) {
        // Only a port conflict during the bind phase is retryable. Once the
        // server is listening, every error (a late EADDRINUSE included) is fatal.
        if (!listening && retriesLeft > 0 && error.code === 'EADDRINUSE') {
          onRetry({ retriesLeft, error });
          scheduleRetry(() => tryListen(retriesLeft - 1));
          return;
        }

        onFatalError(error);
      }
    });
  }

  tryListen(maxRetries);
}

module.exports = {
  createStdIoState,
  getErrorOwnProperties,
  handleProcessFailure,
  handleStdIoStreamError,
  isBrokenPipeError,
  isEnvFlagEnabled,
  killPortProcess,
  parsePortHolderPids,
  startListeningWithRetry
};

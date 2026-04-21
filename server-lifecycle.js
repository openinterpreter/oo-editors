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

module.exports = {
  createStdIoState,
  getErrorOwnProperties,
  handleProcessFailure,
  handleStdIoStreamError,
  isBrokenPipeError
};

import { describe, test, expect } from 'bun:test';
import { spawn } from 'child_process';
import path from 'path';
import {
  createStdIoState,
  handleProcessFailure,
  handleStdIoStreamError,
  isBrokenPipeError,
  isEnvFlagEnabled,
  killPortProcess,
  parsePortHolderPids,
  startListeningWithRetry
} from '../server-lifecycle.js';

function addrInUseError() {
  const error = new Error('listen EADDRINUSE: address already in use :::38123');
  error.code = 'EADDRINUSE';
  return error;
}

// attemptListen fake that replays a queue of outcomes ('listen' | Error) so each
// bind attempt resolves deterministically without real sockets or timers.
function createReplayListen(outcomes) {
  let attempts = 0;
  function attemptListen({ onListening, onError }) {
    attempts += 1;
    const outcome = outcomes.shift();
    if (outcome === 'listen') {
      onListening();
    } else {
      onError(outcome);
    }
  }
  return { attemptListen, getAttempts: () => attempts };
}

function createBrokenPipeError() {
  const error = new Error('write EPIPE');
  error.code = 'EPIPE';
  error.errno = 'EPIPE';
  return error;
}

function runBrokenPipeFixture(fixturePath, timeoutMs = 2000) {
  return new Promise((resolve, reject) => {
    const child = spawn('node', [fixturePath, 'stdout-broken'], {
      cwd: process.cwd(),
      stdio: ['ignore', 'pipe', 'pipe']
    });
    let stdout = '';
    let stderr = '';
    let closedStdout = false;
    let settled = false;

    const timeout = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`child timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    function closeStdoutAfterFirstLine() {
      if (closedStdout || !stdout.includes('\n')) {
        return;
      }

      closedStdout = true;
      child.stdout.pause();
      child.stdout.destroy();
    }

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString();
      closeStdoutAfterFirstLine();
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString();
    });

    child.on('error', (error) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeout);
      reject(error);
    });

    child.on('close', (code, signal) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeout);
      resolve({ code, signal, stdout, stderr, closedStdout });
    });
  });
}

describe('isBrokenPipeError', () => {
  test('detects common broken-pipe shapes', () => {
    expect(isBrokenPipeError(createBrokenPipeError())).toBe(true);
    expect(isBrokenPipeError({ errno: 'EPIPE' })).toBe(true);
    expect(isBrokenPipeError(new Error('stream failed with EPIPE'))).toBe(true);
  });

  test('ignores unrelated errors', () => {
    expect(isBrokenPipeError(new Error('permission denied'))).toBe(false);
    expect(isBrokenPipeError({ code: 'ENOENT' })).toBe(false);
  });
});

describe('createStdIoState', () => {
  test('suppresses repeated writes after an stdout broken pipe and emits telemetry once', () => {
    const telemetry = [];
    const stdioState = createStdIoState({
      onBrokenPipeTelemetry: (details) => telemetry.push(details)
    });
    const fakeConsole = {
      log() {
        throw createBrokenPipeError();
      }
    };

    expect(stdioState.writeLog('log', 'stdout', 'first', [], fakeConsole)).toBe(false);
    expect(stdioState.writeLog('log', 'stdout', 'second', [], fakeConsole)).toBe(false);
    expect(telemetry).toHaveLength(1);
    expect(stdioState.state).toMatchObject({
      stdoutBrokenPipe: true,
      stderrBrokenPipe: false,
      brokenPipeTelemetrySent: true
    });
  });
});

describe('handleStdIoStreamError', () => {
  test('gracefully shuts down on stdout EPIPE', () => {
    const telemetry = [];
    const shutdownCalls = [];
    const stdioState = createStdIoState({
      onBrokenPipeTelemetry: (details) => telemetry.push(details)
    });

    handleStdIoStreamError('stdout', createBrokenPipeError(), {
      stdioState,
      addBreadcrumb: () => {},
      captureException: () => {},
      shutdown: (reason, code) => shutdownCalls.push({ reason, code })
    });

    expect(shutdownCalls).toEqual([{ reason: 'stdout-epipe', code: 0 }]);
    expect(telemetry).toHaveLength(1);
  });

  test('keeps non-broken-pipe stdio errors fatal', () => {
    const breadcrumbs = [];
    const exceptions = [];
    const shutdownCalls = [];
    const error = new Error('stream failure');

    handleStdIoStreamError('stderr', error, {
      stdioState: createStdIoState(),
      addBreadcrumb: (...args) => breadcrumbs.push(args),
      captureException: (...args) => exceptions.push(args),
      shutdown: (reason, code) => shutdownCalls.push({ reason, code })
    });

    expect(breadcrumbs).toHaveLength(1);
    expect(exceptions).toHaveLength(1);
    expect(shutdownCalls).toEqual([{ reason: 'stderr-error', code: 1 }]);
  });
});

describe('handleProcessFailure', () => {
  test('keeps uncaught EPIPE fatal outside stdio handlers', () => {
    const shutdownCalls = [];
    const exceptions = [];
    const logCalls = [];

    handleProcessFailure('uncaughtException', createBrokenPipeError(), {
      addBreadcrumb: () => {},
      captureException: (...args) => exceptions.push(args),
      logError: (...args) => logCalls.push(args),
      shutdown: (reason, code) => shutdownCalls.push({ reason, code })
    });

    expect(exceptions).toHaveLength(1);
    expect(logCalls).toHaveLength(1);
    expect(shutdownCalls).toEqual([{ reason: 'uncaughtException', code: 1 }]);
  });

  test('keeps unhandledRejection EPIPE fatal outside stdio handlers', () => {
    const shutdownCalls = [];
    const exceptions = [];
    const logCalls = [];

    handleProcessFailure('unhandledRejection', createBrokenPipeError(), {
      addBreadcrumb: () => {},
      captureException: (...args) => exceptions.push(args),
      logError: (...args) => logCalls.push(args),
      shutdown: (reason, code) => shutdownCalls.push({ reason, code })
    });

    expect(exceptions).toHaveLength(1);
    expect(logCalls).toHaveLength(1);
    expect(shutdownCalls).toEqual([{ reason: 'unhandledRejection', code: 1 }]);
  });
});

describe('startListeningWithRetry', () => {
  const runImmediately = (fn) => fn();

  test('rides out a transient EADDRINUSE and binds on a later attempt', () => {
    const retries = [];
    const fatals = [];
    const listened = [];
    const { attemptListen, getAttempts } = createReplayListen([
      addrInUseError(),
      addrInUseError(),
      'listen'
    ]);

    startListeningWithRetry({
      attemptListen,
      onListening: () => listened.push(true),
      onRetry: ({ retriesLeft }) => retries.push(retriesLeft),
      onFatalError: (error) => fatals.push(error),
      scheduleRetry: runImmediately,
      maxRetries: 5
    });

    expect(getAttempts()).toBe(3);
    expect(listened).toHaveLength(1);
    expect(retries).toEqual([5, 4]);
    expect(fatals).toHaveLength(0);
  });

  test('gives up after exhausting retries on a stuck port', () => {
    const fatals = [];
    const listened = [];
    const { attemptListen, getAttempts } = createReplayListen([
      addrInUseError(),
      addrInUseError(),
      addrInUseError()
    ]);

    startListeningWithRetry({
      attemptListen,
      onListening: () => listened.push(true),
      onRetry: () => {},
      onFatalError: (error) => fatals.push(error),
      scheduleRetry: runImmediately,
      maxRetries: 2
    });

    expect(getAttempts()).toBe(3);
    expect(listened).toHaveLength(0);
    expect(fatals).toHaveLength(1);
    expect(fatals[0].code).toBe('EADDRINUSE');
  });

  test('does not retry non-port-conflict listen errors', () => {
    const retries = [];
    const fatals = [];
    const permissionError = new Error('listen EACCES');
    permissionError.code = 'EACCES';
    const { attemptListen, getAttempts } = createReplayListen([permissionError]);

    startListeningWithRetry({
      attemptListen,
      onListening: () => {},
      onRetry: ({ retriesLeft }) => retries.push(retriesLeft),
      onFatalError: (error) => fatals.push(error),
      scheduleRetry: runImmediately,
      maxRetries: 5
    });

    expect(getAttempts()).toBe(1);
    expect(retries).toHaveLength(0);
    expect(fatals).toEqual([permissionError]);
  });

  test('treats a late EADDRINUSE after a successful listen as fatal, never re-listening', () => {
    const fatals = [];
    const retries = [];
    let attempts = 0;
    let capturedOnError = null;

    startListeningWithRetry({
      attemptListen({ onListening, onError }) {
        attempts += 1;
        capturedOnError = onError;
        onListening();
      },
      onListening: () => {},
      onRetry: ({ retriesLeft }) => retries.push(retriesLeft),
      onFatalError: (error) => fatals.push(error),
      scheduleRetry: runImmediately,
      maxRetries: 5
    });

    // The server is already listening; a later runtime EADDRINUSE must not
    // re-enter the bind-retry path (which would call attemptListen again).
    capturedOnError(addrInUseError());

    expect(attempts).toBe(1);
    expect(retries).toHaveLength(0);
    expect(fatals).toHaveLength(1);
    expect(fatals[0].code).toBe('EADDRINUSE');
  });
});

describe('isEnvFlagEnabled', () => {
  test('treats 1/true/yes (case- and space-insensitive) as enabled', () => {
    expect(isEnvFlagEnabled('1')).toBe(true);
    expect(isEnvFlagEnabled('true')).toBe(true);
    expect(isEnvFlagEnabled('YES')).toBe(true);
    expect(isEnvFlagEnabled('  true  ')).toBe(true);
  });

  test('treats everything else as disabled', () => {
    expect(isEnvFlagEnabled('0')).toBe(false);
    expect(isEnvFlagEnabled('false')).toBe(false);
    expect(isEnvFlagEnabled('')).toBe(false);
    expect(isEnvFlagEnabled(undefined)).toBe(false);
  });
});

describe('parsePortHolderPids', () => {
  test('reads one listener PID per line from lsof output', () => {
    expect(parsePortHolderPids('123\n456\n', 'darwin', 38123)).toEqual([123, 456]);
  });

  test('keeps only LISTENING rows for the port from netstat output', () => {
    const netstat = [
      '  TCP    0.0.0.0:38123     0.0.0.0:0         LISTENING    4321',
      '  TCP    [::]:38123        [::]:0            LISTENING    4321',
      '  TCP    10.0.0.5:55012    10.0.0.9:38123    ESTABLISHED  7777', // client -> :38123, must be ignored
      '  TCP    0.0.0.0:5173      0.0.0.0:0         LISTENING    8888'  // different port, must be ignored
    ].join('\n');
    expect(parsePortHolderPids(netstat, 'win32', 38123)).toEqual([4321, 4321]);
  });
});

describe('killPortProcess', () => {
  function recordingKill() {
    const killed = [];
    return { kill: (pid) => killed.push(pid), killed };
  }

  test('kills every PID holding the port (posix)', () => {
    const { kill, killed } = recordingKill();
    const result = killPortProcess(38123, {
      platform: 'darwin',
      exec: () => '123\n456\n',
      kill,
      excludePids: [999]
    });

    expect(killed).toEqual([123, 456]);
    expect(result).toEqual([123, 456]);
  });

  test('never kills excluded PIDs (self/parent)', () => {
    const { kill, killed } = recordingKill();
    const result = killPortProcess(38123, {
      platform: 'darwin',
      exec: () => '123\n456\n',
      kill,
      excludePids: [123]
    });

    expect(killed).toEqual([456]);
    expect(result).toEqual([456]);
  });

  test('returns nothing when no process holds the port', () => {
    const { kill, killed } = recordingKill();
    const result = killPortProcess(38123, {
      platform: 'darwin',
      exec: () => { throw new Error('lsof exit 1'); },
      kill,
      excludePids: []
    });

    expect(killed).toEqual([]);
    expect(result).toEqual([]);
  });

  test('keeps going when killing one PID throws', () => {
    const killed = [];
    const result = killPortProcess(38123, {
      platform: 'darwin',
      exec: () => '123\n456\n',
      kill: (pid) => {
        if (pid === 123) throw new Error('ESRCH');
        killed.push(pid);
      },
      excludePids: []
    });

    expect(killed).toEqual([456]);
    expect(result).toEqual([456]);
  });

  test('kills only the listener, never a client connected to the port (windows)', () => {
    const { kill, killed } = recordingKill();
    const netstat = [
      '  TCP    0.0.0.0:38123    0.0.0.0:0        LISTENING    4321',
      '  TCP    10.0.0.5:55012   10.0.0.9:38123   ESTABLISHED  7777'
    ].join('\n');

    const result = killPortProcess(38123, {
      platform: 'win32',
      exec: () => netstat,
      kill,
      excludePids: []
    });

    expect(killed).toEqual([4321]);
    expect(result).toEqual([4321]);
  });

  test('never throws when the lookup tool throws synchronously', () => {
    let result;
    expect(() => {
      result = killPortProcess(38123, {
        platform: 'darwin',
        exec: () => { throw new Error('ENOENT: lsof missing'); },
        kill: () => {},
        excludePids: []
      });
    }).not.toThrow();
    expect(result).toEqual([]);
  });

  test('never throws even with malformed options', () => {
    let result;
    expect(() => {
      // excludePids is not iterable -- must still degrade to [], not crash.
      result = killPortProcess(38123, { excludePids: 42, exec: () => '123\n' });
    }).not.toThrow();
    expect(result).toEqual([]);
  });
});

describe('broken pipe integration', () => {
  test('exits cleanly after stdout closes', async () => {
    const fixturePath = path.join(process.cwd(), '__tests__', 'fixtures', 'stdio-lifecycle-child.js');
    const result = await runBrokenPipeFixture(fixturePath);
    expect(result.closedStdout).toBe(true);
    expect(result.stdout.startsWith('first\n')).toBe(true);
    expect(result.code).toBe(0);
    expect(result.stderr).toContain('shutdown:stdout-epipe:0');
  });
});

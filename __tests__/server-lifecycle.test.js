import { describe, test, expect } from 'bun:test';
import { spawn } from 'child_process';
import path from 'path';
import {
  createStdIoState,
  handleProcessFailure,
  handleStdIoStreamError,
  isBrokenPipeError,
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

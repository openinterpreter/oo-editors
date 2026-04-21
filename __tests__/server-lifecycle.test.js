import { describe, test, expect } from 'bun:test';
import { spawn } from 'child_process';
import path from 'path';
import {
  createStdIoState,
  handleProcessFailure,
  handleStdIoStreamError,
  isBrokenPipeError
} from '../server-lifecycle.js';

function createBrokenPipeError() {
  const error = new Error('write EPIPE');
  error.code = 'EPIPE';
  error.errno = 'EPIPE';
  return error;
}

function waitForChild(child, timeoutMs = 2000) {
  return new Promise((resolve, reject) => {
    let stdout = '';
    let stderr = '';
    let settled = false;

    const timeout = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`child timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString();
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

    child.on('exit', (code, signal) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeout);
      resolve({ code, signal, stdout, stderr });
    });
  });
}

function createBrokenPipeProbe(fixturePath) {
  if (process.platform === 'win32') {
    const quotedFixturePath = fixturePath.replace(/'/g, "''");
    return {
      command: 'powershell.exe',
      args: [
        '-NoProfile',
        '-Command',
        `node '${quotedFixturePath}' stdout-broken | Select-Object -First 1 | Out-Null; exit $LASTEXITCODE`
      ]
    };
  }

  const shellPath = process.env.SHELL || '/bin/zsh';
  return {
    command: shellPath,
    args: [
      '-lc',
      `set -o pipefail; node ${JSON.stringify(fixturePath)} stdout-broken | head -n 1 >/dev/null`
    ]
  };
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

describe('broken pipe integration', () => {
  test('exits cleanly after stdout closes', async () => {
    const fixturePath = path.join(process.cwd(), '__tests__', 'fixtures', 'stdio-lifecycle-child.js');
    const probe = createBrokenPipeProbe(fixturePath);
    const child = spawn(probe.command, probe.args, {
      cwd: process.cwd(),
      stdio: ['ignore', 'pipe', 'pipe']
    });

    const result = await waitForChild(child);
    expect(result.code).toBe(0);
    expect(result.stderr).toContain('shutdown:stdout-epipe:0');
  });
});

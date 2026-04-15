const { version: OO_EDITORS_VERSION } = require('./package.json');

const OO_EDITOR_SENTRY_DSN = process.env.OO_EDITOR_SENTRY_DSN || '';
const DEFAULT_FLUSH_TIMEOUT_MS = 1500;
const BREADCRUMB_CUTOFF = 1024;
const STACK_CUTOFF = 4096;
const MAX_ARRAY_ITEMS = 20;
const MAX_OBJECT_KEYS = 20;
const MAX_DEPTH = 4;

const runtime = typeof Bun !== 'undefined' && typeof Bun.version === 'string'
  ? `bun-${Bun.version}`
  : `node-${process.versions.node}`;

let loadAttempted = false;
let sentrySdk = null;
let initAttempted = false;
let enabled = false;

function truncate(value, maxLength) {
  if (typeof value !== 'string' || value.length <= maxLength) {
    return value;
  }

  return `${value.slice(0, maxLength - 3)}...`;
}

function sanitizeValue(value, depth = 0) {
  if (value == null || typeof value === 'boolean' || typeof value === 'number') {
    return value;
  }

  if (typeof value === 'string') {
    return truncate(value, BREADCRUMB_CUTOFF);
  }

  if (value instanceof Error) {
    return {
      name: value.name,
      message: truncate(value.message, BREADCRUMB_CUTOFF),
      stack: truncate(value.stack || '', STACK_CUTOFF)
    };
  }

  if (depth >= MAX_DEPTH) {
    return Array.isArray(value) ? '[array]' : '[object]';
  }

  if (Array.isArray(value)) {
    return value.slice(0, MAX_ARRAY_ITEMS).map((item) => sanitizeValue(item, depth + 1));
  }

  if (typeof value === 'object') {
    const sanitized = {};
    for (const key of Object.keys(value).slice(0, MAX_OBJECT_KEYS)) {
      sanitized[key] = sanitizeValue(value[key], depth + 1);
    }
    return sanitized;
  }

  return truncate(String(value), BREADCRUMB_CUTOFF);
}

function loadSentrySdk() {
  if (loadAttempted) {
    return sentrySdk;
  }

  loadAttempted = true;
  const packageName = typeof Bun !== 'undefined' && typeof Bun.version === 'string'
    ? '@sentry/bun'
    : '@sentry/node';

  try {
    sentrySdk = require(packageName);
  } catch (error) {
    console.error(`[oo-editors:SENTRY] failed to load ${packageName}:`, error);
    sentrySdk = null;
  }

  return sentrySdk;
}

function applyScopeContext(scope, context = {}) {
  if (context.level) {
    scope.setLevel(context.level);
  }

  if (context.tags) {
    for (const [key, value] of Object.entries(context.tags)) {
      if (value != null) {
        scope.setTag(key, String(value));
      }
    }
  }

  if (context.fingerprint) {
    scope.setFingerprint(context.fingerprint.map((part) => String(part)));
  }

  if (context.data) {
    scope.setContext('oo_editors', sanitizeValue(context.data));
  }
}

function initOoEditorsSentry() {
  if (initAttempted) {
    return enabled;
  }

  initAttempted = true;
  if (!OO_EDITOR_SENTRY_DSN) {
    return false;
  }

  const Sentry = loadSentrySdk();
  if (!Sentry) {
    return false;
  }

  try {
    Sentry.init({
      dsn: OO_EDITOR_SENTRY_DSN,
      release: `oo-editors@${OO_EDITORS_VERSION}`,
      environment: process.env.NODE_ENV || 'development',
      defaultIntegrations: false,
      disableInstrumentationWarnings: true,
      maxBreadcrumbs: 100,
      shutdownTimeout: DEFAULT_FLUSH_TIMEOUT_MS,
      beforeBreadcrumb(breadcrumb) {
        if (!breadcrumb) {
          return null;
        }

        return {
          ...breadcrumb,
          data: sanitizeValue(breadcrumb.data || {})
        };
      },
      initialScope(scope) {
        scope.setTag('service', 'oo-editors');
        scope.setTag('runtime', runtime);
        scope.setTag('version', OO_EDITORS_VERSION);
        return scope;
      }
    });

    enabled = true;
    addLifecycleBreadcrumb('sentry initialized', {
      runtime,
      release: `oo-editors@${OO_EDITORS_VERSION}`
    });
    return true;
  } catch (error) {
    console.error('[oo-editors:SENTRY] init failed:', error);
    enabled = false;
    return false;
  }
}

function addLifecycleBreadcrumb(message, data = {}, options = {}) {
  if (!enabled) {
    return false;
  }

  const Sentry = loadSentrySdk();
  if (!Sentry) {
    return false;
  }

  Sentry.addBreadcrumb({
    category: options.category || 'oo-editors.lifecycle',
    type: options.type || 'default',
    level: options.level || 'info',
    message,
    data: sanitizeValue(data)
  });

  return true;
}

function captureLifecycleMessage(message, context = {}) {
  if (!enabled) {
    return false;
  }

  const Sentry = loadSentrySdk();
  if (!Sentry) {
    return false;
  }

  Sentry.withScope((scope) => {
    applyScopeContext(scope, context);
    Sentry.captureMessage(message);
  });

  return true;
}

function captureLifecycleException(error, context = {}) {
  if (!enabled) {
    return false;
  }

  const Sentry = loadSentrySdk();
  if (!Sentry) {
    return false;
  }

  Sentry.withScope((scope) => {
    const normalizedError = error instanceof Error ? error : new Error(String(error));
    const details = error instanceof Error
      ? context.data
      : {
          rawError: sanitizeValue(error),
          ...(context.data || {})
        };

    applyScopeContext(scope, {
      ...context,
      data: details
    });
    Sentry.captureException(normalizedError);
  });

  return true;
}

async function flushSentry(timeoutMs = DEFAULT_FLUSH_TIMEOUT_MS) {
  if (!enabled) {
    return true;
  }

  const Sentry = loadSentrySdk();
  if (!Sentry) {
    return false;
  }

  try {
    return await Sentry.flush(timeoutMs);
  } catch (error) {
    console.error('[oo-editors:SENTRY] flush failed:', error);
    return false;
  }
}

function isSentryEnabled() {
  return enabled;
}

module.exports = {
  initOoEditorsSentry,
  isSentryEnabled,
  addLifecycleBreadcrumb,
  captureLifecycleMessage,
  captureLifecycleException,
  flushSentry
};

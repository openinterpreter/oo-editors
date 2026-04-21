const { createStdIoState, handleStdIoStreamError } = require('../../server-lifecycle');

const mode = process.argv[2];
const stdioState = createStdIoState();

function shutdown(reason, code) {
  process.stderr.write(`shutdown:${reason}:${code}\n`);
  process.exit(code);
}

process.stdout.on('error', (error) => {
  handleStdIoStreamError('stdout', error, {
    stdioState,
    addBreadcrumb: () => {},
    captureException: () => {},
    shutdown
  });
});

process.stderr.on('error', (error) => {
  handleStdIoStreamError('stderr', error, {
    stdioState,
    addBreadcrumb: () => {},
    captureException: () => {},
    shutdown
  });
});

if (mode === 'stdout-broken') {
  console.log('first');
  const interval = setInterval(() => {
    console.log(`tick-${Date.now()}`);
  }, 10);

  setTimeout(() => {
    clearInterval(interval);
    process.exit(99);
  }, 500);
}

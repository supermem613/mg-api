#!/usr/bin/env node
'use strict';

const { main } = require('../src/mg-api-core');

main(process.argv.slice(2), {
  stdout: process.stdout,
  stderr: process.stderr,
  exit: code => { process.exitCode = code; },
}).catch(err => {
  process.stdout.write(JSON.stringify({
    ok: false,
    command: 'mg-api',
    data: null,
    error: { code: 'UNHANDLED_ERROR', message: err.message },
    meta: { schemaVersion: '0.1.0' },
  }, null, 2) + '\n');
  process.exitCode = 1;
});

#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');

const AUTH_FILE = path.join(os.homedir(), '.microsoft-graph-skill', 'auth.json');

const GRAPH_SCOPES = 'https://graph.microsoft.com/.default';
const OUTLOOK_SCOPES = 'https://outlook.office.com/.default';

function loadAuthFile() {
  try {
    return JSON.parse(fs.readFileSync(AUTH_FILE, 'utf8'));
  } catch {
    return {};
  }
}

function resolve(envName, fileData) {
  return process.env[envName] || fileData[envName] || '';
}

const fileData = loadAuthFile();

module.exports = {
  GRAPH_TOKEN: resolve('GRAPH_TOKEN', fileData),
  GRAPH_CHAT_TOKEN: resolve('GRAPH_CHAT_TOKEN', fileData),
  OUTLOOK_TOKEN: resolve('OUTLOOK_TOKEN', fileData),
  GRAPH_SCOPES,
  OUTLOOK_SCOPES,
  AUTH_FILE,
};

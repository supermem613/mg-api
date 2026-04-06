#!/usr/bin/env node
'use strict';
// Syncs .github/skills/ → .claude/skills/ (canonical → Claude Code copy)
// Run: node scripts/sync-skills.js [--check]

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, '.github', 'skills');
const DST = path.join(ROOT, '.claude', 'skills');
const CHECK = process.argv.includes('--check');

function getFiles(dir) {
  const results = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...getFiles(full));
    } else {
      results.push(full);
    }
  }
  return results;
}

const srcFiles = getFiles(SRC);
let outOfSync = 0;
let copied = 0;

for (const srcFile of srcFiles) {
  const rel = path.relative(SRC, srcFile);
  const dstFile = path.join(DST, rel);
  const srcContent = fs.readFileSync(srcFile);

  let needsSync = false;
  if (!fs.existsSync(dstFile)) {
    needsSync = true;
  } else {
    const dstContent = fs.readFileSync(dstFile);
    needsSync = !srcContent.equals(dstContent);
  }

  if (!needsSync) {
    console.log(`  ✓ ${rel} (in sync)`);
    continue;
  }

  if (CHECK) {
    console.log(`  ✗ ${rel} (out of sync)`);
    outOfSync++;
  } else {
    fs.mkdirSync(path.dirname(dstFile), { recursive: true });
    fs.copyFileSync(srcFile, dstFile);
    console.log(`  → ${rel} (copied)`);
    copied++;
  }
}

if (CHECK) {
  if (outOfSync > 0) {
    console.log(`\n${outOfSync} file(s) out of sync. Run: node scripts/sync-skills.js`);
    process.exit(1);
  }
  console.log('\nAll skill files in sync.');
} else {
  console.log(`\n${copied} file(s) copied, ${srcFiles.length - copied} already in sync.`);
}

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const html = fs.readFileSync(new URL('../taskpane.html', import.meta.url), 'utf8');
const css = fs.readFileSync(new URL('../taskpane.css', import.meta.url), 'utf8');

test('task pane exposes the compact workbook scope and overflow UI contract', () => {
  for (const id of ['scopeButton', 'scopeMenu', 'overflowButton', 'overflowMenu', 'jobStatus', 'attachButton', 'externalApproval', 'approveOnceButton', 'approveSessionButton', 'denyExternalButton']) {
    assert.match(html, new RegExp(`id=["']${id}["']`));
  }
  assert.match(html, /Context: Workbook/);
  assert.match(html, /Review edits/);
});

test('task pane keeps narrow panes responsive and avoids header button overflow', () => {
  assert.match(css, /@media\s*\(max-width:\s*360px\)/);
  assert.match(css, /\.topbar-actions[^{]*\{[\s\S]*?flex-direction:\s*column/);
  assert.match(css, /\.scope-strip/);
});

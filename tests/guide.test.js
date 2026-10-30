const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");
const guideDir = path.join(root, "guide");
const html = fs.readFileSync(path.join(guideDir, "index.html"), "utf8");

test("見出しの id が6つそろっている", () => {
  for (const id of ["about", "monthly", "annual-fetch", "annual-colors", "settings", "patterns"]) {
    assert.match(html, new RegExp(`id="${id}"`), id);
  }
});

// 拡張のページはCSPでインラインスクリプトが止まる
test("src の無い script がない", () => {
  const scripts = [...html.matchAll(/<script\b([^>]*)>/gi)];
  for (const m of scripts) assert.match(m[1], /\bsrc=/, "inline script: " + m[0]);
});

test("参照するファイルがすべて存在する", () => {
  const refs = [
    ...[...html.matchAll(/<img\b[^>]*\bsrc="([^"]+)"/gi)].map((m) => m[1]),
    ...[...html.matchAll(/<link\b[^>]*\bhref="([^"]+)"/gi)].map((m) => m[1]),
    ...[...html.matchAll(/<script\b[^>]*\bsrc="([^"]+)"/gi)].map((m) => m[1]),
  ];
  assert.ok(refs.length > 0);
  const hasKagkag = fs.existsSync(path.join(root, "kagkag"));
  for (const ref of refs) {
    assert.doesNotMatch(ref, /^https?:/, "外部参照: " + ref);
    if (ref.startsWith("../kagkag/") && !hasKagkag) continue;
    assert.ok(fs.existsSync(path.join(guideDir, ref)), "ファイルが無い: " + ref);
  }
});

test("本文に 2.6.0 が無い", () => {
  assert.doesNotMatch(html, /2\.6\.0/);
});

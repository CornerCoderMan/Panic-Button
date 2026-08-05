"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { PUBLIC_FILES, contentTypeFor } = require("../server.js");

test("serves the task pane at the site root", () => {
  assert.equal(PUBLIC_FILES.get("/"), "taskpane.html");
  assert.equal(PUBLIC_FILES.get("/taskpane.js"), "taskpane.js");
});

test("does not expose any backend API route (add-in is client-side only)", () => {
  assert.equal(PUBLIC_FILES.has("/api/analyze"), false);
});

test("maps file extensions to the expected content types", () => {
  assert.equal(contentTypeFor("taskpane.html"), "text/html; charset=utf-8");
  assert.equal(contentTypeFor("taskpane.js"), "text/javascript; charset=utf-8");
  assert.equal(contentTypeFor("assets/icon-80.png"), "image/png");
  assert.equal(contentTypeFor("mystery.bin"), "application/octet-stream");
});

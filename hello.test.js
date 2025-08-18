const assert = require("assert");
const { hello } = require("./index.js");

// Ein ganz einfacher Test
assert.strictEqual(hello("Welt"), "Hallo, Welt! Willkommen bei meinem ersten Codex-Projekt.");

console.log("✅ Alle Tests erfolgreich!");

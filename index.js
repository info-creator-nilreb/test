function hello(name) {
  return `Hallo, ${name}! Willkommen bei meinem ersten Codex-Projekt.`;
}

// Testausgabe
console.log(hello("Welt"));

// Export für Tests
module.exports = { hello };

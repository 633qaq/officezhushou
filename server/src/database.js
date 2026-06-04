const fs = require('fs');
const path = require('path');
const initSqlJs = require('sql.js');
const config = require('./config');
const logger = require('./utils/logger');

let db = null;

async function initDatabase() {
  const dbPath = config.db.path;
  const dataDir = path.dirname(dbPath);

  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
  }

  const SQL = await initSqlJs();
  logger.info('sql.js engine loaded');

  if (fs.existsSync(dbPath)) {
    db = new SQL.Database(fs.readFileSync(dbPath));
    logger.info(`Loaded database from ${dbPath}`);
  } else {
    db = new SQL.Database();
    logger.info(`Created new database at ${dbPath}`);
  }

  db.run(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      display_name TEXT DEFAULT '',
      role TEXT DEFAULT 'user' CHECK(role IN ('user', 'admin')),
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS user_settings (
      user_id TEXT PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
      provider TEXT DEFAULT 'gemini',
      endpoint TEXT DEFAULT '',
      model TEXT DEFAULT '',
      style TEXT DEFAULT 'business',
      settings TEXT DEFAULT '{}',
      updated_at TEXT DEFAULT (datetime('now'))
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS chat_history (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      role TEXT NOT NULL CHECK(role IN ('user', 'assistant', 'system')),
      content TEXT NOT NULL,
      provider TEXT DEFAULT '',
      model TEXT DEFAULT '',
      metadata TEXT DEFAULT '{}',
      created_at TEXT DEFAULT (datetime('now'))
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS documents (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      title TEXT NOT NULL DEFAULT 'Untitled document',
      type TEXT DEFAULT 'text' CHECK(type IN ('text', 'outline', 'slides', 'docx', 'pptx')),
      content TEXT DEFAULT '',
      summary TEXT DEFAULT '',
      tags TEXT DEFAULT '[]',
      source TEXT DEFAULT 'manual',
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS sessions (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      token TEXT UNIQUE NOT NULL,
      ip TEXT DEFAULT '',
      user_agent TEXT DEFAULT '',
      expires_at TEXT NOT NULL,
      created_at TEXT DEFAULT (datetime('now'))
    )
  `);

  db.run('CREATE INDEX IF NOT EXISTS idx_chat_user ON chat_history(user_id, created_at)');
  db.run('CREATE INDEX IF NOT EXISTS idx_doc_user ON documents(user_id, updated_at)');
  db.run('CREATE INDEX IF NOT EXISTS idx_sessions_token ON sessions(token)');
  db.run('CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id)');

  saveDatabase();
  return db;
}

function saveDatabase() {
  if (!db) {
    return;
  }
  fs.writeFileSync(config.db.path, Buffer.from(db.export()));
}

function getDb() {
  if (!db) {
    throw new Error('Database has not been initialized');
  }
  return db;
}

function queryAll(sql, params = []) {
  const statement = getDb().prepare(sql);
  if (params.length > 0) {
    statement.bind(params);
  }

  const rows = [];
  while (statement.step()) {
    rows.push(statement.getAsObject());
  }
  statement.free();
  return rows;
}

function queryOne(sql, params = []) {
  return queryAll(sql, params)[0] || null;
}

function execute(sql, params = []) {
  getDb().run(sql, params);
  const changes = getDb().getRowsModified();
  saveDatabase();
  return { changes };
}

function closeDatabase() {
  if (!db) {
    return;
  }

  saveDatabase();
  db.close();
  db = null;
}

module.exports = {
  initDatabase,
  saveDatabase,
  getDb,
  queryAll,
  queryOne,
  execute,
  closeDatabase,
};

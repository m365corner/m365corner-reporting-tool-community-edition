
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

// Determine correct DB path based on packaging
const isPkg = typeof process.pkg !== 'undefined';
const dbDir = isPkg ? path.dirname(process.execPath) : __dirname;

const dbFileName = 'mocha.db';
const dbPath = path.join(dbDir, dbFileName);

// Log resolved path to verify correctness
console.log(`[db.js] Using database at: ${dbPath}`);

// Check if DB exists
if (!fs.existsSync(dbPath)) {
    console.error(`[db.js] ❌ Database file not found at: ${dbPath}`);
} else {
    console.log(`[db.js] ✅ Database file found.`);
}

// Initialize DB
const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error(`[db.js] ❌ Failed to open DB:`, err.message);
    } else {
        console.log(`[db.js] ✅ Database opened successfully.`);
    }
});

// Table creation logic
const initQueries = [
    `CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY,
    displayName TEXT,
    userPrincipalName TEXT,
    email TEXT,
    licenseStatus TEXT,
    department TEXT,
    jobTitle TEXT,
    firstName TEXT,
    lastName TEXT,
    mailNickName TEXT,
    userAddedTime TEXT,
    alternateEmail TEXT,
    officePhone TEXT,
    mobilePhone TEXT,
    faxNumber TEXT,
    city TEXT,
    country TEXT,
    postalCode TEXT,
    state TEXT,
    streetAddress TEXT,
    companyName TEXT,
    usageLocation TEXT,
    office TEXT,
    signInStatus TEXT,
    mfaStatus TEXT,
    preferredLanguage TEXT,
    directReportsCount INTEGER
  )`,

    `CREATE TABLE IF NOT EXISTS user_admin_roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT,
    role_name TEXT,
    FOREIGN KEY (user_id) REFERENCES users(id)
  )`,

    `CREATE TABLE IF NOT EXISTS delta_tokens (
    resource TEXT PRIMARY KEY,
    delta_link TEXT,
    last_synced_at TEXT
  )`,

    `CREATE TABLE IF NOT EXISTS sync_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    resource TEXT,
    synced_at TEXT,
    added INTEGER,
    updated INTEGER,
    deleted INTEGER,
    status TEXT,
    error_message TEXT
  )`,

    /* === Groups (existing) === */
    `CREATE TABLE IF NOT EXISTS groups (
    id TEXT PRIMARY KEY,
    displayName TEXT,
    createdDateTime TEXT,
    groupTypes TEXT,
    mail TEXT,
    visibility TEXT,
    securityEnabled INTEGER,
    mailEnabled INTEGER,
    membersCount INTEGER DEFAULT 0,
    ownersCount INTEGER DEFAULT 0
  )`,

    `CREATE TABLE IF NOT EXISTS group_owners (
    id TEXT PRIMARY KEY,
    groupId TEXT,
    userId TEXT,
    displayName TEXT,
    userPrincipalName TEXT,
    department TEXT,
    jobTitle TEXT,
    signInStatus TEXT,
    FOREIGN KEY(groupId) REFERENCES groups(id)
  )`,

    `CREATE TABLE IF NOT EXISTS group_members (
    id TEXT PRIMARY KEY,
    groupId TEXT,
    userId TEXT,
    displayName TEXT,
    userPrincipalName TEXT,
    department TEXT,
    jobTitle TEXT,
    signInStatus TEXT,
    FOREIGN KEY(groupId) REFERENCES groups(id)
  )`,

    /* === Teams (new) === */
    `CREATE TABLE IF NOT EXISTS teams (
    id TEXT PRIMARY KEY,
    displayName TEXT,
    description TEXT,
    visibility TEXT,
    isArchived INTEGER,
    createdDateTime TEXT,
    ownersCount INTEGER DEFAULT 0,
    membersCount INTEGER DEFAULT 0,
    privateChannelsCount INTEGER DEFAULT 0,
    standardChannelsCount INTEGER DEFAULT 0,
    sharedChannelsCount INTEGER DEFAULT 0
  )`,

    `CREATE TABLE IF NOT EXISTS team_members (
    id TEXT PRIMARY KEY,
    teamId TEXT,
    userId TEXT,
    displayName TEXT,
    userPrincipalName TEXT,
    department TEXT,
    jobTitle TEXT,
    signInStatus TEXT,
    role TEXT,
    FOREIGN KEY(teamId) REFERENCES teams(id)
  )`,

    `CREATE TABLE IF NOT EXISTS team_owners (
    id TEXT PRIMARY KEY,
    teamId TEXT,
    userId TEXT,
    displayName TEXT,
    userPrincipalName TEXT,
    department TEXT,
    jobTitle TEXT,
    signInStatus TEXT,
    FOREIGN KEY(teamId) REFERENCES teams(id)
  )`
];

initQueries.forEach((query) => {
    db.run(query, (err) => {
        if (err) console.error(`[db.js] ❌ Table init error:`, err.message);
    });
});

module.exports = db;

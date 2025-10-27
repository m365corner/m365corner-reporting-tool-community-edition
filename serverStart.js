const fs = require("fs");
const path = require("path");

require("dotenv").config();
const express = require("express");
const { getTokenAppOnly } = require("./auth");
const reportRoutes = require("./routes/reportRoutes");
const { exec } = require("child_process");
const { syncUsers } = require("./syncUsers");
const { syncGroups } = require('./syncGroups');
const { syncTeams } = require('./syncTeams');

const cors = require("cors");

logStep("STEP: Starting serverStart.js");

// Detect if running inside pkg
const isPkg = typeof process.pkg !== "undefined";

// Runtime paths based on pkg status
const baseDir = isPkg ? path.dirname(process.execPath) : __dirname;

const dbFile = path.join(baseDir, "mocha.db");
const publicDir = path.join(baseDir, "public");

// Validate dependencies
if (!fs.existsSync(dbFile)) {
  logError("Missing database file", new Error(`DB not found at ${dbFile}`));
  process.exit(1);
}
if (!fs.existsSync(publicDir)) {
  logError("Missing public folder", new Error(`Public folder not found at ${publicDir}`));
  process.exit(1);
}

// Dynamic DB init AFTER validation
require("./db");

const app = express();
const PORT = 3001;
const launchUrl = `http://localhost:${PORT}`;
global.authToken = null;

// Middleware
app.use(express.json({ limit: "2gb" }));
app.use(cors());
app.use(reportRoutes);

// Static files
app.use(express.static(publicDir));

// Homepage
app.get("/", (req, res) => {
  res.sendFile(path.join(publicDir, "index.html"));
});

// Login
app.get("/login", async (req, res) => {
  try {
    if (!global.authToken) {
      const token = await getTokenAppOnly();
      global.authToken = token;
    }
    res.json({ status: "success", message: global.authToken });
  } catch (err) {
    logError("Login failed", err);
    res.status(500).json({ status: "error", message: "Login failed" });
  }
});

// Sync users
app.get("/sync", async (req, res) => {
  try {
    if (!global.authToken) {
      return res.status(401).json({ message: "Not authenticated. Please login." });
    }
    const summary = await syncUsers(global.authToken);
    res.json(summary);
  } catch (error) {
    logError("❌ Sync route failed", error);
    res.status(500).json({ message: `Sync failed: ${error.message}` });
  }
});


// Sync groups
app.get("/syncGroups", async (req, res) => {
    try {
        if (!global.authToken) {
            return res.status(401).json({ message: "Not authenticated. Please login." });
        }
        const summary = await syncGroups(global.authToken);
        res.json(summary);
    } catch (error) {
        logError("❌ Sync route failed", error);
        res.status(500).json({ message: `Sync failed: ${error.message}` });
    }
});




app.get('/sync-groups', async (req, res) => {
  try {
    console.log("[/sync-groups] 🔁 Group sync started...");
    await syncGroups();
    res.status(200).send("✅ Group sync completed.");
  } catch (err) {
    console.error("[/sync-groups] ❌ Error:", err.message);
    res.status(500).send("❌ Sync failed.");
  }
});

app.get('/sync-teams', async (req, res) => {
    try {
        console.log('[/sync-teams] 🔁 Teams sync started...');
        await syncTeams();
        res.status(200).send('✅ Teams sync completed.');
    } catch (e) {
        console.error('[/sync-teams] ❌ Error:', e.message);
        res.status(500).send('❌ Teams sync failed.');
    }
});



// Shutdown
app.post("/shutdown", (req, res) => {
  res.sendStatus(200);
  setTimeout(() => process.exit(0), 100);
});

// Start server
app.listen(PORT, () => {
  const msg = `🚀 Server running on ${launchUrl}`;
  console.log(msg);
  logMessage(msg);
  openBrowser(launchUrl);
});

// Open browser
function openBrowser(url) {
  try {
    const command =
      process.platform === "win32"
        ? `start ${url}`
        : process.platform === "darwin"
        ? `open ${url}`
        : `xdg-open ${url}`;

    exec(command, { windowsHide: true }, (err) => {
      if (err) logError("Error launching browser", err);
    });
  } catch (err) {
    logError("Exception in openBrowser()", err);
  }
}

// Logging
function logError(prefix, error) {
  const log = `[${new Date().toISOString()}] ${prefix}: ${error?.stack || error?.message || error}\n`;
  try {
    fs.appendFileSync("error.log", log);
  } catch (e) {
    console.error("⚠ Failed to write to error.log", e);
  }
  console.error(log);
}

function logMessage(message) {
  const log = `[${new Date().toISOString()}] ${message}\n`;
  try {
    fs.appendFileSync("error.log", log);
  } catch (e) {
    console.error("⚠ Failed to write to error.log", e);
  }
}

function logStep(step) {
  logMessage(step);
}

// Fail-safe logging
process.on("uncaughtException", (err) => logError("Uncaught Exception", err));
process.on("unhandledRejection", (reason) => logError("Unhandled Rejection", reason));

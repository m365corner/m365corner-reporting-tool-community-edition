const fs = require("fs");
const path = require("path");

function log(message) {
    const logEntry = `[${new Date().toISOString()}] ${message}\n`;
    try {
        fs.appendFileSync("error.log", logEntry);
    } catch (err) {
        console.error("⚠ Failed to write to error.log", err);
    }
    console.log(logEntry);
}

function main() {
    log("server.js booted");

    try {
        // Instead of full path, just require the embedded module
        require("./serverStart");
    } catch (err) {
        log(`Error loading serverStart: ${err.stack || err.message}`);
        process.exit(1);
    }
}

main();

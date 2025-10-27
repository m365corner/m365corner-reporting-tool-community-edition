const express = require('express');
const router = express.Router();
const db = require('../db');
const { buildQuery } = require('../utils/queryBuilder');
const { getTokenAppOnly } = require('../auth');
const { exportToCSVFile, exportToCSVString } = require('../utils/csvExporter');
const { sendReportByEmail } = require('../utils/mailer');



// POST /auth/save-credentials
router.post('/auth/save-credentials', (req, res) => {
    const { tenantId, clientId, clientSecret } = req.body;
    if (!tenantId || !clientId || !clientSecret) {
        return res.status(400).json({ success: false, error: "Missing values" });
    }

    global.tenantId = tenantId;
    global.clientId = clientId;
    global.clientSecret = clientSecret;

    console.log("🔐 Credentials updated:", { tenantId, clientId });
    return res.json({ success: true });
});






// Middleware to ensure token for all /report routes
async function ensureToken(req, res, next) {
  try {
    const token = await getTokenAppOnly();
    if (!token) return res.status(401).json({ error: 'Unauthorized: Token not acquired' });
    req.token = token;
    next();
  } catch (err) {
    console.error('❌ Token error:', err);
    return res.status(401).json({ error: 'Unauthorized: Token error' });
  }
}

router.use('/report', ensureToken);

// Generic query execution helper
function runQueryAsync(sql, values = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, values, (err, rows) => {
      if (err) reject(err);
      else resolve(rows);
    });
  });
}

function runGetAsync(sql, values = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, values, (err, row) => {
      if (err) reject(err);
      else resolve(row);
    });
  });
}



// ALL USERS -- NEW
router.get('/report/all', async (req, res) => {
    try {
        const { whereClause, values } = buildQuery(req.query, {
            searchFields: ['displayName', 'userPrincipalName', 'firstName', 'lastName', 'mailNickName', 'email'],
            filters: {
                department: 'department',
                jobTitle: 'jobTitle',
                signInStatus: 'signInStatus',
                licenseStatus: 'licenseStatus'
            }
        });

        const isFullExport = req.query.page === "all";
        const limit = isFullExport ? null : 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const selectSQL = `
      SELECT displayName, userPrincipalName, signInStatus, licenseStatus, department,
             jobTitle, city, state, country, firstName, lastName, mailNickName, email
      FROM users
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM users ${whereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, values)
            : await runQueryAsync(selectSQL, [...values, limit, offset]);

        const totalRow = await runGetAsync(countSQL, values);
        const total = totalRow.total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });
    } catch (err) {
        console.error("❌ Query failed:", err);
        res.status(500).json({ error: 'Query failed' });
    }
});




// DISABLED USERS NEW
router.get('/report/disabled', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = parseInt(req.query.page) || 1;
        const offset = (page - 1) * limit;

        // Enforce Disabled filter
        const mergedQuery = { ...req.query, signInStatus: 'Disabled' };

        const { whereClause, values } = buildQuery(mergedQuery, {
            searchFields: ['displayName', 'userPrincipalName', 'firstName', 'lastName', 'mailNickName', 'email'],
            filters: {
                department: 'department',
                jobTitle: 'jobTitle',
                licenseStatus: 'licenseStatus',
                signInStatus: 'signInStatus' // required for injection
            }
        });

        const selectSQL = `
      SELECT displayName, userPrincipalName, signInStatus, licenseStatus, department,
             jobTitle, city, state, country, firstName, lastName, mailNickName, email
      FROM users
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM users ${whereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, values)
            : await runQueryAsync(selectSQL, [...values, limit, offset]);

        const total = (await runGetAsync(countSQL, values)).total;

        res.json({
            page,
            totalPages: Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });
    } catch (error) {
        console.error("❌ Failed to fetch disabled users:", error);
        res.status(500).send("Internal server error.");
    }
});




// ENABLED USERS -- NEW
router.get('/report/enabled', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = parseInt(req.query.page) || 1;
        const offset = (page - 1) * limit;

        // Enforce Enabled filter
        const mergedQuery = { ...req.query, signInStatus: 'Enabled' };

        const { whereClause, values } = buildQuery(mergedQuery, {
            searchFields: ['displayName', 'userPrincipalName', 'firstName', 'lastName', 'mailNickName', 'email'],
            filters: {
                department: 'department',
                jobTitle: 'jobTitle',
                licenseStatus: 'licenseStatus',
                signInStatus: 'signInStatus' // required for injection
            }
        });

        const selectSQL = `
      SELECT displayName, userPrincipalName, signInStatus, licenseStatus, department,
             jobTitle, city, state, country, firstName, lastName, mailNickName, email
      FROM users
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM users ${whereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, values)
            : await runQueryAsync(selectSQL, [...values, limit, offset]);

        const total = (await runGetAsync(countSQL, values)).total;

        res.json({
            page,
            totalPages: Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });
    } catch (error) {
        console.error("❌ Failed to fetch enabled users:", error);
        res.status(500).send("Internal server error.");
    }
});





// LICENSED USERS -- NEW
router.get('/report/licensed', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = parseInt(req.query.page) || 1;
        const offset = (page - 1) * limit;

        // Enforce Licensed filter
        const mergedQuery = { ...req.query, licenseStatus: 'Licensed' };

        const { whereClause, values } = buildQuery(mergedQuery, {
            searchFields: ['displayName', 'userPrincipalName', 'firstName', 'lastName', 'mailNickName', 'email'],
            filters: {
                department: 'department',
                jobTitle: 'jobTitle',
                signInStatus: 'signInStatus',
                licenseStatus: 'licenseStatus'
            }
        });

        const selectSQL = `
      SELECT displayName, userPrincipalName, signInStatus, department, jobTitle, licenseStatus,
             city, state, country, firstName, lastName, mailNickName, email
      FROM users
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM users ${whereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, values)
            : await runQueryAsync(selectSQL, [...values, limit, offset]);

        const total = (await runGetAsync(countSQL, values)).total;

        res.json({
            page,
            totalPages: Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });
    } catch (error) {
        console.error("❌ Failed to fetch licensed users:", error);
        res.status(500).send("Internal server error.");
    }
});





// UNLICENSED USERS -- NEW
router.get('/report/unlicensed', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = parseInt(req.query.page) || 1;
        const offset = (page - 1) * limit;

        // Force filter for licenseStatus: 'Unlicensed'
        const mergedQuery = { ...req.query, licenseStatus: 'Unlicensed' };

        const { whereClause, values } = buildQuery(mergedQuery, {
            searchFields: ['displayName', 'userPrincipalName', 'firstName', 'lastName', 'mailNickName', 'email'],
            filters: {
                department: 'department',
                jobTitle: 'jobTitle',
                signInStatus: 'signInStatus',
                licenseStatus: 'licenseStatus'
            }
        });

        const selectSQL = `
      SELECT displayName, userPrincipalName, department, licenseStatus, jobTitle,
             city, state, country, firstName, lastName, mailNickName, email, signInStatus
      FROM users
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM users ${whereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, values)
            : await runQueryAsync(selectSQL, [...values, limit, offset]);

        const total = (await runGetAsync(countSQL, values)).total;

        res.json({
            page,
            totalPages: Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });
    } catch (error) {
        console.error("❌ Failed to fetch unlicensed users:", error);
        res.status(500).send("Internal server error.");
    }
});


// DOWNLOAD CSV
router.post('/report/download', async (req, res) => {
  try {
    const data = req.body?.data;

    const users = Array.isArray(data) && data.length > 0
      ? data
      : await runQueryAsync('SELECT * FROM users');

    if (!users || users.length === 0) {
      return res.status(400).send("No data available to generate CSV file.");
    }

    const filePath = await exportToCSVFile(users, 'all_users_report');
    res.download(filePath);
  } catch (error) {
    console.error("❌ Download failed:", error);
    res.status(500).send("Error generating CSV file.");
  }
});

// EMAIL CSV
router.post('/report/email', async (req, res) => {
  const { data, recipient } = req.body;

  if (!recipient) {
    return res.status(400).json({ status: "error", message: "Recipient email is required." });
  }

  try {
    const users = Array.isArray(data) && data.length > 0
      ? data
      : await runQueryAsync("SELECT * FROM users");

    if (!users || users.length === 0) {
      return res.status(400).json({ status: "error", message: "No user data available." });
    }

    const { csv, filename } = await exportToCSVString(users, 'all_users_report');

    if (!csv || typeof csv !== 'string') {
      throw new Error("Invalid CSV string generated.");
    }

    await sendReportByEmail(recipient, csv, filename);
    res.json({ status: "success", message: "Report sent successfully!" });
  } catch (err) {
    console.error("❌ Email failed:", err.message);
    res.status(500).json({ status: "error", message: "Failed to send report" });
  }
});

/******************************* GROUPS RELATED REPORTS ***********************************/



router.get('/report/groups/all', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [];
        const values = [];

        // Search support
        if (req.query.search) {
            const s = req.query.search.trim();
            const isGuid = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(s);

            if (isGuid) {
                // Exact match on id for dropdown-driven searches
                whereParts.push(`id = ?`);
                values.push(s);
            } else {
                // Fuzzy match for manual text searches
                whereParts.push(`(displayName LIKE ? OR mail LIKE ? OR id LIKE ?)`);
                values.push(`%${s}%`, `%${s}%`, `%${s}%`);
            }
        }

        // Filter: groupTypes (case-insensitive)
        if (req.query.groupTypes) {
            whereParts.push(`LOWER(groupTypes) = LOWER(?)`);
            values.push(req.query.groupTypes.trim());
        }

        // Filter: securityEnabled
        if (req.query.securityEnabled) {
            const boolValue = req.query.securityEnabled === 'true' ? 1 : 0;
            whereParts.push(`securityEnabled = ?`);
            values.push(boolValue);
        }

        // Filter: mailEnabled
        if (req.query.mailEnabled) {
            const boolValue = req.query.mailEnabled === 'true' ? 1 : 0;
            whereParts.push(`mailEnabled = ?`);
            values.push(boolValue);
        }

        const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

        const selectSQL = `
      SELECT displayName, createdDateTime, groupTypes, mail, id,
             securityEnabled, mailEnabled, membersCount, ownersCount
      FROM groups
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) AS total FROM groups ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const result = rows.map(r => ({
                        ...r,
                        securityEnabled: Boolean(r.securityEnabled),
                        mailEnabled: Boolean(r.mailEnabled)
                    }));
                    resolve(result);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Groups query failed:", err);
        res.status(500).json({ error: 'Failed to fetch group records' });
    }
});



// POST /report/groups/download
router.post('/report/groups/download', async (req, res) => {
    try {
        const data = req.body?.data;

        const groups = Array.isArray(data) && data.length > 0
            ? data
            : await runQueryAsync('SELECT * FROM groups');

        if (!groups || groups.length === 0) {
            return res.status(400).send("No group data available to generate CSV file.");
        }

        const filePath = await exportToCSVFile(groups, 'all_groups_report');
        res.download(filePath);
    } catch (error) {
        console.error("❌ Group CSV download failed:", error);
        res.status(500).send("Error generating CSV file for groups.");
    }
});




router.post('/report/groups/email', async (req, res) => {
    const { data, recipient } = req.body;

    if (!recipient) {
        return res.status(400).json({ status: "error", message: "Recipient email is required." });
    }

    try {
        const groups = Array.isArray(data) && data.length > 0
            ? data
            : await runQueryAsync("SELECT * FROM groups");

        if (!groups || groups.length === 0) {
            return res.status(400).json({ status: "error", message: "No group data available." });
        }

        const { csv, filename } = await exportToCSVString(groups, 'all_groups_report');

        if (!csv || typeof csv !== 'string') {
            throw new Error("Invalid CSV string generated.");
        }

        await sendReportByEmail(recipient, csv, filename);
        res.json({ status: "success", message: "Group report sent successfully!" });
    } catch (err) {
        console.error("❌ Email failed:", err.message);
        res.status(500).json({ status: "error", message: "Failed to send group report" });
    }
});



router.get('/report/groups/unified', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const { whereClause, values } = buildQuery(req.query, {
            searchFields: ['displayName', 'mail', 'id'],
            filters: {
                visibility: 'visibility',
                groupTypes: 'groupTypes',
                membersCount: 'membersCount'
            },
            additionalConditions: [`groupTypes LIKE '%Unified%'`]
        });

        const selectSQL = `
      SELECT displayName, createdDateTime, groupTypes, mail, id,
             visibility, membersCount, ownersCount
      FROM groups
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `
      SELECT COUNT(*) AS total FROM groups
      ${whereClause}
    `;

        const selectParams = isFullExport ? values : [...values, limit, offset];

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, selectParams, (err, rows) => {
                if (err) reject(err);
                else resolve(rows);
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Unified groups query failed:", err);
        res.status(500).json({ error: 'Failed to fetch unified group records' });
    }
});



// GET /report/groups/distribution
router.get('/report/groups/distribution', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Always enforce groupTypes = 'Distribution'
        const modifiedQuery = { ...req.query, groupTypes: "Distribution" };

        const { whereClause, values } = buildQuery(
            modifiedQuery,
            {
                searchFields: ['displayName', 'mail','id'],
                filters: {
                    groupTypes: 'groupTypes'
                }
            }
        );

        // Handle custom membersCount filter
        let finalWhereClause = whereClause;
        const extraValues = [];

        if (req.query.membersCount === 'true') {
            finalWhereClause += finalWhereClause ? ' AND membersCount > ?' : 'WHERE membersCount > ?';
            extraValues.push(0);
        } else if (req.query.membersCount === 'false') {
            finalWhereClause += finalWhereClause ? ' AND membersCount = ?' : 'WHERE membersCount = ?';
            extraValues.push(0);
        }

        const selectSQL = `
      SELECT displayName, createdDateTime, groupTypes, mail, id, membersCount, ownersCount
      FROM groups
      ${finalWhereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM groups ${finalWhereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, [...values, ...extraValues])
            : await runQueryAsync(selectSQL, [...values, ...extraValues, limit, offset]);

        const total = (await runGetAsync(countSQL, [...values, ...extraValues])).total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });
    } catch (err) {
        console.error("❌ Failed to fetch distribution groups:", err);
        res.status(500).json({ error: "Failed to fetch distribution group records." });
    }
});



// GET /report/groups/security-enabled
router.get('/report/groups/security-enabled', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Normalize groupTypes filter
        const inputGroupTypes = req.query.groupTypes
            ? Array.isArray(req.query.groupTypes)
                ? req.query.groupTypes
                : [req.query.groupTypes]
            : ['Security', 'Mail-enabled Security']; // default

        // Build base filters and search
        const { whereClause: baseWhere, values } = buildQuery(
            req.query,
            {
                searchFields: ['displayName', 'mail', 'id'],
                filters: {} // exclude groupTypes from buildQuery
            }
        );

        // Inject groupTypes filtering using IN (?, ?, ?...)
        const groupPlaceholders = inputGroupTypes.map(() => '?').join(', ');
        const groupTypeClause = `groupTypes IN (${groupPlaceholders})`;
        const groupValues = inputGroupTypes;

        let finalWhereClause = '';
        let allValues = [...groupValues, ...values];

        if (baseWhere) {
            finalWhereClause = `WHERE ${groupTypeClause} AND ${baseWhere.replace(/^WHERE\s*/, '')}`;
        } else {
            finalWhereClause = `WHERE ${groupTypeClause}`;
        }

        // Handle custom membersCount filter
        const extraValues = [];
        const countFilter = req.query.membersCount?.toLowerCase();

        if (countFilter === 'true') {
            finalWhereClause += ' AND membersCount > ?';
            extraValues.push(0);
        } else if (countFilter === 'false') {
            finalWhereClause += ' AND membersCount = ?';
            extraValues.push(0);
        }

        const selectSQL = `
            SELECT displayName, createdDateTime, groupTypes, mail, id, membersCount, ownersCount
            FROM groups
            ${finalWhereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM groups ${finalWhereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, [...allValues, ...extraValues])
            : await runQueryAsync(selectSQL, [...allValues, ...extraValues, limit, offset]);

        const total = (await runGetAsync(countSQL, [...allValues, ...extraValues])).total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });

    } catch (err) {
        console.error("❌ Failed to fetch security-enabled groups:", err);
        res.status(500).json({ error: "Failed to fetch security-enabled group records." });
    }
});


// GET /report/groups/mail-enabled-security
router.get('/report/groups/mail-enabled-security', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Always enforce groupTypes = 'Mail-enabled Security'
        const modifiedQuery = { ...req.query, groupTypes: "Mail-enabled Security" };

        const { whereClause, values } = buildQuery(
            modifiedQuery,
            {
                searchFields: ['displayName', 'mail', 'id'],
                filters: {
                    groupTypes: 'groupTypes'
                }
            }
        );

        // Handle custom membersCount filter
        let finalWhereClause = whereClause;
        const extraValues = [];

        const membersCount = String(req.query.membersCount || '').toLowerCase();
        if (membersCount === 'true') {
            finalWhereClause += finalWhereClause ? ' AND membersCount > ?' : 'WHERE membersCount > ?';
            extraValues.push(0);
        } else if (membersCount === 'false') {
            finalWhereClause += finalWhereClause ? ' AND membersCount = ?' : 'WHERE membersCount = ?';
            extraValues.push(0);
        }

        const selectSQL = `
      SELECT displayName, createdDateTime, groupTypes, mail, id, membersCount, ownersCount
      FROM groups
      ${finalWhereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM groups ${finalWhereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, [...values, ...extraValues])
            : await runQueryAsync(selectSQL, [...values, ...extraValues, limit, offset]);

        const total = (await runGetAsync(countSQL, [...values, ...extraValues])).total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });

    } catch (err) {
        console.error("❌ Failed to fetch mail-enabled security groups:", err);
        res.status(500).json({ error: "Failed to fetch mail-enabled security group records." });
    }
});



// GET /report/groups/empty
router.get('/report/groups/empty', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const modifiedQuery = { ...req.query };

        // Extract groupTypes manually (used in WHERE clause below)
        const groupTypesFilter = modifiedQuery.groupTypes?.trim();

        // Remove groupTypes from modifiedQuery so it doesn't interfere with buildQuery
        delete modifiedQuery.groupTypes;

        const { whereClause, values } = buildQuery(
            modifiedQuery,
            {
                searchFields: ['displayName', 'mail'],
                filters: {} // handled manually
            }
        );

        // Begin constructing WHERE clause
        let finalWhereClause = whereClause;
        const extraValues = [];

        // Add groupTypes filtering with LIKE for composite values
        if (groupTypesFilter) {
            finalWhereClause += finalWhereClause ? ' AND groupTypes LIKE ?' : 'WHERE groupTypes LIKE ?';
            extraValues.push(`%${groupTypesFilter}%`);
        }

        // Enforce membersCount = 0 (empty groups only)
        finalWhereClause += finalWhereClause ? ' AND membersCount = ?' : 'WHERE membersCount = ?';
        extraValues.push(0);

        const selectSQL = `
      SELECT displayName, createdDateTime, groupTypes, mail, membersCount, ownersCount
      FROM groups
      ${finalWhereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `SELECT COUNT(*) as total FROM groups ${finalWhereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, [...values, ...extraValues])
            : await runQueryAsync(selectSQL, [...values, ...extraValues, limit, offset]);

        const total = (await runGetAsync(countSQL, [...values, ...extraValues])).total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });

    } catch (err) {
        console.error("❌ Failed to fetch empty groups:", err);
        res.status(500).json({ error: "Failed to fetch empty group records." });
    }
});





// GET /report/groups/recently-created
router.get('/report/groups/recently-created', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Calculate ISO timestamp for 30 days ago
        const dateThreshold = new Date();
        dateThreshold.setDate(dateThreshold.getDate() - 30);
        const isoThreshold = dateThreshold.toISOString();

        // 🔁 Coerce securityEnabled & mailEnabled to integer values (1 or 0)
        const modifiedQuery = {
            ...req.query,
            ...(req.query.securityEnabled !== undefined && {
                securityEnabled: req.query.securityEnabled === 'true' ? 1 : 0
            }),
            ...(req.query.mailEnabled !== undefined && {
                mailEnabled: req.query.mailEnabled === 'true' ? 1 : 0
            })
        };

        // Build search & filters
        const { whereClause: baseWhere, values } = buildQuery(
            modifiedQuery,
            {
                searchFields: ['displayName', 'mail', 'id'],
                filters: {
                    securityEnabled: 'securityEnabled',
                    mailEnabled: 'mailEnabled'
                }
            }
        );

        // Enforce 30-day recent filter
        let finalWhereClause = '';
        const allValues = [...values];
        const extraValues = [];

        if (baseWhere) {
            finalWhereClause = `${baseWhere} AND createdDateTime >= ?`;
        } else {
            finalWhereClause = `WHERE createdDateTime >= ?`;
        }
        extraValues.push(isoThreshold);

        const selectSQL = `
            SELECT displayName, createdDateTime, groupTypes, mail, id, membersCount, ownersCount,
                   securityEnabled, mailEnabled
            FROM groups
            ${finalWhereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM groups ${finalWhereClause}`;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, [...allValues, ...extraValues])
            : await runQueryAsync(selectSQL, [...allValues, ...extraValues, limit, offset]);

        const total = (await runGetAsync(countSQL, [...allValues, ...extraValues])).total;

        const formattedData = data.map(row => ({
            ...row,
            securityEnabled: Boolean(row.securityEnabled),
            mailEnabled: Boolean(row.mailEnabled)
        }));

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: formattedData
        });

    } catch (err) {
        console.error("❌ Failed to fetch recently created groups:", err);
        res.status(500).json({ error: "Failed to fetch recently created group records." });
    }
});




// GET /report/groups/members
router.get('/report/groups/members', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Step 1: Base query builder for search only
        const { whereClause: baseWhere, values } = buildQuery(
            req.query,
            {
                searchFields: [
                    'gm.userPrincipalName',
                    'gm.department',
                    'gm.jobTitle',
                    'g.mail',
                    'g.displayName',
                    'g.id'
                ],
                filters: {} // groupTypes handled below
            }
        );

        // Step 2: Manually append groupTypes filter (exact match only)
        let finalWhereClause = baseWhere;
        const groupTypeValues = [];

        if (req.query.groupTypes) {
            const groupTypes = Array.isArray(req.query.groupTypes)
                ? req.query.groupTypes
                : [req.query.groupTypes];

            // Use exact match via IN clause
            const placeholders = groupTypes.map(() => '?').join(', ');
            const groupTypeClause = `g.groupTypes IN (${placeholders})`;
            groupTypeValues.push(...groupTypes);

            if (finalWhereClause) {
                finalWhereClause += ` AND ${groupTypeClause}`;
            } else {
                finalWhereClause = `WHERE ${groupTypeClause}`;
            }
        }

        const allValues = [...values, ...groupTypeValues];

        // Common FROM clause
        const baseSQL = `
      FROM group_members gm
      JOIN groups g ON gm.groupId = g.id
    `;

        const selectSQL = `
      SELECT gm.displayName AS memberName,
             gm.userPrincipalName,
             gm.department,
             gm.jobTitle,
             gm.signInStatus,
             g.displayName AS groupName,
             g.mail,
             g.displayName,
             g.id,
             g.groupTypes
      ${baseSQL}
      ${finalWhereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `
      SELECT COUNT(*) AS total
      ${baseSQL}
      ${finalWhereClause}
    `;

        const data = isFullExport
            ? await runQueryAsync(selectSQL, allValues)
            : await runQueryAsync(selectSQL, [...allValues, limit, offset]);

        const total = (await runGetAsync(countSQL, allValues)).total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });

    } catch (err) {
        console.error("❌ Failed to fetch group members:", err);
        res.status(500).json({ error: "Failed to fetch group membership records." });
    }
});


// GET /report/groups/owners
router.get('/report/groups/owners', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Build base query
        const { whereClause: baseWhere, values: baseValues } = buildQuery(
            req.query,
            {
                searchFields: [
                    'go.userPrincipalName',
                    'go.department',
                    'go.jobTitle',
                    'g.displayName',
                    'g.mail',
                    'g.id'
                ],
                filters: {
                    userPrincipalName: 'go.userPrincipalName',
                    displayName: 'g.displayName'
                }
            }
        );

        // Handle exact match groupTypes filter manually
        let groupTypeClause = '';
        const groupTypeValues = [];
        if (req.query.groupTypes) {
            groupTypeClause = 'g.groupTypes = ?';
            groupTypeValues.push(req.query.groupTypes);
        }

        let whereClause = '';
        const allValues = [...groupTypeValues, ...baseValues];

        if (groupTypeClause && baseWhere) {
            whereClause = `WHERE ${groupTypeClause} AND ${baseWhere.replace(/^WHERE\s*/, '')}`;
        } else if (groupTypeClause) {
            whereClause = `WHERE ${groupTypeClause}`;
        } else if (baseWhere) {
            whereClause = baseWhere;
        }

        const selectSQL = `
      SELECT 
        go.displayName AS ownerDisplayName,
        go.userPrincipalName,
        go.department,
        go.jobTitle,
        go.signInStatus,
        g.displayName AS groupDisplayName,
        g.mail,
        g.groupTypes,
        g.id
      FROM group_owners go
      JOIN groups g ON go.groupId = g.id
      ${whereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `
      SELECT COUNT(*) as total
      FROM group_owners go
      JOIN groups g ON go.groupId = g.id
      ${whereClause}
    `;

        const records = isFullExport
            ? await runQueryAsync(selectSQL, allValues)
            : await runQueryAsync(selectSQL, [...allValues, limit, offset]);

        const total = (await runGetAsync(countSQL, allValues)).total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Failed to fetch group owners:", err);
        res.status(500).json({ error: "Failed to fetch group owners" });
    }
});



// GET /report/groups/disabled-members
router.get('/report/groups/disabled-members', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Do not inject signInStatus into buildQuery; we’ll add it manually later
        const { whereClause, values } = buildQuery(
            req.query,
            {
                searchFields: [
                    "gm.userPrincipalName",
                    "gm.department",
                    "gm.jobTitle",
                    "gm.displayName",
                    "g.displayName",
                    "g.mail"
                ],
                filters: {
                    "gm.displayName": "gm.displayName",
                    "g.displayName": "g.displayName"
                }
            }
        );

        // Inject gm.signInStatus = 'Disabled' manually into WHERE clause
        let finalWhereClause = whereClause;
        const extraValues = [];

        if (finalWhereClause) {
            finalWhereClause += " AND gm.signInStatus = ?";
        } else {
            finalWhereClause = "WHERE gm.signInStatus = ?";
        }
        extraValues.push("Disabled");

        const selectSQL = `
      SELECT gm.displayName AS memberDisplayName,
             gm.userPrincipalName,
             gm.department,
             gm.jobTitle,
             gm.signInStatus,
             g.displayName AS groupDisplayName,
             g.mail,
             g.membersCount
      FROM group_members gm
      LEFT JOIN groups g ON gm.groupId = g.id
      ${finalWhereClause}
      ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
    `;

        const countSQL = `
      SELECT COUNT(*) AS total
      FROM group_members gm
      LEFT JOIN groups g ON gm.groupId = g.id
      ${finalWhereClause}
    `;

        const queryValues = [...values, ...extraValues];
        const data = isFullExport
            ? await runQueryAsync(selectSQL, queryValues)
            : await runQueryAsync(selectSQL, [...queryValues, limit, offset]);

        const total = (await runGetAsync(countSQL, queryValues)).total;

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: data
        });
    } catch (err) {
        console.error("❌ Failed to fetch disabled group members:", err);
        res.status(500).json({ error: "Failed to fetch disabled group member records." });
    }
});

/********* end of groups related reports *****************************/

/********* teams related reports *****************************/

router.get('/report/teams/all', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [];
        const values = [];

        // Search: displayName or id (fuzzy match on displayName, exact/fuzzy on id)
        if (req.query.search) {
            const s = req.query.search.trim();
            const isGuid = /^[0-9a-fA-F\-]{10,}$/.test(s); // rough check for UUID
            if (isGuid) {
                whereParts.push(`id LIKE ?`);
                values.push(`%${s}%`);
            } else {
                whereParts.push(`displayName LIKE ?`);
                values.push(`%${s}%`);
            }
        }

        // Filter: visibility (case-insensitive)
        if (req.query.visibility) {
            whereParts.push(`LOWER(visibility) = LOWER(?)`);
            values.push(req.query.visibility.trim());
        }

        // Filter: isArchived (convert true/false to 1/0)
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push(`isArchived = ?`);
            values.push(boolValue);
        }

        const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

        const selectSQL = `
            SELECT id, displayName, description, visibility, membersCount, ownersCount,
                   privateChannelsCount, standardChannelsCount, sharedChannelsCount,
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const formatted = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(formatted);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Teams query failed:", err);
        res.status(500).json({ error: 'Failed to fetch team records' });
    }
});


// GET /report/teams/public
router.get('/report/teams/public', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [];
        const values = [];

        // Always enforce visibility = 'Public'
        whereParts.push(`LOWER(visibility) = LOWER(?)`);
        values.push('Public');

        // Search: displayName (fuzzy match)
        if (req.query.search) {
            const s = req.query.search.trim();
            whereParts.push(`displayName LIKE ?`);
            values.push(`%${s}%`);
        }

        // Filter: isArchived (convert true/false to 1/0)
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push(`isArchived = ?`);
            values.push(boolValue);
        }

        const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

        const selectSQL = `
            SELECT id, displayName, description, visibility, membersCount, ownersCount,
                   privateChannelsCount, standardChannelsCount, sharedChannelsCount,
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const formatted = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(formatted);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Public Teams query failed:", err);
        res.status(500).json({ error: 'Failed to fetch public team records' });
    }
});




// GET /report/teams/private
router.get('/report/teams/private', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [];
        const values = [];

        // Enforce visibility = 'Private'
        whereParts.push(`LOWER(visibility) = LOWER(?)`);
        values.push('Private');

        // Search: displayName (fuzzy match)
        if (req.query.search) {
            const s = req.query.search.trim();
            whereParts.push(`displayName LIKE ?`);
            values.push(`%${s}%`);
        }

        // Filter: isArchived (convert true/false to 1/0)
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push(`isArchived = ?`);
            values.push(boolValue);
        }

        const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

        const selectSQL = `
            SELECT id, displayName, description, visibility, membersCount, ownersCount,
                   privateChannelsCount, standardChannelsCount, sharedChannelsCount,
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const formatted = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(formatted);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Private Teams query failed:", err);
        res.status(500).json({ error: 'Failed to fetch private team records' });
    }
});


// GET /report/teams/teams-without-description

router.get('/report/teams/teams-without-description', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [
            "(description IS NULL OR TRIM(description) = '')"  // ✅ TRIM to catch whitespace-only strings
        ];
        const values = [];

        // 🔍 Search by displayName
        if (req.query.search) {
            const s = req.query.search.trim();
            whereParts.push("displayName LIKE ?");
            values.push(`%${s}%`);
        }

        // 🔽 Filter by visibility
        if (req.query.visibility) {
            whereParts.push("LOWER(visibility) = LOWER(?)");
            values.push(req.query.visibility.trim());
        }

        // 🔽 Filter by isArchived
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push("isArchived = ?");
            values.push(boolValue);
        }

        const whereClause = `WHERE ${whereParts.join(' AND ')}`;

        const selectSQL = `
            SELECT id, displayName, description, visibility,
                   membersCount, ownersCount, privateChannelsCount,
                   standardChannelsCount, sharedChannelsCount,
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const result = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(result);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Failed to fetch teams without description:", err);
        res.status(500).json({ error: "Failed to fetch teams without description records." });
    }
});


// GET /report/teams/archived
router.get('/report/teams/archived', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        // Always filter for archived teams only
        const modifiedQuery = { ...req.query, isArchived: 1 };

        const { whereClause, values } = buildQuery(modifiedQuery, {
            searchFields: ['displayName'],
            filters: {
                visibility: 'visibility',
                isArchived: 'isArchived'
            }
        });

        const selectSQL = `
            SELECT id, displayName, description, visibility,
                   membersCount, ownersCount, privateChannelsCount,
                   standardChannelsCount, sharedChannelsCount,
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const result = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(result);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Failed to fetch archived teams:", err);
        res.status(500).json({ error: 'Failed to fetch archived team records' });
    }
});

// GET /report/teams/teams-private-channels
router.get('/report/teams/teams-private-channels', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = ['privateChannelsCount > 0'];
        const values = [];

        // Search: displayName (fuzzy match)
        if (req.query.search) {
            const s = req.query.search.trim();
            whereParts.push(`displayName LIKE ?`);
            values.push(`%${s}%`);
        }

        // Filter: visibility (case-insensitive)
        if (req.query.visibility) {
            whereParts.push(`LOWER(visibility) = LOWER(?)`);
            values.push(req.query.visibility.trim());
        }

        // Filter: isArchived (convert true/false to 1/0)
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push(`isArchived = ?`);
            values.push(boolValue);
        }

        const whereClause = `WHERE ${whereParts.join(' AND ')}`;

        const selectSQL = `
            SELECT id, displayName, description, visibility, membersCount, ownersCount,
                   privateChannelsCount, 
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport
            ? [...values, limit, offset]
            : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const formatted = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(formatted);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Failed to fetch teams with private channels:", err);
        res.status(500).json({ error: 'Failed to fetch records.' });
    }
});


// GET /report/teams/teams-shared-channels
router.get('/report/teams/teams-shared-channels', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [`sharedChannelsCount > 0`]; // core condition
        const values = [];

        // Search: displayName
        if (req.query.search) {
            const s = req.query.search.trim();
            whereParts.push(`displayName LIKE ?`);
            values.push(`%${s}%`);
        }

        // Filter: visibility (case-insensitive)
        if (req.query.visibility) {
            whereParts.push(`LOWER(visibility) = LOWER(?)`);
            values.push(req.query.visibility.trim());
        }

        // Filter: isArchived (Boolean)
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push(`isArchived = ?`);
            values.push(boolValue);
        }

        const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

        const selectSQL = `
            SELECT id, displayName, description, visibility, membersCount, ownersCount,
                   sharedChannelsCount,createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const formatted = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(formatted);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Teams with shared channels query failed:", err);
        res.status(500).json({ error: 'Failed to fetch shared channel team records' });
    }
});


// GET /report/teams/recently-created-teams
router.get('/report/teams/recently-created-teams', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [];
        const values = [];

        // Calculate date 90 days ago
        const currentDate = new Date();
        const ninetyDaysAgo = new Date(currentDate.setDate(currentDate.getDate() - 90));
        const iso90DaysAgo = ninetyDaysAgo.toISOString();

        // Always enforce createdDateTime >= 90 days ago
        whereParts.push(`createdDateTime >= ?`);
        values.push(iso90DaysAgo);

        // Search: displayName (fuzzy match)
        if (req.query.search) {
            const s = req.query.search.trim();
            whereParts.push(`displayName LIKE ?`);
            values.push(`%${s}%`);
        }

        // Filter: visibility (case-insensitive)
        if (req.query.visibility) {
            whereParts.push(`LOWER(visibility) = LOWER(?)`);
            values.push(req.query.visibility.trim());
        }

        // Filter: isArchived (convert true/false to 1/0)
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push(`isArchived = ?`);
            values.push(boolValue);
        }

        const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

        const selectSQL = `
            SELECT id, displayName, description, visibility, membersCount, ownersCount,
                   privateChannelsCount, standardChannelsCount, sharedChannelsCount,
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = isFullExport ? values : [...values, limit, offset];

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const formatted = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(formatted);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Recently created teams query failed:", err);
        res.status(500).json({ error: 'Failed to fetch recently created team records' });
    }
});


// GET /report/teams/hidden-memberships
router.get('/report/teams/hidden-memberships', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const whereParts = [`LOWER(visibility) = LOWER(?)`];
        const values = ['HiddenMembership'];

        // Search support
        if (req.query.search) {
            const s = req.query.search.trim();
            whereParts.push(`displayName LIKE ?`);
            values.push(`%${s}%`);
        }

        // isArchived filter
        if (req.query.isArchived) {
            const boolValue = req.query.isArchived === 'true' ? 1 : 0;
            whereParts.push(`isArchived = ?`);
            values.push(boolValue);
        }

        const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

        const selectSQL = `
            SELECT id, displayName, description, visibility, membersCount, ownersCount,
                   privateChannelsCount, standardChannelsCount, sharedChannelsCount,
                   createdDateTime, isArchived
            FROM teams
            ${whereClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `SELECT COUNT(*) as total FROM teams ${whereClause}`;
        const queryParams = !isFullExport ? [...values, limit, offset] : values;

        const records = await new Promise((resolve, reject) => {
            db.all(selectSQL, queryParams, (err, rows) => {
                if (err) reject(err);
                else {
                    const formatted = rows.map(r => ({
                        ...r,
                        isArchived: Boolean(r.isArchived)
                    }));
                    resolve(formatted);
                }
            });
        });

        const total = await new Promise((resolve, reject) => {
            db.get(countSQL, values, (err, row) => {
                if (err) reject(err);
                else resolve(row.total);
            });
        });

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records
        });

    } catch (err) {
        console.error("❌ Hidden Memberships Teams query failed:", err);
        res.status(500).json({ error: 'Failed to fetch hidden membership teams records' });
    }
});

router.get('/report/teams/teams-owners', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const searchValue = req.query.search?.trim();
        let searchCondition = "";
        const values = [];

        if (searchValue) {
            const isGuid = /^[0-9a-fA-F\-]{10,}$/.test(searchValue);
            if (isGuid) {
                searchCondition = "WHERE t.id LIKE ?";
                values.push(`%${searchValue}%`);
            } else {
                searchCondition = "WHERE t.displayName LIKE ?";
                values.push(`%${searchValue}%`);
            }
        }

        const { whereClause: filterClause, values: filterValues } = buildQuery(
            req.query,
            {
                searchFields: [], // Already handled above
                filters: {
                    "own.userPrincipalName": "own.userPrincipalName",
                    "t.displayName": "t.displayName",
                    "own.department": "own.department",
                    "t.visibility": "t.visibility",
                    "t.isArchived": "t.isArchived"
                }
            }
        );

        const finalWhereClause = [
            searchCondition.replace(/^WHERE\s*/, ""),
            filterClause.replace(/^WHERE\s*/, "")
        ].filter(Boolean).join(" AND ");

        const fullClause = finalWhereClause ? `WHERE ${finalWhereClause}` : "";

        const selectSQL = `
            SELECT 
                own.userId,
                own.userPrincipalName,
                own.department,
                own.jobTitle,
                own.signInStatus,
                own.displayName AS ownerDisplayName,
                t.id,
                t.displayName AS teamDisplayName,
                t.description,
                t.createdDateTime,
                t.visibility,
                t.membersCount,
                t.ownersCount,
                t.privateChannelsCount,
                t.standardChannelsCount,
                t.sharedChannelsCount,
                t.isArchived
            FROM team_owners own
            LEFT JOIN teams t ON own.teamId = t.id
            ${fullClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `
            SELECT COUNT(*) AS total
            FROM team_owners own
            LEFT JOIN teams t ON own.teamId = t.id
            ${fullClause}
        `;

        const queryParams = isFullExport
            ? [...values, ...filterValues]
            : [...values, ...filterValues, limit, offset];

        const data = await runQueryAsync(selectSQL, queryParams);
        const total = (await runGetAsync(countSQL, [...values, ...filterValues])).total;

        const formatted = data.map(row => ({
            ...row,
            isArchived: Boolean(row.isArchived)
        }));

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: formatted
        });

    } catch (err) {
        console.error("❌ Failed to fetch team owners:", err);
        res.status(500).json({ error: "Failed to fetch team owners data." });
    }
});

router.get('/report/teams/teams-members', async (req, res) => {
    try {
        const isFullExport = req.query.page === "all";
        const limit = 20;
        const page = isFullExport ? 1 : parseInt(req.query.page) || 1;
        const offset = isFullExport ? null : (page - 1) * limit;

        const searchValue = req.query.search?.trim();
        let searchCondition = "";
        const values = [];

        if (searchValue) {
            const isGuid = /^[0-9a-fA-F\-]{10,}$/.test(searchValue);
            if (isGuid) {
                searchCondition = "WHERE t.id LIKE ?";
                values.push(`%${searchValue}%`);
            } else {
                searchCondition = "WHERE t.displayName LIKE ?";
                values.push(`%${searchValue}%`);
            }
        }

        const { whereClause: filterClause, values: filterValues } = buildQuery(
            req.query,
            {
                searchFields: [], // handled above
                filters: {
                    "tm.userPrincipalName": "tm.userPrincipalName",
                    "t.displayName": "t.displayName",
                    "t.visibility": "t.visibility",
                    "t.isArchived": "t.isArchived"
                }
            }
        );

        const finalWhereClause = [
            searchCondition.replace(/^WHERE\s*/, ""),
            filterClause.replace(/^WHERE\s*/, "")
        ].filter(Boolean).join(" AND ");

        const fullClause = finalWhereClause ? `WHERE ${finalWhereClause}` : "";

        const selectSQL = `
            SELECT 
                tm.userId,
                tm.userPrincipalName,
                tm.department,
                tm.jobTitle,
                tm.signInStatus,
                tm.displayName AS memberDisplayName,
                t.id,
                t.displayName AS teamDisplayName,
                t.description,
                t.createdDateTime,
                t.visibility,
                t.membersCount,
                t.ownersCount,
                t.privateChannelsCount,
                t.standardChannelsCount,
                t.sharedChannelsCount,
                t.isArchived
            FROM team_members tm
            LEFT JOIN teams t ON tm.teamId = t.id
            ${fullClause}
            ${!isFullExport ? "LIMIT ? OFFSET ?" : ""}
        `;

        const countSQL = `
            SELECT COUNT(*) as total
            FROM team_members tm
            LEFT JOIN teams t ON tm.teamId = t.id
            ${fullClause}
        `;

        const queryParams = !isFullExport
            ? [...values, ...filterValues, limit, offset]
            : [...values, ...filterValues];

        const data = await runQueryAsync(selectSQL, queryParams);
        const total = (await runGetAsync(countSQL, [...values, ...filterValues])).total;

        const formatted = data.map(record => ({
            ...record,
            isArchived: Boolean(record.isArchived)
        }));

        res.json({
            page,
            totalPages: isFullExport ? 1 : Math.ceil(total / limit),
            totalRecords: total,
            records: formatted
        });

    } catch (err) {
        console.error("❌ Failed to fetch team members:", err);
        res.status(500).json({ error: "Failed to fetch team member records." });
    }
});




module.exports = router;

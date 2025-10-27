const path = require("path");
const db = require("./db");
const BATCH_SIZE = 20;

const fetch = global.fetch;

function runAsync(stmt, params = []) {
  return new Promise((resolve, reject) => {
    stmt.run(params, function (err) {
      if (err) reject(err);
      else resolve(this);
    });
  });
}

function getAsync(stmt, params = []) {
  return new Promise((resolve, reject) => {
    stmt.get(params, (err, row) => {
      if (err) reject(err);
      else resolve(row);
    });
  });
}

function allAsync(stmt, params = []) {
  return new Promise((resolve, reject) => {
    stmt.all(params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows);
    });
  });
}

async function fetchUserEnrichmentBatch(userIds, headers) {
  const batchRequests = userIds.map((id, index) => ({
    id: index.toString(),
    method: "GET",
      url: `/users/${id}?$select=id,accountEnabled,department,assignedLicenses,createdDateTime,mailNickname,city,country,state,postalCode,streetAddress,preferredLanguage,mobilePhone,businessPhones,givenName,surname&$top=999`
  }));

  const batchBody = { requests: batchRequests };

  try {
    const res = await fetch("https://graph.microsoft.com/v1.0/$batch", {
      method: "POST",
      headers: {
        ...headers,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(batchBody)
    });

    const data = await res.json();
    const results = data.responses;

    const enrichedMap = new Map();
    for (const item of results) {
      if (item.status === 200 && item.body && item.body.id) {
        enrichedMap.set(item.body.id, item.body);
      }
    }

    return enrichedMap;
  } catch (err) {
    console.error("❌ Batch enrichment failed:", err);
    return new Map();
  }
}

async function syncUsers(token) {
  const headers = { Authorization: `Bearer ${token}` };
  const now = new Date().toISOString();
  let added = 0, updated = 0, deleted = 0;
  let newDeltaLink = "";

  try {
    const deltaRes = await getAsync(db.prepare("SELECT delta_link FROM delta_tokens WHERE resource = 'users'"));
    let deltaUrl = deltaRes?.delta_link || "https://graph.microsoft.com/v1.0/users/delta";

    const allChanges = [];
    let nextLink = deltaUrl;

    while (nextLink) {
      const res = await fetch(nextLink, { headers });
      const data = await res.json();
      allChanges.push(...data.value);

      if (data["@odata.deltaLink"]) {
        newDeltaLink = data["@odata.deltaLink"];
        break;
      } else if (data["@odata.nextLink"]) {
        nextLink = data["@odata.nextLink"];
      } else {
        break;
      }
    }

    const dbUsers = await allAsync(db.prepare("SELECT * FROM users"));
    const dbUserMap = new Map(dbUsers.map(u => [u.id, u]));

    const toDelete = [];
    const toEnrich = [];
    const deltaRawMap = new Map();

    for (const item of allChanges) {
      if (item["@removed"]) {
        toDelete.push(item.id);
      } else {
        toEnrich.push(item.id);
        deltaRawMap.set(item.id, item);
      }
    }

    const enrichedMap = new Map();
    for (let i = 0; i < toEnrich.length; i += BATCH_SIZE) {
      const chunk = toEnrich.slice(i, i + BATCH_SIZE);
      const chunkResults = await fetchUserEnrichmentBatch(chunk, headers);
      for (const [id, data] of chunkResults) {
        enrichedMap.set(id, data);
      }
    }

    const normalize = (v) => (v ?? '').toString().trim().toLowerCase();
    const finalUsers = [];

    for (const id of toEnrich) {
      const dbUser = dbUserMap.get(id);
      const deltaData = deltaRawMap.get(id) || {};
      const enriched = enrichedMap.get(id) || {};

      const user = {
        id,
        displayName: deltaData.displayName || dbUser?.displayName || '',
        userPrincipalName: deltaData.userPrincipalName || dbUser?.userPrincipalName || '',
        mail: deltaData.mail || dbUser?.mail || '',
        jobTitle: deltaData.jobTitle || dbUser?.jobTitle || '',
        licenseStatus: Array.isArray(enriched.assignedLicenses) && enriched.assignedLicenses.length > 0 ? 'Licensed' : 'Unlicensed',
        mailNickname: enriched.mailNickname || dbUser?.mailNickName || '',
        userAddedTime: enriched.createdDateTime || dbUser?.userAddedTime || '',
        officePhone: enriched.businessPhones?.[0] || dbUser?.officePhone || '',
        mobilePhone: enriched.mobilePhone || dbUser?.mobilePhone || '',
        city: enriched.city || dbUser?.city || '',
        country: enriched.country || dbUser?.country || '',
        postalCode: enriched.postalCode || dbUser?.postalCode || '',
        state: enriched.state || dbUser?.state || '',
        streetAddress: enriched.streetAddress || dbUser?.streetAddress || '',
        preferredLanguage: enriched.preferredLanguage || dbUser?.preferredLanguage || '',
        department: enriched.department || dbUser?.department || '',
        firstName: enriched.givenName || dbUser?.firstName || '',
        lastName: enriched.surname || dbUser?.lastName || '',
        signInStatus: enriched.accountEnabled === false ? 'Disabled' : 'Enabled'
      };

      const hasChanges = !dbUser || Object.keys(user).some(key =>
        normalize(user[key]) !== normalize(dbUser[key])
      );

      if (hasChanges) finalUsers.push(user);
    }

    const deleteStmt = db.prepare("DELETE FROM users WHERE id = ?");
    const upsertStmt = db.prepare(`
      INSERT INTO users (
        id, displayName, userPrincipalName, email, licenseStatus, department,
        jobTitle, firstName, lastName, mailNickName, userAddedTime,
        alternateEmail, officePhone, mobilePhone, faxNumber,
        city, country, postalCode, state, streetAddress,
        companyName, usageLocation, office, signInStatus, mfaStatus,
        preferredLanguage, directReportsCount
      ) VALUES (
        ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?,
        '', ?, ?, '', 
        ?, ?, ?, ?, ?, 
        '', '', '', ?, '', ?, NULL
      )
      ON CONFLICT(id) DO UPDATE SET
        displayName = excluded.displayName,
        userPrincipalName = excluded.userPrincipalName,
        email = excluded.email,
        licenseStatus = excluded.licenseStatus,
        department = excluded.department,
        jobTitle = excluded.jobTitle,
        firstName = excluded.firstName,
        lastName = excluded.lastName,
        mailNickName = excluded.mailNickName,
        userAddedTime = excluded.userAddedTime,
        alternateEmail = excluded.alternateEmail,
        officePhone = excluded.officePhone,
        mobilePhone = excluded.mobilePhone,
        faxNumber = excluded.faxNumber,
        city = excluded.city,
        country = excluded.country,
        postalCode = excluded.postalCode,
        state = excluded.state,
        streetAddress = excluded.streetAddress,
        companyName = excluded.companyName,
        usageLocation = excluded.usageLocation,
        office = excluded.office,
        signInStatus = excluded.signInStatus,
        mfaStatus = excluded.mfaStatus,
        preferredLanguage = excluded.preferredLanguage,
        directReportsCount = excluded.directReportsCount;
    `);

    for (const id of toDelete) {
      await runAsync(deleteStmt, [id]);
      deleted++;
    }

    for (const user of finalUsers) {
      await runAsync(upsertStmt, [
        user.id, user.displayName, user.userPrincipalName, user.mail, user.licenseStatus, user.department,
        user.jobTitle, user.firstName, user.lastName, user.mailNickname, user.userAddedTime,
        user.officePhone, user.mobilePhone,
        user.city, user.country, user.postalCode, user.state, user.streetAddress,
        user.signInStatus, user.preferredLanguage
      ]);
      if (!dbUserMap.has(user.id)) added++;
      else updated++;
    }

    if (newDeltaLink) {
      await runAsync(
        db.prepare("INSERT OR REPLACE INTO delta_tokens (resource, delta_link, last_synced_at) VALUES ('users', ?, ?)"),
        [newDeltaLink, now]
      );
    }

    await runAsync(
      db.prepare("INSERT INTO sync_logs (resource, synced_at, added, updated, deleted, status, error_message) VALUES (?, ?, ?, ?, ?, ?, ?)"),
      ['users', now, added, updated, deleted, 'SUCCESS', '']
    );

    console.log(`🌀 Delta Sync | ➕ Added: ${added}, 📝 Updated: ${updated}, ❌ Deleted: ${deleted}`);
    return { added, updated, deleted };

  } catch (err) {
    await runAsync(
      db.prepare("INSERT INTO sync_logs (resource, synced_at, added, updated, deleted, status, error_message) VALUES (?, ?, ?, ?, ?, ?, ?)"),
      ['users', now, added, updated, deleted, 'FAILED', err.message]
    );
    console.error("❌ Sync failed:", err.message);
    throw err;
  }
}

module.exports = { syncUsers };

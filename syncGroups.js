

const db = require('./db');
const { getTokenAppOnly } = require('./auth');

async function fetchJSON(url, token, headers = {}) {
    const res = await fetch(url, {
        headers: { Authorization: `Bearer ${token}`, ...headers }
    });
    if (!res.ok) throw new Error(`[fetchJSON] ${res.status} - ${await res.text()}`);
    return await res.json();
}

async function fetchCount(url, token) {
    try {
        const res = await fetch(url, {
            headers: { Authorization: `Bearer ${token}`, ConsistencyLevel: 'eventual' }
        });
        if (!res.ok) throw new Error("Count fetch failed");
        return parseInt(await res.text()) || 0;
    } catch (err) {
        console.warn(`[fetchCount] ⚠️ ${err.message}`);
        return 0;
    }
}

async function syncGroups() {
    console.log("[syncGroups.js] 🔁 Group delta sync initiated...");
    const token = await getTokenAppOnly();
    const timestamp = new Date().toISOString();
    const resource = 'groups';

    let added = 0, updated = 0, deleted = 0, deltaLink = null;

    const existingToken = await new Promise(resolve => {
        db.get("SELECT delta_link FROM delta_tokens WHERE resource = ?", [resource], (err, row) => {
            resolve(row?.delta_link || null);
        });
    });

    let url = existingToken || 'https://graph.microsoft.com/v1.0/groups/delta?$select=id,displayName,createdDateTime,groupTypes,mail,visibility,securityEnabled,mailEnabled,resourceProvisioningOptions';

    try {
        while (url && typeof url === "string") {
            console.log(`[syncGroups.js] 🔄 Fetching: ${url}`);
            const result = await fetchJSON(url, token);
            if (!result?.value || !Array.isArray(result.value)) break;

            for (const group of result.value) {
                const groupId = group.id;

                if (Array.isArray(group.resourceProvisioningOptions) && group.resourceProvisioningOptions.includes('Team')) {
                    console.log(`[syncGroups.js] ⏭️ Skipping Teams group: ${group.displayName}`);
                    continue;
                }

                if (group["@removed"]) {
                    db.run("DELETE FROM groups WHERE id = ?", [groupId]);
                    db.run("DELETE FROM group_members WHERE groupId = ?", [groupId]);
                    db.run("DELETE FROM group_owners WHERE groupId = ?", [groupId]);
                    deleted++;
                    continue;
                }

                let groupTypes = '';
                let visibility = group.visibility || '';
                if (Array.isArray(group.groupTypes) && group.groupTypes.length > 0) {
                    groupTypes = group.groupTypes.join(',');
                } else {
                    if (group.mailEnabled && !group.securityEnabled) {
                        groupTypes = 'Distribution';
                        visibility = 'Distribution';
                    } else if (!group.mailEnabled && group.securityEnabled) {
                        groupTypes = 'Security';
                        visibility = 'Security';
                    } else if (group.mailEnabled && group.securityEnabled) {
                        groupTypes = 'Mail-enabled Security';
                        visibility = 'Security';
                    }
                }

                const exists = await new Promise(resolve => {
                    db.get("SELECT 1 FROM groups WHERE id = ?", [groupId], (err, row) => {
                        resolve(!!row);
                    });
                });

                db.run(`
                    INSERT INTO groups (id, displayName, createdDateTime, groupTypes, mail, visibility, securityEnabled, mailEnabled, ownersCount, membersCount)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, 0)
                    ON CONFLICT(id) DO UPDATE SET
                        displayName=excluded.displayName,
                        createdDateTime=excluded.createdDateTime,
                        groupTypes=excluded.groupTypes,
                        mail=excluded.mail,
                        visibility=excluded.visibility,
                        securityEnabled=excluded.securityEnabled,
                        mailEnabled=excluded.mailEnabled
                `, [
                    groupId,
                    group.displayName || '',
                    group.createdDateTime || '',
                    groupTypes,
                    group.mail || '',
                    visibility,
                    group.securityEnabled ? 1 : 0,
                    group.mailEnabled ? 1 : 0
                ]);

                exists ? updated++ : added++;

                console.log(`[syncGroups.js] ✅ Synced group metadata: ${group.displayName}`);
            }

            if (result["@odata.nextLink"]) {
                url = result["@odata.nextLink"];
            } else if (result["@odata.deltaLink"]) {
                deltaLink = result["@odata.deltaLink"];
                break;
            } else break;
        }

        // === 🔁 Post-sync: update owner/member count & user details ===
        db.all("SELECT id, displayName FROM groups", async (err, rows) => {
            if (err || !rows) return console.warn("[syncGroups.js] ⚠️ Error fetching groups for count update.");
            for (const { id: groupId, displayName } of rows) {
                const owners = await fetchJSON(`https://graph.microsoft.com/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,department,jobTitle,accountEnabled&$top=999`, token);
                const members = await fetchJSON(`https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,department,jobTitle,accountEnabled&$top=999`, token);

                // Count + update groups table
                const ownersCount = Array.isArray(owners.value) ? owners.value.length : 0;
                const membersCount = Array.isArray(members.value) ? members.value.length : 0;
                db.run("UPDATE groups SET ownersCount = ?, membersCount = ? WHERE id = ?", [ownersCount, membersCount, groupId]);

                // Update group_owners table
                for (const o of owners.value || []) {
                    const compositeId = `${groupId}_${o.id}`;
                    db.run(`
                        INSERT INTO group_owners (id, groupId, userId, displayName, userPrincipalName, department, jobTitle, signInStatus)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(id) DO UPDATE SET
                            displayName=excluded.displayName,
                            userPrincipalName=excluded.userPrincipalName,
                            department=excluded.department,
                            jobTitle=excluded.jobTitle,
                            signInStatus=excluded.signInStatus
                    `, [
                        compositeId,
                        groupId,
                        o.id,
                        o.displayName || '',
                        o.userPrincipalName || '',
                        o.department || '',
                        o.jobTitle || '',
                        o.accountEnabled ? 'Enabled' : 'Disabled'
                    ]);
                }

                // Update group_members table
                for (const m of members.value || []) {
                    const compositeId = `${groupId}_${m.id}`;
                    db.run(`
                        INSERT INTO group_members (id, groupId, userId, displayName, userPrincipalName, department, jobTitle, signInStatus)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(id) DO UPDATE SET
                            displayName=excluded.displayName,
                            userPrincipalName=excluded.userPrincipalName,
                            department=excluded.department,
                            jobTitle=excluded.jobTitle,
                            signInStatus=excluded.signInStatus
                    `, [
                        compositeId,
                        groupId,
                        m.id,
                        m.displayName || '',
                        m.userPrincipalName || '',
                        m.department || '',
                        m.jobTitle || '',
                        m.accountEnabled ? 'Enabled' : 'Disabled'
                    ]);
                }

                console.log(`[syncGroups.js] 🔁 Updated counts and members/owners for: ${displayName}`);
            }
        });

        if (deltaLink) {
            db.run(`
                INSERT INTO delta_tokens (resource, delta_link, last_synced_at)
                VALUES (?, ?, ?)
                ON CONFLICT(resource) DO UPDATE SET
                    delta_link=excluded.delta_link,
                    last_synced_at=excluded.last_synced_at
            `, [resource, deltaLink, timestamp]);
        }

        db.run(`
            INSERT INTO sync_logs (resource, synced_at, added, updated, deleted, status)
            VALUES (?, ?, ?, ?, ?, ?)
        `, [resource, timestamp, added, updated, deleted, 'Success']);

        console.log(`[syncGroups.js] ✅ Delta sync complete: Added ${added}, Updated ${updated}, Deleted ${deleted}`);
        return { added, updated, deleted }; /******************************************************/
    } catch (err) {
        console.error(`[syncGroups.js] ❌ Sync failed: ${err.message}`);
        db.run(`
            INSERT INTO sync_logs (resource, synced_at, added, updated, deleted, status, error_message)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        `, [resource, timestamp, added, updated, deleted, 'Failed', err.message]);
    }
}

module.exports = { syncGroups };


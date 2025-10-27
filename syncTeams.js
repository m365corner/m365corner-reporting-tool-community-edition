const db = require('./db');
const { getTokenAppOnly } = require('./auth');

/* ---------- sqlite helpers ---------- */
function runAsync(sql, params = []) {
    return new Promise((resolve, reject) => {
        db.run(sql, params, function (err) {
            if (err) return reject(err);
            resolve(this);
        });
    });
}
function getAsync(sql, params = []) {
    return new Promise((resolve) => {
        db.get(sql, params, (err, row) => resolve(row || null));
    });
}
function allAsync(sql, params = []) {
    return new Promise((resolve) => {
        db.all(sql, params, (err, rows) => resolve(rows || []));
    });
}

/* ---------- HTTP helpers ---------- */
async function fetchJSON(url, token, headers = {}) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, ...headers } });
    if (!res.ok) throw new Error(`[fetchJSON] ${res.status} - ${await res.text()} (${url})`);
    return res.json();
}

// Generic pager over @odata.nextLink
async function paginate(url, token, headers = {}) {
    const results = [];
    let next = url;
    while (next) {
        const page = await fetchJSON(next, token, headers);
        if (Array.isArray(page.value)) results.push(...page.value);
        next = page['@odata.nextLink'] || null;
    }
    return results;
}

/* ---------- Channels counting: NO $top, follow @odata.nextLink ---------- */
async function countChannelsByType(teamId, membershipType, token) {
    let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels?$filter=membershipType eq '${membershipType}'`;
    let total = 0;
    while (url) {
        try {
            const page = await fetchJSON(url, token);
            const arr = Array.isArray(page.value) ? page.value : [];
            total += arr.length;
            url = page['@odata.nextLink'] || null;
        } catch (e) {
            console.warn(`[channels] ⚠️ ${membershipType} count failed for ${teamId}: ${e.message}`);
            return total;
        }
    }
    return total;
}

/* ---------- Owners/Members: Teams API first; fallback to Groups owners/members (both paginated) ---------- */
async function getOwnersAndMembers(teamId, token) {
    // Prefer Teams members (has roles + includes guests). Page through if needed.
    try {
        const all = await paginate(`https://graph.microsoft.com/v1.0/teams/${teamId}/members`, token);
        const owners = [];
        const members = [];
        for (const m of all) {
            const roles = Array.isArray(m.roles) ? m.roles : [];
            const entry = {
                userId: m.userId || m.id || '',
                displayName: m.displayName || '',
                email: m.email || ''
            };
            if (roles.includes('owner')) owners.push(entry);
            else members.push(entry);
        }
        return { owners, members, usedFallback: false };
    } catch (err) {
        console.warn(`[members] ⚠️ Teams members API failed for ${teamId}, falling back to Groups: ${err.message}`);
        // Fallback: groups owners/members (also paginated)
        let owners = [], members = [];
        try {
            const own = await paginate(
                `https://graph.microsoft.com/v1.0/groups/${teamId}/owners?$select=id,displayName,userPrincipalName`,
                token
            );
            owners = (own || []).map(o => ({
                userId: o.id || '',
                displayName: o.displayName || '',
                email: o.userPrincipalName || ''
            }));
        } catch (e) {
            console.warn(`[owners] ⚠️ Groups owners fetch failed for ${teamId}: ${e.message}`);
        }
        try {
            const mem = await paginate(
                `https://graph.microsoft.com/v1.0/groups/${teamId}/members?$select=id,displayName,userPrincipalName`,
                token
            );
            members = (mem || []).map(m => ({
                userId: m.id || '',
                displayName: m.displayName || '',
                email: m.userPrincipalName || ''
            }));
        } catch (e) {
            console.warn(`[members] ⚠️ Groups members fetch failed for ${teamId}: ${e.message}`);
        }
        return { owners, members, usedFallback: true };
    }
}

/* ---------- Local enrichment from users table ---------- */
async function enrichFromLocalUsers(userId, fallbackEmail = '') {
    const row = await getAsync(
        `SELECT userPrincipalName, department, jobTitle, signInStatus FROM users WHERE id = ?`,
        [userId]
    );
    return {
        userPrincipalName: row?.userPrincipalName || fallbackEmail || '',
        department: row?.department || '',
        jobTitle: row?.jobTitle || '',
        signInStatus: row?.signInStatus || ''
    };
}

/* ---------- Upserts ---------- */
async function upsertTeamRow(row) {
    const sql = `
    INSERT INTO teams (
      id, displayName, description, visibility, isArchived, createdDateTime,
      ownersCount, membersCount, privateChannelsCount, standardChannelsCount, sharedChannelsCount
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      displayName=excluded.displayName,
      description=excluded.description,
      visibility=excluded.visibility,
      isArchived=excluded.isArchived,
      createdDateTime=excluded.createdDateTime,
      ownersCount=excluded.ownersCount,
      membersCount=excluded.membersCount,
      privateChannelsCount=excluded.privateChannelsCount,
      standardChannelsCount=excluded.standardChannelsCount,
      sharedChannelsCount=excluded.sharedChannelsCount
  `;
    await runAsync(sql, [
        row.id,
        row.displayName || '',
        row.description || '',
        row.visibility || '',
        row.isArchived ? 1 : 0,
        row.createdDateTime || '',
        row.ownersCount || 0,
        row.membersCount || 0,
        row.privateChannelsCount || 0,
        row.standardChannelsCount || 0,
        row.sharedChannelsCount || 0
    ]);
}

async function upsertTeamOwner(teamId, entry, enrich) {
    const compositeId = `${teamId}_${entry.userId || entry.email || entry.displayName || 'unknown'}`;
    const sql = `
    INSERT INTO team_owners (
      id, teamId, userId, displayName, userPrincipalName, department, jobTitle, signInStatus
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      displayName=excluded.displayName,
      userPrincipalName=excluded.userPrincipalName,
      department=excluded.department,
      jobTitle=excluded.jobTitle,
      signInStatus=excluded.signInStatus
  `;
    await runAsync(sql, [
        compositeId,
        teamId,
        entry.userId || '',
        entry.displayName || '',
        enrich.userPrincipalName || entry.email || '',
        enrich.department || '',
        enrich.jobTitle || '',
        enrich.signInStatus || ''
    ]);
}

async function upsertTeamMember(teamId, entry, enrich, role) {
    const compositeId = `${teamId}_${entry.userId || entry.email || entry.displayName || 'unknown'}`;
    const sql = `
    INSERT INTO team_members (
      id, teamId, userId, displayName, userPrincipalName, department, jobTitle, signInStatus, role
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      displayName=excluded.displayName,
      userPrincipalName=excluded.userPrincipalName,
      department=excluded.department,
      jobTitle=excluded.jobTitle,
      signInStatus=excluded.signInStatus,
      role=excluded.role
  `;
    await runAsync(sql, [
        compositeId,
        teamId,
        entry.userId || '',
        entry.displayName || '',
        enrich.userPrincipalName || entry.email || '',
        enrich.department || '',
        enrich.jobTitle || '',
        enrich.signInStatus || '',
        role
    ]);
}

/* ---------- Per-team full refresh (counts + members/owners + channels + description/isArchived) ---------- */
async function refreshTeamDetails(teamId, token) {
    // Grab previous to preserve description if we can’t fetch it this run
    const prev = await getAsync(`SELECT description, isArchived FROM teams WHERE id = ?`, [teamId]);

    // description from group (cheap); fall back to previous if not present
    let description = prev?.description || '';
    try {
        const g = await fetchJSON(`https://graph.microsoft.com/v1.0/groups/${teamId}?$select=description`, token);
        if (typeof g?.description === 'string') description = g.description;
    } catch { /* keep prev */ }

    // isArchived from teams (fall back to previous)
    let isArchived = prev?.isArchived ? 1 : 0;
    try {
        const t = await fetchJSON(`https://graph.microsoft.com/v1.0/teams/${teamId}?$select=isArchived`, token);
        isArchived = t?.isArchived ? 1 : 0;
    } catch { /* keep prev */ }

    // Owners/Members (paginated; with fallback)
    const { owners, members } = await getOwnersAndMembers(teamId, token);

    // Channel counts (no $top; follow nextLink)
    const [stdCnt, privCnt, sharedCnt] = await Promise.all([
        countChannelsByType(teamId, 'standard', token),
        countChannelsByType(teamId, 'private', token),
        countChannelsByType(teamId, 'shared', token)
    ]);

    // Update team row counts and flags
    await runAsync(`
    UPDATE teams
       SET description = ?,
           isArchived = ?,
           ownersCount = ?,
           membersCount = ?,
           privateChannelsCount = ?,
           standardChannelsCount = ?,
           sharedChannelsCount = ?
     WHERE id = ?`,
        [description, isArchived, owners.length, members.length, privCnt, stdCnt, sharedCnt, teamId]
    );

    // Refresh owner/member detail tables
    await runAsync(`DELETE FROM team_owners WHERE teamId = ?`, [teamId]);
    await runAsync(`DELETE FROM team_members WHERE teamId = ?`, [teamId]);

    // Upsert owners
    for (const o of owners) {
        const enr = await enrichFromLocalUsers(o.userId, o.email);
        await upsertTeamOwner(teamId, o, enr);
    }
    // Upsert members
    for (const m of members) {
        const enr = await enrichFromLocalUsers(m.userId, m.email);
        await upsertTeamMember(teamId, m, enr, 'member');
    }
}

/* ---------- Small promise pool for parallel post‑pass ---------- */
async function withConcurrency(items, limit, worker) {
    const queue = [...items];
    let active = 0;
    let resolveAll, rejectAll;
    const done = new Promise((res, rej) => { resolveAll = res; rejectAll = rej; });

    const next = () => {
        if (!queue.length && active === 0) return resolveAll();
        while (active < limit && queue.length) {
            const item = queue.shift();
            active++;
            worker(item).then(() => {
                active--; next();
            }).catch(err => {
                console.warn(`[post-pass] ⚠️ worker failed: ${err.message}`);
                active--; next();
            });
        }
    };
    next();
    return done;
}

/* ---------- MAIN ---------- */
async function syncTeams() {
    console.log('[syncTeams.js] 🔁 Teams delta sync started...');
    const token = await getTokenAppOnly();
    const resource = 'teams';
    const timestamp = new Date().toISOString();

    let added = 0, updated = 0, deleted = 0, deltaLink = null;

    // Load existing delta token for the "teams" resource (we use groups/delta as carrier)
    const existing = await getAsync("SELECT delta_link FROM delta_tokens WHERE resource = ?", [resource]);

    // Include description so new Teams get it on first page when present
    let url = existing?.delta_link ||
        'https://graph.microsoft.com/v1.0/groups/delta?$select=id,displayName,description,createdDateTime,visibility,groupTypes,mailEnabled,securityEnabled,resourceProvisioningOptions';

    try {
        while (url && typeof url === 'string') {
            console.log(`[syncTeams.js] 🔄 Fetching: ${url}`);
            const page = await fetchJSON(url, token);
            const rows = Array.isArray(page?.value) ? page.value : [];

            for (const grp of rows) {
                const groupId = grp.id;
                const rpo = Array.isArray(grp.resourceProvisioningOptions) ? grp.resourceProvisioningOptions : [];
                const isTeam = rpo.includes('Team');

                // Handle deletions
                if (grp['@removed']) {
                    if (isTeam) {
                        await runAsync(`DELETE FROM teams WHERE id = ?`, [groupId]);
                        await runAsync(`DELETE FROM team_members WHERE teamId = ?`, [groupId]);
                        await runAsync(`DELETE FROM team_owners WHERE teamId = ?`, [groupId]);
                        deleted++;
                        console.log(`[syncTeams.js] 🗑️ Deleted team: ${groupId}`);
                    }
                    continue;
                }

                if (!isTeam) continue;

                const existed = await getAsync(`SELECT 1 FROM teams WHERE id = ?`, [groupId]);

                // Insert/Update base metadata (counts will be filled by refresh step)
                await runAsync(`
          INSERT INTO teams (
            id, displayName, description, visibility, isArchived, createdDateTime,
            ownersCount, membersCount, privateChannelsCount, standardChannelsCount, sharedChannelsCount
          )
          VALUES (?, ?, ?, ?, 0, ?, 0, 0, 0, 0, 0)
          ON CONFLICT(id) DO UPDATE SET
            displayName=excluded.displayName,
            description=CASE WHEN excluded.description != '' THEN excluded.description ELSE teams.description END,
            visibility=excluded.visibility,
            createdDateTime=excluded.createdDateTime
        `, [
                    groupId,
                    grp.displayName || '',
                    grp.description || '',               // keep if present; else preserve current
                    grp.visibility || '',
                    grp.createdDateTime || ''
                ]);

                // Per-team refresh ensures counts + isArchived + description are correct
                await refreshTeamDetails(groupId, token);

                existed ? updated++ : added++;
                console.log(`[syncTeams.js] ✅ Synced team: ${grp.displayName || groupId}`);
            }

            if (page['@odata.nextLink']) {
                url = page['@odata.nextLink'];
            } else if (page['@odata.deltaLink']) {
                deltaLink = page['@odata.deltaLink'];
                break;
            } else {
                break;
            }
        }

        // Save delta token
        if (deltaLink) {
            await runAsync(`
        INSERT INTO delta_tokens (resource, delta_link, last_synced_at)
        VALUES (?, ?, ?)
        ON CONFLICT(resource) DO UPDATE SET
          delta_link=excluded.delta_link,
          last_synced_at=excluded.last_synced_at
      `, [resource, deltaLink, timestamp]);
        }

        // POST-PASS: refresh ALL teams (ensures older teams’ counts & flags stay current)
        const allTeams = await allAsync(`SELECT id FROM teams`);
        // Tune this concurrency to trade speed vs. API pressure
        const CONCURRENCY = 5;
        await withConcurrency(allTeams, CONCURRENCY, async (t) => {
            await refreshTeamDetails(t.id, token);
        });

        // Log
        await runAsync(`
      INSERT INTO sync_logs (resource, synced_at, added, updated, deleted, status)
      VALUES (?, ?, ?, ?, ?, ?)
    `, [resource, timestamp, added, updated, deleted, 'Success']);

        console.log(`[syncTeams.js] ✅ Delta sync complete: Added ${added}, Updated ${updated}, Deleted ${deleted}`);
    } catch (err) {
        console.error(`[syncTeams.js] ❌ Sync failed: ${err.message}`);
        await runAsync(`
      INSERT INTO sync_logs (resource, synced_at, added, updated, deleted, status, error_message)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `, [resource, timestamp, added, updated, deleted, 'Failed', err.message]);
        throw err;
    }
}

module.exports = { syncTeams };

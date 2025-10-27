




document.getElementById("loginBtn").addEventListener("click", async () => {
    try {
        const res = await fetch("/login");
        const data = await res.json();
        alert(data.message);
    } catch (e) {
        alert("Login failed.");
    }
});

document.getElementById("logoutBtn").addEventListener("click", () => {
    localStorage.clear();
    alert("Logged out");
});

document.getElementById("syncBtn").addEventListener("click", async () => {
    document.getElementById("statusMessage").innerText = "⏳ Syncing...";
    try {
        const res = await fetch("http://localhost:3001/sync");
        const result = await res.text();
        document.getElementById("statusMessage").innerText = result;
    } catch (err) {
        console.error("Sync failed", err);
        document.getElementById("statusMessage").innerText = "❌ Sync failed.";
    }
});

document.getElementById("syncBtnGroups").addEventListener("click", async () => {
    document.getElementById("statusMessage").innerText = "⏳ Syncing...";
    try {
        const res = await fetch("http://localhost:3001/syncGroups");
        const result = await res.text();
        document.getElementById("statusMessage").innerText = result;
    } catch (err) {
        console.error("Sync failed", err);
        document.getElementById("statusMessage").innerText = "❌ Sync failed.";
    }
});

document.getElementById("syncBtnTeams").addEventListener("click", async () => {
    document.getElementById("statusMessage").innerText = "⏳ Syncing...";
    try {
        const res = await fetch("http://localhost:3001/sync-teams");
        const result = await res.text();
        document.getElementById("statusMessage").innerText = result;
    } catch (err) {
        console.error("Sync failed", err);
        document.getElementById("statusMessage").innerText = "❌ Sync failed.";
    }
});


// ------------------ CONFIG ------------------

const REPORT_FILTERS = {
    all: ["department", "jobTitle", "signInStatus", "licenseStatus"],
    enabled: ["department", "jobTitle", "licenseStatus"],
    disabled: ["department", "jobTitle", "licenseStatus"],
    licensed: ["department", "jobTitle", "signInStatus"],
    unlicensed: ["department", "jobTitle", "signInStatus"],
    "groups/distribution": ["membersCount"],
    "groups/security-enabled": ["membersCount"],
    "groups/mail-enabled-security": ["membersCount"],
    "groups/empty": ["groupTypes"],
    "groups/recently-created": ["securityEnabled", "mailEnabled"],
    "groups/all": ["groupTypes", "securityEnabled", "mailEnabled"],
    "groups/unified": ["visibility", "groupTypes", "membersCount"],
    "groups/members": ["groupTypes"],
    "groups/owners": ["groupTypes"],
    "teams/all": ["visibility", "isArchived"],
    "teams/public": ["isArchived"],
    "teams/private": ["isArchived"],
    "teams/hidden-memberships": ["isArchived"],
    "teams/archived": ["visibility"],
    "teams/teams-without-description": ["isArchived"],
    "teams/teams-private-channels": ["visibility", "isArchived"],
    "teams/teams-shared-channels": ["visibility", "isArchived"],
    "teams/recently-created-teams": ["isArchived"],
    "teams/team-owners": ["isArchived"]
};

// ✅ User report routes that need hydrated department & jobTitle
const USER_ROUTES = new Set(["all", "enabled", "disabled", "licensed", "unlicensed"]);

let currentRoute = "all";
let currentSearch = "";
let currentFilters = {};
let lastFetchedRecords = [];

/** Routes where we show the Group Display Name selector */
const GROUP_NAME_ROUTES = [
    "groups/all",
    "groups/unified",
    "groups/distribution",
    "groups/security-enabled",
    "groups/mail-enabled-security",
    "groups/recently-created"
];

/** Routes where membersCount should be clickable */
const MEMBER_COUNT_ROUTES = [
    "groups/unified",
    "groups/distribution",
    "groups/security-enabled",
    "groups/mail-enabled-security",
    "groups/all",
    "groups/recently-created"
];

// Teams routes where memberCount / ownerCount should be clickable
const TEAM_COUNT_ROUTES = [
    "teams/all",
    "teams/public",
    "teams/private",
    "teams/hidden-memberships",
    "teams/teams-without-description",
    "teams/archived",
    "teams/teams-private-channels",
    "teams/teams-shared-channels",
    "teams/recently-created-teams"
];

/** Two-entity dropdowns (Group + Person) config per route */
const TWO_ENTITY_FILTER_ROUTES = {
    "groups/members": {
        leftLabel: "Select a group",
        rightLabel: "Select a member",
        // left => submit group displayName; right => submit userPrincipalName
        getPair(route, which, r) {
            if (which === "left") {
                const text = r.groupName || r.displayName;
                return text ? { text, value: text } : null;
            } else {
                const text = r.memberName || r.userPrincipalName;
                const value = r.userPrincipalName;
                return value ? { text, value } : null;
            }
        }
    },
    "groups/owners": {
        leftLabel: "Select a group",
        rightLabel: "Select an owner",
        getPair(route, which, r) {
            if (which === "left") {
                const text = r.groupDisplayName || r.displayName;
                return text ? { text, value: text } : null;
            } else {
                const text = r.ownerDisplayName || r.userPrincipalName;
                const value = r.userPrincipalName;
                return value ? { text, value } : null;
            }
        }
    },
    "groups/disabled-members": {
        leftLabel: "Select a group",
        rightLabel: "Select a disabled member",
        getPair(route, which, r) {
            if (which === "left") {
                const text = r.groupDisplayName || r.displayName;
                return text ? { text, value: text } : null;
            } else {
                const text = r.memberDisplayName || r.userPrincipalName;
                const value = r.userPrincipalName;
                return value ? { text, value } : null;
            }
        }
    }
};


// ------------------ FILTER UI ------------------

function createFilterUI(records, route) {
    const container = document.createElement("div");
    container.className = "mb-3";

    // Text search
    const searchInput = document.createElement("input");
    searchInput.type = "text";
    searchInput.className = "form-control mb-2";
    searchInput.placeholder = "Search....";
    searchInput.id = "searchBox";
    searchInput.value = currentSearch;
    container.appendChild(searchInput);

    // --- Two-entity filters (Group + Person) for specific routes ---
    let entityLeftSelect = null;
    let entityRightSelect = null;

    if (TWO_ENTITY_FILTER_ROUTES[route]) {
        const { leftLabel, rightLabel } = TWO_ENTITY_FILTER_ROUTES[route];

        // Left = Group selector
        entityLeftSelect = document.createElement("select");
        entityLeftSelect.className = "form-select mb-2";
        entityLeftSelect.id = "entityLeftSelect";
        entityLeftSelect.setAttribute("aria-label", leftLabel);
        const phL = document.createElement("option");
        phL.value = "";
        phL.textContent = leftLabel;
        entityLeftSelect.appendChild(phL);
        container.appendChild(entityLeftSelect);
        populateTwoEntitySelect(entityLeftSelect, route, "left");

        // Right = Person selector
        entityRightSelect = document.createElement("select");
        entityRightSelect.className = "form-select mb-2";
        entityRightSelect.id = "entityRightSelect";
        entityRightSelect.setAttribute("aria-label", rightLabel);
        const phR = document.createElement("option");
        phR.value = "";
        phR.textContent = rightLabel;
        entityRightSelect.appendChild(phR);
        container.appendChild(entityRightSelect);
        populateTwoEntitySelect(entityRightSelect, route, "right");

        // UX: selecting one clears the other
        entityLeftSelect.addEventListener("change", () => {
            if (entityLeftSelect.value) entityRightSelect.value = "";
        });
        entityRightSelect.addEventListener("change", () => {
            if (entityRightSelect.value) entityLeftSelect.value = "";
        });
    }

    // --- Group Display Name selector for group routes (existing feature) ---
    let groupSelect = null;
    if (typeof GROUP_NAME_ROUTES !== "undefined" && GROUP_NAME_ROUTES.includes(route)) {
        groupSelect = document.createElement("select");
        groupSelect.className = "form-select mb-2";
        groupSelect.id = "groupSelect";
        groupSelect.setAttribute("aria-label", "Group Display Name");

        const ph = document.createElement("option");
        ph.value = "";
        ph.textContent = "Select a group";
        groupSelect.appendChild(ph);

        container.appendChild(groupSelect);
        populateGroupSelect(groupSelect, route); // hydrates from page=all
    }

    // --- Route-specific filters (existing & patched) ---
    const filters = REPORT_FILTERS[route] || [];
    filters.forEach((field) => {
        const select = document.createElement("select");
        select.className = "form-select mb-2 filter-select";
        select.setAttribute("data-field", field);

        const option = document.createElement("option");
        option.value = "";
        option.textContent = `Filter by ${field}`;
        select.appendChild(option);

        // ✅ HYDRATE user facets (department/jobTitle) across the tenant for user routes
        if (USER_ROUTES.has(route) && (field === "department" || field === "jobTitle")) {
            (async () => {
                try {
                    await populateUserFacetOptions(select, route, field);
                } catch (e) {
                    console.error(`Failed to populate ${field} options:`, e);
                }
            })();
        }
        // ✅ Existing hydration for groupTypes (groups reports)
        else if (field === "groupTypes") {
            (async () => {
                try {
                    await populateGroupTypesOptions(select, route);
                } catch (e) {
                    console.error("Failed to populate groupTypes options:", e);
                }
            })();
        }
        // Fallback: use current page records (legacy behavior)
        else {
            const uniqueValues =
                field === "membersCount"
                    ? ["true", "false"]
                    : (field === "securityEnabled" || field === "mailEnabled")
                        ? ["true", "false"]
                        : (field === "visibility" && route === "groups/unified")
                            ? ["Public", "Private", "HiddenMembership"]
                            : [...new Set(
                                records
                                    .map(r => r[field])
                                    .filter(v => v !== undefined && v !== null)
                            )].sort();

            uniqueValues.forEach(val => {
                const opt = document.createElement("option");
                opt.value = val;
                opt.textContent =
                    field === "membersCount"
                        ? (val === "true" ? "Groups with Members" : "Groups with No Members")
                        : val;
                if (currentFilters[field] === val) opt.selected = true;
                select.appendChild(opt);
            });
        }

        container.appendChild(select);
    });

    // --- Buttons row ---
    const btnRow = document.createElement("div");
    btnRow.className = "d-flex gap-2 mt-2";

    const applyBtn = document.createElement("button");
    applyBtn.textContent = "Search";
    applyBtn.className = "btn btn-primary";
    applyBtn.onclick = () => {
        // Priority: two-entity routes first
        if (TWO_ENTITY_FILTER_ROUTES[route]) {
            const leftVal = document.getElementById("entityLeftSelect")?.value || "";
            const rightVal = document.getElementById("entityRightSelect")?.value || "";

            if (leftVal) {
                // group-based search: submit displayName
                currentSearch = leftVal;
                const sb = document.getElementById("searchBox"); if (sb) sb.value = "";
            } else if (rightVal) {
                // person-based search: submit userPrincipalName
                currentSearch = rightVal;
                const sb = document.getElementById("searchBox"); if (sb) sb.value = "";
            } else {
                // fall back to text or group-id dropdown (if present)
                const selectedGroupId = document.getElementById("groupSelect")?.value || "";
                currentSearch = selectedGroupId || document.getElementById("searchBox").value.trim();
            }
        } else {
            // Existing behavior for other routes
            const selectedGroupId = document.getElementById("groupSelect")?.value || "";
            if (selectedGroupId) {
                currentSearch = selectedGroupId; // exact id match
                const sb = document.getElementById("searchBox"); if (sb) sb.value = "";
            } else {
                currentSearch = document.getElementById("searchBox").value.trim();
            }
        }

        currentFilters = {};
        document.querySelectorAll(".filter-select").forEach(el => {
            const val = el.value.trim();
            if (val) currentFilters[el.getAttribute("data-field")] = val;
        });

        fetchReport(route, 1);
    };

    const clearBtn = document.createElement("button");
    clearBtn.textContent = "Clear";
    clearBtn.className = "btn btn-outline-secondary";
    clearBtn.onclick = () => {
        currentSearch = "";
        currentFilters = {};
        const gs = document.getElementById("groupSelect"); if (gs) gs.value = "";
        const ls = document.getElementById("entityLeftSelect"); if (ls) ls.value = "";
        const rs = document.getElementById("entityRightSelect"); if (rs) rs.value = "";
        fetchReport(route, 1);
    };

    btnRow.appendChild(applyBtn);
    btnRow.appendChild(clearBtn);
    container.appendChild(btnRow);

    return container;
}


// ------------------ HELPERS ------------------

function createPagination(currentPage, totalPages, onPageClick) {
    const wrapper = document.createElement("div");
    wrapper.className = "d-flex justify-content-center flex-wrap mt-4";

    const pagination = document.createElement("nav");
    pagination.setAttribute("aria-label", "Report navigation");

    const ul = document.createElement("ul");
    ul.className = "pagination justify-content-center flex-wrap";

    const prevLi = document.createElement("li");
    prevLi.className = `page-item ${currentPage === 1 ? "disabled" : ""}`;
    prevLi.innerHTML = `<a class="page-link" href="#">Prev</a>`;
    prevLi.onclick = (e) => {
        e.preventDefault();
        if (currentPage > 1) onPageClick(currentPage - 1);
    };
    ul.appendChild(prevLi);

    for (let i = 1; i <= totalPages; i++) {
        const li = document.createElement("li");
        li.className = `page-item ${i === currentPage ? "active" : ""}`;
        li.innerHTML = `<a class="page-link" href="#">${i}</a>`;
        li.onclick = (e) => {
            e.preventDefault();
            onPageClick(i);
        };
        ul.appendChild(li);
    }

    const nextLi = document.createElement("li");
    nextLi.className = `page-item ${currentPage === totalPages ? "disabled" : ""}`;
    nextLi.innerHTML = `<a class="page-link" href="#">Next</a>`;
    nextLi.onclick = (e) => {
        e.preventDefault();
        if (currentPage < totalPages) onPageClick(currentPage + 1);
    };
    ul.appendChild(nextLi);

    pagination.appendChild(ul);
    wrapper.appendChild(pagination);
    return wrapper;
}

function buildQueryParams(page = 1) {
    const params = new URLSearchParams();
    if (page !== null) params.append("page", page); // allow 'all'

    if (currentSearch) params.append("search", currentSearch);
    for (const [field, value] of Object.entries(currentFilters)) {
        params.append(field, value);
    }

    return params.toString();
}

/** Hydrate the Group Display Name selector with ALL records in-context (filters respected, no text search) */
async function populateGroupSelect(selectEl, route) {
    try {
        const params = new URLSearchParams();
        params.append("page", "all");
        // Respect currentFilters so the list reflects the current report context
        for (const [field, value] of Object.entries(currentFilters)) {
            if (value !== undefined && value !== null && `${value}`.trim() !== "") {
                params.append(field, value);
            }
        }

        const url = `http://localhost:3001/report/${route}?${params.toString()}`;
        const res = await fetch(url);
        const data = await res.json();
        if (!data?.records?.length) return;

        const seen = new Set();
        data.records.forEach(r => {
            if (!r || !r.id || !r.displayName) return;
            if (seen.has(r.id)) return;
            seen.add(r.id);

            const opt = document.createElement("option");
            opt.value = r.id;                 // submit ID
            opt.textContent = r.displayName;  // show display name
            selectEl.appendChild(opt);
        });
    } catch (e) {
        console.error("Failed to populate group select:", e);
    }
}

/** Hydrate groupTypes filter with ALL records in-context (respect filters except groupTypes, ignore text search) */
async function populateGroupTypesOptions(selectEl, route) {
    try {
        const params = new URLSearchParams();
        params.append("page", "all");

        // Respect other filters, but DO NOT include groupTypes itself (to avoid self-filtering)
        for (const [field, value] of Object.entries(currentFilters)) {
            if (field === "groupTypes") continue;
            if (value !== undefined && value !== null && `${value}`.trim() !== "") {
                params.append(field, value);
            }
        }

        // Intentionally do not include currentSearch
        const url = `http://localhost:3001/report/${route}?${params.toString()}`;
        const res = await fetch(url);
        const data = await res.json();
        if (!data?.records?.length) return;

        const set = new Set();
        data.records.forEach(r => {
            if (!r || r.groupTypes == null) return;
            set.add(String(r.groupTypes));
        });

        [...set].sort().forEach(val => {
            const opt = document.createElement("option");
            opt.value = val;
            opt.textContent = val;
            if (currentFilters["groupTypes"] === val) opt.selected = true;
            selectEl.appendChild(opt);
        });
    } catch (e) {
        console.error("Failed to populate groupTypes options:", e);
    }
}


/** Generic hydrator: populate a "Group/Person" select for two-entity routes (page=all; respect filters; ignore text search) */
async function populateTwoEntitySelect(selectEl, route, which) {
    try {
        const cfg = TWO_ENTITY_FILTER_ROUTES[route];
        if (!cfg) return;

        const params = new URLSearchParams();
        params.append("page", "all");
        // Respect other active filters (don’t include currentSearch)
        for (const [field, value] of Object.entries(currentFilters)) {
            if (value !== undefined && value !== null && `${value}`.trim() !== "") {
                params.append(field, value);
            }
        }

        const url = `http://localhost:3001/report/${route}?${params.toString()}`;
        const res = await fetch(url);
        const data = await res.json();
        if (!data?.records?.length) return;

        const map = new Map(); // dedupe by value
        data.records.forEach(r => {
            const pair = cfg.getPair(route, which, r);
            if (!pair || !pair.value) return;
            if (!map.has(pair.value)) map.set(pair.value, pair.text || pair.value);
        });

        [...map.entries()]
            .sort((a, b) => String(a[1]).localeCompare(String(b[1])))
            .forEach(([value, text]) => {
                const opt = document.createElement("option");
                opt.value = value;
                opt.textContent = text;
                selectEl.appendChild(opt);
            });
    } catch (e) {
        console.error("Failed to populate two-entity select:", e);
    }
}



/** ✅ Hydrate department / jobTitle options with ALL tenant values for user routes */
async function populateUserFacetOptions(selectEl, route, field /* 'department' | 'jobTitle' */) {
    const params = new URLSearchParams();
    params.append("page", "all");

    // Respect other active filters, but DO NOT include this field (avoid self-filtering)
    for (const [f, v] of Object.entries(currentFilters)) {
        if (f === field) continue;
        if (v !== undefined && v !== null && `${v}`.trim() !== "") {
            params.append(f, v);
        }
    }

    // Intentionally ignore currentSearch so the facet list is complete for the report context
    const url = `http://localhost:3001/report/${route}?${params.toString()}`;
    const res = await fetch(url);
    const data = await res.json();
    if (!data?.records?.length) return;

    const values = new Set();
    let hasEmpty = false;

    data.records.forEach(r => {
        const val = r?.[field];
        if (val === undefined || val === null || `${val}`.trim() === "") {
            hasEmpty = true;
            return;
        }
        values.add(String(val));
    });

    // Optional: expose a filter for empty values, if present
    if (hasEmpty) {
        const optEmpty = document.createElement("option");
        optEmpty.value = "__EMPTY__";
        optEmpty.textContent = "(No value)";
        if (currentFilters[field] === "__EMPTY__") optEmpty.selected = true;
        selectEl.appendChild(optEmpty);
    }

    [...values].sort((a, b) => a.localeCompare(b)).forEach(val => {
        const opt = document.createElement("option");
        opt.value = val;
        opt.textContent = val;
        if (currentFilters[field] === val) opt.selected = true;
        selectEl.appendChild(opt);
    });
}

function renderTable(data, containerId, route, onPageClick) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    const { records, page, totalPages } = data;
    lastFetchedRecords = records;

    if (!records || records.length === 0) {
        container.innerHTML = "<p>No data available.</p>";
        return;
    }

    container.appendChild(createFilterUI(records, route));

    // Responsive wrapper
    const tableWrapper = document.createElement("div");
    tableWrapper.className = "table-responsive";

    const table = document.createElement("table");
    table.className = "table table-bordered table-striped mb-4";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    const headers = Object.keys(records[0]);

    headers.forEach(key => {
        const th = document.createElement("th");
        th.textContent = key
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/^./, str => str.toUpperCase());
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    records.forEach(record => {
        const row = document.createElement("tr");
        headers.forEach(key => {
            const td = document.createElement("td");

            if (TEAM_COUNT_ROUTES.includes(route) && (key === "membersCount" || key === "ownersCount")) {
                const count = record[key] ?? 0;
                const teamId = record.id;
                if (count > 0 && teamId) {
                    const cls = key === "membersCount" ? "team-members-trigger" : "team-owners-trigger";
                    td.innerHTML = `<button class="btn btn-link p-0 ${cls}" data-teamid="${teamId}">${count}</button>`;
                } else {
                    td.innerHTML = `<span class="text-muted">${count}</span>`;
                }
            } else if (key === "membersCount" && MEMBER_COUNT_ROUTES.includes(route)) {
                // Existing GROUPS behavior
                const count = record[key] ?? 0;
                if (count > 0 && record.id) {
                    td.innerHTML = `<button class="btn btn-link p-0 member-trigger" data-groupid="${record.id}">${count}</button>`;
                } else {
                    td.innerHTML = `<span class="text-muted">${count}</span>`;
                }
            } else if (key === "ownersCount" && MEMBER_COUNT_ROUTES.includes(route)) {
                const count = record[key] ?? 0;
                if (count > 0 && record.id) {
                    td.innerHTML = `<button class="btn btn-link p-0 owner-trigger" data-groupid="${record.id}">${count}</button>`;
                } else {
                    td.innerHTML = `<span class="text-muted">${count}</span>`;
                }
            } else {
                td.textContent = record[key] ?? "";
            }

            row.appendChild(td);
        });
        tbody.appendChild(row);
    });

    table.appendChild(tbody);
    tableWrapper.appendChild(table);
    container.appendChild(tableWrapper);

    // Action buttons
    const buttonRow = document.createElement("div");
    buttonRow.className = "d-flex gap-2 mb-3";

    const downloadBtn = document.createElement("button");
    downloadBtn.className = "btn btn-outline-primary";
    downloadBtn.textContent = "Download CSV";
    downloadBtn.onclick = () => fetchFullReportData("download");

    const emailBtn = document.createElement("button");
    emailBtn.className = "btn btn-outline-success";
    emailBtn.textContent = "Email Report";
    emailBtn.onclick = () => fetchFullReportData("email");

    buttonRow.appendChild(downloadBtn);
    buttonRow.appendChild(emailBtn);
    container.appendChild(buttonRow);

    // Pagination
    container.appendChild(createPagination(page, totalPages, onPageClick));
}


// ------------------ MODALS & CLICK HANDLERS ------------------

document.addEventListener("click", async (e) => {
    if (e.target && e.target.classList.contains("member-trigger")) {
        const groupId = e.target.getAttribute("data-groupid");
        if (!groupId) return;

        try {
            const res = await fetch(`http://localhost:3001/report/groups/members?search=${groupId}&page=all`);
            const data = await res.json();
            if (!data || !data.records || data.records.length === 0) {
                alert("No members found for this group.");
                return;
            }
            showGroupMembersPopup(data.records);
        } catch (err) {
            console.error("Error fetching group members", err);
            alert("Failed to fetch group members.");
        }
    }
});

// GROUPS: owner count -> open owners modal
document.addEventListener("click", async (e) => {
    if (e.target && e.target.classList.contains("owner-trigger")) {
        const groupId = e.target.getAttribute("data-groupid");
        if (!groupId) return;

        try {
            const res = await fetch(`http://localhost:3001/report/groups/owners?search=${groupId}&page=all`);
            const data = await res.json();
            if (!data?.records?.length) {
                alert("No owners found for this group.");
                return;
            }
            showTeamPeoplePopup("Group Owners", data.records, "ownerDisplayName");
        } catch (err) {
            console.error("Error fetching group owners", err);
            alert("Failed to fetch group owners.");
        }
    }
});

// TEAMS: member count -> open members modal
document.addEventListener("click", async (e) => {
    if (e.target && e.target.classList.contains("team-members-trigger")) {
        const teamId = e.target.getAttribute("data-teamid");
        if (!teamId) return;
        try {
            const res = await fetch(`http://localhost:3001/report/teams/teams-members?search=${teamId}&page=all`);
            const data = await res.json();
            if (!data?.records?.length) {
                alert("No team members found for this team.");
                return;
            }
            showTeamPeoplePopup("Team Members", data.records, "memberDisplayName");
        } catch (err) {
            console.error("Error fetching team members", err);
            alert("Failed to fetch team members.");
        }
    }
});

// TEAMS: owner count -> open owners modal
document.addEventListener("click", async (e) => {
    if (e.target && e.target.classList.contains("team-owners-trigger")) {
        const teamId = e.target.getAttribute("data-teamid");
        if (!teamId) return;
        try {
            const res = await fetch(`http://localhost:3001/report/teams/teams-owners?page=all&search=${teamId}`);
            const data = await res.json();
            if (!data?.records?.length) {
                alert("No team owners found for this team.");
                return;
            }
            showTeamPeoplePopup("Team Owners", data.records, "ownerDisplayName");
        } catch (err) {
            console.error("Error fetching team owners", err);
            alert("Failed to fetch team owners.");
        }
    }
});

async function fetchFullReportData(mode = "download") {
    const params = buildQueryParams("all"); // pass page='all'
    const url = `http://localhost:3001/report/${currentRoute}?${params}`;
    try {
        const res = await fetch(url);
        const fullData = await res.json();
        if (!fullData.records || fullData.records.length === 0) {
            alert("❌ No records found.");
            return;
        }

        if (mode === "download") {
            await triggerCSVDownload(fullData.records);
        } else if (mode === "email") {
            await triggerEmailSend(fullData.records);
        }
    } catch (err) {
        alert("❌ Failed to retrieve full report data.");
        console.error(err);
    }
}

async function triggerCSVDownload(data) {
    try {
        const res = await fetch("http://localhost:3001/report/download", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ data })
        });

        if (!res.ok) throw new Error("Failed to generate CSV.");

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "report.csv";
        a.click();
        window.URL.revokeObjectURL(url);
    } catch (err) {
        alert("❌ CSV download failed.");
        console.error(err);
    }
}

async function triggerEmailSend(data) {
    const recipient = prompt("Enter recipient email:");
    if (!recipient) return;

    try {
        const res = await fetch("http://localhost:3001/report/email", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ data, recipient })
        });

        const result = await res.json();
        if (res.ok) {
            alert("✅ " + result.message);
        } else {
            alert("❌ " + result.message);
        }
    } catch (err) {
        alert("❌ Failed to send report.");
        console.error(err);
    }
}

function fetchReport(route, page = 1) {
    currentRoute = route;
    const params = buildQueryParams(page);
    fetch(`http://localhost:3001/report/${route}?${params}`)
        .then(res => res.json())
        .then(data => renderTable(data, "reportContent", route, (newPage) => fetchReport(route, newPage)))
        .catch(err => {
            console.error("Fetch error:", err);
            document.getElementById("reportContent").innerHTML = "<p class='text-danger'>Failed to load report.</p>";
        });
}


// ------------------ REPORT BUTTONS (unchanged) ------------------

document.getElementById("allUsersBtn").addEventListener("click", () => {
    currentSearch = "";
    currentFilters = {};
    fetchReport("all");
    document.getElementById("reportName").innerText = "All Users Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, User Prinicipal Name, First Name, Last Name, Mail Nick Name, Email";
});
document.getElementById("disabledUsersBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("disabled");
    document.getElementById("reportName").innerText = "Disabled Users Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, User Prinicipal Name, First Name, Last Name, Mail Nick Name, Email";
});
document.getElementById("enabledUsersBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("enabled");
    document.getElementById("reportName").innerText = "Enabled Users Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, User Prinicipal Name, First Name, Last Name, Mail Nick Name, Email";
});
document.getElementById("licensedUsersBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("licensed");
    document.getElementById("reportName").innerText = "Licensed Users Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, User Prinicipal Name, First Name, Last Name, Mail Nick Name, Email";
});
document.getElementById("unlicensedUsersBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("unlicensed");
    document.getElementById("reportName").innerText = "Unlicensed Users Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, User Prinicipal Name, First Name, Last Name, Mail Nick Name, Email";
});

document.getElementById("distributionGroupsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/distribution");
    document.getElementById("reportName").innerText = "Distribution Groups Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, Mail";
});

document.getElementById("securityGroupsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/security-enabled");
    document.getElementById("reportName").innerText = "Security Groups Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, Mail";
});

document.getElementById("mailEnabledSecurityGroupsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/mail-enabled-security");
    document.getElementById("reportName").innerText = "Mail-Enabled Security Groups Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, Mail";
});

document.getElementById("emptyGroupsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/empty");
    document.getElementById("reportName").innerText = "Empty Groups Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, Mail";
});

document.getElementById("recentlyCreatedGroupsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/recently-created");
    document.getElementById("reportName").innerText = "Recently Created Groups Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, Mail";
});

document.getElementById("allGroupsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/all");
    document.getElementById("reportName").innerText = "All Groups Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, Mail";
});

document.getElementById("unifiedGroupsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/unified");
    document.getElementById("reportName").innerText = "Unified Groups Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name, Mail";
});

document.getElementById("groupMembersBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/members");
    document.getElementById("reportName").innerText = "Group Members Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Group Name, User Principal Name";
});

document.getElementById("btn-groups-owners").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/owners");
    document.getElementById("reportName").innerText = "Group Owners Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Group Display Name, User Principal Name";
});

document.getElementById("disabled-user-groups").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("groups/disabled-members");
    document.getElementById("reportName").innerText = "Groups With Disabled Users Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Group Display Name, User Principal Name";
});

document.getElementById("allTeamsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/all");
    document.getElementById("reportName").innerText = "All Teams Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("publicTeamsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/public");
    document.getElementById("reportName").innerText = "Public Teams Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("privateTeamsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/private");
    document.getElementById("reportName").innerText = "Private Teams Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("hiddenTeamsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/hidden-memberships");
    document.getElementById("reportName").innerText = "Hidden Memberships Teams Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("archivedTeamsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/archived");
    document.getElementById("reportName").innerText = "Archived Teams Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("nodescTeamsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/teams-without-description");
    document.getElementById("reportName").innerText = "Teams With No Description Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("teamsPvtChannels").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/teams-private-channels");
    document.getElementById("reportName").innerText = "Teams With Private Channels Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("teamsSharedChannels").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/teams-shared-channels");
    document.getElementById("reportName").innerText = "Teams With Shared Channels Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("recentTeamsBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/recently-created-teams");
    document.getElementById("reportName").innerText = "Recently Created Teams Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Display Name";
});

document.getElementById("teamOwnersBtn").addEventListener("click", () => {
    document.getElementById("sfields").innerHTML = "";
    currentSearch = "";
    currentFilters = {};
    fetchReport("teams/teams-owners");
    document.getElementById("reportName").innerText = "Team Owners Report";
    document.getElementById("sfields").innerHTML = "<b>Searchable Fields</b>: Team Display Name";
});


// ------------------ POPUPS & SHUTDOWN ------------------

function showGroupMembersPopup(members) {
    // Remove any existing modal
    const existing = document.getElementById("memberModal");
    if (existing) existing.remove();

    // Apply blur to background
    document.getElementById("reportContent").style.filter = "blur(5px)";

    // Create modal wrapper
    const modal = document.createElement("div");
    modal.id = "memberModal";
    modal.style.position = "fixed";
    modal.style.top = "0";
    modal.style.left = "0";
    modal.style.width = "100vw";
    modal.style.height = "100vh";
    modal.style.backgroundColor = "rgba(0, 0, 0, 0.4)";
    modal.style.display = "flex";
    modal.style.alignItems = "center";
    modal.style.justifyContent = "center";
    modal.style.zIndex = "9999";
    modal.style.padding = "10px"; // for mobile padding

    // Inner modal content
    const modalContent = document.createElement("div");
    modalContent.style.background = "#fff";
    modalContent.style.width = "100%";
    modalContent.style.maxWidth = "800px";
    modalContent.style.maxHeight = "80vh";
    modalContent.style.overflowY = "auto";
    modalContent.style.borderRadius = "8px";
    modalContent.style.padding = "20px";
    modalContent.style.position = "relative";
    modalContent.style.boxShadow = "0 4px 12px rgba(0,0,0,0.3)";

    // Close button
    const closeBtn = document.createElement("button");
    closeBtn.textContent = "×";
    closeBtn.style.position = "absolute";
    closeBtn.style.top = "10px";
    closeBtn.style.right = "15px";
    closeBtn.style.fontSize = "28px";
    closeBtn.style.fontWeight = "bold";
    closeBtn.style.color = "#dc3545"; // Bootstrap red
    closeBtn.title = "Close";
    closeBtn.style.border = "none";
    closeBtn.style.background = "transparent";
    closeBtn.style.cursor = "pointer";
    closeBtn.onclick = closeModal;

    // Build table
    const table = document.createElement("table");
    table.className = "table table-bordered table-sm";
    table.innerHTML = `
    <thead>
      <tr>
        <th>Member Name</th>
        <th>User Principal Name</th>
        <th>Department</th>
        <th>Job Title</th>
        <th>Sign-In Status</th>
      </tr>
    </thead>
    <tbody>
      ${members.map(member => `
        <tr>
          <td>${member.memberName || ""}</td>
          <td>${member.userPrincipalName || ""}</td>
          <td>${member.department || ""}</td>
          <td>${member.jobTitle || ""}</td>
          <td>${member.signInStatus || ""}</td>
        </tr>
      `).join("")}
    </tbody>
  `;

    modalContent.appendChild(closeBtn);
    modalContent.appendChild(table);
    modal.appendChild(modalContent);
    document.body.appendChild(modal);

    // Clicking outside modalContent closes the modal
    modal.addEventListener("click", (e) => {
        if (e.target === modal) closeModal();
    });

    function closeModal() {
        document.getElementById("reportContent").style.filter = "none";
        modal.remove();
    }
}

function showTeamPeoplePopup(title, rows, nameKey /* "memberDisplayName" | "ownerDisplayName" */) {
    // Remove any existing modal
    const existing = document.getElementById("teamPeopleModal");
    if (existing) existing.remove();

    // Blur background content only
    const reportEl = document.getElementById("reportContent");
    if (reportEl) reportEl.style.filter = "blur(5px)";

    // Overlay
    const modal = document.createElement("div");
    modal.id = "teamPeopleModal";
    modal.style.position = "fixed";
    modal.style.top = "0";
    modal.style.left = "0";
    modal.style.width = "100vw";
    modal.style.height = "100vh";
    modal.style.backgroundColor = "rgba(0,0,0,0.4)";
    modal.style.display = "flex";
    modal.style.alignItems = "center";
    modal.style.justifyContent = "center";
    modal.style.zIndex = "9999";
    modal.style.padding = "10px";

    // Content
    const panel = document.createElement("div");
    panel.style.background = "#fff";
    panel.style.width = "100%";
    panel.style.maxWidth = "900px";
    panel.style.maxHeight = "80vh";
    panel.style.overflowY = "auto";
    panel.style.borderRadius = "8px";
    panel.style.padding = "20px";
    panel.style.position = "relative";
    panel.style.boxShadow = "0 4px 12px rgba(0,0,0,0.3)";

    const header = document.createElement("div");
    header.className = "d-flex align-items-center justify-content-between mb-3";
    const h = document.createElement("h5");
    h.textContent = title;
    h.className = "m-0";
    const x = document.createElement("button");
    x.textContent = "×";
    x.style.fontSize = "28px";
    x.style.fontWeight = "bold";
    x.style.color = "#dc3545";
    x.style.border = "none";
    x.style.background = "transparent";
    x.style.cursor = "pointer";
    x.title = "Close";
    x.onclick = closeModal;
    header.appendChild(h);
    header.appendChild(x);

    const tableWrap = document.createElement("div");
    tableWrap.className = "table-responsive";
    const table = document.createElement("table");
    table.className = "table table-bordered table-sm";
    table.innerHTML = `
    <thead>
      <tr>
        <th>Name</th>
        <th>User Principal Name</th>
        <th>Department</th>
        <th>Job Title</th>
        <th>Sign-In Status</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(r => `
        <tr>
          <td>${r[nameKey] || ""}</td>
          <td>${r.userPrincipalName || ""}</td>
          <td>${r.department || ""}</td>
          <td>${r.jobTitle || ""}</td>
          <td>${r.signInStatus || ""}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
    tableWrap.appendChild(table);

    panel.appendChild(header);
    panel.appendChild(tableWrap);
    modal.appendChild(panel);
    document.body.appendChild(modal);

    // Click outside closes
    modal.addEventListener("click", (e) => {
        if (e.target === modal) closeModal();
    });

    function closeModal() {
        if (reportEl) reportEl.style.filter = "none";
        modal.remove();
    }
}

window.addEventListener("beforeunload", function () {
    navigator.sendBeacon("/shutdown");
});


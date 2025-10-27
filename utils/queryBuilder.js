



// utils/queryBuilder.js -- Enhanced Version 2

function buildQuery(reqQuery, options = {}) {
    const whereParts = [];
    const values = [];

    const {
        searchFields = [],
        filters = {},
        additionalConditions = []
    } = options;

    // Search logic
    if (reqQuery.search && searchFields.length > 0) {
        const searchConditions = searchFields.map(field => `${field} LIKE ?`).join(' OR ');
        whereParts.push(`(${searchConditions})`);
        searchFields.forEach(() => values.push(`%${reqQuery.search}%`));
    }

    // Filter logic
    Object.entries(filters).forEach(([field, paramName]) => {
        let value = reqQuery[paramName || field];
        if (value !== undefined) {
            // Normalize visibility values to match DB format
            if (field === 'visibility') {
                value = normalizeVisibility(value);
            }

            // Handle boolean filtering for membersCount
            if (field === 'membersCount') {
                if (value === 'true') {
                    whereParts.push(`membersCount > 0`);
                } else if (value === 'false') {
                    whereParts.push(`membersCount = 0`);
                }
            } else {
                whereParts.push(`${field} = ?`);
                values.push(value);
            }
        }
    });

    // Add any fixed WHERE conditions
    if (additionalConditions.length > 0) {
        whereParts.push(...additionalConditions);
    }

    const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';
    return { whereClause, values };
}

// Utility to normalize visibility filters
function normalizeVisibility(input) {
    const map = {
        public: 'Public',
        private: 'Private',
        hiddenmembership: 'HiddenMembership'
    };

    return map[input.toLowerCase()] || input;
}




module.exports = { buildQuery };


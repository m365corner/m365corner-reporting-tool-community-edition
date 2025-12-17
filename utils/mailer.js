const nodemailer = require("nodemailer");

const REPORT_META = {
    // ============================
    // USER REPORTS
    // ============================
    all: {
        subject: "All Users Report",
        fileBaseName: "all_users_report",
        body: "Hi,\n\nPlease find attached the latest All Users Report.\n\nRegards,\nM365Corner"
    },
    enabled: {
        subject: "Enabled Users Report",
        fileBaseName: "enabled_users_report",
        body: "Hi,\n\nPlease find attached the latest Enabled Users Report.\n\nRegards,\nM365Corner"
    },
    disabled: {
        subject: "Disabled Users Report",
        fileBaseName: "disabled_users_report",
        body: "Hi,\n\nPlease find attached the latest Disabled Users Report.\n\nRegards,\nM365Corner"
    },
    licensed: {
        subject: "Licensed Users Report",
        fileBaseName: "licensed_users_report",
        body: "Hi,\n\nPlease find attached the latest Licensed Users Report.\n\nRegards,\nM365Corner"
    },
    unlicensed: {
        subject: "Unlicensed Users Report",
        fileBaseName: "unlicensed_users_report",
        body: "Hi,\n\nPlease find attached the latest Unlicensed Users Report.\n\nRegards,\nM365Corner"
    },

    // ============================
    // GROUP REPORTS
    // ============================
    "groups/all": {
        subject: "All Groups Report",
        fileBaseName: "all_groups_report",
        body: "Hi,\n\nAttached is your All Groups Report.\n\nRegards,\nM365Corner"
    },
    "groups/unified": {
        subject: "Unified Groups Report",
        fileBaseName: "unified_groups_report",
        body: "Hi,\n\nAttached is your Unified Groups Report.\n\nRegards,\nM365Corner"
    },
    "groups/distribution": {
        subject: "Distribution Groups Report",
        fileBaseName: "distribution_groups_report",
        body: "Hi,\n\nAttached is your Distribution Groups Report.\n\nRegards,\nM365Corner"
    },
    "groups/security-enabled": {
        subject: "Security Groups Report",
        fileBaseName: "security_enabled_groups_report",
        body: "Hi,\n\nAttached is your Security Groups Report.\n\nRegards,\nM365Corner"
    },
    "groups/mail-enabled-security": {
        subject: "Mail-enabled Security Groups Report",
        fileBaseName: "mail_enabled_security_groups_report",
        body: "Hi,\n\nAttached is your Mail-enabled Security Groups Report.\n\nRegards,\nM365Corner"
    },
    "groups/empty": {
        subject: "Empty Groups Report",
        fileBaseName: "empty_groups_report",
        body: "Hi,\n\nAttached is your Empty Groups Report.\n\nRegards,\nM365Corner"
    },
    "groups/recently-created": {
        subject: "Recently Created Groups Report",
        fileBaseName: "recently_created_groups_report",
        body: "Hi,\n\nAttached is your Recently Created Groups Report.\n\nRegards,\nM365Corner"
    },
    "groups/members": {
        subject: "Group Members Report",
        fileBaseName: "group_members_report",
        body: "Hi,\n\nAttached is your Group Members Report.\n\nRegards,\nM365Corner"
    },
    "groups/owners": {
        subject: "Group Owners Report",
        fileBaseName: "group_owners_report",
        body: "Hi,\n\nAttached is your Group Owners Report.\n\nRegards,\nM365Corner"
    },
    "groups/disabled-members": {
        subject: "Groups With Disabled Members Report",
        fileBaseName: "groups_with_disabled_members_report",
        body: "Hi,\n\nAttached is your Groups With Disabled Members Report.\n\nRegards,\nM365Corner"
    },

    // ============================
    // TEAMS REPORTS
    // ============================
    "teams/all": {
        subject: "All Teams Report",
        fileBaseName: "all_teams_report",
        body: "Hi,\n\nAttached is your All Teams Report.\n\nRegards,\nM365Corner"
    },
    "teams/public": {
        subject: "Public Teams Report",
        fileBaseName: "public_teams_report",
        body: "Hi,\n\nAttached is your Public Teams Report.\n\nRegards,\nM365Corner"
    },
    "teams/private": {
        subject: "Private Teams Report",
        fileBaseName: "private_teams_report",
        body: "Hi,\n\nAttached is your Private Teams Report.\n\nRegards,\nM365Corner"
    },
    "teams/hidden-memberships": {
        subject: "Hidden Membership Teams Report",
        fileBaseName: "hidden_membership_teams_report",
        body: "Hi,\n\nAttached is your Hidden Membership Teams Report.\n\nRegards,\nM365Corner"
    },
    "teams/archived": {
        subject: "Archived Teams Report",
        fileBaseName: "archived_teams_report",
        body: "Hi,\n\nAttached is your Archived Teams Report.\n\nRegards,\nM365Corner"
    },
    "teams/teams-without-description": {
        subject: "Teams Without Description Report",
        fileBaseName: "teams_without_description_report",
        body: "Hi,\n\nAttached is your Teams Without Description Report.\n\nRegards,\nM365Corner"
    },
    "teams/teams-private-channels": {
        subject: "Teams With Private Channels Report",
        fileBaseName: "teams_private_channels_report",
        body: "Hi,\n\nAttached is your Teams With Private Channels Report.\n\nRegards,\nM365Corner"
    },
    "teams/teams-shared-channels": {
        subject: "Teams With Shared Channels Report",
        fileBaseName: "teams_shared_channels_report",
        body: "Hi,\n\nAttached is your Teams With Shared Channels Report.\n\nRegards,\nM365Corner"
    },
    "teams/recently-created-teams": {
        subject: "Recently Created Teams Report",
        fileBaseName: "recently_created_teams_report",
        body: "Hi,\n\nAttached is your Recently Created Teams Report.\n\nRegards,\nM365Corner"
    },
    "teams/teams-owners": {
        subject: "Team Owners Report",
        fileBaseName: "team_owners_report",
        body: "Hi,\n\nAttached is your Team Owners Report.\n\nRegards,\nM365Corner"
    },
    "teams/teams-members": {
        subject: "Team Members Report",
        fileBaseName: "team_members_report",
        body: "Hi,\n\nAttached is your Team Members Report.\n\nRegards,\nM365Corner"
    }
};



async function sendReportByEmail({ recipient, csvData, filename, reportKey }) {
    const meta = REPORT_META[reportKey] || {
        subject: "M365 Report",
        fileBaseName: "m365_report",
        body: "Hi,\n\nPlease find your requested report attached.\n\nRegards,\nM365Corner"
    };

    const transporter = nodemailer.createTransport({
        service: "gmail",
        auth: {
            user: process.env.REPORT_EMAIL,
            pass: process.env.REPORT_PASS
        }
    });

    const mailOptions = {
        from: process.env.REPORT_EMAIL,
        to: recipient,
        subject: meta.subject,
        text: meta.body,
        attachments: [
            {
                filename,
                content: Buffer.from(csvData, "utf-8"),
                contentType: "text/csv",
                encoding: "base64"
            }
        ]
    };

    return transporter.sendMail(mailOptions);
}

module.exports = { sendReportByEmail };

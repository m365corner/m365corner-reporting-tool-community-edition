

const nodemailer = require("nodemailer");
require('dotenv').config(); // Ensure this is at the top

async function sendReportByEmail(recipient, csvData, filename) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: process.env.REPORT_EMAIL,
      pass: process.env.REPORT_PASS,
    },
  });

  const mailOptions = {
    from: process.env.REPORT_EMAIL,
    to: recipient,
    subject: "Unlicensed M365 Users Report",
    text: "Should fix the filename which is hardcoded as all_users_report",
    attachments: [
      {
        filename,
        content: Buffer.from(csvData, 'utf-8'), // ✅ Buffer conversion
        contentType: 'text/csv',
        encoding: 'base64' // ✅ Ensure email systems treat it properly
      },
    ],
  };

  return transporter.sendMail(mailOptions);
}

module.exports = { sendReportByEmail };


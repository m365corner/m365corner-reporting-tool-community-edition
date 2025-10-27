
const { URLSearchParams } = require("url");

// 🔁 Replace with your real values

const tenantId = process.env.tenantId;
const clientId = process.env.clientId
const clientSecret = process.env.clientSecret;




const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

async function getTokenAppOnly() {
  const params = new URLSearchParams();
  params.append("grant_type", "client_credentials");
  params.append("client_id", clientId);
  params.append("client_secret", clientSecret);
  params.append("scope", "https://graph.microsoft.com/.default");

  try {
    const res = await fetch(tokenEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString()
    });

    if (!res.ok) {
      const error = await res.text();
      throw new Error(`Token request failed: ${res.status} ${res.statusText} - ${error}`);
    }

    const data = await res.json();
    console.log("✅ App-only token acquired.");
    return data.access_token;
  } catch (err) {
    console.error("❌ Failed to acquire app-only token:", err.message);
    throw err;
  }
}

module.exports = { getTokenAppOnly };

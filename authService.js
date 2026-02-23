const axios = require("axios");
const qs = require("querystring");
const { saveTokens } = require("./tokenService");

async function refreshAccessToken(tokenStore) {
  try {
    const response = await axios.post(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      qs.stringify({
        client_id: process.env.CLIENT_ID,
        scope: "offline_access Chat.Read ChatMessage.Send User.Read",
        refresh_token: tokenStore.refresh_token,
        grant_type: "refresh_token",
        client_secret: process.env.CLIENT_SECRET
      }),
      {
        headers: { "Content-Type": "application/x-www-form-urlencoded" }
      }
    );

    const newTokens = response.data;
    newTokens.created_at = Math.floor(Date.now() / 1000);
    saveTokens(newTokens);
    console.log("[AUTH] Token refreshed successfully");
    return newTokens;
  } catch (error) {
    console.error("[AUTH] Token refresh failed:", error.response?.data || error.message);
    throw error;
  }
}

module.exports = { refreshAccessToken };

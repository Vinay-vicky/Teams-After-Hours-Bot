require("dotenv").config();
const express = require("express");
const axios = require("axios");
const qs = require("querystring");
const dayjs = require("dayjs");
const utc = require("dayjs/plugin/utc");
const timezone = require("dayjs/plugin/timezone");
const { saveTokens, loadTokens } = require("./tokenService");
const { refreshAccessToken } = require("./authService");

// Configure dayjs for IST timezone
dayjs.extend(utc);
dayjs.extend(timezone);

const app = express();
app.use(express.json());

const {
  PORT,
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  TARGET_CHAT_ID,
  SACHI_USER_ID,
  VIGNESH_USER_ID,
  TRIGGER_USER_ID,
  TELEGRAM_BOT_TOKEN,
  TELEGRAM_CHAT_ID,
  PUBLIC_BASE_URL,
  TEST_MODE,
  OPENCLAW_GO_LIVE_DATE
} = process.env;

const REDIRECT_URI = "http://localhost:3000/auth/callback";

// Persistent token storage - loads from file on startup
let tokenStore = loadTokens() || {};
if (tokenStore.access_token) {
  console.log("Tokens loaded from file (persisted session)");
}

// --------------------
// Helper: Check if token is expired
// --------------------
function isTokenExpired() {
  if (!tokenStore.expires_in || !tokenStore.access_token) return true;
  
  const now = Math.floor(Date.now() / 1000);
  const issuedAt = tokenStore.created_at || now;
  
  // Refresh 60 seconds before actual expiry
  return now >= issuedAt + tokenStore.expires_in - 60;
}

// --------------------
// Helper: Ensure valid token before Graph API calls
// --------------------
async function ensureValidToken() {
  if (!tokenStore.access_token) {
    throw new Error("Not authenticated");
  }
  
  if (isTokenExpired()) {
    console.log("[AUTH] Token expired, refreshing...");
    tokenStore = await refreshAccessToken(tokenStore);
  }
  
  return tokenStore.access_token;
}

// Idempotency: Track processed message IDs to prevent duplicate processing
const processedMessages = new Set();

// Subscription tracking for renewal
let currentSubscription = null;
let subscriptionRenewalTimer = null;

// Auto-reply message
const AUTO_REPLY_MESSAGE = "I am out of office for the day. I will respond to it next working day.";

// Log environment variables at startup (for debugging)
console.log("Environment loaded:");
console.log("  TELEGRAM_CHAT_ID:", TELEGRAM_CHAT_ID);
console.log("  TELEGRAM_BOT_TOKEN:", TELEGRAM_BOT_TOKEN ? "✓ Set" : "✗ Missing");
console.log("  TARGET_CHAT_ID:", TARGET_CHAT_ID ? "✓ Set" : "✗ Missing");
console.log("  SACHI_USER_ID:", SACHI_USER_ID);
console.log("  TRIGGER_USER_ID:", TRIGGER_USER_ID);
console.log("  PUBLIC_BASE_URL:", PUBLIC_BASE_URL);
console.log("  TEST_MODE:", TEST_MODE === "true" ? "ENABLED (bypassing time check)" : "disabled");
console.log("  OPENCLAW_GO_LIVE_DATE:", OPENCLAW_GO_LIVE_DATE || "NOT SET");
if (OPENCLAW_GO_LIVE_DATE) {
  const activationDate = dayjs(OPENCLAW_GO_LIVE_DATE).add(1, "month").format("YYYY-MM-DD");
  console.log("  Auto-reply activates on:", activationDate);
}

// --------------------
// Helper: Check if time is after 6PM IST (or TEST_MODE enabled)
// --------------------
function isAfterHours(dateTimeString) {
  // TEST_MODE bypasses time check for testing
  if (TEST_MODE === "true") {
    console.log("[TEST_MODE] Bypassing time check");
    return true;
  }
  const istTime = dayjs(dateTimeString).tz("Asia/Kolkata");
  const hour = istTime.hour();
  return hour >= 18; // 6PM or later
}

// --------------------
// Helper: Check if 1 month has passed since OpenClaw go-live date
// --------------------
function isOpenClawMonthComplete() {
  if (!OPENCLAW_GO_LIVE_DATE) {
    console.log("[WARNING] OPENCLAW_GO_LIVE_DATE not set, defaulting to eligible");
    return true;
  }
  
  const goLiveDate = dayjs(OPENCLAW_GO_LIVE_DATE);
  const activationDate = goLiveDate.add(1, "month");
  const today = dayjs();
  
  const isEligible = today.isAfter(activationDate) || today.isSame(activationDate, "day");
  
  if (!isEligible) {
    console.log(`[DATE CHECK] OpenClaw go-live: ${OPENCLAW_GO_LIVE_DATE}, Activation: ${activationDate.format("YYYY-MM-DD")}, Today: ${today.format("YYYY-MM-DD")}`);
  }
  
  return isEligible;
}

// --------------------
// Helper: Format IST time for logging
// --------------------
function formatIST(dateTimeString) {
  return dayjs(dateTimeString).tz("Asia/Kolkata").format("YYYY-MM-DD HH:mm:ss IST");
}

// --------------------
// Health Check
// --------------------
app.get("/health", (req, res) => {
  res.json({ status: "OK" });
});

// --------------------
// Start OAuth
// --------------------
app.get("/auth/start", (req, res) => {
  const authUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize` +
    `?client_id=${CLIENT_ID}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&response_mode=query` +
    `&scope=${encodeURIComponent("offline_access Chat.Read ChatMessage.Send User.Read")}`;

  res.redirect(authUrl);
});

// --------------------
// OAuth Callback
// --------------------
app.get("/auth/callback", async (req, res) => {
  const code = req.query.code;

  if (!code) {
    return res.status(400).send("No authorization code received");
  }

  try {
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      qs.stringify({
        client_id: CLIENT_ID,
        scope: "offline_access Chat.Read ChatMessage.Send User.Read",
        code: code,
        redirect_uri: REDIRECT_URI,
        grant_type: "authorization_code",
        client_secret: CLIENT_SECRET
      }),
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded"
        }
      }
    );

    tokenStore = tokenResponse.data;
    tokenStore.created_at = Math.floor(Date.now() / 1000);
    saveTokens(tokenStore);

    console.log("Access token received and saved to file");
    res.send("Authentication successful. You can close this window.");
  } catch (error) {
    console.error("Token error:", error.response?.data || error.message);
    res.status(500).send("Authentication failed");
  }
});

// --------------------
// List My Chats
// --------------------
app.get("/chats", async (req, res) => {
  try {
    const accessToken = await ensureValidToken();

    const response = await axios.get(
      "https://graph.microsoft.com/v1.0/me/chats",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      }
    );

    res.json(response.data);
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).send("Failed to fetch chats");
  }
});

// --------------------
// Get Members of a Chat
// --------------------
app.get("/chat-members", async (req, res) => {
  try {
    const chatId = req.query.id;

    if (!chatId) {
      return res.status(400).send("Chat ID is required");
    }

    const accessToken = await ensureValidToken();

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/chats/${chatId}/members`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      }
    );

    res.json(response.data);
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).send("Failed to fetch members");
  }
});

// --------------------
// Webhook endpoint (receives notifications from Microsoft Graph)
// --------------------
app.post("/webhook", async (req, res) => {
  // Handle validation request from Microsoft
  if (req.query.validationToken) {
    console.log("[WEBHOOK] Validation request received");
    res.set("Content-Type", "text/plain");
    return res.send(req.query.validationToken);
  }

  console.log("[WEBHOOK] Notification received");

  // Always respond with 202 immediately (Microsoft requires fast response)
  res.status(202).send();

  const notifications = req.body.value || [];

  for (const notification of notifications) {
    // Check if it's a chat message notification
    if (notification.resourceData) {
      const messageId = notification.resourceData.id;

      // ========== IDEMPOTENCY CHECK ==========
      if (processedMessages.has(messageId)) {
        console.log(`[SKIP] Message ${messageId} already processed (idempotency)`);
        continue;
      }
      processedMessages.add(messageId);

      // Fetch the actual message content
      try {
        const accessToken = await ensureValidToken();

        const messageResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(TARGET_CHAT_ID)}/messages/${messageId}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`
            }
          }
        );

        const message = messageResponse.data;
        const messageTime = message.createdDateTime;
        const istTimeFormatted = formatIST(messageTime);
        const senderId = message.from?.user?.id;
        const senderName = message.from?.user?.displayName || "Unknown";

        // ========== LOGGING ==========
        console.log("---------------------------------------------");
        console.log(`[INFO] Message ID: ${messageId}`);
        console.log(`[INFO] Timestamp IST: ${istTimeFormatted}`);
        console.log(`[INFO] From: ${senderName} (${senderId})`);

        // ========== SENDER CHECK ==========
        if (senderId !== TRIGGER_USER_ID) {
          console.log(`[DECISION] IGNORED - Not from target sender (sender: ${senderId})`);
          console.log("---------------------------------------------");
          continue;
        }

        // ========== DATE CHECK (1 Month After OpenClaw) ==========
        if (!isOpenClawMonthComplete()) {
          console.log(`[DECISION] IGNORED - 1 month not completed since OpenClaw go-live`);
          console.log("---------------------------------------------");
          continue;
        }

        // ========== TIME CHECK (6PM IST Rule) ==========
        if (!isAfterHours(messageTime)) {
          console.log(`[DECISION] IGNORED - Message before 6PM IST`);
          console.log("---------------------------------------------");
          continue;
        }

        console.log(`[DECISION] TRIGGERED - 1 month complete + After 6PM IST + From Sachi`);

        // ========== SEND TEAMS AUTO-REPLY ==========
        let teamsResult = "FAILED";
        try {
          const replyToken = await ensureValidToken();
          await axios.post(
            `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(TARGET_CHAT_ID)}/messages`,
            {
              body: {
                content: AUTO_REPLY_MESSAGE
              }
            },
            {
              headers: {
                Authorization: `Bearer ${replyToken}`,
                "Content-Type": "application/json"
              }
            }
          );
          teamsResult = "SUCCESS";
          console.log(`[TEAMS] Auto-reply sent: ${teamsResult}`);
        } catch (teamsError) {
          console.error(`[TEAMS] Auto-reply failed:`, teamsError.response?.data || teamsError.message);
        }

        // ========== SEND TELEGRAM NOTIFICATION ==========
        let telegramResult = "FAILED";
        try {
          const cleanContent = message.body?.content?.replace(/<[^>]*>/g, "") || "";
          await axios.post(
            `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
            {
              chat_id: TELEGRAM_CHAT_ID,
              text: `📩 Teams message from ${senderName} (after hours):\n\n${cleanContent}\n\n⏰ ${istTimeFormatted}`
            }
          );
          telegramResult = "SUCCESS";
          console.log(`[TELEGRAM] Notification sent: ${telegramResult}`);
        } catch (telegramError) {
          console.error(`[TELEGRAM] Notification failed:`, telegramError.response?.data || telegramError.message);
        }

        console.log(`[SUMMARY] Teams: ${teamsResult}, Telegram: ${telegramResult}`);
        console.log("---------------------------------------------");

      } catch (graphError) {
        console.error("[GRAPH API ERROR]", graphError.response?.data || graphError.message);
      }
    }
  }
});

// --------------------
// Helper: Create or renew subscription
// --------------------
async function createSubscription() {
  const accessToken = await ensureValidToken();

  if (!PUBLIC_BASE_URL) {
    console.log("[SUBSCRIPTION] No PUBLIC_BASE_URL set in .env");
    return null;
  }

  // Subscription expires in 60 minutes (max for chat messages)
  const expirationDateTime = new Date(Date.now() + 60 * 60 * 1000).toISOString();

  const subscription = {
    changeType: "created",
    notificationUrl: `${PUBLIC_BASE_URL}/webhook`,
    resource: `/chats/${TARGET_CHAT_ID}/messages`,
    expirationDateTime: expirationDateTime,
    clientState: "teams-after-hours-bot"
  };

  const response = await axios.post(
    "https://graph.microsoft.com/v1.0/subscriptions",
    subscription,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      }
    }
  );

  return response.data;
}

// --------------------
// Helper: Renew existing subscription
// --------------------
async function renewSubscription(subscriptionId) {
  const accessToken = await ensureValidToken();

  const expirationDateTime = new Date(Date.now() + 60 * 60 * 1000).toISOString();

  const response = await axios.patch(
    `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
    { expirationDateTime },
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      }
    }
  );

  return response.data;
}

// --------------------
// Helper: Schedule subscription renewal (10 minutes before expiry)
// --------------------
function scheduleSubscriptionRenewal() {
  if (subscriptionRenewalTimer) {
    clearTimeout(subscriptionRenewalTimer);
  }

  if (!currentSubscription) {
    return;
  }

  // Renew 10 minutes before expiry (50 minutes from now for 60-min subscription)
  const renewalTime = 50 * 60 * 1000; // 50 minutes

  subscriptionRenewalTimer = setTimeout(async () => {
    console.log("[SUBSCRIPTION] Auto-renewing subscription...");
    try {
      const renewed = await renewSubscription(currentSubscription.id);
      currentSubscription = renewed;
      console.log(`[SUBSCRIPTION] Renewed successfully. Expires: ${renewed.expirationDateTime}`);
      scheduleSubscriptionRenewal(); // Schedule next renewal
    } catch (error) {
      console.error("[SUBSCRIPTION] Renewal failed:", error.response?.data || error.message);
      // Try to create new subscription
      try {
        console.log("[SUBSCRIPTION] Attempting to create new subscription...");
        currentSubscription = await createSubscription();
        if (currentSubscription) {
          console.log("[SUBSCRIPTION] New subscription created after renewal failure");
          scheduleSubscriptionRenewal();
        }
      } catch (createError) {
        console.error("[SUBSCRIPTION] Failed to create new subscription:", createError.response?.data || createError.message);
      }
    }
  }, renewalTime);

  console.log(`[SUBSCRIPTION] Renewal scheduled in 50 minutes`);
}

// --------------------
// Create subscription to chat messages
// --------------------
app.get("/subscribe", async (req, res) => {
  try {
    // ensureValidToken() is called inside createSubscription()
    if (!PUBLIC_BASE_URL) {
      return res.status(400).send("PUBLIC_BASE_URL not set in .env file.");
    }

    console.log("[SUBSCRIPTION] Creating new subscription...");
    currentSubscription = await createSubscription();

    // Schedule automatic renewal
    scheduleSubscriptionRenewal();

    console.log("[SUBSCRIPTION] Created:", currentSubscription);
    res.json({
      message: "Subscription created successfully! Auto-renewal enabled.",
      subscription: currentSubscription
    });
  } catch (error) {
    console.error("[SUBSCRIPTION] Error:", error.response?.data || error.message);
    res.status(500).json({
      error: "Failed to create subscription",
      details: error.response?.data || error.message
    });
  }
});

// --------------------
// List active subscriptions
// --------------------
app.get("/subscriptions", async (req, res) => {
  try {
    const accessToken = await ensureValidToken();

    const response = await axios.get(
      "https://graph.microsoft.com/v1.0/subscriptions",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      }
    );

    res.json({
      activeSubscriptions: response.data,
      currentTracked: currentSubscription ? {
        id: currentSubscription.id,
        expirationDateTime: currentSubscription.expirationDateTime,
        renewalScheduled: !!subscriptionRenewalTimer
      } : null
    });
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).send("Failed to fetch subscriptions");
  }
});

// --------------------
// Delete all subscriptions
// --------------------
app.get("/delete-subscriptions", async (req, res) => {
  try {
    const accessToken = await ensureValidToken();

    // Get all subscriptions
    const response = await axios.get(
      "https://graph.microsoft.com/v1.0/subscriptions",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      }
    );

    const subscriptions = response.data.value || [];
    console.log(`[CLEANUP] Found ${subscriptions.length} subscriptions to delete`);

    const results = [];
    for (const sub of subscriptions) {
      try {
        await axios.delete(
          `https://graph.microsoft.com/v1.0/subscriptions/${sub.id}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`
            }
          }
        );
        console.log(`[CLEANUP] Deleted subscription ${sub.id}`);
        results.push({ id: sub.id, status: "deleted" });
      } catch (err) {
        console.error(`[CLEANUP] Failed to delete ${sub.id}:`, err.response?.data || err.message);
        results.push({ id: sub.id, status: "failed", error: err.response?.data || err.message });
      }
    }

    // Clear tracked subscription
    currentSubscription = null;
    if (subscriptionRenewalTimer) {
      clearTimeout(subscriptionRenewalTimer);
      subscriptionRenewalTimer = null;
    }

    res.json({ message: "Cleanup complete", results });
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).send("Failed to delete subscriptions");
  }
});

// --------------------
// Start Server
// --------------------
const server = app.listen(PORT || 3000, () => {
  console.log(`Server running on port ${PORT || 3000}`);
});

// --------------------
// Graceful Shutdown Handler
// --------------------
function gracefulShutdown(signal) {
  console.log(`\n${signal} received. Shutting down gracefully...`);

  // Clear subscription renewal timer
  if (subscriptionRenewalTimer) {
    clearTimeout(subscriptionRenewalTimer);
    console.log("Subscription renewal timer cleared.");
  }

  server.close(() => {
    console.log("Closed out remaining connections.");
    process.exit(0);
  });

  // Force exit after 10 seconds
  setTimeout(() => {
    console.error("Forcing shutdown...");
    process.exit(1);
  }, 10000);
}

process.on("SIGTERM", () => gracefulShutdown("SIGTERM"));
process.on("SIGINT", () => gracefulShutdown("SIGINT"));
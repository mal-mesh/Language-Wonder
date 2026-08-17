/* ==========================================================
   Language Wonder Academy Suite
   Module  : WhatsApp Engine
   File    : Whatsapp.gs
   Status  : Phase 1 Foundation
========================================================== */

const WHATSAPP = {
  provider: "",      // e.g. Meta, Twilio, Interakt, WATI
  apiUrl: "",
  apiKey: "",
  phoneNumberId: "",
  enabled: false
};

/* ==========================================================
   WHATSAPP CONFIGURATION
========================================================== */
function getWhatsAppConfig() {
  return {
    enabled: WHATSAPP.enabled,
    provider: WHATSAPP.provider,
    configured: !!(
      WHATSAPP.apiUrl &&
      WHATSAPP.apiKey &&
      WHATSAPP.phoneNumberId
    )
  };
}
/* ==========================================================
   WHATSAPP AVAILABILITY
========================================================== */
function isWhatsAppReady() {
  const config = getWhatsAppConfig();

  return (
    config.enabled &&
    config.configured
  );
}
/* ==========================================================
   BUILD MESSAGE PAYLOAD
========================================================== */
function buildWhatsAppMessage(phoneNumber, message, options) {
  options = options || {};

  return {
    to: String(phoneNumber || "").trim(),
    message: String(message || "").trim(),
    type: options.type || "text",
    template: options.template || "",
    mediaUrl: options.mediaUrl || "",
    filename: options.filename || "",
    metadata: options.metadata || {}
  };
}
/* ==========================================================
   SEND WHATSAPP MESSAGE (ENTRY POINT)
========================================================== */
function sendWhatsApp(payload) {

  if (!isWhatsAppReady()) {
    throw new Error("WhatsApp module is not configured.");
  }

  if (!payload || !payload.to || !payload.message) {
    throw new Error("Invalid WhatsApp payload.");
  }

  switch ((WHATSAPP.provider || "").toLowerCase()) {

    case "meta":
      return sendMetaWhatsApp(payload);

    default:
      throw new Error(
        "Unsupported WhatsApp provider: " + WHATSAPP.provider
      );
  }
}
/* ==========================================================
   META WHATSAPP CLOUD API
========================================================== */
function sendMetaWhatsApp(payload) {

  const endpoint =
    WHATSAPP.apiUrl.replace(/\/$/, "") +
    "/" +
    WHATSAPP.phoneNumberId +
    "/messages";

  const requestBody = {
    messaging_product: "whatsapp",
    recipient_type: "individual",
    to: payload.to,
    type: "text",
    text: {
      preview_url: false,
      body: payload.message
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + WHATSAPP.apiKey
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);

  const statusCode = response.getResponseCode();
  const responseBody = response.getContentText();

  return {
    success: statusCode >= 200 && statusCode < 300,
    statusCode: statusCode,
    response: responseBody
  };
}
// config.example.js
// Copy this file to config.js and fill in your credentials

module.exports = {
    SMTP_HOST: "your-smtp-server.com",
    SMTP_PORT: 465,
    SMTP_USER: "your-email@example.com",
    SMTP_PASS: "your-password-here", // ⚠️ Never commit real passwords!
    FROM_EMAIL: "sender@example.com",
    FROM_NAME: "Your Company Name",
    SUBJECT: "Default Email Subject",
    DELAY_MS: 15000,
    BATCH_SIZE: 10,
    BATCH_DELAY: 3600000,
};
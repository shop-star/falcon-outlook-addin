/* Configuration, read from the Worker's environment variables.
 *
 * Secrets (FERGUS_API_TOKEN, ANTHROPIC_API_KEY) live only here, on the
 * server — the task pane never sees them. Everything else has a sensible
 * default so a local run needs nothing but those two plus AUTH_DISABLED=true
 * (in .dev.vars).
 */

function list(value, fallback = []) {
  if (!value) return fallback;
  return value
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

function bool(value) {
  return /^(1|true|yes)$/i.test(value || "");
}

// The job types Fergus accepts on POST /jobs.
export const JOB_TYPES = ["Quote", "Estimate", "Charge Up"];

// `env` is the Worker's environment: vars and secrets from wrangler.jsonc
// and the Cloudflare dashboard.
export function loadConfig(env) {
  const defaultJobType = env.FERGUS_DEFAULT_JOB_TYPE || "Quote";
  if (!JOB_TYPES.includes(defaultJobType)) {
    throw new Error(`FERGUS_DEFAULT_JOB_TYPE must be one of: ${JOB_TYPES.join(", ")}`);
  }

  const config = {
    allowedOrigins: list(env.ALLOWED_ORIGINS, ["https://shop-star.github.io"]),
    // Our own staff's email domains, so a forwarded email's forwarder isn't
    // mistaken for the customer.
    ownEmailDomains: list(env.OWN_EMAIL_DOMAINS),

    fergus: {
      apiToken: env.FERGUS_API_TOKEN || "",
      baseUrl: (env.FERGUS_BASE_URL || "https://api.fergus.com").replace(/\/+$/, ""),
      // Where a person opens a job in the Fergus web app. {jobId} and
      // {jobNo} are substituted. Check this against a real job's URL.
      jobUrlTemplate: env.FERGUS_JOB_URL_TEMPLATE || "https://app.fergus.com/jobs/view/{jobId}",
      // Business-unit groups Claude may suggest from. Fergus's public API
      // exposes no groups endpoint, so this list is ours to keep in step
      // with the Fergus account.
      groups: list(env.FERGUS_GROUPS),
      defaultJobType,
    },

    anthropic: {
      apiKey: env.ANTHROPIC_API_KEY || "",
      model: env.ANTHROPIC_MODEL || "claude-opus-5-5",
    },

    auth: {
      // Local testing only: skips the Microsoft sign-in check entirely.
      disabled: bool(env.AUTH_DISABLED),
      tenantId: env.ENTRA_TENANT_ID || "",
      clientId: env.ENTRA_CLIENT_ID || "",
    },
  };

  const missing = [];
  if (!config.fergus.apiToken) missing.push("FERGUS_API_TOKEN");
  if (!config.anthropic.apiKey) missing.push("ANTHROPIC_API_KEY");
  if (!config.auth.disabled) {
    if (!config.auth.tenantId) missing.push("ENTRA_TENANT_ID");
    if (!config.auth.clientId) missing.push("ENTRA_CLIENT_ID");
  }
  if (missing.length) {
    throw new Error(
      `Missing required environment variable(s): ${missing.join(", ")}. ` +
        "See the README's \"Create Fergus job\" section."
    );
  }

  return config;
}

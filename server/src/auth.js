/* Checks that each request comes from someone signed in to the
 * organisation's Microsoft 365 tenant.
 *
 * The task pane gets an access token for this API from Outlook itself
 * (nested app authentication, see taskpane/fergus.js) and sends it as a
 * Bearer token. Without this check, anyone who found the backend URL could
 * create Fergus jobs and spend Claude credits.
 */
import { createRemoteJWKSet, jwtVerify } from "jose";
import { AppError } from "./errors.js";

export function createAuthenticator({ disabled, tenantId, clientId }) {
  if (disabled) {
    return async () => ({ name: "local testing", email: "" });
  }

  const jwks = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`));
  const verifyOptions = {
    // v2 tokens carry the client ID; v1 tokens carry the Application ID URI.
    audience: [clientId, `api://${clientId}`],
    issuer: [`https://login.microsoftonline.com/${tenantId}/v2.0`, `https://sts.windows.net/${tenantId}/`],
  };

  return async function authenticate(request) {
    const match = /^Bearer (.+)$/i.exec(request.headers.get("Authorization") || "");
    if (!match) {
      throw new AppError(401, "UNAUTHENTICATED", "Sign in to Microsoft 365 to use this.");
    }
    let payload;
    try {
      ({ payload } = await jwtVerify(match[1], jwks, verifyOptions));
    } catch {
      throw new AppError(401, "UNAUTHENTICATED", "Your Microsoft 365 sign-in has expired or isn't valid here. Close and reopen the pane.");
    }
    const scopes = String(payload.scp || "").split(" ");
    if (!scopes.includes("access_as_user")) {
      throw new AppError(403, "FORBIDDEN", "This sign-in isn't allowed to use the Fergus job API.");
    }
    return { name: payload.name || "", email: payload.preferred_username || payload.upn || "" };
  };
}

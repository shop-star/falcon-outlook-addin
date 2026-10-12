/* Settings for the "Create Fergus job" task pane. None of these are secrets:
 * the Fergus token and Claude API key live only on the backend.
 *
 * Fill these in after deploying the backend and registering the add-in in
 * Microsoft Entra ID — see the README's "Create Fergus job" section.
 */
window.FERGUS_ADDIN_CONFIG = {
  // Where server/ is deployed, without a trailing slash.
  backendUrl: "https://YOUR-BACKEND-HOST",

  // Entra ID app registration for the backend API. Leave clientId empty
  // only when testing against a backend running with AUTH_DISABLED=true.
  entraClientId: "",
  entraTenantId: "",
};

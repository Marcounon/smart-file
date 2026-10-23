/* Smart File - sign-in and Microsoft Graph tokens.
 * Primary path: nested app authentication (NAA) via MSAL, which reuses the account Outlook is already signed in with.
 * Fallback:     an Office dialog that runs a normal MSAL redirect flow (fallbackauth.html) for hosts without NAA.
 */
(function (global) {
  "use strict";

  const cfg = () => global.SmartFileConfig;

  class AuthRequired extends Error {
    constructor(cause) {
      super("Sign-in is needed before Smart File can look at your folders.");
      this.name = "AuthRequired";
      this.cause = cause;
    }
  }

  let pcaPromise = null;
  let cached = null; // { accessToken, expiresOn (ms) }

  function authority() {
    return "https://login.microsoftonline.com/" + (cfg().tenantId || "common");
  }

  function msalConfig() {
    return {
      auth: { clientId: cfg().clientId, authority: authority() },
      cache: { cacheLocation: "localStorage" }
    };
  }

  function getPca() {
    if (!pcaPromise) pcaPromise = global.msal.createNestablePublicClientApplication(msalConfig());
    return pcaPromise;
  }

  function naaSupported() {
    try { return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"); } catch (e) { return false; }
  }

  function remember(result) {
    const expiresOn = result.expiresOn ? new Date(result.expiresOn).getTime() : Date.now() + 50 * 60 * 1000;
    cached = { accessToken: result.accessToken, expiresOn: expiresOn };
    return cached;
  }

  function cachedIsFresh() {
    return cached && cached.expiresOn - Date.now() > 2 * 60 * 1000;
  }

  let inflight = null;

  /** Get a token without user interaction. Throws AuthRequired when the user has to sign in or consent.
   *  Concurrent callers share one request: MSAL does not merge NAA requests itself. */
  function getToken() {
    if (cachedIsFresh()) return Promise.resolve(cached);
    if (!inflight) inflight = acquireSilently().finally(() => { inflight = null; });
    return inflight;
  }

  async function acquireSilently() {
    if (!naaSupported()) throw new AuthRequired(new Error("Nested app authentication isn\u2019t available in this Outlook client."));
    try {
      const pca = await getPca();
      const result = await pca.acquireTokenSilent({ scopes: cfg().scopes });
      return remember(result);
    } catch (e) {
      throw new AuthRequired(e);
    }
  }

  /** Get a token with user interaction (call this from a click handler). */
  async function getTokenInteractive() {
    if (naaSupported()) {
      const pca = await getPca();
      const result = await pca.acquireTokenPopup({ scopes: cfg().scopes });
      return remember(result);
    }
    const result = await getTokenViaDialog();
    return remember(result);
  }

  /** Fallback: open fallbackauth.html in an Office dialog and wait for it to send a token back. */
  function getTokenViaDialog() {
    return new Promise((resolve, reject) => {
      const url = new URL("fallbackauth.html", global.location.href).href;
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 30, promptBeforeOpen: false }, (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) { reject(new Error(res.error.message)); return; }
        const dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          let msg;
          try { msg = JSON.parse(arg.message); } catch (e) { msg = { error: "Unexpected reply from the sign-in window." }; }
          dialog.close();
          if (msg.accessToken) resolve(msg); else reject(new Error(msg.error || "Sign-in didn\u2019t complete."));
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          reject(new Error(arg.error === 12006 ? "The sign-in window was closed before finishing." : "Sign-in window error " + arg.error));
        });
      });
    });
  }

  function describeError(e) {
    const code = (e && (e.errorCode || e.code)) || "";
    const text = (e && e.message) || String(e);
    if (/AADSTS65001|consent_required|invalid_grant/i.test(text + code)) return "Permission hasn\u2019t been granted yet. An admin can grant consent in Entra, or you can sign in and accept the prompt.";
    if (/AADSTS700016|unauthorized_client/i.test(text + code)) return "The app registration ID in config.js doesn\u2019t exist in this tenant. Check clientId and tenantId.";
    if (/AADSTS50011|redirect/i.test(text + code)) return "The redirect URI isn\u2019t registered. Add brk-multihub://" + global.location.host + " as a single-page application redirect in the Entra app.";
    if (/user_cancelled|user_cancelled_error/i.test(text + code)) return "Sign-in was cancelled.";
    return text;
  }

  global.SmartFileAuth = { AuthRequired, getToken, getTokenInteractive, naaSupported, describeError, authority };
})(typeof window !== "undefined" ? window : globalThis);

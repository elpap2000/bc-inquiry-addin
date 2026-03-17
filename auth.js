(function () {
  const cfg = window.ADDIN_CONFIG || {};

  function isPlaceholder(v) {
    return !v || /PASTE-|YOUR-/i.test(v);
  }

  function normalizeError(err) {
    if (!err) return "Unknown authentication error.";
    if (typeof err === "string") return err;
    return err.message || err.errorMessage || JSON.stringify(err);
  }

  function showError(err) {
    const el = document.getElementById("error");
    if (el) {
      el.style.display = "";
      el.textContent = normalizeError(err);
    }
  }

  function sendParent(payload) {
    Office.context.ui.messageParent(JSON.stringify(payload));
  }

  Office.onReady(async function () {
    try {
      if (isPlaceholder(cfg.tenantId) || isPlaceholder(cfg.clientId)) {
        throw new Error(
          "Add-in not configured. Update app-config.js with real tenantId and clientId."
        );
      }

      const msalConfig = {
        auth: {
          clientId: cfg.clientId,
          authority: "https://login.microsoftonline.com/" + cfg.tenantId,
          redirectUri: cfg.redirectUri || (window.location.origin + window.location.pathname)
        },
        cache: {
          cacheLocation: "localStorage",
          storeAuthStateInCookie: true
        }
      };

      const tokenRequest = {
        scopes: [cfg.bcScope]
      };

      const msalInstance = new msal.PublicClientApplication(msalConfig);
      await msalInstance.initialize();

      const redirectResult = await msalInstance.handleRedirectPromise();
      if (redirectResult && redirectResult.accessToken) {
        sendParent({
          type: "auth-success",
          accessToken: redirectResult.accessToken,
          expiresOn: redirectResult.expiresOn ? redirectResult.expiresOn.toISOString() : null
        });
        return;
      }

      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        try {
          const silentResult = await msalInstance.acquireTokenSilent({
            ...tokenRequest,
            account: accounts[0]
          });

          sendParent({
            type: "auth-success",
            accessToken: silentResult.accessToken,
            expiresOn: silentResult.expiresOn ? silentResult.expiresOn.toISOString() : null
          });
          return;
        } catch (_) {
          // continue to interactive
        }
      }

      await msalInstance.acquireTokenRedirect({
        ...tokenRequest,
        loginHint:
          Office.context?.mailbox?.userProfile?.emailAddress || undefined
      });
    } catch (err) {
      showError(err);
      sendParent({
        type: "auth-error",
        message: normalizeError(err)
      });
    }
  });
})();

(function () {
  const cfg = window.ADDIN_CONFIG;

  const msalConfig = {
    auth: {
      clientId: cfg.clientId,
      authority: "https://login.microsoftonline.com/" + cfg.tenantId,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true
    }
  };

  const tokenRequest = {
    scopes: [cfg.bcScope]
  };

  let msalInstance;

  Office.onReady(async function () {
    try {
      msalInstance = new msal.PublicClientApplication(msalConfig);
      await msalInstance.initialize();
      await authenticate();
    } catch (err) {
      showError(err);
      sendParent({
        type: "auth-error",
        message: normalizeError(err)
      });
    }
  });

  async function authenticate() {
    // 1) Return from redirect flow
    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult && redirectResult.accessToken) {
      return sendSuccess(redirectResult);
    }

    // 2) Try silent token if an account already exists
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      try {
        const silentResult = await msalInstance.acquireTokenSilent({
          ...tokenRequest,
          account: accounts[0]
        });
        return sendSuccess(silentResult);
      } catch (e) {
        // Silent failed -> continue interactive
      }
    }

    // 3) Interactive sign-in/token
    await msalInstance.acquireTokenRedirect(tokenRequest);
  }

  function sendSuccess(result) {
    sendParent({
      type: "auth-success",
      accessToken: result.accessToken,
      expiresOn: result.expiresOn ? result.expiresOn.toISOString() : null
    });
  }

  function sendParent(payload) {
    Office.context.ui.messageParent(JSON.stringify(payload));
  }

  function showError(err) {
    const el = document.getElementById("error");
    el.style.display = "";
    el.textContent = normalizeError(err);
  }

  function normalizeError(err) {
    if (!err) return "Unknown authentication error.";
    if (typeof err === "string") return err;
    return err.message || err.errorMessage || JSON.stringify(err);
  }
})();

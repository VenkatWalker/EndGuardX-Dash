import { useCallback, useRef } from "react";
import { useAuthContext } from "@/context/AuthContext";

type SSOProvider = {
  id: "azure" | "google" | string;
  label: string;
  client_id?: string;
  tenant_id?: string;
};

type ProvidersInfo = {
  local_login: boolean;
  providers: SSOProvider[];
  reachable: boolean;
};

// Google script loader singleton
let googleScriptPromise: Promise<void> | null = null;
function loadGoogleScript(): Promise<void> {
  if (typeof window === "undefined") return Promise.resolve();
  if ((window as any).google?.accounts?.id) return Promise.resolve();
  if (googleScriptPromise) return googleScriptPromise;
  googleScriptPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.async = true;
    s.defer = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Failed to load Google script"));
    document.head.appendChild(s);
  });
  return googleScriptPromise;
}

export function useAuth() {
  const { login, managerUrl } = useAuthContext();
  const msalInstanceRef = useRef<any>(null);

  // Get or create MSAL instance (singleton per provider config)
  const getMsalInstance = useCallback(async (p: SSOProvider) => {
    if (!msalInstanceRef.current) {
      const { PublicClientApplication } = await import("@azure/msal-browser");
      const msal = new PublicClientApplication({
        auth: {
          clientId: p.client_id!,
          authority: `https://login.microsoftonline.com/${p.tenant_id!}`,
          redirectUri:
            `${window.location.origin}/auth-redirect.html` +
            `?clientId=${encodeURIComponent(p.client_id!)}` +
            `&tenantId=${encodeURIComponent(p.tenant_id!)}`,
        },
      });
      await msal.initialize();
      await msal.handleRedirectPromise().catch(() => null);
      msalInstanceRef.current = msal;
    }
    return msalInstanceRef.current;
  }, []);

  // Fetch available auth providers from manager
  const fetchProviders = useCallback(
    async (url: string): Promise<ProvidersInfo> => {
      if (!url) return { local_login: true, providers: [], reachable: true };
      try {
        const res = await fetch(`${url}/api/v1/auth/providers`);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        return {
          local_login: data?.local_login !== false,
          providers: Array.isArray(data?.providers) ? data.providers : [],
          reachable: true,
        };
      } catch {
        return { local_login: true, providers: [], reachable: false };
      }
    },
    []
  );

  // Local login handler
  const handleLocalLogin = useCallback(
    async (
      username: string,
      password: string,
      onError: (msg: string) => void,
      onSuccess: () => void
    ) => {
      try {
        const res = await fetch(`${managerUrl}/api/v1/auth/login`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ username, password }),
        });
        if (res.status === 401) {
          onError("Invalid username or password");
          return;
        }
        if (!res.ok) {
          onError(`Login failed (HTTP ${res.status})`);
          return;
        }
        const data = await res.json();
        login(data.token || "", data.username || username, data.role || "");
        onSuccess();
      } catch {
        onError("Cannot reach manager. Check URL.");
      }
    },
    [managerUrl, login]
  );

  // Azure SSO login handler
  const handleAzureLogin = useCallback(
    async (
      provider: SSOProvider,
      onError: (msg: string) => void,
      onSuccess: () => void
    ) => {
      try {
        if (!provider.client_id || !provider.tenant_id) {
          throw new Error("Azure provider misconfigured — check sso_config.json");
        }
        const msal = await getMsalInstance(provider);
        const result = await msal.loginPopup({
          scopes: ["openid", "profile", "email"],
        });
        const idToken = result.idToken;
        if (!idToken) {
          throw new Error(
            "No idToken in MSAL result — check Azure app registration scopes"
          );
        }

        const res = await fetch(`${managerUrl}/api/v1/auth/sso/verify`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ provider: "azure", access_token: idToken }),
        });
        if (!res.ok) {
          const errText = await res.text().catch(() => "");
          onError(`SSO failed (HTTP ${res.status}) — ${errText.slice(0, 120)}`);
          return;
        }
        const data = await res.json();
        login(data.token || "", data.username || "", data.role || "");
        onSuccess();
      } catch (err: any) {
        if (err?.errorCode === "interaction_in_progress") {
          msalInstanceRef.current = null; // reset stuck instance
          onError("Login already in progress — please refresh and try again");
        } else if (err?.errorCode === "user_cancelled") {
          onError("Login cancelled");
        } else {
          onError(err?.message || "Microsoft sign-in failed");
        }
      }
    },
    [managerUrl, login, getMsalInstance]
  );

  // Google SSO credential handler
  const handleGoogleCredential = useCallback(
    async (
      idToken: string,
      onError: (msg: string) => void,
      onSuccess: () => void
    ) => {
      try {
        const res = await fetch(`${managerUrl}/api/v1/auth/sso/verify`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ provider: "google", access_token: idToken }),
        });
        if (!res.ok) {
          const errText = await res.text().catch(() => "");
          onError(`Google SSO failed (HTTP ${res.status}) — ${errText.slice(0, 120)}`);
          return;
        }
        const data = await res.json();
        login(data.token || "", data.username || "", data.role || "");
        onSuccess();
      } catch (err: any) {
        onError(err?.message || "Google sign-in failed");
      }
    },
    [managerUrl, login]
  );

  // Initialize Google Sign-In button
  const initGoogleButton = useCallback(
    async (
      googleProvider: SSOProvider | undefined,
      buttonElementId: string,
      theme: "dark" | "light",
      onCredential: (token: string) => void
    ) => {
      if (!googleProvider || !googleProvider.client_id) return;
      let active = true;
      try {
        await loadGoogleScript();
        if (!active) return;
        const google = (window as any).google;
        if (!google?.accounts?.id) return;
        google.accounts.id.initialize({
          client_id: googleProvider.client_id,
          callback: (resp: any) => {
            if (resp?.credential) onCredential(resp.credential);
          },
          auto_select: false,
          cancel_on_tap_outside: true,
        });
        const btnElem = document.getElementById(buttonElementId);
        if (btnElem) {
          google.accounts.id.renderButton(btnElem, {
            theme: theme === "dark" ? "filled_black" : "outline",
            size: "large",
            width: btnElem.clientWidth || 280,
          });
        }
      } catch (err) {
        console.error("Google Identity Services load failed:", err);
      }
      return () => {
        active = false;
      };
    },
    []
  );

  // Reset MSAL instance (call on logout)
  const resetMsal = useCallback(() => {
    msalInstanceRef.current = null;
  }, []);

  return {
    fetchProviders,
    handleLocalLogin,
    handleAzureLogin,
    handleGoogleCredential,
    initGoogleButton,
    resetMsal,
  };
}

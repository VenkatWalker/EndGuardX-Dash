import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useCallback, useEffect, useMemo, useState } from "react";
import { useAuthContext } from "@/context/AuthContext";
import { useAuth } from "@/hooks/useAuth";
import Threads from "@/components/endguardx/Threads.jsx";

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

function LoginComponent() {
  const navigate = useNavigate();
  const { token, managerUrl, setManagerUrl, login, isAuthenticated } = useAuthContext();
  const {
    fetchProviders,
    handleLocalLogin,
    handleAzureLogin,
    handleGoogleCredential,
    initGoogleButton,
  } = useAuth();

  // Redirect if already authenticated
  useEffect(() => {
    if (isAuthenticated && token) {
      navigate({ to: "/dashboard" });
    }
  }, [isAuthenticated, token, navigate]);

  // UI state
  const [theme, setTheme] = useState<"dark" | "light">(() => {
    if (typeof window === "undefined") return "dark";
    return (localStorage.getItem("gx-theme") as any) || "dark";
  });
  useEffect(() => {
    document.documentElement.setAttribute("data-gx-theme", theme);
    try {
      localStorage.setItem("gx-theme", theme);
    } catch {
      /* ignore */
    }
  }, [theme]);

  // Login form state
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [loginErr, setLoginErr] = useState("");
  const [loggingIn, setLoggingIn] = useState(false);
  const [ssoBusy, setSsoBusy] = useState<string>("");

  // Providers state
  const [providersInfo, setProvidersInfo] = useState<ProvidersInfo>({
    local_login: true,
    providers: [],
    reachable: true,
  });
  const [providersLoading, setProvidersLoading] = useState(false);

  // Time display
  const [nowStr, setNowStr] = useState<string>(() =>
    new Date().toLocaleString()
  );
  useEffect(() => {
    const id = setInterval(() => setNowStr(new Date().toLocaleString()), 1000);
    return () => clearInterval(id);
  }, []);

  // Fetch providers when manager URL changes
  const doFetchProviders = useCallback(async () => {
    setProvidersLoading(true);
    const info = await fetchProviders(managerUrl);
    setProvidersInfo(info);
    setProvidersLoading(false);
  }, [fetchProviders, managerUrl]);

  useEffect(() => {
    void doFetchProviders();
  }, [doFetchProviders]);

  // Handle local login
  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginErr("");
    setLoggingIn(true);
    await handleLocalLogin(
      loginUser,
      loginPass,
      (err) => setLoginErr(err),
      () => {
        setLoginErr("");
        navigate({ to: "/dashboard" });
      }
    );
    setLoggingIn(false);
  };

  // Handle Azure login
  const handleAzureClick = async (p: SSOProvider) => {
    setLoginErr("");
    setSsoBusy("azure");
    await handleAzureLogin(
      p,
      (err) => setLoginErr(err),
      () => {
        setLoginErr("");
        navigate({ to: "/dashboard" });
      }
    );
    setSsoBusy("");
  };

  // Handle Google credential
  const handleGoogleClick = useCallback(
    async (idToken: string) => {
      setLoginErr("");
      setSsoBusy("google");
      await handleGoogleCredential(
        idToken,
        (err) => setLoginErr(err),
        () => {
          setLoginErr("");
          navigate({ to: "/dashboard" });
        }
      );
      setSsoBusy("");
    },
    [handleGoogleCredential, navigate]
  );

  // Initialize Google button when providers load
  useEffect(() => {
    const googleP = providersInfo.providers.find((x) => x.id === "google");
    if (!googleP) return;
    const cleanup = initGoogleButton(
      googleP,
      "google-signin-button",
      theme,
      handleGoogleClick
    );
    return cleanup;
  }, [providersInfo.providers, theme, initGoogleButton, handleGoogleClick]);

  // Computed values
  const azureP = useMemo(
    () => providersInfo.providers.find((x) => x.id === "azure"),
    [providersInfo.providers]
  );
  const googleP = useMemo(
    () => providersInfo.providers.find((x) => x.id === "google"),
    [providersInfo.providers]
  );
  const showLocal = providersInfo.local_login || !providersInfo.reachable;
  const showDivider = showLocal && (azureP || googleP);

  return (
    <div className="gx-root">
      <div className="gx-login-overlay">
        <div
          style={{
            position: "absolute",
            inset: 0,
            zIndex: 0,
            pointerEvents: "none",
            opacity: theme === "dark" ? 0.85 : 0.75,
          }}
        >
          <Threads
            color={theme === "dark" ? [0, 0.85, 1] : [0.05, 0.35, 0.55]}
            amplitude={1.6}
            distance={0.3}
            enableMouseInteraction={true}
          />
        </div>
        <div
          style={{
            position: "absolute",
            top: 16,
            right: 20,
            display: "flex",
            alignItems: "center",
            gap: 16,
            fontFamily: "Share Tech Mono, monospace",
            fontSize: 11,
            color: "var(--gx-text)",
            letterSpacing: "0.1em",
            zIndex: 1,
          }}
        >
          <span
            style={{
              color: providersInfo.reachable ? "var(--gx-green)" : "var(--gx-red)",
            }}
          >
            ● {providersLoading ? "CHECKING..." : providersInfo.reachable ? "MANAGER ONLINE" : "MANAGER OFFLINE"}
          </span>
          <span>{nowStr}</span>
          <span
            onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
            style={{ cursor: "pointer", color: "var(--gx-cyan-text)" }}
            title="Toggle theme"
          >
            ☾ {theme.toUpperCase()}
          </span>
        </div>

        <form className="gx-login-card" onSubmit={handleLogin}>
          <h2>
            Endguard<span style={{ color: "var(--gx-green)" }}>X</span>
          </h2>
          <div className="sub">SECURE ACCESS</div>

          <label>MANAGER URL</label>
          <input
            type="text"
            value={managerUrl}
            onChange={(e) => setManagerUrl(e.target.value)}
            placeholder="https://manager:8443"
          />

          {!providersInfo.reachable && (
            <div
              style={{
                fontSize: 10,
                color: "var(--gx-amber, #ffb454)",
                margin: "4px 0 8px",
                letterSpacing: "0.1em",
              }}
            >
              ⚠ MANAGER UNREACHABLE — LOCAL LOGIN ONLY
            </div>
          )}

          {showLocal && (
            <>
              <label>USERNAME</label>
              <input
                type="text"
                value={loginUser}
                onChange={(e) => setLoginUser(e.target.value)}
                autoFocus
              />
              <label>PASSWORD</label>
              <input
                type="password"
                value={loginPass}
                onChange={(e) => setLoginPass(e.target.value)}
              />
              <button type="submit" disabled={loggingIn || !!ssoBusy}>
                {loggingIn ? "SIGNING IN..." : "SIGN IN"}
              </button>
            </>
          )}

          {showDivider && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                margin: "14px 0 10px",
                color: "var(--gx-fg-dim, #6b7b8d)",
                fontSize: 10,
                letterSpacing: "0.3em",
              }}
            >
              <div
                style={{
                  flex: 1,
                  height: 1,
                  background: "currentColor",
                  opacity: 0.3,
                }}
              />
              OR
              <div
                style={{
                  flex: 1,
                  height: 1,
                  background: "currentColor",
                  opacity: 0.3,
                }}
              />
            </div>
          )}

          {azureP && (
            <button
              type="button"
              onClick={() => handleAzureClick(azureP)}
              disabled={!!ssoBusy || loggingIn}
              style={{ marginTop: 6 }}
            >
              {ssoBusy === "azure" ? "OPENING MICROSOFT..." : (azureP.label || "LOGIN WITH MICROSOFT").toUpperCase()}
            </button>
          )}

          {googleP && (
            <div
              id="google-signin-button"
              style={{
                marginTop: 8,
                display: "flex",
                justifyContent: "center",
                width: "100%",
                minHeight: 40,
              }}
            />
          )}

          <div className="gx-login-err">{loginErr}</div>
        </form>
      </div>
    </div>
  );
}

export const Route = createFileRoute("/login")({
  head: () => ({
    meta: [
      { title: "EndguardX - Login" },
      { name: "description", content: "EndguardX endpoint control platform login" },
    ],
  }),
  component: LoginComponent,
  ssr: false,
});

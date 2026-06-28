import React, { createContext, useContext, useState, useCallback, ReactNode, useEffect } from "react";

type AuthContextType = {
  token: string;
  username: string;
  role: string;
  managerUrl: string;
  setManagerUrl: (url: string) => void;
  login: (token: string, username: string, role?: string) => void;
  logout: () => void;
  isAuthenticated: boolean;
};

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export function AuthProvider({ children }: { children: ReactNode }) {
  const [token, setToken] = useState<string>("");
  const [username, setUsername] = useState<string>("");
  const [role, setRole] = useState<string>("");
  const [managerUrl, setManagerUrl] = useState<string>(() => {
    if (typeof window === "undefined") return "https://192.168.1.5:8443";
    return localStorage.getItem("gx-last-manager") || "https://192.168.1.5:8443";
  });

  // Restore auth state from sessionStorage on mount
  useEffect(() => {
    if (typeof window === "undefined") return;
    const storedToken = sessionStorage.getItem("gx-dash-token") || "";
    const storedUser = sessionStorage.getItem("gx-dash-user") || "";
    const storedManager = sessionStorage.getItem("gx-dash-manager");
    
    if (storedToken) {
      setToken(storedToken);
      setUsername(storedUser);
      if (storedManager) setManagerUrl(storedManager);
    }
  }, []);

  // Persist managerUrl to localStorage when it changes
  useEffect(() => {
    if (typeof window === "undefined") return;
    try {
      localStorage.setItem("gx-last-manager", managerUrl);
    } catch {
      /* ignore */
    }
  }, [managerUrl]);

  const login = useCallback((newToken: string, newUsername: string, newRole?: string) => {
    setToken(newToken);
    setUsername(newUsername);
    setRole(newRole || "");

    // Persist to sessionStorage
    if (typeof window !== "undefined") {
      try {
        sessionStorage.setItem("gx-dash-token", newToken);
        sessionStorage.setItem("gx-dash-user", newUsername);
        sessionStorage.setItem("gx-dash-manager", managerUrl);
      } catch {
        /* ignore */
      }
    }
  }, [managerUrl]);

  const logout = useCallback(() => {
    setToken("");
    setUsername("");
    setRole("");

    // Clear sessionStorage
    if (typeof window !== "undefined") {
      try {
        sessionStorage.removeItem("gx-dash-token");
        sessionStorage.removeItem("gx-dash-manager");
        sessionStorage.removeItem("gx-dash-user");
      } catch {
        /* ignore */
      }
    }

    // Reset MSAL instance ref globally if it exists (will be handled in useAuth hook)
  }, []);

  return (
    <AuthContext.Provider
      value={{
        token,
        username,
        role,
        managerUrl,
        setManagerUrl,
        login,
        logout,
        isAuthenticated: !!token,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
}

export function useAuthContext() {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error("useAuthContext must be used within AuthProvider");
  }
  return context;
}

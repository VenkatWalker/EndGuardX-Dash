import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect } from "react";
import { useAuthContext } from "@/context/AuthContext";

function IndexComponent() {
  const navigate = useNavigate();
  const { isAuthenticated, token } = useAuthContext();

  useEffect(() => {
    if (isAuthenticated && token) {
      navigate({ to: "/dashboard", replace: true });
    } else {
      navigate({ to: "/login", replace: true });
    }
  }, [isAuthenticated, token, navigate]);

  return null;
}

export const Route = createFileRoute("/")({
  component: IndexComponent,
  ssr: false,
});

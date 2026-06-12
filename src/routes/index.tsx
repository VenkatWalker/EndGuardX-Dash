import { createFileRoute } from "@tanstack/react-router";
import EndguardX from "@/components/endguardx/EndguardX";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "EndguardX - Endpoint Control Platform" },
      { name: "description", content: "EndguardX endpoint control platform — monitor agents, events, alerts and policy violations from a single dashboard." },
      { property: "og:title", content: "EndguardX - Endpoint Control Platform" },
      { property: "og:description", content: "Monitor endpoint agents, events, alerts and policy violations in real time." },
    ],
  }),
  component: EndguardX,
  ssr: false,
});

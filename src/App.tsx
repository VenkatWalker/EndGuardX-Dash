import { RouterProvider, createRouter } from "@tanstack/react-router";
import { AuthProvider } from "@/context/AuthContext";
import { routeTree } from "@/routeTree.gen";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";

// Create router
const router = createRouter({
  routeTree,
  context: { queryClient: new QueryClient() },
  scrollRestoration: true,
  defaultPreloadStaleTime: 0,
});

declare module "@tanstack/react-router" {
  interface Register {
    router: typeof router;
  }
}

export function App() {
  return (
    <QueryClientProvider client={router.context.queryClient}>
      <AuthProvider>
        <RouterProvider router={router} />
      </AuthProvider>
    </QueryClientProvider>
  );
}

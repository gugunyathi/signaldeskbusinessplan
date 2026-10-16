// @lovable.dev/vite-tanstack-config already includes the following — do NOT add them manually
// or the app will break with duplicate plugins:
//   - TanStack devtools (dev-only, first), tanstackStart, viteReact, tailwindcss, tsConfigPaths,
//     nitro (build-only using cloudflare as a default target), VITE_* env injection, @ path alias,
//     React/TanStack dedupe, error logger plugins, and sandbox detection (port/host/strictPort).
// You can pass additional config via defineConfig({ vite: { ... }, etc... }) if needed.
import { defineConfig } from "@lovable.dev/vite-tanstack-config";

export default defineConfig({
  vite: {
    plugins: [
      {
        name: "download-headers",
        configureServer(server) {
          server.middlewares.use((req, res, next) => {
            if (req.url?.includes("Signal_Desk_Business_Plan.pdf")) {
              res.setHeader(
                "Content-Disposition",
                'attachment; filename="Signal_Desk_Business_Plan.pdf"',
              );
              res.setHeader("Content-Type", "application/pdf");
            } else if (req.url?.includes("Signal_Desk_Business_Plan.pptx")) {
              res.setHeader(
                "Content-Disposition",
                'attachment; filename="Signal_Desk_Business_Plan.pptx"',
              );
              res.setHeader(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              );
            }
            next();
          });
        },
      },
    ],
  },
  tanstackStart: {
    // Redirect TanStack Start's bundled server entry to src/server.ts (our SSR error wrapper).
    // nitro/vite builds from this
    server: { entry: "server" },
  },
});

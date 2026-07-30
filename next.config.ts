import type { NextConfig } from "next";
import { appVersion } from "./appVersion";

const nextConfig: NextConfig = {
  // Build-time app version (git commit count + hash) inlined for the client;
  // shown as the bottom-right badge. See appVersion.ts.
  env: {
    NEXT_PUBLIC_APP_VERSION: appVersion(),
  },
  // Pin the workspace root to this project. Without this, a stray
  // package-lock.json in a parent directory makes Next infer the wrong root.
  turbopack: {
    root: __dirname,
  },
};

export default nextConfig;

import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  images: { unoptimized: true },
  env: {
    // Baked at build time. lib/deployState.ts compares this against the
    // DFE_USERS_JSON env var's updatedAt to tell whether the user list this
    // deployment is serving is stale. Must stay a build-time value — reading
    // the clock at request time would always look current.
    DFE_BUILD_TIME: new Date().toISOString(),
  },
};

export default nextConfig;

import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  reactCompiler: true,
  // Force webpack bundler (Turbopack doesn't support Node.js module fallbacks
  // that ExcelJS requires). The --webpack flag is passed via package.json scripts.
  webpack: (config, { isServer }) => {
    if (!isServer) {
      config.resolve.fallback = {
        ...config.resolve.fallback,
        fs:     false,
        path:   false,
        stream: false,
        zlib:   false,
        crypto: false,
        buffer: false,
      };
    }
    return config;
  },
};

export default nextConfig;

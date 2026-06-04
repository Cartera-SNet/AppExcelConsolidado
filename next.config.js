/** @type {import('next').NextConfig} */
const nextConfig = {
  // Allow up to 100MB uploads (Vercel Pro; free tier is 4.5MB)
  experimental: {
    serverActions: {
      bodySizeLimit: '100mb',
    },
  },
};

module.exports = nextConfig;

/** @type {import('next').NextConfig} */

// In production the app is exported to static HTML and served by FastAPI from
// the same origin, so the relative /api paths in the app resolve to the
// backend. `next dev` serves the app on its own port, where those same paths
// hit Next instead and 404 -- so in development they are proxied across.
const isDev = process.env.NODE_ENV === 'development'
const apiTarget = process.env.API_PROXY_TARGET || 'http://127.0.0.1:8000'

const nextConfig = {
  // `output: 'export'` disables rewrites, so it is only applied for builds.
  ...(isDev ? {} : { output: 'export' }),
  trailingSlash: true,
  images: {
    unoptimized: true,
  },
  ...(isDev && {
    async rewrites() {
      return {
        beforeFiles: [
          { source: '/api/:path*', destination: `${apiTarget}/api/:path*` },
        ],
      }
    },
  }),
}

module.exports = nextConfig

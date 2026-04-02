/**
 * Production deployment URL for the add-in (HTTPS only, no trailing slash).
 *
 * Examples:
 *   "https://lower-third-builder.netlify.app"
 *   "https://cdn.contoso.com/lower-third"
 *
 * After you set this, run: npm run build
 * Then upload the contents of /dist to that host (same paths as localhost: taskpane.html, manifest.xml, assets/*).
 *
 * Override without editing this file:
 *   OFFICE_ADDIN_ORIGIN=https://your.host npm run build
 *
 * On Vercel, webpack prefers `VERCEL_PROJECT_PRODUCTION_URL` (stable `*.vercel.app`),
 * then `VERCEL_URL`. Set `OFFICE_ADDIN_ORIGIN` to pin a custom domain in the manifest.
 */
module.exports = {
  productionOrigin: "https://localhost:3000"
};

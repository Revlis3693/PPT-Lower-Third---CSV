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
 * On Vercel, `VERCEL_URL` is set automatically and webpack uses it for the manifest
 * unless OFFICE_ADDIN_ORIGIN is set (use that for a custom domain).
 */
module.exports = {
  productionOrigin: "https://localhost:3000"
};

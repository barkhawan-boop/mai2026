# Monthly activity app

The app is public. Local edits remain saved in the current browser, and pressing Save also replaces that signed-in user's private Cloudflare D1 record. Use **Sync data** on another device to sign in with the same email and load that copy.

The `/api/*` routes are protected by Cloudflare Access with email one-time PIN authentication. The Worker validates the Access JWT, hashes the verified email to choose a D1 row, and never accepts anonymous reads or writes. A first Save signs in and then uploads the current device's data; Sync data downloads the server copy when one exists.

## Deploy

The D1 database, table, and Access application are provisioned in Cloudflare. Access protects only `mai2026.barkhawan.workers.dev/api/*`; the app page remains public. The Access app uses Cloudflare's default email one-time PIN and scopes its JWT cookie to the API path.

Deploy updates from this folder with `wrangler deploy`.

# Monthly activity app

The app is public and saves changes in the current browser's local storage. Pressing Save overwrites the saved copy for this browser profile. Data is not uploaded or shared between devices. Separate browser profiles keep separate copies.

## Deploy

The static app is served by a Cloudflare Worker with Static Assets. Deploy it with `wrangler deploy`. The app does not use a server-side database or account sign-in.

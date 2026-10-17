# Boosthis browser monitoring (experimental)

This integration uses the plain HTML instructions for the Boosthis project
"Outlook automate signature". Production activation requires the steps below.
The signature source is unchanged. Manifest 1.0.0.4 points to this repository's new Pages host, aminalzouabi.github.io; the previous manifest used amin-alz.github.io, which belongs to a separate site. An administrator must upload the updated manifest to the existing Microsoft 365 installation.

## Credentials

Store only the browser-scoped project key as the GitHub Actions repository secret
`BOOSTHIS_PROJECT_KEY` (Settings > Secrets and variables > Actions). Do not put it
in source, issues, pull requests, logs, or this document. Never use an account or
MCP API credential here. The build injects the browser key into `_site/autorun.html`;
that generated folder is ignored and only uploaded as the Pages artifact.

The browser key is visible in the published HTML, its deployment artifact, and
network requests. GitHub Secrets protects source history, not published browser
credentials. Boosthis says a copied browser key can submit fabricated readings.
If no publicly readable key is acceptable, do not activate this integration.

## Activate and verify

1. First verify the existing signature in Outlook without monitoring. The current
   README still lists insertion testing as incomplete.
2. Add the browser-scoped key to the repository secret above.
3. Review and merge the integration branch. Change Settings > Pages > Source to
   **GitHub Actions**, then run **Deploy signature add-in** on `main`.
4. Check the successful Pages deployment and its reported URL. Preserve the current
   manifest URLs unless the deployed URL actually differs; changing a manifest
   runtime URL requires a version increase and an administrator update.
5. Open the deployed `autorun.html` in a normal browser and check Boosthis for a
   reading. This checks delivery only, not Outlook signature insertion.
6. In Outlook on the web and new Outlook on Windows, test new, reply, reply-all,
   and forward messages, the default brand and at least two mapped brands. Verify
   the name, formatting, and no duplicate signature. Repeat with Boosthis requests
   blocked. Test Mac separately; Boosthis lists Chromium as its tested browser.
7. Test classic Outlook on Windows: the manifest loads the unchanged `signature.js`
   directly, so it intentionally has no browser monitoring.
8. Verify actual Outlook check-ins in Boosthis. Do not interpret no data or a good
   blank-page score as proof of signature performance. Event runtimes are short
   lived and may finish before deferred initialization or telemetry upload.

The kit loads after the existing signature scripts and does not gate their
initialization on its download. Its floating bubble is disabled. No mailbox data
is deliberately passed to Boosthis by our integration. The third-party script
executes inside the HTML runtime, so review its current collection policy before
rollout; this hosted URL can change independently of this repository.

## Scope and rollback

This measures the hidden HTML runtime, not Outlook itself or `setSignatureAsync`
duration. Custom signature timings require a separately verified supported API;
they are not invented here. The browser kit is not imported in the JS-only file.

To switch off monitoring, remove `BOOSTHIS_PROJECT_KEY` from GitHub Actions secrets
and rerun the deployment. A build without the key publishes the original HTML.
An existing published artifact retains its key until replaced, and Outlook caches
can delay removal. Revoke the browser key in Boosthis if it must stop being usable.
Keep the previous working deployment available for immediate rollback.

Sources: https://www.boosthis.com/docs/before-you-install and
https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes

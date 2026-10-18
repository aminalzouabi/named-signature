# Ripple Collective Signature

An Outlook add-in that inserts a branded email signature when an assigned user starts a new message, reply, reply-all, or forward.

The add-in reads the user's Outlook profile name and email address, selects a brand by email domain, and inserts the signature automatically. It does not need a scheduled Power Automate flow or a signature button.

## Project status

The current manifest is version **1.0.0.4**. Microsoft 365 accepted the corrected manifest during initial deployment. Successful automatic signature insertion in Outlook still needs to be confirmed before wider rollout.

## How it works

1. An administrator uploads `manifest.xml` to Microsoft 365 and assigns users.
2. An assigned user starts composing a message in a supported Outlook client.
3. Outlook loads the hosted runtime from GitHub Pages.
4. `signature.js` reads the profile name and email domain, builds the signature, and calls `setSignatureAsync`.

On Mac, Outlook on the web, and new Outlook on Windows, the HTML runtime loads the JavaScript. Classic Outlook on Windows uses the JavaScript runtime specified in the manifest.

## Files

| File | Purpose |
| --- | --- |
| [`manifest.xml`](manifest.xml) | Declares the add-in's identity, permissions, supported event, and hosted runtime URLs. This is the file uploaded to Microsoft 365. |
| [`autorun.html`](autorun.html) | Loads Microsoft's Office.js library and the signature script. A blank page when opened directly is expected. |
| [`signature.js`](signature.js) | Contains the brand mappings, signature templates, and insertion logic. |

Uploading the XML does not upload the other files. Outlook retrieves them from their hosted URLs when needed.

## Brand mapping

| Profile email domain | Signature |
| --- | --- |
| `slash.ae` | Slash |
| `fabrica.ae` | Fabrica |
| `fountconcept.ae` | Fount |
| `central.ae` | Central |
| `basearchitecture.ae` | Base |
| `ripplecollective.ae` and all unlisted domains | Ripple Collective |

Each signature includes the user's display name, brand details, and website links. Sub-brand signatures also include a Ripple Collective footer.

The code uses `Office.context.mailbox.userProfile`, not the message's selected From address. The current design assumes each employee sends from their own work mailbox. Shared mailboxes, aliases, and switching the From account need additional handling and testing.

## Hosting

GitHub Pages serves the files from this repository's `main` branch and root folder.

- [Manifest](https://aminalzouabi.github.io/named-signature/manifest.xml)
- [Runtime page](https://aminalzouabi.github.io/named-signature/autorun.html)
- [Signature script](https://aminalzouabi.github.io/named-signature/signature.js)

No build step or package installation is required for these static files. Changes can be committed to `main` or merged into it from another branch. Wait for the Pages deployment to finish before testing.

## Deployment

1. Confirm the hosted files are accessible over HTTPS.
2. Download the current `manifest.xml` as an XML file, not an HTML page or text file.
3. In Microsoft 365 admin center, open **Settings > Integrated apps > Upload custom apps** and select **Office Add-in**. Menu labels may vary.
4. Upload the manifest and assign a small test group or the intended test mailbox.
5. If updating an existing installation, use its update action rather than creating a duplicate.
6. Allow time for deployment to reach Outlook. Microsoft advises that availability can take up to 72 hours.
7. Restart Outlook and test a newly created email and a reply.

If the administrator signs in with a separate admin account, **Just me** may assign the add-in to that account instead of the intended work mailbox.

## Compatibility and limits

- Intended for supported Microsoft 365/Exchange Online accounts and Outlook clients supporting Mailbox requirement set 1.10 and event-based activation.
- The current manifest targets desktop/web clients. Outlook mobile requires additional manifest configuration and testing.
- `OnNewMessageCompose` covers replies and forwards as well as new messages. It does not run merely because an existing draft is reopened.
- `OnNewReplyCompose` is not a supported launch event. Only `OnNewMessageCompose` is registered.
- Deployment success does not verify that the signature code has executed successfully in every Outlook client.

## Permissions and data use

The current manifest requests `ReadWriteMailbox`. The signature API requires only `ReadWriteItem`, so reducing the requested permission is a planned improvement. This README does not change the deployed permission.

The current script reads the Outlook profile name and email address and writes a signature to the composed message. It contains no code to send messages or upload mailbox contents to an external service. Outlook still downloads Office.js and the hosted runtime files.

## Updating the add-in

**Signature wording, websites, or brands:** edit `signature.js` and publish the change. A manifest upload is generally unnecessary if the runtime URLs and permissions stay the same. Outlook and browser caches may delay the update.

**Permissions, events, or runtime URLs:** edit `manifest.xml`, increase its version, keep the same add-in ID, and upload the revised manifest through the existing installation's update action.

Opening `signature.js?v=5` in a browser only requests that URL. It does not automatically update installed Outlook references or guarantee an Outlook cache refresh. If versioned script URLs are used, keep the HTML script reference and the manifest's JavaScript runtime URL consistent.

## Troubleshooting

| Symptom | Check |
| --- | --- |
| Signature does not appear | Confirm assignment to the actual mailbox, client support, deployment time, and access to the hosted files. Start a new message rather than opening a saved draft. |
| Wrong brand | Check the profile email domain against `BRANDS` in `signature.js`. Unlisted domains use Ripple Collective. |
| Wrong name | Check the user's Microsoft 365 display name and the profile Outlook returns. |
| Old signature text | Check the published script and Pages deployment, then allow for client caching. |
| Deployment error | Check valid XML, supported events, runtime URLs, icon URLs, and the full admin-center error details. A generic error does not establish a single cause. |
| Conflicting signatures | Check existing automatic signatures and other signature add-ins during testing. |

## Remaining checks

- [ ] Confirm insertion on Outlook for Mac and Outlook on the web.
- [ ] Test new messages, replies, reply-all, and forwards.
- [ ] Test at least two sub-brands and the Ripple Collective default.
- [ ] Reduce the manifest permission to `ReadWriteItem` and retest.
- [x] Remove the unused reply handler and complete events on handled API errors and exceptions.
- [ ] Expand user assignment after the tests pass.

## Microsoft references

- [Signature add-in sample](https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/outlook-add-in-set-signature/)
- [Event-based activation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation)
- [Outlook add-in permissions](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)
- [Body and signature API](https://learn.microsoft.com/en-us/javascript/api/outlook/office.body?view=outlook-js-preview)
- [Centralized deployment FAQ](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-faq?view=o365-worldwide)

## Signature font update (1.0.0.4)

The generated signature uses scoped internal CSS to request Helvetica at 10pt with a 14pt line height for its text and links. Each text segment has a font-family fallback attribute, and the user's name and company remain bold. It avoids inline CSS, which Microsoft documents as unsupported by setSignatureAsync. The style is scoped to the signature and does not set the message body's font.

Helvetica must be available on the viewing device. Arial and the system sans-serif font are fallbacks; the add-in cannot force remote recipients to install Helvetica. Plain-text emails receive a plain-text signature and cannot preserve font formatting.

The HTML loader and JavaScript runtime URL both use version 1.0.0.4. The manifest now targets this repository owner's GitHub Pages site. Before uploading the updated manifest, enable GitHub Pages on main/root and confirm autorun.html and signature.js are served successfully. Keep the add-in ID when updating an existing installation. Existing installations pointing to amin-alz.github.io continue to use the original site until their manifest is updated.

Run the regression checks with Node.js:

```sh
node tests/signature.test.cjs
```

These checks exercise signature generation and mocked Office callbacks, not Outlook's rendering engine. Before broad rollout, send new messages, replies and forwards from Outlook for Windows, Mac and web, then inspect the received messages for consistent size and font fallback. Confirm the CSS survives insertion and sending; internal CSS can also be rewritten by mail clients.

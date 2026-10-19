# Signature Manager

An in-house replacement for Exclaimer / CodeTwo. It puts one standard email signature on every
Children's Alliance mailbox and fills it in from each person's Microsoft 365 (Entra ID)
profile. A change to the template, or to someone's job title in Entra ID, reaches Outlook
automatically. Nobody edits their own signature.

## How it works

```
                 Entra ID (name, job title, department, phones, address, extension attributes)
                         │ Microsoft Graph
        ┌────────────────┴──────────────────┐
        ▼                                   ▼
 Outlook add-in (every user)          Admin console (IT / comms)
 addin/launchevent.js                 index.html
  · runs on every new email,           · browse all staff, preview each signature
    reply, forward and meeting         · edit brand / template, campaign banners
  · reads the sender's /me profile     · flags people with missing profile data
  · renders template.json              · exports template.json and per-user HTML
  · inserts the signature, turns off           │
    Outlook's own signature                    ▼
        ▲                             deploy/Deploy-Signatures.ps1 (optional)
        │ template.json                 · server-side push to Outlook on the web / new Outlook
        └──── published on GitHub Pages
```

One renderer (`src/renderer.js`) is shared by the add-in, the admin console and the tests, so the
preview is exactly what people send.

| Feature you get from Exclaimer/CodeTwo | Here |
|---|---|
| Signature built from Microsoft 365 profile fields | `user.*` placeholders, with custom fields from extension attributes via `fieldMap` |
| Lines hidden when a field is empty | `{{#user.mobilePhone}}…{{/user.mobilePhone}}` |
| Different signature for replies/forwards | `replyHtml` |
| Scheduled, department-targeted campaign banners | `campaigns` (start/end dates, departments, exclusions) |
| Works in Outlook web, new Outlook, Mac, classic Windows | Event-based Outlook add-in, deployed centrally |
| Signature appears in the sender's Sent Items | Yes: it's added in the client, not by a mail flow rule |
| Data-quality report | "Download data gaps (CSV)" in the admin console |

## Files

| Path | What it is |
|---|---|
| `template.json` | **The live signature.** Brand settings, HTML template, reply template, campaigns. |
| `index.html`, `admin.css`, `dist/admin.js` | Admin console |
| `addin/manifest.xml` | Outlook add-in manifest to upload in the Microsoft 365 admin center |
| `addin/commands.html`, `addin/launchevent.js` | Add-in runtime (built) |
| `deploy/Deploy-Signatures.ps1` | Optional server-side push via Exchange Online PowerShell |
| `src/` | Sources: renderer, add-in, admin console, `config.js` |
| `assets/` | Logo, banner and add-in icons (**placeholders**, replace with real artwork) |
| `test/` | `npm test`: renderer tests, plus the built add-in run against a fake Outlook |

## Setup (one-off, about 30 minutes)

### 1. Register the app in Entra ID

Entra admin center > App registrations > New registration:

- Name: `Signature Manager`. Supported account types: this organisation only.
- **Authentication > Add a platform > Single-page application**, with these redirect URIs:
  - `https://danbrimmicombe.co.uk/signatures/auth-redirect.html` (admin console)
  - `brk-multihub://danbrimmicombe.co.uk` (the Outlook add-in's nested app authentication broker)
- **API permissions** (Microsoft Graph, delegated): `User.Read` and `User.Read.All`.
  Click **Grant admin consent**. The add-in can't show a sign-in prompt while it runs, so it
  only works with tenant-wide consent.
- Copy the **Application (client) ID** and the **Directory (tenant) ID**.

`User.Read.All` is only used when an admin signs in to the console. The add-in only asks for
`User.Read`, so each person reads only their own profile.

### 2. Configure and build

Edit `src/config.js` with the client ID, tenant ID and the URL the folder is published at. Then:

```sh
cd signatures
npm install
npm run build
npm test
```

Commit the `dist/` and `addin/launchevent.js` output. GitHub Pages serves the built files as they
are, with no build step on the server.

> **Hosting:** everything is currently addressed at `https://danbrimmicombe.co.uk/signatures`.
> For production, consider an organisation-owned domain or repository, such as
> `signatures.childrensalliance.org.uk`. If you move it, update `baseUrl` in `src/config.js`, every
> URL in `addin/manifest.xml`, `logoUrl`/`imageUrl` in `template.json` and the Entra redirect URIs.
> Keep "Enforce HTTPS" on in the GitHub Pages settings, because Outlook refuses add-ins served
> over plain HTTP.

### 3. Fill in the profile data

The signature only shows what Entra ID holds. Sign in to the admin console and click
**Download data gaps (CSV)** to see who is missing a job title, phone number and so on. Fix the
gaps in Entra ID or your HR sync. Pronouns are read from `extensionAttribute1` by default: set them
with `Set-Mailbox -CustomAttribute1 "she/her"` in Exchange Online, or change `fieldMap` in
`template.json` to point at another attribute.

### 4. Pilot, then roll out the add-in

1. Microsoft 365 admin center > Settings > **Integrated apps** > Upload custom apps >
   Office Add-in > upload `addin/manifest.xml`.
2. Assign it to a **pilot group** first, for example IT plus a few people from each team.
3. Deployment can take up to 24 hours (sometimes 72) to reach Outlook. In the pilot, check:
   - [ ] Outlook on the web and new Outlook: new email, reply, forward, new meeting
   - [ ] Classic Outlook for Windows (see "Known limitations")
   - [ ] Outlook for Mac
   - [ ] Signatures from a person with no mobile, no job title or pronouns look right
   - [ ] Signatures received in Gmail and on a phone render properly
4. Assign to everyone, then remove Exclaimer/CodeTwo and ask staff to delete their old personal
   signatures, or leave them: the add-in turns off Outlook's own signature on each message it handles.

Optional, for the switchover day: export from the console and run
`deploy/Deploy-Signatures.ps1 -ExportPath signatures-export.json -WhatIf`, then without `-WhatIf`.
This sets the signature server-side for Outlook on the web, but only while roaming signatures are
postponed for the tenant (`Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true`).
The script warns you if they aren't.

## Day-to-day

- **Change the design or wording:** open the console, edit, check a few people in the preview,
  click **Download template.json** and commit it over `signatures/template.json`. Clients pick it up
  within `cacheMinutes` (60 by default), or straight away with the **Refresh signature** button on
  the compose ribbon.
- **Someone's details change:** update them in Entra ID. Their next email picks the change up once
  the cache expires.
- **Run a campaign banner:** add an entry to `campaigns` with `enabled: true`, `start`, `end`, and
  optionally `departments`. Use the console's "As on" date to preview it. The first live campaign in
  the list wins.
- **New starter:** nothing to do. The add-in is assigned to everyone and reads their profile.

Test the templates in real clients. Outlook for Windows renders HTML with Word: stick to tables,
inline styles, PNG/JPG images at absolute HTTPS URLs (no SVG), and web-safe fonts.

## Known limitations and risks

- **Classic Outlook for Windows.** Event-based add-ins there run in a JavaScript-only runtime.
  Microsoft says nested app authentication is generally available in classic Outlook, but it isn't
  certain that it works inside that runtime. If it doesn't, the add-in still inserts a signature but
  only has the sender's name and email (plus any `defaults`). **Check this first in the pilot.**
  If it's a problem, the usual fix is an Intune/logon script that writes the rendered HTML into
  `%APPDATA%\Microsoft\Signatures` for classic Outlook. The "Export all" output contains what that
  script needs.
- **Mobile.** Outlook for iOS/Android supports event-based add-ins only partially. Check the
  mobile signature in the pilot.
- **No server.** Everything runs in Outlook or the admin's browser, so there is nothing to patch or
  pay for. The flip side is that `template.json` is public, so keep personal data out of it
  (profile data is only ever fetched at send time).
- **Mail sent from other devices or apps** (scanners, CRM, phones without the add-in) doesn't get
  the signature. Exclaimer's server-side mode covers that. An Exchange transport-rule disclaimer is
  the free fallback if you need it.

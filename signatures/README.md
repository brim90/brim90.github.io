# Water Babies email signature

Puts the Water Babies signature on every email sent from Outlook, filled in from each person's
Microsoft 365 profile. It replaces Exclaimer/CodeTwo for one job only: rolling out this signature.

It works in classic Outlook for Windows, new Outlook, Outlook for Mac, Outlook on the web, and
Outlook for iOS and Android. New emails get the full signature. Replies and forwards get a short,
text-only version, so long threads aren't full of repeated logos.

## How it works

```
Microsoft 365 profile ──► Outlook add-in ──► signature inserted when an email is started
(name, job title,         (addin/)            (Outlook's own signature is switched off
 phone, mobile)              ▲                 for that email)
                             │
              signatures.json + designs/*.html + assets/
              (published on GitHub Pages)
```

| What you change | Where |
|---|---|
| Company name, number, address, switchboard, website | `signatures.json` > `company` |
| The design | `designs/standard.html` (new emails) and `designs/standard-reply.html` (replies and forwards) |
| Logo and icons | `assets/logo.png` and `assets/logo-white.png` (208×148, shown at 104×74), and `assets/icons.png` (188×52, shown at 94×26). Replace them with same-size exports, keeping the file names. |
| Someone's name, job title or numbers | Their Microsoft 365 profile (Entra ID) |

The design follows the *Water Babies Email Signature 2026* guide: 500px wide, three columns,
Arial, brand blue `#0070C0`, cyan `#00B0F0`, orange `#FF6600` and panel `#E5F4F8`. It is built
from tables and inline styles because that's what renders consistently in Outlook for Windows,
which uses Word to display email, and in mobile apps. The logo is the blue version on white with
one "W" width of clear space, as the brand guidelines require. The starfish, splash and anchor
icons are the ones from the guide.

Where it deliberately differs from the guide, for better display and delivery:
- **Hosted images, not embedded ones.** Gmail strips embedded images and Outlook turns them into
  attachments (paperclip icons, higher spam scores).
- **Two images instead of four.** The three icons are one strip, so there are fewer requests,
  fewer "download pictures?" placeholders and a better image-to-text ratio.
- **Double-resolution logo,** so it's sharp on phones and modern laptops.
  **The icons are still 1x:** the guide only has them at 26px. Swap in a 2x export from
  marketing's originals when you get them.
- **Readable legal text:** 7pt in brand blue instead of 4.5pt cyan, with the company number and
  registered office (a legal requirement for UK company emails).
- **An underlined website link.**
- **Dark mode:** apps that support it (Apple Mail, Outlook for iOS/Android/Mac, Outlook.com)
  switch to `assets/logo-white.png` and lighten the dark blue text. Every other app, including
  classic Outlook for Windows and Gmail, shows the blue logo on its white box.

Rules applied to everyone:
- **T:** is their own office number from Microsoft 365, or the 0330 223 6777 switchboard if they
  don't have one.
- **M:** only appears for people who have a mobile in Microsoft 365.

## Preview a change

Open `https://signatures.waterbabies.co.uk/` (or run `npx serve` in this folder locally). It
shows the signature exactly as Outlook will insert it, and you can type any name, title or
numbers to test with. **Copy signature** lets you paste it into a real Outlook email.

## Changing or adding designs

- **Change the design:** edit `designs/standard.html` or `designs/standard-reply.html`, check it in
  the preview, commit. Outlook
  picks it up within an hour (`cacheMinutes` in `src/config.js`).
- **Add another design**, for example for franchisees or another country:
  1. Copy `designs/standard.html` to `designs/franchise.html` and edit it. A reply version is
     optional: without one, replies get the full signature.
  2. List it in `signatures.json`, and add a rule saying who gets it. The first matching rule wins.
     Everyone else gets `defaultDesign`.

     ```json
     "rules": [
       { "design": "franchise", "when": { "department": ["Franchise Support"] } }
     ],
     "designs": {
       "standard": { "new": "designs/standard.html", "reply": "designs/standard-reply.html" },
       "franchise": { "new": "designs/franchise.html" }
     }
     ```

     Rules can match `department`, `officeLocation`, `companyName` or `email`.
- **Switch to a new design for everyone:** point `defaultDesign` at it.

Placeholders available in designs: `{{user.displayName}}`, `{{user.jobTitle}}`,
`{{user.phone}}`, `{{user.phoneHref}}`, `{{user.mobilePhone}}`, `{{user.mobilePhoneHref}}`,
`{{user.email}}`, every `{{company.*}}` value, and `{{assets}}` for image URLs. Wrap a line in
`{{#user.mobilePhone}}…{{/user.mobilePhone}}` to drop it when the field is empty.

Run `npm test` after any change. It checks every design for leftover placeholders, missing
images, and images without a width and height (Outlook for Windows needs both).

## One-off setup

### 1. Register the app in Entra ID (5 minutes)

Entra admin center > App registrations > New registration:

- Name `Water Babies Signature`, single tenant.
- **Authentication** > Add a platform > **Single-page application**, redirect URI
  `brk-multihub://signatures.waterbabies.co.uk`.
- **API permissions:** Microsoft Graph > Delegated > `User.Read`, then **Grant admin consent**.
  Each person's add-in can only read their own profile.
- Copy the Application (client) ID and Directory (tenant) ID into `src/config.js`, then run:

  ```sh
  cd signatures && npm install && npm run build && npm test
  ```

  Commit the rebuilt `addin/launchevent.js`.

### 2. Publish

Everything is served from `https://signatures.waterbabies.co.uk`. Images and links on the
sender's own domain matter for delivery: spam filters treat a mismatch as a phishing signal.

1. Put the contents of this folder at the root of its own GitHub repository. `CNAME` and
   `.nojekyll` are already here.
2. In the repository, go to Settings > Pages > Deploy from a branch > `main`, folder `/ (root)`,
   with custom domain `signatures.waterbabies.co.uk`.
3. DNS for waterbabies.co.uk: add a `CNAME` record for `signatures` pointing to
   `<github-owner>.github.io` (the account or organisation that owns the repository).
4. Recommended: verify the domain in GitHub (owner Settings > Pages > Add a verified domain, then
   add the TXT record it gives you). This stops anyone else's GitHub site claiming
   `*.waterbabies.co.uk`.
5. Once the certificate is issued (usually within an hour), tick **Enforce HTTPS**.

To host it somewhere else, change `baseUrl` in `src/config.js` and run `npm run build`. The build
regenerates `addin/manifest.xml` with the new URLs. Update the Entra redirect URI to match.
If Outlook already has the add-in, upload the new manifest in Integrated apps.

### 3. Roll out to Head Office first

1. Create a group for the pilot. Either:
   - an assigned **Microsoft 365 or security group**, "Signature - Head Office", with the Winslade
     House staff added, or
   - a **dynamic group** (needs Entra ID P1), e.g. `(user.officeLocation -eq "Winslade House")`.
     Check the attribute is actually filled in first.
2. Microsoft 365 admin center > Settings > **Integrated apps** > Upload custom apps > Office Add-in
   > upload `addin/manifest.xml`, and assign it to **that group only**.
3. Allow up to 24 hours (occasionally longer) for it to reach Outlook. Then check:
   - [ ] Classic Outlook for Windows: new email, reply, forward
   - [ ] New Outlook for Windows / Outlook on the web
   - [ ] Outlook for Mac
   - [ ] Outlook for iOS and Android (new email, reply)
   - [ ] Replies and forwards get the short version everywhere
   - [ ] Someone with a mobile, someone without, and someone with a long job title
   - [ ] What recipients see in Outlook, Gmail and the iPhone Mail app, including dark mode
4. Ask pilot users to delete their old personal signature. The add-in switches it off per email
   anyway, but a clean slate avoids confusion.
5. Once happy, change the add-in's assignment in Integrated apps from the group to everyone.
   Remove Exclaimer/CodeTwo at the same time, so nobody ends up with two signatures.

## Known limitations

- **Classic Outlook for Windows:** the add-in signs in to read the person's profile using nested
  app authentication. Microsoft says this is supported in classic Outlook, but the pilot is
  where to confirm it. If it fails, the signature still goes in, with only the person's name.
  Check that job titles appear for pilot users on classic Outlook.
- **Mobile:** Outlook for iOS/Android runs the add-in when an email is started, but the 500px
  signature is scaled down to fit the phone screen, as the guide's own signature is today. If the
  app can't tell the add-in whether the email is a reply, it uses the full signature.
- **Editing a draft** doesn't re-insert the signature (Outlook doesn't fire the event for drafts).
- **Email sent from outside Outlook** (CRM, scanners, other mail apps) won't get the signature.

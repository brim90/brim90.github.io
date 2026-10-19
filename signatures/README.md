# Water Babies email signature

Puts the Water Babies signature on every email sent from Outlook, filled in from each person's
Microsoft 365 profile. It replaces Exclaimer/CodeTwo for one job only: rolling out this signature.

It works in classic Outlook for Windows, new Outlook, Outlook for Mac, Outlook on the web, and
Outlook for iOS and Android. It covers new emails, replies and forwards.

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
| The design | `designs/standard.html` |
| Logo and icons | `assets/` |
| Someone's name, job title or numbers | Their Microsoft 365 profile (Entra ID) |

The design follows the *Water Babies Email Signature 2026* guide: 500px wide, three columns,
Arial, brand blue `#0070C0`, cyan `#00B0F0`, orange `#FF6600` and panel `#E5F4F8`. It is built
from tables and inline styles because that's what renders consistently in Outlook for Windows,
which uses Word to display email, and in mobile apps. The logo is the blue version on white with
one "W" width of clear space, as the brand guidelines require. The starfish, splash and anchor
icons are the ones from the guide.

Rules applied to everyone:
- **T:** is their own office number from Microsoft 365, or the 0330 223 6777 switchboard if they
  don't have one.
- **M:** only appears for people who have a mobile in Microsoft 365.

## Preview a change

Open `https://danbrimmicombe.co.uk/signatures/` (or run `npx serve` in this folder locally). It
shows the signature exactly as Outlook will insert it, and you can type any name, title or
numbers to test with. **Copy signature** lets you paste it into a real Outlook email.

## Changing or adding designs

- **Change the design:** edit `designs/standard.html`, check it in the preview, commit. Outlook
  picks it up within an hour (`cacheMinutes` in `src/config.js`).
- **Add another design**, for example for franchisees or another country:
  1. Copy `designs/standard.html` to `designs/franchise.html` and edit it.
  2. List it in `signatures.json`, and add a rule saying who gets it. The first matching rule wins.
     Everyone else gets `defaultDesign`.

     ```json
     "rules": [
       { "design": "franchise", "when": { "department": ["Franchise Support"] } }
     ],
     "designs": {
       "standard": "designs/standard.html",
       "franchise": "designs/franchise.html"
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
  `brk-multihub://danbrimmicombe.co.uk`.
- **API permissions:** Microsoft Graph > Delegated > `User.Read`, then **Grant admin consent**.
  Each person's add-in can only read their own profile.
- Copy the Application (client) ID and Directory (tenant) ID into `src/config.js`, then run:

  ```sh
  cd signatures && npm install && npm run build && npm test
  ```

  Commit the rebuilt `addin/launchevent.js`.

### 2. Publish

The folder is served by GitHub Pages from the default branch. Keep **Enforce HTTPS** on.

If you'd rather host it on a Water Babies domain such as `signatures.waterbabies.co.uk`, change
`baseUrl` in `src/config.js`, every URL in `addin/manifest.xml`, and the redirect URI.

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
  signature is scaled down to fit the phone screen, as the guide's own signature is today.
- **Editing a draft** doesn't re-insert the signature (Outlook doesn't fire the event for drafts).
- **Email sent from outside Outlook** (CRM, scanners, other mail apps) won't get the signature.

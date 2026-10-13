# CloudNest

## QUICK START (this zip is the whole project)

1. Create a **Public** GitHub repo named exactly **mrvirussofficial.github.io**
2. Extract this zip on a laptop and upload EVERYTHING (including the `.github`, `.well-known` and `android` folders) to the repo.
3. Settings -> Pages -> Source: **GitHub Actions**.
4. Wait for both workflows in the Actions tab to turn green.
5. Releases -> download `CloudNest.apk` and install it. It opens full screen (no browser bar).
6. Register these addresses:
   - Google (Authorized JavaScript origins): `https://mrvirussofficial.github.io`
   - Microsoft (SPA redirect URI): `https://mrvirussofficial.github.io/`
7. Open the app -> gear icon -> paste both Client IDs -> Save & reload.

---

One app for all your Google Drive and OneDrive accounts. Connect as many accounts as you like, then browse, preview, upload (with progress), download, create folders and delete, all from one screen.

Everything runs in the browser. You sign in through the official Google / Microsoft popups, and files travel directly between this page and Google / Microsoft. There is no CloudNest server or database.

## What works

| | Google Drive | OneDrive |
|---|---|---|
| Connect several accounts | yes | yes |
| Storage used / total per account | yes | yes |
| Browse folders, grid / list, sort | yes | yes |
| Upload (many files, progress, cancel, retry, drag & drop) | yes (resumable) | yes (chunked for big files) |
| New folder | yes | yes |
| Delete | moves to Drive trash | moves to recycle bin |
| Download | yes (Docs/Sheets/Slides export to docx/xlsx/pptx) | yes |
| Preview | images, PDF, DOCX, XLSX/CSV, PPTX (text), text/code, audio, video, Google Docs/Sheets/Slides | same |

## Step 1 - get your two Client IDs (one time, free)

Google and Microsoft only allow sign-in for apps that are registered, so you register CloudNest once under your own account. Client IDs are public identifiers, not passwords.

You first need to know the address where the app will live:
- GitHub Pages: `https://YOUR-USERNAME.github.io/YOUR-REPO/`
- Local testing: `http://localhost:8080/`

(Open the app, tap the gear icon, and it shows the exact addresses to register.)

### Google

1. Go to https://console.cloud.google.com and create a project.
2. APIs & Services -> Library -> search **Google Drive API** -> Enable.
3. APIs & Services -> **OAuth consent screen** (now called "Google Auth Platform"):
   - User type **External**, fill in app name (CloudNest) and your email.
   - Add the scope `https://www.googleapis.com/auth/drive`.
   - Under **Test users**, add every Gmail address you want to connect. Leave the app in **Testing** mode.
4. APIs & Services -> Credentials -> Create credentials -> **OAuth client ID** -> type **Web application**.
   - **Authorized JavaScript origins**: add only the origin, e.g. `https://YOUR-USERNAME.github.io` and `http://localhost:8080` (no path, no trailing slash).
   - Leave redirect URIs empty. There is no client secret to keep.
5. Copy the **Client ID** (ends in `.apps.googleusercontent.com`).

While the app is in Testing mode Google shows "Google hasn't verified this app". Tap **Advanced -> Go to CloudNest**. That is expected for a personal app.

### Microsoft

1. Go to https://portal.azure.com -> **Microsoft Entra ID** -> **App registrations** -> **New registration**. (A personal Microsoft account works; app registrations are free.)
2. Name: CloudNest. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**.
3. Redirect URI: platform **Single-page application (SPA)**, value = the full address with trailing slash, e.g. `https://YOUR-USERNAME.github.io/YOUR-REPO/`. Register.
4. Copy the **Application (client) ID**.
5. API permissions -> Add -> Microsoft Graph -> **Delegated** -> `Files.ReadWrite.All` (and `User.Read`, usually already there).
   Work/school tenants may need an admin to approve this; personal accounts do not.
6. Authentication -> add another SPA redirect URI `http://localhost:8080/` if you also test locally.

## Step 2 - put the IDs in the app

Either:
- open the app, tap the **gear icon**, paste both IDs, **Save & reload** (stored in that browser only), or
- edit `config.js` and commit it (safe, they are public IDs).

## Step 3 - run it

**GitHub Pages (easiest, also gives you the phone app)**
1. Create a GitHub repository and upload all files in this folder (keep the structure).
2. Settings -> Pages -> Source: **Deploy from a branch** -> `main` / `(root)`. (Or choose "GitHub Actions" to use the included workflow.)
3. Open `https://YOUR-USERNAME.github.io/YOUR-REPO/`.

**On your computer**
```
cd cloudnest
python3 -m http.server 8080
```
Then open http://localhost:8080. Opening `index.html` by double-click will NOT work for sign-in.

**Install like an app**
- Android: Chrome menu -> Install app / Add to Home screen.
- Windows: Chrome or Edge -> install icon in the address bar.

## Good to know

- **Google sessions last 1 hour.** Google's browser-only sign-in does not hand out refresh tokens. After an hour, tapping the account (or the Reconnect button) opens a popup that closes by itself. OneDrive renews silently.
- Allow popups for the site, otherwise sign-in is blocked.
- Delete is never permanent: items go to the provider's trash / recycle bin.
- Folders cannot be dragged in yet; upload files, or create the folder first.
- Google downloads are held in browser memory, so the app stops at about 1.5 GB per file. Use "Open in Google Drive" for bigger ones.
- "Shared with me" and shared drives are not listed yet.
- Scope `drive` (full access) is needed to see files you did not create with this app. If you only want files uploaded by CloudNest, change `GOOGLE_SCOPES` in `config.js` to `https://www.googleapis.com/auth/drive.file`.

## Troubleshooting

| Message | Fix |
|---|---|
| Google: `origin_mismatch` | The page address is not listed under Authorized JavaScript origins (origin only, exact match). |
| Google: "Access blocked / app not verified" | Add your Gmail under Test users, then Advanced -> Go to CloudNest. |
| Microsoft: `AADSTS50011` redirect mismatch | The SPA redirect URI must equal the address in the gear dialog exactly, including the trailing `/`. |
| Microsoft: `AADSTS700016` / "unauthorized_client" | Account types must include personal accounts, and the platform must be SPA. |
| Sign-in popup never appears | Allow popups for this site. |

## Files

```
index.html     the UI (your CloudNest design)
app.js         sign-in, Drive/Graph calls, uploads, previews
config.js      Client IDs
sw.js, manifest.webmanifest, icons/   installable app support
.github/workflows/pages.yml           optional GitHub Pages deploy
```

## Android APK (automatic)

Every push to `main` builds an APK with GitHub Actions (`.github/workflows/android.yml`).

1. Repo -> **Settings -> Pages -> Source: GitHub Actions** (once).
2. Push the files. Wait for both workflows (Actions tab) to turn green.
3. Repo -> **Releases** -> latest -> download `CloudNest.apk` and install it (allow "install unknown apps").
4. Register the website address (see Step 1) with Google / Microsoft.

The APK is a Trusted Web Activity: it opens your GitHub Pages site in Chrome, so Google / Microsoft sign-in works.
The APK is signed with `android/debug.keystore`, so new builds install over the old one.
Updating the site (index.html, app.js ...) updates the app automatically; you only need a new APK if you change `android/`.

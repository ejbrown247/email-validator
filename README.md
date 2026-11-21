# Email Validator for Mac

A free, simple macOS desktop app that opens Excel (`.xlsx`) or CSV files, checks email **syntax and domain mail-routing DNS**, and exports a new Excel workbook with results. **It cannot confirm that an individual inbox exists.** No API key or paid verification service is needed; DNS checks require an internet connection.

## For your friend

1. Download the `EmailValidator-arm64` build artifact from GitHub Actions (for a recent Apple Silicon Mac), unzip the downloaded artifact, then unzip `EmailValidator-arm64.zip` inside it.
2. Move `EmailValidator.app` into Applications if desired.
3. Open the app. Because it is **unsigned and not notarized**, macOS may block the first launch. In Finder, Control-click the app and choose **Open** if available; otherwise go to **System Settings → Privacy & Security → Open Anyway** after attempting to open it. Only bypass this warning for software you trust.
4. Choose a spreadsheet, select the email column, click **Validate and Save Results…**, and choose an output file.

Statuses:
- **DOMAIN ACCEPTS MAIL**: syntax and DNS mail-routing checks passed; **mailbox existence unknown**.
- **INVALID SYNTAX**: email format invalid.
- **DOMAIN CANNOT RECEIVE MAIL**: mail-routing DNS checks failed.
- **INCONCLUSIVE**: temporary DNS failure or another uncertainty.
- **MISSING**: empty email cell.

No email messages are sent, and no SMTP recipient probing is performed. Do not treat results as proof that an individual person or mailbox exists.

## Build from Windows using GitHub Actions

1. Create a GitHub repository (private is fine). Upload all project files **including `.github/workflows/build-mac.yml`**, preserving folder structure.
2. Go to the repository **Actions** tab, choose **Build macOS apps**, then **Run workflow**. Pushing to `main` also triggers a build.
3. After both jobs finish, open the run and download the `EmailValidator-arm64` artifact for Apple Silicon. An Intel build is also produced.
4. Send the resulting ZIP to your friend. GitHub Actions hosted-runner usage and artifact retention depend on your GitHub plan and limits.

These builds run on macOS runners; PyInstaller does not cross-compile from Windows. The app is unsigned, so Gatekeeper warnings are expected. Signing/notarization for seamless distribution requires an Apple Developer setup.

## Local development (Mac)

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Build locally:

```bash
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --windowed --name EmailValidator --collect-all email_validator app.py
```

## Privacy

Email addresses are read locally and used for DNS domain checks. They are not uploaded to an AI service. DNS requests necessarily disclose domain names to DNS resolvers. Keep your spreadsheet private.

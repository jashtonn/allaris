# OutlookAgentBridge (Actual Project + Downloadable EXE Build)

You were right to call this out — a `.md` file alone is not useful.

This repo now includes a real Windows app project at:
- `OutlookAgentBridge/OutlookAgentBridge.csproj`
- `OutlookAgentBridge/Program.cs`

It exposes localhost APIs to:
- read unread mail
- create folders
- move mail to folders
- set categories
- draft replies

All operations happen via local Outlook Classic COM automation, so no Microsoft 365 admin consent is required.

---

## What was added so you can get an EXE

1. **Runnable .NET project** (`OutlookAgentBridge/*`).
2. **Build script** (`scripts/build-windows-exe.ps1`).
3. **GitHub Actions workflow** (`.github/workflows/build-windows-exe.yml`) that builds and uploads a downloadable artifact:
   - Artifact name: `OutlookAgentBridge-win-x64`

---

## Fastest way to get a downloadable EXE

### Option A: GitHub Actions (recommended)

1. Push this repo to GitHub.
2. Open **Actions** tab.
3. Run workflow **Build Windows EXE**.
4. Download artifact `OutlookAgentBridge-win-x64` from the workflow run.
5. Extract and run `OutlookAgentBridge.exe` on your Windows machine.

---

## Run the EXE on Windows

In PowerShell:

```powershell
$env:OUTLOOK_BRIDGE_TOKEN = "replace-with-long-random-token"
.\OutlookAgentBridge.exe
```

Service listens on: `http://127.0.0.1:5077`

Use header on every request:
- `X-Bridge-Token: <same token>`

---

## API endpoints (implemented)

- `GET /health`
- `GET /email/unread?limit=25`
- `POST /folder/create`
- `POST /email/move`
- `POST /email/category`
- `POST /email/draft-reply`

---

## Quick test commands (PowerShell)

```powershell
$h = @{ "X-Bridge-Token" = "replace-with-long-random-token" }
Invoke-RestMethod -Headers $h -Uri "http://127.0.0.1:5077/email/unread?limit=5"
```

Create folder:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" `
  -Uri "http://127.0.0.1:5077/folder/create" `
  -Body '{"name":"AI - Needs Review"}'
```

Move email:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" `
  -Uri "http://127.0.0.1:5077/email/move" `
  -Body '{"entryId":"<ENTRY_ID>","targetFolderName":"AI - Needs Review"}'
```

Set categories:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" `
  -Uri "http://127.0.0.1:5077/email/category" `
  -Body '{"entryId":"<ENTRY_ID>","categories":["AI","Follow Up"]}'
```

Draft reply:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" `
  -Uri "http://127.0.0.1:5077/email/draft-reply" `
  -Body '{"entryId":"<ENTRY_ID>","body":"Thanks — I reviewed this and will reply in detail by tomorrow."}'
```

---

## If you want to build locally on Windows

```powershell
./scripts/build-windows-exe.ps1
```

Output:

```text
OutlookAgentBridge/bin/Release/net8.0-windows/win-x64/publish/
```

---

## Notes

- You must run as the same Windows user/profile that has Outlook configured.
- Outlook Classic desktop app must be installed and signed in.
- COM calls run on STA threads in the implementation.

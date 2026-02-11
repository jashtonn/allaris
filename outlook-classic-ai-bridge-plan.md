# OutlookAgentBridge (Real App + Direct Download ZIP Release)

You are absolutely right: you should not have to run `dotnet new` just to use this.

This repo now has:
- a real app project (`OutlookAgentBridge/`)
- a CI build workflow
- a **Release workflow** that publishes a downloadable ZIP containing `OutlookAgentBridge.exe`

---

## What to do now (no SDK on your PC)

### 1) Download ready-to-run app from GitHub Releases

1. Open this repo on GitHub.
2. Go to **Actions** → run **Release Windows EXE**.
3. For `tag`, enter something like `v0.1.0`.
4. Wait for completion.
5. Go to **Releases** and download `OutlookAgentBridge-win-x64.zip`.
6. Unzip anywhere (example: `C:\outlook-agent-bridge\app`).

This path does **not** require .NET SDK on your machine.

---

## Run the app on Windows

Open PowerShell in the extracted folder:

```powershell
$env:OUTLOOK_BRIDGE_TOKEN = "replace-with-long-random-token"
.\OutlookAgentBridge.exe
```

API base URL:
- `http://127.0.0.1:5077`

Header required on all calls:
- `X-Bridge-Token: <same token>`

---

## Implemented endpoints

- `GET /health`
- `GET /email/unread?limit=25`
- `POST /folder/create`
- `POST /email/move`
- `POST /email/category`
- `POST /email/draft-reply`

These cover what you asked for:
- read emails
- sort/move into folders
- categorize
- create folders
- draft replies

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

## If you are building (optional)

You only need this if you want to compile locally.

```powershell
./scripts/build-windows-exe.ps1
```

If `dotnet` is missing, install SDK first:

```powershell
winget install Microsoft.DotNet.SDK.8
```

---

## Files added for this

- App: `OutlookAgentBridge/OutlookAgentBridge.csproj`, `OutlookAgentBridge/Program.cs`
- Build script: `scripts/build-windows-exe.ps1`
- CI build artifact workflow: `.github/workflows/build-windows-exe.yml`
- Release ZIP workflow: `.github/workflows/release-windows-exe.yml`

---

## Important notes

- Run the bridge as the same Windows user profile that has Outlook configured.
- Outlook Classic desktop app must be installed and signed in.
- COM calls are handled on STA threads in the implementation.

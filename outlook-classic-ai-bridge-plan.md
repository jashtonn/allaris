# OutlookAgentBridge — Working Setup (No .NET SDK Required on Your PC)

You are right: the previous flow was too easy to break.

## What you want
- Download/run an **unsigned EXE**
- No `dotnet new`
- No SDK install on your machine
- Bridge can read/sort/move/categorize/create folders/draft replies in Outlook Classic

This is supported.

---

## A) Easiest path: download unsigned EXE from GitHub Release

1. In GitHub, run workflow **Release Windows EXE**.
2. It creates a Release asset: `OutlookAgentBridge-win-x64.zip`.
3. Download and unzip.
4. Run:

```powershell
$env:OUTLOOK_BRIDGE_TOKEN = "replace-with-long-random-token"
.\OutlookAgentBridge.exe
```

No .NET SDK required for this path.

---

## B) One-command installer/runner (PowerShell, no SDK)

Use this script from the repo on your Windows machine:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-and-run-latest.ps1 `
  -RepoOwner "YOUR_GITHUB_OWNER" `
  -RepoName "YOUR_REPO" `
  -Token "replace-with-long-random-token"
```

What it does:
- Creates install folder with `-Force` (so existing folder is OK)
- Downloads latest Release ZIP
- Extracts
- Starts `OutlookAgentBridge.exe`

---

## C) API quick test

```powershell
$h = @{ "X-Bridge-Token" = "replace-with-long-random-token" }
Invoke-RestMethod -Headers $h -Uri "http://127.0.0.1:5077/health"
Invoke-RestMethod -Headers $h -Uri "http://127.0.0.1:5077/email/unread?limit=5"
```

Create folder:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" `
  -Uri "http://127.0.0.1:5077/folder/create" `
  -Body '{"name":"AI - Needs Review"}'
```

Move mail:

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

## D) About your previous errors

- `mkdir ... already exists` is harmless; this is now handled with `New-Item -Force`.
- `No .NET SDKs were found` only matters for local compile. It is **not needed** when running prebuilt EXE.
- PowerShell parser errors usually come from copy/paste of broken multiline text. Updated scripts are strict and standalone `.ps1` files.

---

## Included files

- App: `OutlookAgentBridge/OutlookAgentBridge.csproj`, `OutlookAgentBridge/Program.cs`
- Build script (for builders): `scripts/build-windows-exe.ps1`
- No-SDK installer/runner: `scripts/install-and-run-latest.ps1`
- CI artifact: `.github/workflows/build-windows-exe.yml`
- Release ZIP workflow: `.github/workflows/release-windows-exe.yml`

---

## Important runtime notes

- Run as same Windows user profile that has Outlook configured.
- Outlook Classic desktop app must be installed and signed in.
- API is local: `http://127.0.0.1:5077`

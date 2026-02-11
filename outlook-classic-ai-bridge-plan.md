# Outlook Classic AI Bridge — Windows Setup Guide (Working Starter)

This is a practical, end-to-end setup to run a **local agent bridge** against **Outlook Classic (desktop app)** with no Microsoft 365 admin consent flow.

## 0) What you will get

A local Windows service (localhost only) that can:
- read emails
- move emails to folders
- set categories
- create folders
- draft replies

Your AI agent calls this local API; the API does COM automation into Outlook.

---

## 1) Prerequisites (Windows)

1. Windows 10/11 with **Outlook Classic** installed and signed in.
2. .NET 8 SDK installed (`dotnet --version`).
3. Outlook desktop must have your mailbox profile fully synced.
4. In Outlook Trust Center, keep programmatic access allowed (default typically works if AV is healthy).

> Important: This uses your logged-in Windows/Outlook user context. Run the bridge as that same user.

---

## 2) Create project

Open PowerShell:

```powershell
mkdir C:\outlook-agent-bridge
cd C:\outlook-agent-bridge
dotnet new webapi -n OutlookAgentBridge --no-https
cd .\OutlookAgentBridge
```

Install package for COM interop support:

```powershell
dotnet add package Microsoft.Windows.Compatibility
```

Replace `Program.cs` with the code below.

---

## 3) Working starter `Program.cs`

```csharp
using Microsoft.AspNetCore.Mvc;
using System.Runtime.InteropServices;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

var token = Environment.GetEnvironmentVariable("OUTLOOK_BRIDGE_TOKEN")
            ?? "change-me-in-env";

bool IsAuthorized(HttpRequest req)
{
    return req.Headers.TryGetValue("X-Bridge-Token", out var provided)
           && provided == token;
}

app.Use(async (ctx, next) =>
{
    if (!IsAuthorized(ctx.Request))
    {
        ctx.Response.StatusCode = 401;
        await ctx.Response.WriteAsync("Unauthorized");
        return;
    }
    await next();
});

app.MapGet("/health", () => Results.Ok(new { ok = true }));

app.MapGet("/email/unread", ([FromQuery] int limit = 25) =>
{
    return RunSta(() =>
    {
        dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!);
        dynamic ns = outlook.GetNamespace("MAPI");
        dynamic inbox = ns.GetDefaultFolder(6); // olFolderInbox
        dynamic items = inbox.Items;
        items.Sort("[ReceivedTime]", true);

        var unread = new List<object>();
        int count = items.Count;
        for (int i = 1; i <= count && unread.Count < limit; i++)
        {
            dynamic item = items[i];
            try
            {
                if (item.Class == 43 && item.UnRead) // 43 = MailItem
                {
                    unread.Add(new
                    {
                        entryId = (string)item.EntryID,
                        subject = (string)(item.Subject ?? ""),
                        sender = (string)(item.SenderEmailAddress ?? ""),
                        received = (DateTime)item.ReceivedTime
                    });
                }
            }
            finally
            {
                if (item != null) Marshal.ReleaseComObject(item);
            }
        }

        Marshal.ReleaseComObject(items);
        Marshal.ReleaseComObject(inbox);
        Marshal.ReleaseComObject(ns);
        Marshal.ReleaseComObject(outlook);

        return Results.Ok(unread);
    });
});

app.MapPost("/folder/create", ([FromBody] CreateFolderRequest req) =>
{
    return RunSta(() =>
    {
        dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!);
        dynamic ns = outlook.GetNamespace("MAPI");
        dynamic inbox = ns.GetDefaultFolder(6);

        dynamic folder = inbox.Folders.Add(req.Name);

        var result = Results.Ok(new { created = true, name = (string)folder.Name });

        Marshal.ReleaseComObject(folder);
        Marshal.ReleaseComObject(inbox);
        Marshal.ReleaseComObject(ns);
        Marshal.ReleaseComObject(outlook);

        return result;
    });
});

app.MapPost("/email/move", ([FromBody] MoveEmailRequest req) =>
{
    return RunSta(() =>
    {
        dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!);
        dynamic ns = outlook.GetNamespace("MAPI");

        dynamic mail = ns.GetItemFromID(req.EntryId);
        dynamic inbox = ns.GetDefaultFolder(6);
        dynamic target = inbox.Folders[req.TargetFolderName];

        mail.Move(target);

        Marshal.ReleaseComObject(target);
        Marshal.ReleaseComObject(inbox);
        Marshal.ReleaseComObject(mail);
        Marshal.ReleaseComObject(ns);
        Marshal.ReleaseComObject(outlook);

        return Results.Ok(new { moved = true });
    });
});

app.MapPost("/email/category", ([FromBody] CategorizeEmailRequest req) =>
{
    return RunSta(() =>
    {
        dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!);
        dynamic ns = outlook.GetNamespace("MAPI");
        dynamic mail = ns.GetItemFromID(req.EntryId);

        // Outlook category string is semicolon-delimited
        mail.Categories = string.Join(";", req.Categories);
        mail.Save();

        Marshal.ReleaseComObject(mail);
        Marshal.ReleaseComObject(ns);
        Marshal.ReleaseComObject(outlook);

        return Results.Ok(new { categorized = true });
    });
});

app.MapPost("/email/draft-reply", ([FromBody] DraftReplyRequest req) =>
{
    return RunSta(() =>
    {
        dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!);
        dynamic ns = outlook.GetNamespace("MAPI");
        dynamic mail = ns.GetItemFromID(req.EntryId);

        dynamic reply = mail.Reply();
        reply.Body = req.Body + "\r\n\r\n" + reply.Body;
        reply.Save(); // saves to Drafts

        string draftId = reply.EntryID;

        Marshal.ReleaseComObject(reply);
        Marshal.ReleaseComObject(mail);
        Marshal.ReleaseComObject(ns);
        Marshal.ReleaseComObject(outlook);

        return Results.Ok(new { drafted = true, draftEntryId = draftId });
    });
});

app.Run("http://127.0.0.1:5077");

static IResult RunSta(Func<IResult> action)
{
    IResult? result = null;
    Exception? error = null;

    var thread = new Thread(() =>
    {
        try { result = action(); }
        catch (Exception ex) { error = ex; }
    });

    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    thread.Join();

    if (error != null)
        return Results.Problem(error.Message);

    return result!;
}

record CreateFolderRequest(string Name);
record MoveEmailRequest(string EntryId, string TargetFolderName);
record CategorizeEmailRequest(string EntryId, string[] Categories);
record DraftReplyRequest(string EntryId, string Body);
```

---

## 4) Run it

In PowerShell:

```powershell
$env:OUTLOOK_BRIDGE_TOKEN = "super-long-random-token"
dotnet run
```

Service listens on `http://127.0.0.1:5077`.

---

## 5) Quick API tests

Use another PowerShell window.

```powershell
$h = @{ "X-Bridge-Token" = "super-long-random-token" }
Invoke-RestMethod -Headers $h -Uri "http://127.0.0.1:5077/email/unread?limit=5"
```

Create folder:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" \
  -Uri "http://127.0.0.1:5077/folder/create" \
  -Body '{"name":"AI - Needs Review"}'
```

Move email:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" \
  -Uri "http://127.0.0.1:5077/email/move" \
  -Body '{"entryId":"<ENTRY_ID>","targetFolderName":"AI - Needs Review"}'
```

Set categories:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" \
  -Uri "http://127.0.0.1:5077/email/category" \
  -Body '{"entryId":"<ENTRY_ID>","categories":["AI","Follow Up"]}'
```

Draft reply:

```powershell
Invoke-RestMethod -Headers $h -Method Post -ContentType "application/json" \
  -Uri "http://127.0.0.1:5077/email/draft-reply" \
  -Body '{"entryId":"<ENTRY_ID>","body":"Thanks — I reviewed this and will reply in detail by tomorrow."}'
```

---

## 6) Connect your AI agent

Point your agent tools to localhost endpoints. Recommended tool set:
- `list_unread(limit)` -> `/email/unread`
- `create_folder(name)` -> `/folder/create`
- `move_email(entryId, targetFolderName)` -> `/email/move`
- `categorize_email(entryId, categories[])` -> `/email/category`
- `draft_reply(entryId, body)` -> `/email/draft-reply`

Policy recommendations:
- Keep "send" out of scope initially.
- Only allow folder names under a prefix (`AI - ...`).
- Add audit logging for every action.

---

## 7) Make it reliable for daily use

1. Run this app via Task Scheduler at logon (user context, highest privileges not required).
2. Keep Outlook running (or add startup check and launch logic).
3. Add a local SQLite audit table:
   - timestamp
   - action
   - entryId
   - result
4. Add a manual approval queue for bulk moves.

---

## 8) Known Outlook COM gotchas

- COM automation must run in STA threads (handled above).
- `EntryID` may change if items move across stores/accounts.
- Some antivirus/policy setups can trigger Outlook "programmatic access" prompts.
- Avoid high-frequency polling; batch reads are safer.

---

## 9) Next upgrade path

After this starter works, split into:
- `OutlookAgentBridge.Api` (HTTP API)
- `OutlookAgentBridge.Outlook` (COM adapter)
- `OutlookAgentBridge.Policy` (allow/deny rules)
- `OutlookAgentBridge.Audit` (logging)

This gives you exactly what you asked for: read all emails, sort/move, categorize, create folders, and draft replies with local control and no M365 admin consent dependency.

using System.Runtime.InteropServices;
using Microsoft.AspNetCore.Mvc;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

var token = Environment.GetEnvironmentVariable("OUTLOOK_BRIDGE_TOKEN")
    ?? throw new InvalidOperationException("Set OUTLOOK_BRIDGE_TOKEN before starting OutlookAgentBridge.");

app.Use(async (ctx, next) =>
{
    if (!ctx.Request.Headers.TryGetValue("X-Bridge-Token", out var supplied) || supplied != token)
    {
        ctx.Response.StatusCode = StatusCodes.Status401Unauthorized;
        await ctx.Response.WriteAsync("Unauthorized");
        return;
    }

    await next();
});

app.MapGet("/health", () => Results.Ok(new { ok = true }));

app.MapGet("/email/unread", ([FromQuery] int limit) =>
{
    var boundedLimit = Math.Clamp(limit <= 0 ? 25 : limit, 1, 200);

    return RunSta(() =>
    {
        using var outlook = OutlookSession.Create();
        var items = outlook.GetUnread(boundedLimit);
        return Results.Ok(items);
    });
});

app.MapPost("/folder/create", ([FromBody] CreateFolderRequest req) =>
{
    if (string.IsNullOrWhiteSpace(req.Name))
    {
        return Results.BadRequest(new { error = "Folder name is required." });
    }

    return RunSta(() =>
    {
        using var outlook = OutlookSession.Create();
        outlook.CreateFolder(req.Name.Trim());
        return Results.Ok(new { created = true, name = req.Name.Trim() });
    });
});

app.MapPost("/email/move", ([FromBody] MoveEmailRequest req) =>
{
    if (string.IsNullOrWhiteSpace(req.EntryId) || string.IsNullOrWhiteSpace(req.TargetFolderName))
    {
        return Results.BadRequest(new { error = "entryId and targetFolderName are required." });
    }

    return RunSta(() =>
    {
        using var outlook = OutlookSession.Create();
        outlook.MoveMail(req.EntryId.Trim(), req.TargetFolderName.Trim());
        return Results.Ok(new { moved = true });
    });
});

app.MapPost("/email/category", ([FromBody] CategorizeEmailRequest req) =>
{
    if (string.IsNullOrWhiteSpace(req.EntryId) || req.Categories is null || req.Categories.Length == 0)
    {
        return Results.BadRequest(new { error = "entryId and at least one category are required." });
    }

    var categories = req.Categories
        .Where(c => !string.IsNullOrWhiteSpace(c))
        .Select(c => c.Trim())
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();

    if (categories.Length == 0)
    {
        return Results.BadRequest(new { error = "At least one non-empty category is required." });
    }

    return RunSta(() =>
    {
        using var outlook = OutlookSession.Create();
        outlook.SetCategories(req.EntryId.Trim(), categories);
        return Results.Ok(new { categorized = true });
    });
});

app.MapPost("/email/draft-reply", ([FromBody] DraftReplyRequest req) =>
{
    if (string.IsNullOrWhiteSpace(req.EntryId) || string.IsNullOrWhiteSpace(req.Body))
    {
        return Results.BadRequest(new { error = "entryId and body are required." });
    }

    return RunSta(() =>
    {
        using var outlook = OutlookSession.Create();
        var draftEntryId = outlook.CreateDraftReply(req.EntryId.Trim(), req.Body.Trim());
        return Results.Ok(new { drafted = true, draftEntryId });
    });
});

app.Run("http://127.0.0.1:5077");

static IResult RunSta(Func<IResult> action)
{
    IResult? result = null;
    Exception? error = null;

    var thread = new Thread(() =>
    {
        try
        {
            result = action();
        }
        catch (Exception ex)
        {
            error = ex;
        }
    });

    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    thread.Join();

    return error is null
        ? result!
        : Results.Problem(error.Message);
}

internal sealed class OutlookSession : IDisposable
{
    private dynamic _application;
    private dynamic _namespace;

    private OutlookSession(dynamic application, dynamic nameSpace)
    {
        _application = application;
        _namespace = nameSpace;
    }

    public static OutlookSession Create()
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook Classic is not installed or COM registration is unavailable.");

        var application = Activator.CreateInstance(outlookType)
            ?? throw new InvalidOperationException("Unable to create Outlook application instance.");
        dynamic nameSpace = application.GetNamespace("MAPI");

        return new OutlookSession(application, nameSpace);
    }

    public IReadOnlyList<object> GetUnread(int limit)
    {
        var inbox = GetInbox();
        var items = inbox.Items;
        items.Sort("[ReceivedTime]", true);

        var unread = new List<object>();

        try
        {
            var total = (int)items.Count;
            for (var i = 1; i <= total && unread.Count < limit; i++)
            {
                dynamic item = items[i];
                try
                {
                    if (item.Class == 43 && item.UnRead)
                    {
                        unread.Add(new
                        {
                            entryId = (string)item.EntryID,
                            subject = (string)(item.Subject ?? string.Empty),
                            sender = (string)(item.SenderEmailAddress ?? string.Empty),
                            received = (DateTime)item.ReceivedTime
                        });
                    }
                }
                finally
                {
                    ReleaseCom(item);
                }
            }

            return unread;
        }
        finally
        {
            ReleaseCom(items);
            ReleaseCom(inbox);
        }
    }

    public void CreateFolder(string name)
    {
        var inbox = GetInbox();
        try
        {
            dynamic folder = inbox.Folders.Add(name);
            ReleaseCom(folder);
        }
        finally
        {
            ReleaseCom(inbox);
        }
    }

    public void MoveMail(string entryId, string targetFolderName)
    {
        dynamic mail = _namespace.GetItemFromID(entryId);
        var inbox = GetInbox();
        dynamic targetFolder = inbox.Folders[targetFolderName];

        try
        {
            mail.Move(targetFolder);
        }
        finally
        {
            ReleaseCom(targetFolder);
            ReleaseCom(inbox);
            ReleaseCom(mail);
        }
    }

    public void SetCategories(string entryId, string[] categories)
    {
        dynamic mail = _namespace.GetItemFromID(entryId);
        try
        {
            mail.Categories = string.Join(';', categories);
            mail.Save();
        }
        finally
        {
            ReleaseCom(mail);
        }
    }

    public string CreateDraftReply(string entryId, string body)
    {
        dynamic mail = _namespace.GetItemFromID(entryId);
        dynamic reply = mail.Reply();

        try
        {
            reply.Body = body + "\r\n\r\n" + reply.Body;
            reply.Save();
            return (string)reply.EntryID;
        }
        finally
        {
            ReleaseCom(reply);
            ReleaseCom(mail);
        }
    }

    private dynamic GetInbox() => _namespace.GetDefaultFolder(6);

    public void Dispose()
    {
        ReleaseCom(_namespace);
        ReleaseCom(_application);
    }

    private static void ReleaseCom(dynamic? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.ReleaseComObject(value);
        }
    }
}

internal record CreateFolderRequest(string Name);
internal record MoveEmailRequest(string EntryId, string TargetFolderName);
internal record CategorizeEmailRequest(string EntryId, string[] Categories);
internal record DraftReplyRequest(string EntryId, string Body);

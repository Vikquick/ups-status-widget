using System;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace UpsStatusWidget;

sealed class LocalStatusServer : IDisposable
{
    readonly HttpListener _listener = new();
    readonly Func<UpsSnapshot?> _getSnapshot;
    readonly Func<int, UpsSnapshot[]> _getHistory;
    readonly Func<UpsRawSnapshot?> _getRaw;
    readonly CancellationTokenSource _cts = new();
    Task _loopTask;

    public LocalStatusServer(
        string prefix,
        Func<UpsSnapshot?> getSnapshot,
        Func<int, UpsSnapshot[]> getHistory,
        Func<UpsRawSnapshot?> getRaw)
    {
        _getSnapshot = getSnapshot;
        _getHistory = getHistory;
        _getRaw = getRaw;
        _listener.Prefixes.Add(prefix);
    }

    public void Start()
    {
        _listener.Start();
        _loopTask = Task.Run(() => Loop(_cts.Token));
    }

    async Task Loop(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested) {
            HttpListenerContext ctx;
            try {
                ctx = await _listener.GetContextAsync();
            }
            catch (ObjectDisposedException) { break; }
            catch (HttpListenerException) { break; }
            _ = Task.Run(() => Handle(ctx), ct);
        }
    }

    void Handle(HttpListenerContext ctx)
    {
        try {
            string path = ctx.Request.Url?.AbsolutePath ?? "/";
            if (path.Equals("/health", StringComparison.OrdinalIgnoreCase)) {
                WriteJson(ctx.Response, new { ok = true, ts = DateTime.Now });
                return;
            }

            if (path.Equals("/status", StringComparison.OrdinalIgnoreCase)) {
                var s = _getSnapshot();
                if (!s.HasValue) {
                    ctx.Response.StatusCode = 503;
                    WriteJson(ctx.Response, new { error = "no_data" });
                    return;
                }
                WriteJson(ctx.Response, s.Value);
                return;
            }

            if (path.Equals("/status/history", StringComparison.OrdinalIgnoreCase)) {
                int limit = ParseLimit(ctx.Request.QueryString["limit"], 120);
                var items = _getHistory(limit);
                WriteJson(ctx.Response, new { count = items.Length, items });
                return;
            }

            if (path.Equals("/status/raw", StringComparison.OrdinalIgnoreCase)) {
                var raw = _getRaw();
                if (!raw.HasValue) {
                    ctx.Response.StatusCode = 503;
                    WriteJson(ctx.Response, new { error = "no_raw_data" });
                    return;
                }
                WriteJson(ctx.Response, raw.Value);
                return;
            }

            ctx.Response.StatusCode = 404;
            WriteJson(ctx.Response, new { error = "not_found" });
        }
        catch {
            try { ctx.Response.StatusCode = 500; } catch { }
        }
        finally {
            try { ctx.Response.OutputStream.Close(); } catch { }
        }
    }

    static void WriteJson(HttpListenerResponse resp, object payload)
    {
        byte[] body = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(payload));
        resp.ContentType = "application/json; charset=utf-8";
        resp.ContentLength64 = body.Length;
        resp.OutputStream.Write(body, 0, body.Length);
    }

    static int ParseLimit(string raw, int fallback)
    {
        if (int.TryParse(raw, out int v)) return Math.Clamp(v, 1, 1000);
        return fallback;
    }

    public void Dispose()
    {
        _cts.Cancel();
        try { _listener.Stop(); } catch { }
        try { _listener.Close(); } catch { }
        try { _loopTask?.Wait(500); } catch { }
        _cts.Dispose();
    }
}


using System;
using System.Collections.Generic;
using System.Text.Json;

namespace UpsStatusWidget;

readonly record struct ReleaseAsset(
    string Name,
    string DownloadUrl,
    string Sha256Hex
);

readonly record struct ReleaseInfo(
    string Tag,
    Version Version,
    string HtmlUrl,
    ReleaseAsset[] Assets
);

static class UpdateCore
{
    public static bool TryParseTagVersion(string tag, out Version version)
    {
        version = new Version(0, 0);
        if (string.IsNullOrWhiteSpace(tag)) return false;

        string s = tag.Trim();
        if (s.StartsWith("v", StringComparison.OrdinalIgnoreCase)) s = s.Substring(1);
        int cut = s.IndexOfAny(new[] { '-', '+' });
        if (cut >= 0) s = s.Substring(0, cut);

        string[] parts = s.Split('.', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2) return false;

        int[] nums = new int[4];
        int used = Math.Min(parts.Length, 4);
        for (int i = 0; i < used; i++) {
            if (!int.TryParse(parts[i], out nums[i])) return false;
        }

        version = used switch {
            2 => new Version(nums[0], nums[1]),
            3 => new Version(nums[0], nums[1], nums[2]),
            _ => new Version(nums[0], nums[1], nums[2], nums[3])
        };
        return true;
    }

    public static bool TryParseReleaseJson(string json, out ReleaseInfo info)
    {
        info = default;
        if (string.IsNullOrWhiteSpace(json)) return false;

        try {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (!root.TryGetProperty("tag_name", out var tagEl)) return false;
            string tag = tagEl.GetString() ?? string.Empty;
            if (!TryParseTagVersion(tag, out var ver)) return false;

            string html = string.Empty;
            if (root.TryGetProperty("html_url", out var htmlEl)) html = htmlEl.GetString() ?? string.Empty;

            var assets = new List<ReleaseAsset>();
            if (root.TryGetProperty("assets", out var assetsEl) && assetsEl.ValueKind == JsonValueKind.Array) {
                foreach (var a in assetsEl.EnumerateArray()) {
                    string name = a.TryGetProperty("name", out var n) ? (n.GetString() ?? string.Empty) : string.Empty;
                    string url = a.TryGetProperty("browser_download_url", out var u) ? (u.GetString() ?? string.Empty) : string.Empty;
                    string digest = a.TryGetProperty("digest", out var d) ? (d.GetString() ?? string.Empty) : string.Empty;
                    TryParseDigestSha256(digest, out string shaHex);
                    assets.Add(new ReleaseAsset(name, url, shaHex));
                }
            }

            info = new ReleaseInfo(tag, ver, html, assets.ToArray());
            return true;
        }
        catch {
            return false;
        }
    }

    public static bool TryParseDigestSha256(string digest, out string sha256Hex)
    {
        sha256Hex = string.Empty;
        if (string.IsNullOrWhiteSpace(digest)) return false;
        string s = digest.Trim();
        if (!s.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase)) return false;
        s = s.Substring("sha256:".Length).Trim();
        if (s.Length != 64) return false;
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            bool isHex = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
            if (!isHex) return false;
        }
        sha256Hex = s.ToLowerInvariant();
        return true;
    }

    public static ReleaseAsset? SelectInstallerAsset(ReleaseInfo release, bool is64BitProcess)
    {
        if (release.Assets == null || release.Assets.Length == 0) return null;
        string primaryArch = is64BitProcess ? "x64" : "x86";
        string secondaryArch = is64BitProcess ? "x86" : "x64";

        ReleaseAsset? best = FindInstallerByArch(release.Assets, primaryArch);
        if (best.HasValue) return best;
        best = FindInstallerByArch(release.Assets, secondaryArch);
        if (best.HasValue) return best;
        return FindAnyInstaller(release.Assets);
    }

    static ReleaseAsset? FindInstallerByArch(ReleaseAsset[] assets, string arch)
    {
        foreach (var a in assets) {
            string n = (a.Name ?? string.Empty).ToLowerInvariant();
            if (!n.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) continue;
            if (!n.Contains(arch)) continue;
            if (n.Contains("setup") || n.Contains("installer")) return a;
        }
        return null;
    }

    static ReleaseAsset? FindAnyInstaller(ReleaseAsset[] assets)
    {
        foreach (var a in assets) {
            string n = (a.Name ?? string.Empty).ToLowerInvariant();
            if (!n.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) continue;
            if (n.Contains("setup") || n.Contains("installer")) return a;
        }
        return null;
    }
}

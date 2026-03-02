#nullable disable
using System;
using System.Collections.Generic;

namespace UpsStatusWidget;

readonly record struct HidReportDelta(
    string Kind,
    byte ReportId,
    int Offset,
    byte OldValue,
    byte NewValue
);

sealed class HidReportDeltaTracker
{
    readonly Dictionary<string, byte[]> _lastByKey = new(StringComparer.Ordinal);

    public IReadOnlyList<HidReportDelta> UpdateAndGetDeltas(string kind, byte reportId, byte[] report)
    {
        if (string.IsNullOrWhiteSpace(kind)) throw new ArgumentException("kind is required", nameof(kind));
        if (report == null) throw new ArgumentNullException(nameof(report));

        string key = BuildKey(kind, reportId);
        if (!_lastByKey.TryGetValue(key, out byte[] prev)) {
            _lastByKey[key] = Clone(report);
            return Array.Empty<HidReportDelta>();
        }

        int max = Math.Min(prev.Length, report.Length);
        var deltas = new List<HidReportDelta>();
        for (int i = 0; i < max; i++) {
            if (prev[i] == report[i]) continue;
            deltas.Add(new HidReportDelta(kind, reportId, i, prev[i], report[i]));
        }

        if (report.Length > prev.Length) {
            for (int i = prev.Length; i < report.Length; i++) {
                deltas.Add(new HidReportDelta(kind, reportId, i, 0, report[i]));
            }
        }

        _lastByKey[key] = Clone(report);
        return deltas;
    }

    static string BuildKey(string kind, byte reportId) => string.Concat(kind, ":", reportId.ToString());

    static byte[] Clone(byte[] report)
    {
        var clone = new byte[report.Length];
        Buffer.BlockCopy(report, 0, clone, 0, report.Length);
        return clone;
    }
}


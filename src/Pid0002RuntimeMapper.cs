#nullable disable
using System;

namespace UpsStatusWidget;

static class Pid0002RuntimeMapper
{
    public static double? TryMapMinutes(ushort? rid12u16, ushort? rid35u16)
    {
        ushort? raw = rid12u16 ?? rid35u16;
        if (!raw.HasValue) return null;
        if (raw.Value < 60 || raw.Value > 20000) return null;

        double minutes = Math.Round(raw.Value / 60.0);
        if (minutes < 1 || minutes > 240) return null;
        return minutes;
    }
}


#nullable disable
using System;

namespace UpsStatusWidget;

static class HidCorrelationFormatter
{
    public static string Format(
        double? load,
        double? runtime,
        double? freq,
        ushort? rid12u16,
        ushort? rid35u16,
        byte? rid49b1,
        byte? rid22b1)
    {
        return "UPS HID CORR " +
               "load=" + Fmt(load) +
               " runtime=" + Fmt(runtime) +
               " freq=" + Fmt(freq) +
               " rid12_u16=" + Fmt(rid12u16) +
               " rid35_u16=" + Fmt(rid35u16) +
               " rid49_b1=" + Fmt(rid49b1) +
               " rid22_b1=" + Fmt(rid22b1);
    }

    static string Fmt(double? value)
    {
        if (!value.HasValue) return "na";
        return Math.Round(value.Value, 2).ToString("0.##");
    }

    static string Fmt(ushort? value) => value.HasValue ? value.Value.ToString() : "na";

    static string Fmt(byte? value) => value.HasValue ? value.Value.ToString() : "na";
}


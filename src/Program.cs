using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;

namespace UpsStatusWidget;

static class Log
{
    const long MaxLogBytes = 1 * 1024 * 1024;
    const int MaxBackups = 3;
    const string SettingsRegPath = @"Software\UPS-Status-Widget";
    const string SettingsLogEnabledName = "LogEnabled";
    static bool _enabled = ReadEnabledFlag();
    static readonly string _dir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "UpsStatusWidget",
        "logs");
    static readonly string _path = Path.Combine(_dir, "ups-status-widget.log");
    static readonly object _lock = new();

    public static bool Enabled { get { lock (_lock) return _enabled; } }
    public static string LogFilePath => _path;
    public static string LogDirectory => _dir;

    public static void W(string msg)
    {
        if (!Enabled) return;
        string line = $"{DateTime.Now:HH:mm:ss.fff} {msg}";
        lock (_lock) {
            try {
                Directory.CreateDirectory(_dir);
                RotateIfNeeded(System.Text.Encoding.UTF8.GetByteCount(line + Environment.NewLine));
                File.AppendAllText(_path, line + Environment.NewLine);
            }
            catch {
                // Logging must never break app flow.
            }
        }
    }

    static bool ReadEnabledFlag()
    {
        string raw = Environment.GetEnvironmentVariable("UPS_HID_LOG");
        if (!string.IsNullOrWhiteSpace(raw)) return ParseBool(raw);
        try {
            using var k = Registry.CurrentUser.OpenSubKey(SettingsRegPath);
            if (k == null) return false;
            object v = k.GetValue(SettingsLogEnabledName);
            if (v == null) return false;
            return v switch {
                int i => i != 0,
                string s => ParseBool(s),
                _ => false
            };
        }
        catch {
            return false;
        }
    }

    public static void SetEnabled(bool enabled)
    {
        lock (_lock) _enabled = enabled;
        if (enabled) W("LoggingEnabled=True");
    }

    static bool ParseBool(string raw)
    {
        raw = raw.Trim().ToLowerInvariant();
        return raw == "1" || raw == "true" || raw == "yes" || raw == "on";
    }

    static void RotateIfNeeded(int incomingBytes)
    {
        if (!File.Exists(_path)) return;
        var fi = new FileInfo(_path);
        if (fi.Length + incomingBytes <= MaxLogBytes) return;

        for (int i = MaxBackups; i >= 1; i--) {
            string src = $"{_path}.{i}";
            string dst = $"{_path}.{i + 1}";
            if (i == MaxBackups && File.Exists(src)) File.Delete(src);
            if (File.Exists(src)) File.Move(src, dst);
        }
        File.Move(_path, $"{_path}.1");
    }
}

static class Win32
{
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int w, int ht, uint f);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", SetLastError = true)] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll", SetLastError = true)] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [StructLayout(LayoutKind.Sequential)] public struct RECT  { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] public struct WINDOWPOS {
        public IntPtr hwnd, hwndInsertAfter;
        public int x, y, cx, cy;
        public uint flags;
    }

    public static readonly IntPtr HWND_TOP     = new(0);
    public static readonly IntPtr HWND_TOPMOST = new(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new(-2);
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const int  WM_WINDOWPOSCHANGING = 0x0046;
    public const int  WM_ACTIVATE          = 0x0006;
    public const int  WM_ACTIVATEAPP       = 0x001C;
    public const int  WM_MOUSEACTIVATE     = 0x0021;
    public const int  WM_HOTKEY            = 0x0312;
    public const int  MA_NOACTIVATE        = 3;
    public const uint MOD_CONTROL          = 0x0002;
    public const uint MOD_SHIFT            = 0x0004;

    public static void ShowWidgetWindow(IntPtr hwnd, int x, int y, int w, int h)
    {
        // Bring window to visible z-order without stealing focus.
        SetWindowPos(hwnd, HWND_TOPMOST, x, y, w, h, SWP_NOACTIVATE | SWP_SHOWWINDOW);
        SetWindowPos(hwnd, HWND_NOTOPMOST, x, y, w, h, SWP_NOACTIVATE | SWP_SHOWWINDOW);
        ShowWindow(hwnd, 5);
        GetWindowRect(hwnd, out RECT r);
        Log.W($"ShowWidgetWindow rect={r.L},{r.T}->{r.R},{r.B} vis={IsWindowVisible(hwnd)}");
    }
}

readonly record struct UpdateInfo(
    string Tag,
    Version Version,
    string Url,
    ReleaseAsset? InstallerAsset
);

static class UpdateChecker
{
    const string LatestReleaseUrl = "https://api.github.com/repos/Vikquick/ups-status-widget/releases/latest";

    public static async System.Threading.Tasks.Task<UpdateInfo?> GetLatestAsync(bool is64BitProcess)
    {
        try {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd("ups-status-widget");
            string json = await client.GetStringAsync(LatestReleaseUrl).ConfigureAwait(false);
            if (!UpdateCore.TryParseReleaseJson(json, out var rel)) return null;
            var installer = UpdateCore.SelectInstallerAsset(rel, is64BitProcess);
            return new UpdateInfo(rel.Tag, rel.Version, rel.HtmlUrl, installer);
        }
        catch {
            return null;
        }
    }

    public static async System.Threading.Tasks.Task<string> DownloadInstallerAsync(UpdateInfo info)
    {
        if (!info.InstallerAsset.HasValue) return null;
        var asset = info.InstallerAsset.Value;
        if (string.IsNullOrWhiteSpace(asset.DownloadUrl) || !asset.DownloadUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) return null;

        string dir = Path.Combine(Path.GetTempPath(), "UpsStatusWidget", "updates", info.Tag);
        Directory.CreateDirectory(dir);
        string fileName = Path.GetFileName(asset.Name);
        if (string.IsNullOrWhiteSpace(fileName)) fileName = $"UPS-Status-Widget-Setup-{info.Tag}.exe";
        string dst = Path.Combine(dir, fileName);

        using var client = new HttpClient();
        client.DefaultRequestHeaders.UserAgent.ParseAdd("ups-status-widget");
        using var resp = await client.GetAsync(asset.DownloadUrl, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        await using (var input = await resp.Content.ReadAsStreamAsync().ConfigureAwait(false))
        await using (var output = File.Create(dst)) {
            await input.CopyToAsync(output).ConfigureAwait(false);
        }

        if (!string.IsNullOrWhiteSpace(asset.Sha256Hex) && !VerifySha256(dst, asset.Sha256Hex)) {
            try { File.Delete(dst); } catch { }
            return null;
        }

        return dst;
    }

    static bool VerifySha256(string path, string expectedHex)
    {
        try {
            using var sha = SHA256.Create();
            using var fs = File.OpenRead(path);
            byte[] hash = sha.ComputeHash(fs);
            string actual = Convert.ToHexString(hash).ToLowerInvariant();
            return string.Equals(actual, expectedHex.Trim().ToLowerInvariant(), StringComparison.Ordinal);
        }
        catch {
            return false;
        }
    }
}

readonly record struct UpsData(
    double Bcharge,
    double? Timeleft,
    double? Linev,
    double? Loadpct,
    double? Freq,
    string Status,
    double? HealthPct = null,
    uint? CycleCount = null,
    uint? FullCapacity = null,
    uint? DesignCapacity = null,
    string Source = null,
    double? PcLoadReport50 = null,
    double? PcRuntimeReport23Minutes = null,
    uint? PcRuntimeReport23Raw = null,
    double? PcFreqReport0F = null,
    double? PcFreqLive = null,
    double? PcChargeReport22 = null,
    uint? PcStatusReport06Raw = null,
    string PcStatusReport06Bits = null,
    bool? PcStatus06Bit19 = null,
    bool? PcStatus06Bit18 = null,
    bool? PcStatus06Bit17 = null,
    uint? PcStatusReport49Raw = null,
    UpsRawSnapshot? RawSnapshot = null
);

static class UpsReader
{
    public static UpsData? Read()
    {
        var h = HidUpsApi.ReadUpsHid();
        if (h.HasValue) return h.Value;

        var d = ReadBatteryClass();
        if (d.HasValue) return d.Value;
        var w = ReadWindowsPower();
        if (!w.HasValue) return null;
        return w.Value;
    }

    static UpsData? ReadBatteryClass()
    {
        try {
            foreach (var path in BatteryApi.EnumerateBatteryDevicePaths()) {
                using var h = BatteryApi.OpenBattery(path);
                if (h.IsInvalid) continue;

                uint tag = BatteryApi.QueryTag(h);
                if (tag == 0) continue;

                if (!BatteryApi.TryQueryInformation(h, tag, BatteryApi.BatteryQueryInformationLevel.BatteryInformation, out BatteryApi.BATTERY_INFORMATION bi)) {
                    continue;
                }

                if (!BatteryApi.TryQueryStatus(h, tag, out BatteryApi.BATTERY_STATUS bs)) {
                    continue;
                }

                double pct = 0;
                if (bi.FullChargedCapacity > 0) {
                    pct = Math.Clamp(bs.Capacity * 100.0 / bi.FullChargedCapacity, 0, 100);
                } else if (bi.DesignedCapacity > 0) {
                    pct = Math.Clamp(bs.Capacity * 100.0 / bi.DesignedCapacity, 0, 100);
                }

                double? timeleft = null;
                if (BatteryApi.TryQueryEstimatedTime(h, tag, bs.Rate, out uint secs) && secs != 0xFFFFFFFF) {
                    timeleft = Math.Round(secs / 60.0);
                }

                double? linev = bs.Voltage > 0 ? Math.Round(bs.Voltage / 1000.0, 1) : null; // mV -> V
                bool online = (bs.PowerState & BatteryApi.BATTERY_POWER_ON_LINE) != 0;
                bool discharging = (bs.PowerState & BatteryApi.BATTERY_DISCHARGING) != 0;
                string status = online ? "ONLINE" : discharging ? "ONBATT" : "UNKNOWN";

                uint? full = bi.FullChargedCapacity > 0 ? bi.FullChargedCapacity : null;
                uint? design = bi.DesignedCapacity > 0 ? bi.DesignedCapacity : null;
                double? health = (full.HasValue && design.HasValue && design.Value > 0)
                    ? Math.Round(full.Value * 100.0 / design.Value, 1)
                    : null;
                uint? cycle = bi.CycleCount > 0 ? bi.CycleCount : null;

                Log.W($"Source: BatteryClass ({path})");
                return new UpsData(pct, timeleft, linev, null, null, status, health, cycle, full, design, "BatteryClass");
            }
        }
        catch (Exception ex) {
            Log.W("BatteryClass: " + ex.Message);
        }

        return null;
    }

    static UpsData? ReadWindowsPower()
    {
        try {
            var ps = SystemInformation.PowerStatus;
            bool noBattery = (ps.BatteryChargeStatus & BatteryChargeStatus.NoSystemBattery) != 0;

            double bcharge = ps.BatteryLifePercent >= 0
                ? Math.Clamp(ps.BatteryLifePercent * 100.0, 0, 100)
                : 0;
            double? timeleft = ps.BatteryLifeRemaining >= 0
                ? Math.Round(ps.BatteryLifeRemaining / 60.0)
                : null;

            string status = ps.PowerLineStatus switch {
                PowerLineStatus.Online => "ONLINE",
                PowerLineStatus.Offline => "ONBATT",
                _ => "UNKNOWN"
            };

            if (noBattery && ps.PowerLineStatus == PowerLineStatus.Unknown) status = "NO UPS";

            Log.W("Source: Windows Power API");
            return new UpsData(bcharge, timeleft, null, null, null, status, null, null, null, null, "Windows Power API");
        }
        catch (Exception ex) {
            Log.W("Power API: " + ex.Message);
            return null;
        }
    }

}

static class HidUpsApi
{
    const ushort UPS_VID = 0x051D;
    const ushort UPS_PID = 0x0002;
    const uint DIGCF_PRESENT = 0x00000002;
    const uint DIGCF_DEVICEINTERFACE = 0x00000010;

    const uint FILE_SHARE_READ = 0x00000001;
    const uint FILE_SHARE_WRITE = 0x00000002;
    const uint GENERIC_READ = 0x80000000;
    const uint GENERIC_WRITE = 0x40000000;
    const uint OPEN_EXISTING = 3;
    const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

    static readonly System.Collections.Generic.Dictionary<string, byte> _usageReportCache = new();
    static readonly bool _diagReportDump = !string.Equals(Environment.GetEnvironmentVariable("UPS_HID_DIAG_DUMP"), "0", StringComparison.OrdinalIgnoreCase);
    static readonly bool _diagDeltaLog = string.Equals(Environment.GetEnvironmentVariable("UPS_HID_DIAG_DELTAS"), "1", StringComparison.OrdinalIgnoreCase);
    static readonly bool _diagCorrelation = string.Equals(Environment.GetEnvironmentVariable("UPS_HID_DIAG_CORR"), "1", StringComparison.OrdinalIgnoreCase);
    static DateTime _lastTrace = DateTime.MinValue;
    static DateTime _lastScan = DateTime.MinValue;
    static DateTime _lastCorr = DateTime.MinValue;
    static readonly TimeSpan _diagInterval = TimeSpan.FromSeconds(2);
    static readonly HidReportDeltaTracker _deltaScanner = new();
    static readonly object _rawLock = new();
    static UpsRawSnapshot? _lastRaw;
    static readonly ConcurrentDictionary<byte, byte[]> _liveInputReports = new();
    static readonly object _streamLock = new();
    static string _streamPath;
    static Thread _streamThread;

    static readonly (ushort Page, ushort Usage)[] ChargeUsages = { (0x0085, 0x0066), (0x0085, 0x0064), (0xFF85, 0x0066), (0xFF85, 0x0064) };
    static readonly (ushort Page, ushort Usage)[] RuntimeUsages = { (0x0085, 0x0068), (0xFF85, 0x0068) };
    static readonly (ushort Page, ushort Usage)[] LineVoltageUsages = { (0x0084, 0x0030), (0xFF84, 0x0030) };
    static readonly (ushort Page, ushort Usage)[] LoadUsages = {
        (0x0084, 0x0035), (0xFF84, 0x0035), // PercentLoad
        (0x0084, 0x0031), (0xFF84, 0x0031)  // Current (fallback for some vendor maps)
    };
    static readonly (ushort Page, ushort Usage)[] FreqUsages = {
        (0x0084, 0x0032), (0xFF84, 0x0032), // Frequency
        (0x0084, 0x0036), (0xFF84, 0x0036)  // Power (fallback for some vendor maps)
    };
    static readonly (ushort Page, ushort Usage)[] AcPresentUsages = { (0x0085, 0x00D0), (0xFF85, 0x00D0) };
    static readonly (ushort Page, ushort Usage)[] DischargingUsages = { (0x0085, 0x0045), (0xFF85, 0x0045) };

    public static UpsData? ReadUpsHid()
    {
        try {
            var paths = EnumerateHidPaths();
            foreach (string path in paths) {
                string lp = path.ToLowerInvariant();
                if (!lp.Contains("vid_051d")) continue;
                if (!lp.Contains("pid_0002")) continue;

                SafeFileHandle h = OpenHid(path, GENERIC_READ | GENERIC_WRITE);
                if (h.IsInvalid) {
                    h.Dispose();
                    h = OpenHid(path, 0);
                }
                using (h) {
                    if (h.IsInvalid) {
                        Log.W($"UPS HID: CreateFile failed err={Marshal.GetLastWin32Error()}");
                        continue;
                    }
                    var attrs = new HIDD_ATTRIBUTES { Size = Marshal.SizeOf<HIDD_ATTRIBUTES>() };
                    if (!HidD_GetAttributes(h, ref attrs)) {
                        continue;
                    }
                    if (attrs.VendorID != UPS_VID || attrs.ProductID != UPS_PID) continue;
                    if (!HidD_GetPreparsedData(h, out IntPtr ppd)) {
                        continue;
                    }

                    try {
                        int st = HidP_GetCaps(ppd, out HIDP_CAPS caps);
                        if (!IsHidpSuccess(st)) {
                            continue;
                        }
                        EnsureLiveReader(path, caps.InputReportByteLength);

                        if (!TryReadValue(h, ppd, caps, ChargeUsages, out int rawCharge)) {
                            continue;
                        }

                        double charge = ClampPercent(rawCharge);
                        double? runtime = null;
                        double? linev = null;
                        double? load = null;
                        double? freq = null;
                        byte? rid50B1 = TryReadFeatureByteAt(h, caps, 0x50, 1);
                        byte? rid22FeatureB1 = TryReadFeatureByteAt(h, caps, 0x22, 1);
                        byte? rid0FfeatureB1 = TryReadFeatureByteAt(h, caps, 0x0F, 1);
                        byte? rid49FeatureB1 = TryReadFeatureByteAt(h, caps, 0x49, 1);
                        double? freqLive = TryReadDeviceFrequencyLive(h, caps);
                        double? linevPc = TryReadDeviceLineVoltage(h, caps);
                        ushort? rid23FeatureU16 = TryReadFeatureU16At(h, caps, 0x23, 1);
                        uint? rid06FeatureU32 = TryReadFeatureU32At(h, caps, 0x06, 1);
                        double? runtimePc = rid23FeatureU16.HasValue && rid23FeatureU16.Value >= 60
                            ? Math.Max(1, Math.Round(rid23FeatureU16.Value / 60.0)) : null;
                        runtime = runtimePc;
                        linev = linevPc;
                        double? loadPc = (rid50B1.HasValue && rid50B1.Value <= 100) ? rid50B1.Value : null;
                        load = loadPc;
                        freq = freqLive;
                        if (_diagReportDump && (!load.HasValue || !freq.HasValue) && (DateTime.Now - _lastTrace).TotalSeconds > 20) {
                            _lastTrace = DateTime.Now;
                            DumpRawReports(h, caps);
                        }
                        MaybeLogReportDeltas(h, caps);
                        MaybeLogCorrelation(h, caps, load, runtime, freq, rid22FeatureB1);

                        bool acPresent = TryReadValue(h, ppd, caps, AcPresentUsages, out int rawAc) && rawAc != 0;
                        bool discharging = TryReadValue(h, ppd, caps, DischargingUsages, out int rawDis) && rawDis != 0;
                        string status = acPresent ? "ONLINE" : discharging ? "ONBATT" : "UNKNOWN";
                        UpdateRawSnapshot(h, caps, charge, runtime, linev, load, freq);
                        var rawSnap = GetLastRawSnapshot();

                        Log.W("Source: UPS HID");
                        return new UpsData(
                            charge, runtime, linev, load, freq, status, null, null, null, null, "UPS HID",
                            rid50B1,
                            runtimePc,
                            rid23FeatureU16,
                            rid0FfeatureB1,
                            freqLive,
                            rid22FeatureB1,
                            rid06FeatureU32,
                            FormatSetBits(rid06FeatureU32),
                            GetBit(rid06FeatureU32, 19),
                            GetBit(rid06FeatureU32, 18),
                            GetBit(rid06FeatureU32, 17),
                            rid49FeatureB1,
                            rawSnap);
                    }
                    finally {
                        HidD_FreePreparsedData(ppd);
                    }
                }
            }
        }
        catch (Exception ex) {
            Log.W("UPS HID: " + ex.Message);
        }

        return null;
    }

    static bool TryReadValue(SafeFileHandle handle, IntPtr ppd, HIDP_CAPS caps, (ushort Page, ushort Usage)[] candidates, out int value)
    {
        foreach (var c in candidates) {
            if (TryReadValue(handle, ppd, caps, HidPReportType.HidP_Feature, c.Page, c.Usage, out value)) return true;
            if (TryReadValue(handle, ppd, caps, HidPReportType.HidP_Input, c.Page, c.Usage, out value)) return true;
        }
        value = 0;
        return false;
    }

    static bool TryReadValue(SafeFileHandle handle, IntPtr ppd, HIDP_CAPS caps, HidPReportType rt, ushort usagePage, ushort usage, out int value)
    {
        string key = $"{(int)rt}:{usagePage:X4}:{usage:X4}";
        if (_usageReportCache.TryGetValue(key, out byte rid)) {
            if (TryReadByReportId(handle, ppd, caps, rt, usagePage, usage, rid, out value)) return true;
        }

        for (byte reportId = 0; reportId < 64; reportId++) {
            if (TryReadByReportId(handle, ppd, caps, rt, usagePage, usage, reportId, out value)) {
                _usageReportCache[key] = reportId;
                return true;
            }
        }

        value = 0;
        return false;
    }

    static bool TryReadByReportId(SafeFileHandle handle, IntPtr ppd, HIDP_CAPS caps, HidPReportType rt, ushort usagePage, ushort usage, byte reportId, out int value)
    {
        value = 0;
        if (!TryGetReport(handle, caps, rt, reportId, out byte[] report)) return false;

        for (ushort linkCollection = 0; linkCollection < 16; linkCollection++) {
            int st = HidP_GetScaledUsageValue(rt, usagePage, linkCollection, usage, out value, ppd, report, (uint)report.Length);
            if (IsHidpSuccess(st)) return true;

            st = HidP_GetUsageValue(rt, usagePage, linkCollection, usage, out uint uval, ppd, report, (uint)report.Length);
            if (IsHidpSuccess(st)) {
                value = unchecked((int)uval);
                return true;
            }
        }
        return false;
    }

    static bool TryGetReport(SafeFileHandle handle, HIDP_CAPS caps, HidPReportType rt, byte reportId, out byte[] report)
    {
        report = Array.Empty<byte>();
        int len = rt switch {
            HidPReportType.HidP_Input => caps.InputReportByteLength,
            HidPReportType.HidP_Feature => caps.FeatureReportByteLength,
            _ => 0
        };
        if (len <= 0) return false;

        report = new byte[len];
        report[0] = reportId;
        bool ok = rt == HidPReportType.HidP_Feature
            ? HidD_GetFeature(handle, report, report.Length)
            : HidD_GetInputReport(handle, report, report.Length);
        return ok;
    }

    static void DumpRawReports(SafeFileHandle handle, HIDP_CAPS caps)
    {
        try {
            Log.W("UPS HID DIAG DUMP BEGIN");
            DumpReportSet(handle, caps, HidPReportType.HidP_Feature);
            DumpReportSet(handle, caps, HidPReportType.HidP_Input);
            Log.W("UPS HID DIAG DUMP END");
        }
        catch (Exception ex) {
            Log.W("UPS HID DIAG DUMP error: " + ex.Message);
        }
    }

    static void DumpReportSet(SafeFileHandle handle, HIDP_CAPS caps, HidPReportType rt)
    {
        string rtName = rt == HidPReportType.HidP_Feature ? "FEATURE" : "INPUT";
        for (byte rid = 0; rid < 64; rid++) {
            if (!TryGetReport(handle, caps, rt, rid, out byte[] report)) continue;
            if (report.Length <= 1) continue;
            bool nonZero = false;
            for (int i = 1; i < report.Length; i++) {
                if (report[i] != 0) { nonZero = true; break; }
            }
            if (!nonZero) continue;
            Log.W($"UPS HID {rtName} rid={rid} data={BitConverter.ToString(report)}");
        }
    }

    static void MaybeLogReportDeltas(SafeFileHandle handle, HIDP_CAPS caps)
    {
        if (!_diagDeltaLog) return;

        DateTime now = DateTime.UtcNow;
        if ((now - _lastScan) < _diagInterval) return;
        _lastScan = now;

        try {
            ScanReportType(handle, caps, HidPReportType.HidP_Input, "INPUT");
            ScanReportType(handle, caps, HidPReportType.HidP_Feature, "FEATURE");
        }
        catch (Exception ex) {
            Log.W("UPS HID DIAG DELTAS error: " + ex.Message);
        }
    }

    static void ScanReportType(SafeFileHandle handle, HIDP_CAPS caps, HidPReportType reportType, string kind)
    {
        for (byte rid = 0; rid < 64; rid++) {
            if (!TryGetReport(handle, caps, reportType, rid, out byte[] report)) continue;
            if (report.Length <= 1) continue;

            var deltas = _deltaScanner.UpdateAndGetDeltas(kind, rid, report);
            for (int i = 0; i < deltas.Count; i++) {
                var d = deltas[i];
                Log.W($"UPS HID DIAG DELTA {d.Kind} rid={d.ReportId} off={d.Offset} {d.OldValue}->{d.NewValue}");
            }
        }
    }

    public static UpsRawSnapshot? GetLastRawSnapshot()
    {
        lock (_rawLock) return _lastRaw;
    }

    static void UpdateRawSnapshot(SafeFileHandle h, HIDP_CAPS caps, double charge, double? runtime, double? linev, double? load, double? freq)
    {
        try {
            var feature = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var input = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (byte rid = 0; rid < 128; rid++) {
                if (TryGetReport(h, caps, HidPReportType.HidP_Feature, rid, out byte[] f) && HasNonZeroPayload(f)) feature[$"0x{rid:X2}"] = BitConverter.ToString(f);
                if (TryGetReport(h, caps, HidPReportType.HidP_Input, rid, out byte[] i) && HasNonZeroPayload(i)) input[$"0x{rid:X2}"] = BitConverter.ToString(i);
            }

            var u16 = new System.Collections.Generic.Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (TryReadRidWord(h, caps, HidPReportType.HidP_Feature, 0x23, out ushort w23)) u16["feature_0x23"] = w23;
            if (TryReadRidWord(h, caps, HidPReportType.HidP_Feature, 0x35, out ushort w35)) u16["feature_0x35"] = w35;
            if (TryReadRidWord(h, caps, HidPReportType.HidP_Input, 0x12, out ushort w12)) u16["input_0x12"] = w12;

            var metrics = new System.Collections.Generic.Dictionary<string, double>(StringComparer.OrdinalIgnoreCase) {
                ["charge"] = charge,
                ["runtime"] = runtime ?? double.NaN,
                ["linev"] = linev ?? double.NaN,
                ["load"] = load ?? double.NaN,
                ["freq"] = freq ?? double.NaN
            };

            var snap = new UpsRawSnapshot(DateTime.Now, feature, input, u16, metrics);
            lock (_rawLock) _lastRaw = snap;
        }
        catch {
            // best-effort telemetry path, ignore failures
        }
    }

    static bool HasNonZeroPayload(byte[] report)
    {
        if (report == null || report.Length <= 1) return false;
        for (int i = 1; i < report.Length; i++) {
            if (report[i] != 0) return true;
        }
        return false;
    }

    static void MaybeLogCorrelation(
        SafeFileHandle handle,
        HIDP_CAPS caps,
        double? load,
        double? runtime,
        double? freq,
        byte? rid22B1)
    {
        if (!_diagCorrelation) return;

        DateTime now = DateTime.UtcNow;
        if ((now - _lastCorr) < _diagInterval) return;
        _lastCorr = now;

        try {
            ushort? rid12 = TryReadInputU16At(handle, caps, 12, 2);
            ushort? rid35 = TryReadFeatureU16At(handle, caps, 35, 1);
            byte? rid49 = TryReadFeatureByteAt(handle, caps, 49, 1);
            Log.W(HidCorrelationFormatter.Format(load, runtime, freq, rid12, rid35, rid49, rid22B1));
        }
        catch (Exception ex) {
            Log.W("UPS HID CORR error: " + ex.Message);
        }
    }

    static ushort? TryReadInputU16At(SafeFileHandle handle, HIDP_CAPS caps, byte rid, int offset)
    {
        if (!TryGetReport(handle, caps, HidPReportType.HidP_Input, rid, out byte[] report)) return null;
        if (offset < 0 || (offset + 1) >= report.Length) return null;
        return (ushort)(report[offset] | (report[offset + 1] << 8));
    }

    static ushort? TryReadFeatureU16At(SafeFileHandle handle, HIDP_CAPS caps, byte rid, int offset)
    {
        if (!TryGetReport(handle, caps, HidPReportType.HidP_Feature, rid, out byte[] report)) return null;
        if (offset < 0 || (offset + 1) >= report.Length) return null;
        return (ushort)(report[offset] | (report[offset + 1] << 8));
    }

    static byte? TryReadFeatureByteAt(SafeFileHandle handle, HIDP_CAPS caps, byte rid, int offset)
    {
        if (!TryGetReport(handle, caps, HidPReportType.HidP_Feature, rid, out byte[] report)) return null;
        if (offset < 0 || offset >= report.Length) return null;
        return report[offset];
    }

    static bool? GetBit(uint? value, int bit)
    {
        if (!value.HasValue || bit < 0 || bit > 31) return null;
        return ((value.Value >> bit) & 0x1u) == 1u;
    }

    static string FormatSetBits(uint? value)
    {
        if (!value.HasValue) return null;
        if (value.Value == 0) return "none";
        var bits = new System.Collections.Generic.List<string>();
        for (int i = 0; i < 32; i++) {
            if (((value.Value >> i) & 0x1u) == 1u) bits.Add(i.ToString());
        }
        return string.Join(",", bits);
    }

    static uint? TryReadFeatureU32At(SafeFileHandle handle, HIDP_CAPS caps, byte rid, int offset)
    {
        if (!TryGetReport(handle, caps, HidPReportType.HidP_Feature, rid, out byte[] report)) return null;
        if (offset < 0 || (offset + 3) >= report.Length) return null;
        return (uint)(report[offset] | (report[offset + 1] << 8) | (report[offset + 2] << 16) | (report[offset + 3] << 24));
    }

    static double? TryReadDeviceLineVoltage(SafeFileHandle h, HIDP_CAPS caps)
    {
        // Traffic-calibrated source priority for VID_051D PID_0002 profile:
        // - Feature 0x31 byte1 tracks live line voltage.
        // No fallback path by design.
        byte? v31 = TryReadFeatureByteAt(h, caps, 0x31, 1);
        if (v31.HasValue) {
            if (v31.Value >= 80) return v31.Value;
        }

        return null;
    }

    static double? TryReadDeviceFrequencyLive(SafeFileHandle h, HIDP_CAPS caps)
    {
        // Traffic-calibrated for VID_051D PID_0002 profile:
        // Feature 0x31 byte1 moves in small steps while load/runtime change.
        // On this device 212 => 50.00Hz, i.e. (b1 - 12) / 4.
        byte? b31 = TryReadFeatureByteAt(h, caps, 0x31, 1);
        if (b31.HasValue) {
            double hz = (b31.Value - 12.0) / 4.0;
            if (hz >= 45 && hz <= 65) return Math.Round(hz, 2);
        }
        return null;
    }

    static bool TryReadRidByte(SafeFileHandle h, HIDP_CAPS caps, HidPReportType rt, byte rid, out byte value)
    {
        value = 0;
        if (rt == HidPReportType.HidP_Input && _liveInputReports.TryGetValue(rid, out byte[] live) && live.Length >= 2) {
            value = live[1];
            return true;
        }
        if (!TryGetReport(h, caps, rt, rid, out byte[] report)) return false;
        if (report.Length < 2) return false;
        value = report[1];
        return true;
    }

    static bool TryReadRidWord(SafeFileHandle h, HIDP_CAPS caps, HidPReportType rt, byte rid, out ushort value)
    {
        value = 0;
        if (rt == HidPReportType.HidP_Input && _liveInputReports.TryGetValue(rid, out byte[] live) && live.Length >= 3) {
            value = (ushort)(live[1] | (live[2] << 8));
            return true;
        }
        if (!TryGetReport(h, caps, rt, rid, out byte[] report)) return false;
        if (report.Length < 3) return false;
        value = (ushort)(report[1] | (report[2] << 8));
        return true;
    }

    static void EnsureLiveReader(string path, int inputLen)
    {
        if (inputLen <= 0) return;
        lock (_streamLock) {
            if (_streamThread != null && _streamThread.IsAlive && string.Equals(_streamPath, path, StringComparison.OrdinalIgnoreCase)) {
                return;
            }
            _streamPath = path;
            _streamThread = new Thread(() => LiveReaderLoop(path, inputLen)) {
                IsBackground = true,
                Name = "UpsHidLiveReader"
            };
            _streamThread.Start();
            Log.W("UPS HID live reader started");
        }
    }

    static void LiveReaderLoop(string path, int inputLen)
    {
        try {
            using var h = OpenHid(path, GENERIC_READ | GENERIC_WRITE);
            if (h.IsInvalid) {
                Log.W($"UPS HID live reader open failed err={Marshal.GetLastWin32Error()}");
                return;
            }
            var buf = new byte[inputLen];
            while (true) {
                if (!ReadFile(h, buf, buf.Length, out int read, IntPtr.Zero)) {
                    int err = Marshal.GetLastWin32Error();
                    Log.W($"UPS HID live reader stopped err={err}");
                    return;
                }
                if (read <= 0) continue;
                byte rid = buf[0];
                var snap = new byte[read];
                Buffer.BlockCopy(buf, 0, snap, 0, read);
                _liveInputReports[rid] = snap;
            }
        }
        catch (Exception ex) {
            Log.W("UPS HID live reader exception: " + ex.Message);
        }
    }

    static string[] EnumerateHidPaths()
    {
        var list = new System.Collections.Generic.List<string>();
        HidD_GetHidGuid(out Guid hidGuid);
        IntPtr info = SetupDiGetClassDevs(ref hidGuid, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
        if (info == INVALID_HANDLE_VALUE) return list.ToArray();
        try {
            uint idx = 0;
            while (true) {
                var ifData = new SP_DEVICE_INTERFACE_DATA { cbSize = (uint)Marshal.SizeOf<SP_DEVICE_INTERFACE_DATA>() };
                if (!SetupDiEnumDeviceInterfaces(info, IntPtr.Zero, ref hidGuid, idx, ref ifData)) break;

                uint need = 0;
                SetupDiGetDeviceInterfaceDetail(info, ref ifData, IntPtr.Zero, 0, out need, IntPtr.Zero);
                IntPtr buf = Marshal.AllocHGlobal((int)need);
                try {
                    Marshal.WriteInt32(buf, IntPtr.Size == 8 ? 8 : 6);
                    if (SetupDiGetDeviceInterfaceDetail(info, ref ifData, buf, need, out _, IntPtr.Zero)) {
                        IntPtr p = IntPtr.Add(buf, 4);
                        string path = Marshal.PtrToStringAuto(p) ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(path)) list.Add(path);
                    }
                }
                finally { Marshal.FreeHGlobal(buf); }
                idx++;
            }
        }
        finally { SetupDiDestroyDeviceInfoList(info); }
        return list.ToArray();
    }

    static SafeFileHandle OpenHid(string path)
    {
        return OpenHid(path, GENERIC_READ | GENERIC_WRITE);
    }

    static SafeFileHandle OpenHid(string path, uint desiredAccess)
    {
        return CreateFile(path, desiredAccess, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero);
    }

    static bool IsHidpSuccess(int status) => (status & unchecked((int)0xFFFF0000)) == 0x00110000;

    static double ClampPercent(int v)
    {
        double p = v;
        if (p > 1000) p /= 100.0;
        else if (p > 100) p /= 10.0;
        return Math.Clamp(p, 0, 100);
    }

    enum HidPReportType
    {
        HidP_Input = 0,
        HidP_Output = 1,
        HidP_Feature = 2
    }

    [StructLayout(LayoutKind.Sequential)]
    struct HIDD_ATTRIBUTES
    {
        public int Size;
        public ushort VendorID;
        public ushort ProductID;
        public ushort VersionNumber;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct HIDP_CAPS
    {
        public short Usage;
        public short UsagePage;
        public short InputReportByteLength;
        public short OutputReportByteLength;
        public short FeatureReportByteLength;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
        public short[] Reserved;
        public short NumberLinkCollectionNodes;
        public short NumberInputButtonCaps;
        public short NumberInputValueCaps;
        public short NumberInputDataIndices;
        public short NumberOutputButtonCaps;
        public short NumberOutputValueCaps;
        public short NumberOutputDataIndices;
        public short NumberFeatureButtonCaps;
        public short NumberFeatureValueCaps;
        public short NumberFeatureDataIndices;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct SP_DEVICE_INTERFACE_DATA
    {
        public uint cbSize;
        public Guid InterfaceClassGuid;
        public uint Flags;
        public IntPtr Reserved;
    }

    static readonly IntPtr INVALID_HANDLE_VALUE = new(-1);

    [DllImport("hid.dll")]
    static extern void HidD_GetHidGuid(out Guid hidGuid);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_GetAttributes(SafeFileHandle hidDeviceObject, ref HIDD_ATTRIBUTES attributes);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_GetPreparsedData(SafeFileHandle hidDeviceObject, out IntPtr preparsedData);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_FreePreparsedData(IntPtr preparsedData);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_GetFeature(SafeFileHandle hidDeviceObject, byte[] reportBuffer, int reportBufferLength);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_GetInputReport(SafeFileHandle hidDeviceObject, byte[] reportBuffer, int reportBufferLength);

    [DllImport("hid.dll")]
    static extern int HidP_GetCaps(IntPtr preparsedData, out HIDP_CAPS capabilities);

    [DllImport("hid.dll")]
    static extern int HidP_GetUsageValue(HidPReportType reportType, ushort usagePage, ushort linkCollection, ushort usage, out uint usageValue, IntPtr preparsedData, byte[] report, uint reportLength);

    [DllImport("hid.dll")]
    static extern int HidP_GetScaledUsageValue(HidPReportType reportType, ushort usagePage, ushort linkCollection, ushort usage, out int usageValue, IntPtr preparsedData, byte[] report, uint reportLength);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern IntPtr SetupDiGetClassDevs(ref Guid classGuid, IntPtr enumerator, IntPtr hwndParent, uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiEnumDeviceInterfaces(IntPtr deviceInfoSet, IntPtr deviceInfoData, ref Guid interfaceClassGuid, uint memberIndex, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr deviceInfoSet, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData, IntPtr deviceInterfaceDetailData, uint deviceInterfaceDetailDataSize, out uint requiredSize, IntPtr deviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess, uint shareMode, IntPtr securityAttributes, uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool ReadFile(SafeFileHandle file, byte[] buffer, int numberOfBytesToRead, out int numberOfBytesRead, IntPtr overlapped);
}

static class BatteryApi
{
    public const uint BATTERY_POWER_ON_LINE = 0x00000001;
    public const uint BATTERY_DISCHARGING   = 0x00000002;
    const uint DIGCF_PRESENT = 0x00000002;
    const uint DIGCF_DEVICEINTERFACE = 0x00000010;

    const uint FILE_SHARE_READ = 0x00000001;
    const uint FILE_SHARE_WRITE = 0x00000002;
    const uint GENERIC_READ = 0x80000000;
    const uint GENERIC_WRITE = 0x40000000;
    const uint OPEN_EXISTING = 3;
    const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

    const uint IOCTL_BATTERY_QUERY_TAG = 0x294040;
    const uint IOCTL_BATTERY_QUERY_INFORMATION = 0x294044;
    const uint IOCTL_BATTERY_QUERY_STATUS = 0x29404C;

    static readonly Guid GUID_DEVINTERFACE_BATTERY = new("72631e54-78A4-11d0-bcf7-00aa00b7b32a");

    public static string[] EnumerateBatteryDevicePaths()
    {
        var list = new System.Collections.Generic.List<string>();
        Guid g = GUID_DEVINTERFACE_BATTERY;
        IntPtr info = SetupDiGetClassDevs(ref g, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
        if (info == INVALID_HANDLE_VALUE) return list.ToArray();
        try {
            uint idx = 0;
            while (true) {
                var ifData = new SP_DEVICE_INTERFACE_DATA { cbSize = (uint)Marshal.SizeOf<SP_DEVICE_INTERFACE_DATA>() };
                g = GUID_DEVINTERFACE_BATTERY;
                if (!SetupDiEnumDeviceInterfaces(info, IntPtr.Zero, ref g, idx, ref ifData)) break;

                uint need = 0;
                SetupDiGetDeviceInterfaceDetail(info, ref ifData, IntPtr.Zero, 0, out need, IntPtr.Zero);
                IntPtr buf = Marshal.AllocHGlobal((int)need);
                try {
                    // cbSize is 8 on x64, 6 on x86
                    Marshal.WriteInt32(buf, IntPtr.Size == 8 ? 8 : 6);
                    if (SetupDiGetDeviceInterfaceDetail(info, ref ifData, buf, need, out _, IntPtr.Zero)) {
                        IntPtr p = IntPtr.Add(buf, IntPtr.Size == 8 ? 8 : 4);
                        string path = Marshal.PtrToStringUni(p) ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(path)) list.Add(path);
                    }
                }
                finally { Marshal.FreeHGlobal(buf); }

                idx++;
            }
        }
        finally { SetupDiDestroyDeviceInfoList(info); }
        return list.ToArray();
    }

    public static SafeFileHandle OpenBattery(string path)
    {
        return CreateFile(path, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero);
    }

    public static uint QueryTag(SafeFileHandle h)
    {
        uint wait = 0;
        uint outTag = 0;
        if (!DeviceIoControl(h, IOCTL_BATTERY_QUERY_TAG, ref wait, (uint)Marshal.SizeOf<uint>(), out outTag, (uint)Marshal.SizeOf<uint>(), out _, IntPtr.Zero)) {
            return 0;
        }
        return outTag;
    }

    public static bool TryQueryInformation<T>(SafeFileHandle h, uint tag, BatteryQueryInformationLevel level, out T value) where T : struct
    {
        var q = new BATTERY_QUERY_INFORMATION { BatteryTag = tag, InformationLevel = level, AtRate = 0 };
        uint outSize = (uint)Marshal.SizeOf<T>();
        IntPtr outBuf = Marshal.AllocHGlobal((int)outSize);
        try {
            bool ok = DeviceIoControl(h, IOCTL_BATTERY_QUERY_INFORMATION, ref q, (uint)Marshal.SizeOf<BATTERY_QUERY_INFORMATION>(), outBuf, outSize, out _, IntPtr.Zero);
            if (!ok) {
                value = default;
                return false;
            }
            value = Marshal.PtrToStructure<T>(outBuf);
            return true;
        }
        finally { Marshal.FreeHGlobal(outBuf); }
    }

    public static bool TryQueryEstimatedTime(SafeFileHandle h, uint tag, int atRate, out uint seconds)
    {
        var q = new BATTERY_QUERY_INFORMATION { BatteryTag = tag, InformationLevel = BatteryQueryInformationLevel.BatteryEstimatedTime, AtRate = atRate };
        seconds = 0;
        return DeviceIoControl(h, IOCTL_BATTERY_QUERY_INFORMATION, ref q, (uint)Marshal.SizeOf<BATTERY_QUERY_INFORMATION>(), out seconds, (uint)Marshal.SizeOf<uint>(), out _, IntPtr.Zero);
    }

    public static bool TryQueryStatus(SafeFileHandle h, uint tag, out BATTERY_STATUS status)
    {
        var wait = new BATTERY_WAIT_STATUS { BatteryTag = tag, Timeout = 0, PowerState = 0, LowCapacity = 0, HighCapacity = 0 };
        status = default;
        return DeviceIoControl(h, IOCTL_BATTERY_QUERY_STATUS, ref wait, (uint)Marshal.SizeOf<BATTERY_WAIT_STATUS>(), out status, (uint)Marshal.SizeOf<BATTERY_STATUS>(), out _, IntPtr.Zero);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SP_DEVICE_INTERFACE_DATA
    {
        public uint cbSize;
        public Guid InterfaceClassGuid;
        public uint Flags;
        public IntPtr Reserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BATTERY_QUERY_INFORMATION
    {
        public uint BatteryTag;
        public BatteryQueryInformationLevel InformationLevel;
        public int AtRate;
    }

    public enum BatteryQueryInformationLevel : int
    {
        BatteryInformation = 0,
        BatteryGranularityInformation = 1,
        BatteryTemperature = 2,
        BatteryEstimatedTime = 3,
        BatteryManufactureDate = 5,
        BatteryManufactureName = 6,
        BatteryUniqueID = 7
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BATTERY_INFORMATION
    {
        public uint Capabilities;
        public byte Technology;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
        public byte[] Reserved;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
        public byte[] Chemistry;
        public uint DesignedCapacity;
        public uint FullChargedCapacity;
        public uint DefaultAlert1;
        public uint DefaultAlert2;
        public uint CriticalBias;
        public uint CycleCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BATTERY_WAIT_STATUS
    {
        public uint BatteryTag;
        public uint Timeout;
        public uint PowerState;
        public uint LowCapacity;
        public uint HighCapacity;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BATTERY_STATUS
    {
        public uint PowerState;
        public uint Capacity;
        public uint Voltage;
        public int Rate;
    }

    static readonly IntPtr INVALID_HANDLE_VALUE = new(-1);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern IntPtr SetupDiGetClassDevs(ref Guid classGuid, IntPtr enumerator, IntPtr hwndParent, uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiEnumDeviceInterfaces(IntPtr deviceInfoSet, IntPtr deviceInfoData, ref Guid interfaceClassGuid, uint memberIndex, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr deviceInfoSet, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData, IntPtr deviceInterfaceDetailData, uint deviceInterfaceDetailDataSize, out uint requiredSize, IntPtr deviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess, uint shareMode, IntPtr securityAttributes, uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool DeviceIoControl(SafeFileHandle device, uint ioControlCode, ref uint inBuffer, uint nInBufferSize, out uint outBuffer, uint nOutBufferSize, out uint bytesReturned, IntPtr overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool DeviceIoControl(SafeFileHandle device, uint ioControlCode, ref BATTERY_QUERY_INFORMATION inBuffer, uint nInBufferSize, IntPtr outBuffer, uint nOutBufferSize, out uint bytesReturned, IntPtr overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool DeviceIoControl(SafeFileHandle device, uint ioControlCode, ref BATTERY_QUERY_INFORMATION inBuffer, uint nInBufferSize, out uint outBuffer, uint nOutBufferSize, out uint bytesReturned, IntPtr overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool DeviceIoControl(SafeFileHandle device, uint ioControlCode, ref BATTERY_WAIT_STATUS inBuffer, uint nInBufferSize, out BATTERY_STATUS outBuffer, uint nOutBufferSize, out uint bytesReturned, IntPtr overlapped);
}

class UpsWidget : Form
{
    const int W = 270, HBase = 186, HDebug = 340;
    const int HotkeyDebugId = 0xA11;

    static readonly Color cBg     = Color.FromArgb(255, 10,  14,  20);
    static readonly Color cBorder = Color.FromArgb(255, 30,  45,  60);
    static readonly Color cAccent = Color.FromArgb(255, 0,   200, 255);
    static readonly Color cGreen  = Color.FromArgb(255, 50,  255, 80);
    static readonly Color cWarn   = Color.FromArgb(255, 255, 180, 0);
    static readonly Color cDanger = Color.FromArgb(255, 255, 60,  60);
    static readonly Color cDim    = Color.FromArgb(255, 80,  100, 120);
    static readonly Color cText   = Color.FromArgb(255, 200, 220, 240);
    static readonly Color cDimBar = Color.FromArgb(255, 22,  32,  44);
    static readonly Color cBrand  = Color.FromArgb(200, 0,   200, 255);
    static readonly Color cDiv    = Color.FromArgb(255, 35,  50,  65);
    static readonly Color cFoot   = Color.FromArgb(255, 55,  70,  85);

    readonly Font fSm  = new("Consolas", 7f);
    readonly Font fMed = new("Consolas", 9f,  FontStyle.Bold);
    readonly Font fBig = new("Consolas", 24f, FontStyle.Bold);
    readonly Font fRt  = new("Consolas", 16f, FontStyle.Bold);
    readonly Font fUn  = new("Consolas", 9f);
    readonly Font fVal = new("Consolas", 13f, FontStyle.Bold);

    UpsSnapshot? _snapshot;
    bool _showDebugPanel;
    ToolStripMenuItem _mDebug;
    bool _dot = true;
    Point _dragStart;
    bool _dragging;
    readonly NotifyIcon _tray;
    readonly UpsCollector _collector;
    readonly LocalStatusServer _statusServer;
    DateTime _lastUpdateCheckUtc = DateTime.MinValue;
    string _lastNotifiedTag;

    const string RegPath = @"Software\UPS-Status-Widget";
    const string RunKey  = @"Software\Microsoft\Windows\CurrentVersion\Run";

    public UpsWidget()
    {
        Log.W($"=== START OS:{Environment.OSVersion} Screen:{Screen.PrimaryScreen!.Bounds.Width}x{Screen.PrimaryScreen.Bounds.Height}");

        Width = W; Height = HBase; Text = "UPS Status Widget";
        FormBorderStyle = FormBorderStyle.None;
        BackColor = cBg;
        TopMost = false; ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        SetStyle(ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.AllPaintingInWmPaint  |
                 ControlStyles.UserPaint, true);
        LoadPos();
        EnsureVisibleOnScreen();
        Log.W($"Pos: {Left},{Top}");

        _tray = new NotifyIcon { Text = "UPS Status Widget", Icon = MkIcon(cGreen), Visible = true };

        var menu  = new ContextMenuStrip();
        var mShow = new ToolStripMenuItem("Show / Hide");
        mShow.Click += (_, _) => { if (Visible) { Hide(); Log.W("Hidden"); } else { Show(); Log.W("Shown"); } };

        var mAuto = new ToolStripMenuItem("Autostart with Windows") { CheckOnClick = true };
        mAuto.Checked = IsAuto();
        mAuto.Click += (_, _) => SetAuto(mAuto.Checked);

        var mResetPos = new ToolStripMenuItem("Reset position");
        mResetPos.Click += (_, _) => {
            ResetToDefaultPosition();
            Show();
            Win32.ShowWidgetWindow(Handle, Left, Top, W, Height);
            Invalidate();
            Log.W($"Position reset to {Left},{Top}");
        };

        var mDiagLog = new ToolStripMenuItem("Debug logging") { CheckOnClick = true };
        mDiagLog.Checked = IsLogEnabledSetting();
        mDiagLog.Click += (_, _) => SetLogEnabled(mDiagLog.Checked);

        var mLog = new ToolStripMenuItem("Open log");
        mLog.Click += (_, _) => {
            try {
                Directory.CreateDirectory(Log.LogDirectory);
                if (Log.Enabled && File.Exists(Log.LogFilePath)) {
                    System.Diagnostics.Process.Start("notepad.exe", Log.LogFilePath);
                } else {
                    System.Diagnostics.Process.Start("explorer.exe", Log.LogDirectory);
                }
            }
            catch { }
        };

        var mUpdate = new ToolStripMenuItem("Check for updates");
        mUpdate.Click += async (_, _) => await CheckForUpdatesAsync(interactive: true);
        var mInstall = new ToolStripMenuItem("Install latest update");
        mInstall.Click += async (_, _) => await InstallLatestUpdateAsync();

        var mExit = new ToolStripMenuItem("Exit");
        mExit.Click += (_, _) => { Log.W("Exit"); _tray.Visible = false; Application.Exit(); };

        _mDebug = new ToolStripMenuItem("Debug Panel (Ctrl+Shift+D)") { CheckOnClick = true };
        _mDebug.CheckedChanged += (_, _) => ToggleDebugPanel(_mDebug.Checked);

        menu.Items.Add(mShow);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(mAuto);
        menu.Items.Add(mResetPos);
        menu.Items.Add(mDiagLog);
        menu.Items.Add(_mDebug);
        menu.Items.Add(mUpdate);
        menu.Items.Add(mInstall);
        menu.Items.Add(mLog);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(mExit);
        _tray.ContextMenuStrip = menu;
        _tray.DoubleClick += (_, _) => { if (!Visible) Show(); Invalidate(); };

        _collector = new UpsCollector(TimeSpan.FromSeconds(1));
        _collector.SnapshotUpdated += s => {
            if (!IsHandleCreated) return;
            BeginInvoke(() => {
                _snapshot = s;
                string tl = s.Timeleft.HasValue ? $"{(int)s.Timeleft.Value}min" : "N/A";
                _tray.Text = $"UPS: {(int)s.Bcharge}% | {tl}";
                _tray.Icon = MkIcon(s.Bcharge > 50 ? cGreen : s.Bcharge > 20 ? cWarn : cDanger);
                Invalidate();
            });
        };

        var bt = new System.Windows.Forms.Timer { Interval = 800 };
        bt.Tick += (_, _) => { _dot = !_dot; Invalidate(); };
        bt.Start();

        var initial = _collector.Latest;
        if (initial.HasValue) _snapshot = initial.Value;

        _statusServer = new LocalStatusServer(
            "http://127.0.0.1:8765/",
            () => _collector.Latest,
            limit => _collector.GetRecent(limit),
            () => HidUpsApi.GetLastRawSnapshot());
        try {
            _statusServer.Start();
            Log.W("Local status endpoint: http://127.0.0.1:8765/status (history: /status/history?limit=120 raw: /status/raw)");
        }
        catch (Exception ex) {
            Log.W("Local status endpoint failed: " + ex.Message);
        }

        _ = CheckForUpdatesAsync(interactive: false);
        var updateTimer = new System.Windows.Forms.Timer { Interval = 6 * 60 * 60 * 1000 };
        updateTimer.Tick += async (_, _) => await CheckForUpdatesAsync(interactive: false);
        updateTimer.Start();

        Log.W("Constructor done");
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        try { Win32.RegisterHotKey(Handle, HotkeyDebugId, Win32.MOD_CONTROL | Win32.MOD_SHIFT, (uint)Keys.D); }
        catch { }
        Log.W($"OnShown at {Left},{Top}");
        Win32.ShowWidgetWindow(Handle, Left, Top, W, Height);
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == Win32.WM_WINDOWPOSCHANGING) {
            var wp = (Win32.WINDOWPOS)Marshal.PtrToStructure(m.LParam, typeof(Win32.WINDOWPOS))!;
            wp.hwndInsertAfter = Win32.HWND_TOP;
            wp.flags |= Win32.SWP_NOACTIVATE;
            Marshal.StructureToPtr(wp, m.LParam, false);
        }
        if (m.Msg == Win32.WM_ACTIVATE || m.Msg == Win32.WM_ACTIVATEAPP) return;
        if (m.Msg == Win32.WM_MOUSEACTIVATE) { m.Result = new IntPtr(Win32.MA_NOACTIVATE); return; }
        if (m.Msg == Win32.WM_HOTKEY && m.WParam.ToInt32() == HotkeyDebugId) {
            ToggleDebugPanel(!_showDebugPanel);
            if (_mDebug != null) _mDebug.Checked = _showDebugPanel;
            return;
        }
        base.WndProc(ref m);
    }

    protected override bool ShowWithoutActivation => true;

    protected override CreateParams CreateParams {
        get {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
            cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW
            return cp;
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        DrawBg(g); DrawHeader(g); DrawBattery(g); DrawMetrics(g);
        if (_showDebugPanel) DrawDebugPanel(g);
        DrawFooter(g);
    }

    void DrawBg(Graphics g)
    {
        using var br  = new SolidBrush(cBg);
        using var pen = new Pen(cBorder, 1f);
        using var ap  = new Pen(Color.FromArgb(200, 0, 200, 255), 1.5f);
        g.FillRectangle(br, 0, 0, Width, Height);
        g.DrawRectangle(pen, 1, 1, Width - 2, Height - 2);
        g.DrawLine(ap, 40, 1, Width - 40, 1);
    }

    void DrawHeader(Graphics g)
    {
        using var bb = new SolidBrush(cBrand);
        g.DrawString("UPS // HID MONITOR", fSm, bb, 14, 8);
        string st = _snapshot?.Status ?? "NO CONNECTION";
        Color dc = st.Contains("BATT") ? cWarn : (st.Contains("NO") || st.Contains("LOW")) ? cDanger : cGreen;
        using var sb = new SolidBrush(dc);
        if (_dot) g.FillEllipse(sb, new Rectangle(14, 22, 9, 9));
        g.DrawString(st, fMed, sb, 28, 21);
    }

    void DrawBattery(Graphics g)
    {
        double bc = _snapshot?.Bcharge ?? 0;
        double? tl = _snapshot?.Timeleft;
        Color c2 = bc > 50 ? cGreen : bc > 20 ? cWarn : cDanger;
        using var bBr  = new SolidBrush(c2);
        using var dBr  = new SolidBrush(cDim);
        using var aBr  = new SolidBrush(cAccent);
        using var dbBr = new SolidBrush(cDimBar);
        string ps = $"{(int)bc}";
        g.DrawString(ps, fBig, bBr, 10, 44);
        float pw = g.MeasureString(ps, fBig).Width;
        using var sB = new SolidBrush(Color.FromArgb(150, c2.R, c2.G, c2.B));
        g.DrawString("%", fUn, sB, pw + 4, 52);
        g.DrawString("BATTERY CHARGE", fSm, dBr, 14, 88);
        g.FillRectangle(dbBr, new Rectangle(14, 96, 160, 5));
        if (bc > 0) { int fw = (int)(160 * bc / 100); if (fw > 0) g.FillRectangle(bBr, new Rectangle(14, 96, fw, 5)); }
        g.DrawString("RUNTIME", fSm, dBr, 192, 44);
        if (tl.HasValue) {
            g.DrawString($"{(int)tl.Value}", fRt, aBr, 192, 56);
            g.DrawString("min", fSm, dBr, 192, 84);
        } else {
            g.DrawString("N/A", fMed, dBr, 192, 64);
        }
    }

    void DrawMetrics(Graphics g)
    {
        double? lv = _snapshot?.Linev, lp = _snapshot?.Loadpct, fq = _snapshot?.Freq;
        using var dBr  = new SolidBrush(cDim);
        using var tBr  = new SolidBrush(cText);
        using var aBr  = new SolidBrush(cAccent);
        using var dbBr = new SolidBrush(cDimBar);
        using var dp   = new Pen(cDiv, 1f);
        g.DrawLine(dp, 14, 108, W - 14, 108);
        string[] lb = { "MAINS", "LOAD", "FREQ" };
        string[] vl = { Fmt(lv), Fmt(lp), FmtHz(fq) };
        string[] un = { "V", "%", "Hz" };
        int[]    xs = { 14, 94, 174 };
        for (int i = 0; i < 3; i++) {
            g.DrawString(lb[i], fSm,  dBr, xs[i], 114);
            g.DrawString(vl[i], fVal, tBr, xs[i], 124);
            if (vl[i] != "N/A") {
                float vw = g.MeasureString(vl[i], fVal).Width;
                g.DrawString(un[i], fSm, dBr, xs[i] + vw - 2, 132);
            }
        }
        g.FillRectangle(dbBr, new Rectangle(94, 141, 74, 3));
        if (lp.HasValue && lp.Value > 0) {
            int lf = (int)(74 * lp.Value / 100);
            if (lf > 0) g.FillRectangle(aBr, new Rectangle(94, 141, lf, 3));
        }
    }

    void DrawFooter(Graphics g)
    {
        using var fb = new SolidBrush(cFoot);
        string ext = BuildPluginFooter(_snapshot);
        if (!string.IsNullOrWhiteSpace(ext)) g.DrawString(ext, fSm, fb, 14, Height - 24);
        string ts = _snapshot?.UpdatedAt.ToString("HH:mm:ss") ?? DateTime.Now.ToString("HH:mm:ss");
        g.DrawString($"updated {ts}", fSm, fb, 14, Height - 14);
        g.DrawString("1s",                               fSm, fb, W - 28, Height - 14);
    }

    void DrawDebugPanel(Graphics g)
    {
        if (!_snapshot.HasValue) return;
        var s = _snapshot.Value;
        int y = 148;
        int panelH = Math.Max(0, Height - y - 28);
        if (panelH < 24) return;

        using var bg = new SolidBrush(Color.FromArgb(180, 15, 22, 30));
        using var br = new Pen(Color.FromArgb(255, 45, 60, 78), 1f);
        using var tx = new SolidBrush(Color.FromArgb(220, 170, 210, 240));
        g.FillRectangle(bg, 12, y, W - 24, panelH);
        g.DrawRectangle(br, 12, y, W - 24, panelH);

        int ty = y + 4;
        void Ln(string t) { g.DrawString(t, fSm, tx, 16, ty); ty += 11; }
        Ln("PLUGIN TELEMETRY");
        Ln($"load {Fmt(s.PcLoadReport50)}% | runtime {Fmt(s.PcRuntimeReport23Minutes)}m | rt_raw {FmtU(s.PcRuntimeReport23Raw)}");
        Ln($"charge {Fmt(s.PcChargeReport22)}% | freq {FmtFreqLive(s.PcFreqLive, s.PcFreqReport0F)} | line {Fmt(s.Linev)}V");
        Ln($"status {s.Status} | s06 0x{FmtHex(s.PcStatusReport06Raw)} | bits {s.PcStatusReport06Bits ?? "na"}");
        Ln($"flags b19={FmtB(s.PcStatus06Bit19)} b18={FmtB(s.PcStatus06Bit18)} b17={FmtB(s.PcStatus06Bit17)} | r49 {FmtU(s.PcStatusReport49Raw)}");
        if (s.RawSnapshot.HasValue) {
            var r = s.RawSnapshot.Value;
            Ln($"raw {r.CapturedAt:HH:mm:ss} | feature {r.FeatureReports.Count} | input {r.InputReports.Count} | u16 {r.ValuesU16.Count}");
        }
    }

    protected override void OnMouseDown(MouseEventArgs e)
        { if (e.Button == MouseButtons.Left) { _dragging = true; _dragStart = e.Location; } }
    protected override void OnMouseMove(MouseEventArgs e)
        { if (_dragging && e.Button == MouseButtons.Left) { Left += e.X - _dragStart.X; Top += e.Y - _dragStart.Y; } }
    protected override void OnMouseUp(MouseEventArgs e)
        { if (e.Button == MouseButtons.Left) { _dragging = false; SavePos(); Log.W($"Moved to {Left},{Top}"); } }

    void LoadPos()
    {
        try {
            using var k = Registry.CurrentUser.OpenSubKey(RegPath);
            if (k != null) { Left = (int)(k.GetValue("X") ?? 30); Top = (int)(k.GetValue("Y") ?? 30); return; }
        } catch { }
        Left = 30; Top = 30;
    }

    void EnsureVisibleOnScreen()
    {
        var rect = new Rectangle(Left, Top, Width, Height);
        foreach (var s in Screen.AllScreens) {
            if (s.WorkingArea.IntersectsWith(rect)) return;
        }
        ResetToDefaultPosition();
    }

    void ResetToDefaultPosition()
    {
        var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        Left = Math.Max(wa.Left + 20, wa.Right - Width - 20);
        Top = Math.Max(wa.Top + 20, wa.Bottom - Height - 60);
        SavePos();
    }

    void SavePos()
    {
        try { using var k = Registry.CurrentUser.CreateSubKey(RegPath); k.SetValue("X", Left); k.SetValue("Y", Top); }
        catch { }
    }

    static bool IsAuto()
    {
        try { using var k = Registry.CurrentUser.OpenSubKey(RunKey); return k?.GetValue("UPS-Status-Widget") != null; }
        catch { return false; }
    }

    static bool IsLogEnabledSetting()
    {
        try {
            using var k = Registry.CurrentUser.OpenSubKey(RegPath);
            if (k == null) return Log.Enabled;
            object v = k.GetValue("LogEnabled");
            if (v == null) return Log.Enabled;
            return v switch {
                int i => i != 0,
                string s => s == "1" || string.Equals(s, "true", StringComparison.OrdinalIgnoreCase),
                _ => Log.Enabled
            };
        }
        catch {
            return Log.Enabled;
        }
    }

    void SetLogEnabled(bool on)
    {
        try {
            using var k = Registry.CurrentUser.CreateSubKey(RegPath);
            k.SetValue("LogEnabled", on ? 1 : 0, RegistryValueKind.DWord);
            if (!on) Log.W("LoggingEnabled=False");
            Log.SetEnabled(on);
        }
        catch { }
    }

    async System.Threading.Tasks.Task CheckForUpdatesAsync(bool interactive)
    {
        try {
            if (!interactive && (DateTime.UtcNow - _lastUpdateCheckUtc) < TimeSpan.FromMinutes(10)) return;
            _lastUpdateCheckUtc = DateTime.UtcNow;

            Version current = ParseCurrentVersion();
            var latest = await UpdateChecker.GetLatestAsync(Environment.Is64BitProcess);
            if (!latest.HasValue) {
                if (interactive) MessageBox.Show("Unable to check updates right now.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool hasUpdate = latest.Value.Version > current;
            if (!hasUpdate) {
                if (interactive) MessageBox.Show($"You are up to date ({FormatVersion(current)}).", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string text = $"New version available: {latest.Value.Tag} (current {FormatVersion(current)})";
            Log.W("UpdateAvailable " + text);

            if (!interactive) {
                if (_lastNotifiedTag == latest.Value.Tag) return;
                _lastNotifiedTag = latest.Value.Tag;
                try {
                    _tray.BalloonTipTitle = "UPS Status Widget";
                    _tray.BalloonTipText = text;
                    _tray.ShowBalloonTip(5000);
                }
                catch { }
                return;
            }

            var action = MessageBox.Show(
                text + "\n\nYes = Install now\nNo = Open release page\nCancel = Later",
                "UPS Status Widget",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Information);
            if (action == DialogResult.Yes) {
                await InstallUpdateAsync(latest.Value, interactive: true);
            } else if (action == DialogResult.No) {
                OpenUrl(latest.Value.Url);
            }
        }
        catch {
            if (interactive) MessageBox.Show("Unable to check updates right now.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    async System.Threading.Tasks.Task InstallLatestUpdateAsync()
    {
        try {
            Version current = ParseCurrentVersion();
            var latest = await UpdateChecker.GetLatestAsync(Environment.Is64BitProcess);
            if (!latest.HasValue) {
                MessageBox.Show("Unable to check updates right now.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (latest.Value.Version <= current) {
                MessageBox.Show($"You are up to date ({FormatVersion(current)}).", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            await InstallUpdateAsync(latest.Value, interactive: true);
        }
        catch {
            MessageBox.Show("Unable to install update right now.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    async System.Threading.Tasks.Task InstallUpdateAsync(UpdateInfo info, bool interactive)
    {
        try {
            if (!info.InstallerAsset.HasValue) {
                if (interactive && MessageBox.Show("Installer asset not found in latest release. Open release page?", "UPS Status Widget", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes) {
                    OpenUrl(info.Url);
                }
                return;
            }

            Cursor oldCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            string installerPath = await UpdateChecker.DownloadInstallerAsync(info);
            Cursor.Current = oldCursor;
            if (string.IsNullOrWhiteSpace(installerPath) || !File.Exists(installerPath)) {
                MessageBox.Show("Failed to download or verify installer.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var psi = new System.Diagnostics.ProcessStartInfo {
                FileName = installerPath,
                UseShellExecute = true
            };
            if (System.Diagnostics.Process.Start(psi) != null) {
                Log.W($"Installer started: {installerPath}");
                MessageBox.Show("Installer started. The app will now close.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _tray.Visible = false;
                Application.Exit();
            } else {
                MessageBox.Show("Unable to start installer.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        catch {
            MessageBox.Show("Unable to install update right now.", "UPS Status Widget", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    static Version ParseCurrentVersion()
    {
        string product = Application.ProductVersion ?? "0.0.0";
        string cleaned = product.Split('+')[0];
        if (Version.TryParse(cleaned, out var v)) return v;
        return new Version(0, 0, 0);
    }

    static string FormatVersion(Version v)
    {
        if (v.Build >= 0) return $"{v.Major}.{v.Minor}.{v.Build}";
        return $"{v.Major}.{v.Minor}";
    }

    static void OpenUrl(string url)
    {
        try {
            if (string.IsNullOrWhiteSpace(url)) return;
            var psi = new System.Diagnostics.ProcessStartInfo {
                FileName = url,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
        }
        catch { }
    }

    void SetAuto(bool on)
    {
        try {
            using var k = Registry.CurrentUser.CreateSubKey(RunKey);
            if (on) k.SetValue("UPS-Status-Widget", $"\"{Application.ExecutablePath}\"");
            else k.DeleteValue("UPS-Status-Widget", false);
            Log.W($"Autostart={on}");
        } catch { }
    }

    static Icon MkIcon(Color c)
    {
        var b = new Bitmap(16, 16);
        using var g = Graphics.FromImage(b);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var br = new SolidBrush(c);
        g.FillEllipse(br, new Rectangle(2, 2, 12, 12));
        return Icon.FromHandle(b.GetHicon());
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        Log.W("Closing");
        try { Win32.UnregisterHotKey(Handle, HotkeyDebugId); } catch { }
        _statusServer.Dispose();
        _collector.Dispose();
        _tray.Visible = false;
        base.OnFormClosing(e);
    }

    void ToggleDebugPanel(bool on)
    {
        _showDebugPanel = on;
        Height = on ? HDebug : HBase;
        Win32.ShowWidgetWindow(Handle, Left, Top, W, Height);
        Log.W("DebugPanel=" + on);
        Invalidate();
    }

    static string Fmt(double? v) => v.HasValue ? $"{(int)v.Value}" : "N/A";
    static string FmtHz(double? v) => v.HasValue ? v.Value.ToString("0.##") : "N/A";
    static string FmtU(uint? v) => v.HasValue ? v.Value.ToString() : "na";
    static string FmtHex(uint? v) => v.HasValue ? v.Value.ToString("X8") : "na";
    static string FmtB(bool? v) => v.HasValue ? (v.Value ? "1" : "0") : "na";
    static string FmtFreq(double? v) => v.HasValue ? $"{(int)v.Value}Hz nom" : "N/A";
    static string FmtFreqLive(double? live, double? nominal)
    {
        if (live.HasValue && nominal.HasValue) return $"{live.Value:0.##}Hz live (nom {(int)nominal.Value}Hz)";
        if (live.HasValue) return $"{live.Value:0.##}Hz live";
        return FmtFreq(nominal);
    }

    static string BuildPluginFooter(UpsSnapshot? d)
    {
        if (!d.HasValue) return string.Empty;
        var dv = d.Value;
        string load = $"plg load {Fmt(dv.PcLoadReport50)}%";
        string runtime = $"rt {Fmt(dv.PcRuntimeReport23Minutes)}m";
        string charge = $"chg {Fmt(dv.PcChargeReport22)}%";
        string freq = $"fr {FmtFreqLive(dv.PcFreqLive, dv.PcFreqReport0F)}";
        string bits = $"s06 {dv.PcStatusReport06Bits ?? "na"}";
        return JoinParts(load, runtime, charge, freq, bits);
    }

    static string JoinParts(params string[] parts)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var p in parts) {
            if (string.IsNullOrWhiteSpace(p)) continue;
            if (sb.Length > 0) sb.Append(" | ");
            sb.Append(p);
        }
        return sb.ToString();
    }

}

static class Program
{
    [STAThread]
    static void Main()
    {
        Log.W("=== PROGRAM START ===");
        Log.W($"OS: {Environment.OSVersion}");
        Log.W($"Screen: {Screen.PrimaryScreen!.Bounds.Width}x{Screen.PrimaryScreen.Bounds.Height}");
        try {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new UpsWidget());
        }
        catch (Exception ex) { Log.W("FATAL: " + ex); }
        Log.W("=== PROGRAM END ===");
    }
}


using System;
using System.Threading;

namespace UpsStatusWidget;

sealed class UpsCollector : IDisposable
{
    readonly object _lock = new();
    readonly Timer _timer;
    readonly SnapshotHistory _history = new(600);
    UpsSnapshot? _latest;
    bool _disposed;

    public event Action<UpsSnapshot> SnapshotUpdated;

    public UpsCollector(TimeSpan pollInterval)
    {
        _timer = new Timer(_ => Poll(), null, TimeSpan.Zero, pollInterval);
    }

    public UpsSnapshot? Latest
    {
        get {
            lock (_lock) return _latest;
        }
    }

    public UpsSnapshot[] GetRecent(int limit) => _history.GetRecent(limit);

    void Poll()
    {
        if (_disposed) return;
        var d = UpsReader.Read();
        if (!d.HasValue) return;

        var v = d.Value;
        var s = new UpsSnapshot(
            v.Bcharge,
            v.Timeleft,
            v.Linev,
            v.Loadpct,
            v.Freq,
            v.Status,
            v.Source,
            DateTime.Now,
            v.PcLoadReport50,
            v.PcRuntimeReport23Minutes,
            v.PcRuntimeReport23Raw,
            v.PcFreqReport0F,
            v.PcFreqLive,
            v.PcChargeReport22,
            v.PcStatusReport06Raw,
            v.PcStatusReport06Bits,
            v.PcStatus06Bit19,
            v.PcStatus06Bit18,
            v.PcStatus06Bit17,
            v.PcStatusReport49Raw,
            v.RawSnapshot
        );

        lock (_lock) _latest = s;
        _history.Add(s);
        SnapshotUpdated?.Invoke(s);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _timer.Dispose();
    }
}


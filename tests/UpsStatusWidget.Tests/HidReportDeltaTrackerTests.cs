using System;
using System.Collections.Generic;
#nullable disable

namespace UpsStatusWidget.Tests;

static class Program
{
    static int Main()
    {
        var tests = new Action[] {
            FirstObservation_ProducesNoDeltas,
            ChangedByte_ProducesSingleExpectedDelta,
            SameReportRepeated_ProducesNoAdditionalDeltas,
            CorrelationLine_FormatsValuesAndNa,
            RuntimeMapper_UsesRid12BySixty,
            SnapshotHistory_RespectsCapacity,
            SnapshotHistory_ReturnsRecentInChronologicalOrder
        };

        try {
            foreach (var test in tests) test();
            Console.WriteLine($"PASS {tests.Length} tests");
            return 0;
        }
        catch (Exception ex) {
            Console.Error.WriteLine("FAIL: " + ex.Message);
            return 1;
        }
    }

    static void FirstObservation_ProducesNoDeltas()
    {
        var tracker = new HidReportDeltaTracker();
        var deltas = tracker.UpdateAndGetDeltas("INPUT", 54, new byte[] { 54, 10, 20, 30 });
        AssertEmpty(deltas, nameof(FirstObservation_ProducesNoDeltas));
    }

    static void ChangedByte_ProducesSingleExpectedDelta()
    {
        var tracker = new HidReportDeltaTracker();
        tracker.UpdateAndGetDeltas("FEATURE", 23, new byte[] { 23, 1, 2, 3 });
        var deltas = tracker.UpdateAndGetDeltas("FEATURE", 23, new byte[] { 23, 1, 9, 3 });
        var delta = AssertSingle(deltas, nameof(ChangedByte_ProducesSingleExpectedDelta));

        AssertEqual("FEATURE", delta.Kind, "Kind");
        AssertEqual((byte)23, delta.ReportId, "ReportId");
        AssertEqual(2, delta.Offset, "Offset");
        AssertEqual((byte)2, delta.OldValue, "OldValue");
        AssertEqual((byte)9, delta.NewValue, "NewValue");
    }

    static void SameReportRepeated_ProducesNoAdditionalDeltas()
    {
        var tracker = new HidReportDeltaTracker();
        tracker.UpdateAndGetDeltas("INPUT", 22, new byte[] { 22, 7, 8 });
        tracker.UpdateAndGetDeltas("INPUT", 22, new byte[] { 22, 7, 9 });
        var deltas = tracker.UpdateAndGetDeltas("INPUT", 22, new byte[] { 22, 7, 9 });
        AssertEmpty(deltas, nameof(SameReportRepeated_ProducesNoAdditionalDeltas));
    }

    static void CorrelationLine_FormatsValuesAndNa()
    {
        string line = HidCorrelationFormatter.Format(
            15,
            42,
            50,
            1200,
            null,
            210,
            22);

        AssertEqual(
            "UPS HID CORR load=15 runtime=42 freq=50 rid12_u16=1200 rid35_u16=na rid49_b1=210 rid22_b1=22",
            line,
            nameof(CorrelationLine_FormatsValuesAndNa));
    }

    static void RuntimeMapper_UsesRid12BySixty()
    {
        double? runtime = Pid0002RuntimeMapper.TryMapMinutes(1532, null);
        AssertEqual(26d, runtime, nameof(RuntimeMapper_UsesRid12BySixty));
    }

    static void SnapshotHistory_RespectsCapacity()
    {
        var h = new SnapshotHistory(3);
        h.Add(MkSnapshot(1));
        h.Add(MkSnapshot(2));
        h.Add(MkSnapshot(3));
        h.Add(MkSnapshot(4));

        var recent = h.GetRecent(10);
        AssertEqual(3, recent.Length, nameof(SnapshotHistory_RespectsCapacity));
        AssertEqual(2d, recent[0].Bcharge, "recent[0]");
        AssertEqual(4d, recent[2].Bcharge, "recent[2]");
    }

    static void SnapshotHistory_ReturnsRecentInChronologicalOrder()
    {
        var h = new SnapshotHistory(10);
        for (int i = 1; i <= 5; i++) h.Add(MkSnapshot(i));

        var recent = h.GetRecent(2);
        AssertEqual(2, recent.Length, nameof(SnapshotHistory_ReturnsRecentInChronologicalOrder));
        AssertEqual(4d, recent[0].Bcharge, "recent[0]");
        AssertEqual(5d, recent[1].Bcharge, "recent[1]");
    }

    static UpsSnapshot MkSnapshot(double bcharge)
    {
        return new UpsSnapshot(
            bcharge,
            null,
            null,
            null,
            null,
            "ONLINE",
            "UPS HID",
            DateTime.Now,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null
        );
    }

    static HidReportDelta AssertSingle(IReadOnlyList<HidReportDelta> deltas, string testName)
    {
        if (deltas.Count != 1) throw new InvalidOperationException($"{testName}: expected 1 delta, got {deltas.Count}");
        return deltas[0];
    }

    static void AssertEmpty(IReadOnlyList<HidReportDelta> deltas, string testName)
    {
        if (deltas.Count != 0) throw new InvalidOperationException($"{testName}: expected 0 deltas, got {deltas.Count}");
    }

    static void AssertEqual<T>(T expected, T actual, string field)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) {
            throw new InvalidOperationException($"{field}: expected {expected}, got {actual}");
        }
    }
}


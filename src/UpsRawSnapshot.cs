#nullable disable
using System;
using System.Collections.Generic;

namespace UpsStatusWidget;

readonly record struct UpsRawSnapshot(
    DateTime CapturedAt,
    Dictionary<string, string> FeatureReports,
    Dictionary<string, string> InputReports,
    Dictionary<string, int> ValuesU16,
    Dictionary<string, double> Metrics
);


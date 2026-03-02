using System;

namespace UpsStatusWidget;

readonly record struct UpsSnapshot(
    double Bcharge,
    double? Timeleft,
    double? Linev,
    double? Loadpct,
    double? Freq,
    string Status,
    string Source,
    DateTime UpdatedAt,
    double? PcLoadReport50,
    double? PcRuntimeReport23Minutes,
    uint? PcRuntimeReport23Raw,
    double? PcFreqReport0F,
    double? PcFreqLive,
    double? PcChargeReport22,
    uint? PcStatusReport06Raw,
    string PcStatusReport06Bits,
    bool? PcStatus06Bit19,
    bool? PcStatus06Bit18,
    bool? PcStatus06Bit17,
    uint? PcStatusReport49Raw,
    UpsRawSnapshot? RawSnapshot
);


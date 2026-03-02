# UPS Status Widget (Windows)

Desktop widget for monitoring UPS telemetry on Windows via HID reports.

## Legal Notice (Important)

- This is an independent community project and is **not affiliated with, endorsed by, or sponsored by** any hardware vendor.
- The project reads UPS telemetry via HID interfaces and public OS APIs.
- No vendor source code, firmware extraction, or NDA materials are included.
- Use at your own risk. This software is provided **"as is"**, without warranties.
- Do not use this software as the only control/alert layer for safety-critical environments.

## Privacy and Data Handling

- The app runs locally on your machine.
- Local API binds to `127.0.0.1` only.
- Logging is disabled by default.
- Enable logging with environment variable: `UPS_HID_LOG=1`.
- Log path (when enabled): `Documents\UpsStatusWidget\logs\ups-status-widget.log` with rotation.
- Logs/raw dumps may contain device identifiers (for example serial-like values). Do not publish them without redaction.

## Current Functionality

- WinForms desktop widget.
- Tray icon with:
  - Show/Hide
  - Autostart
  - Open log
  - Exit
- Optional debug panel: `Ctrl+Shift+D`.
- Local status API:
  - `GET /status`
  - `GET /status/history?limit=120`
  - `GET /status/raw`

## Telemetry Model (Current)

The app currently works in strict report mode for key metrics (no generic metric fallbacks):

- Load: `Feature 0x50, byte1`
- Runtime: `Feature 0x23, u16` -> minutes
- Mains voltage: `Feature 0x31, byte1` (live)
- Frequency (live): calibrated from `Feature 0x31, byte1`
- Charge: `Feature 0x22, byte1`
- Status flags: `Feature 0x06` bits

Note: mapping is empirical and may differ across UPS model/firmware variants.

## Build

```powershell
dotnet restore src/UpsStatusWidget.csproj
dotnet build src/UpsStatusWidget.csproj -c Release
dotnet publish src/UpsStatusWidget.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true --output out/
```

## Tests

```powershell
dotnet run --project tests/UpsStatusWidget.Tests/UpsStatusWidget.Tests.csproj -c Debug
```

## License

See [LICENSE](LICENSE).


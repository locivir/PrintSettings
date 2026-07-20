# Locivir.Printing

Save and restore complete Windows printer configurations — including the driver-specific
`DEVMODE` blob — per application, stored in the registry under `HKEY_CURRENT_USER`.

Each configuration is stored by label and namespaced by a caller-supplied application
identity, so multiple applications (and multiple named presets per application) never
collide.

## Requirements

- Windows (uses the Win32 print spooler and `System.Windows.Forms.PrintDialog`)
- `net8.0-windows`

## Build

```bat
dotnet build -c Release
```

The output `Locivir.Printing.dll` (and `Locivir.Printing.xml` for IntelliSense) is written to
`bin\Release\net8.0-windows\`. Reference the DLL from your application, or add `Printer.cs`
directly to your project.

## Usage

```csharp
using Locivir.Printing;

// Identify your application. Use a fixed GUID...
var cfg = new PrinterConfig(new Guid("f1e2d3c4-b5a6-7890-1234-567890abcdef"));
// ...or a stable name (hashed to a deterministic GUID):
var cfg2 = new PrinterConfig("Bixlers.LabelTool");

// Load the "Invoices" preset and apply it. If it does not exist, or cannot be
// applied (e.g. the printer was removed), the print dialog is shown and the
// user's choice is saved under that label.
PrinterSettings settings = cfg.SetPrinter("Invoices", showDialog: false);

// Capture the current selection under a label:
cfg.SaveCurrent("Invoices");

// Apply a stored preset without any dialog (false if it does not exist):
bool applied = cfg.ApplySaved("Invoices");

// Read-only helpers reflecting the current selection:
string  name     = cfg.PrinterName;
bool    landscape = cfg.Landscape;
Size    paper     = cfg.PaperSize;      // hundredths of an inch, orientation-corrected
Size    printable = cfg.PrintableArea;
```

Override the storage root if you want a different registry location:

```csharp
var cfg = new PrinterConfig("Bixlers.LabelTool", @"HKEY_CURRENT_USER\Software\Bixlers\Printing");
```

## API

| Member | Description |
| --- | --- |
| `PrinterConfig(Guid, string)` | Store keyed by an explicit application id. Throws on `Guid.Empty`. |
| `PrinterConfig(string, string)` | Store keyed by an application name (RFC 4122 v5 hash → stable GUID). |
| `SetPrinter(label, showDialog)` | Load + apply a preset; prompt + save if missing/unusable. |
| `SaveCurrent(label)` | Capture the current printer + `DEVMODE` under a label. |
| `ApplySaved(label)` | Apply a stored preset; `false` if absent. |
| `PrinterSettings`, `PrinterName`, `Landscape`, `PaperSize`, `PrintableArea` | Read the current selection. |
| `PrinterConfigData` | Serializable DTO (`Application`, `Label`, `Printer`, `DevMode`). |

## Notes

- A single `PrinterConfig` instance is **not** thread-safe (it holds mutable printer
  state). Use one instance per thread, or synchronize. Separate instances are independent.
- Configurations are stored under the current user's registry hive. The `DEVMODE` blob is
  validated for internal size consistency before it is applied to a printer.
- `PrinterConfigData` is public only because it is XML-serialized; treat it as data.

## License

Released into the public domain under the [Unlicense](LICENSE).

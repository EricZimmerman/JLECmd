# JLECmd

Parses Windows jump lists, both automatic (`*.automaticDestinations-ms`) and custom (`*.customDestinations-ms`),
including Windows 10/11 ones: the DestList entries, the AppId of the program each jump list belongs to and the shortcuts
(`.lnk`) inside them, with the same details LECmd shows. Results are shown on the console and can be exported to CSV,
JSON and HTML, and the shortcuts can be saved as `.lnk` files.

| | |
|---|---|
| **Version** | 2026.10.5 |
| **Download** | [.NET 10](https://download.ericzimmermanstools.com/net/JLECmd.zip) (recommended) &middot; [.NET Framework 4.7.2](https://download.ericzimmermanstools.com/JLECmd.zip) |
| **All tools** | [ericzimmerman.github.io](https://ericzimmerman.github.io) or [Get-ZimmermanTools](https://download.ericzimmermanstools.com/Get-ZimmermanTools.zip), which downloads them all and keeps them current |
| **Changes** | [ChangeLog](https://ericzimmerman.github.io/changelog/) |
| **Issues** | [EricZimmerman/JLECmd](https://github.com/EricZimmerman/JLECmd/issues) or [EricZimmerman/Issues](https://github.com/EricZimmerman/Issues/issues) |

## Requirements

- **.NET 10 build**: the [.NET 10 runtime](https://dotnet.microsoft.com/download/dotnet/10.0). Run `JLECmd.exe` on Windows,
  or `dotnet JLECmd.dll` on any system with the runtime.
- **.NET Framework build**: .NET Framework 4.7.2 or later (Windows only).

## Command line

```text
Usage:
  JLECmd [options]

Options:
  -f <f>             File to process. Either this or -d is required
  -d <d>             Directory to recursively process. Either this or -f is required
  --all              Process all files in directory vs. only files matching *.automaticDestinations-ms or *.customDestinations-ms
  --csv <csv>        Directory to save CSV formatted results to. Be sure to include the full path in double quotes
  --csvf <csvf>      File name to save CSV formatted results to. When present, overrides default name
  --json <json>      Directory to save JSON formatted results to. Be sure to include the full path in double quotes. Use --pretty for a more human readable layout
  --jsonf <jsonf>    File name to save JSON formatted results to. When present, overrides default name
  --html <html>      Directory to save xhtml formatted results to. Be sure to include the full path in double quotes
  --pretty           When exporting to json, use a more human readable layout
  -q                 Do not dump full details about each file processed. Speeds up processing when using --json or --csv
  --ld               Include more information about lnk files
  --fd               Include full information about lnk files (Alternatively, dump lnk files using --dumpTo and process with LECmd)
  --appIds <appIds>  Path to file containing AppIDs and descriptions (appid|description format). New appIds are added to the built-in list, existing appIds will have their descriptions updated
  --dumpTo <dumpTo>  Directory to save exported lnk files
  --dt <dt>          The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options [default: yyyy-MM-dd HH:mm:ss]
  --mp               When true, display higher precision for timestamps
  --withDir          When true, show contents of Directory not accounted for in DestList entries
  --user <user>      Username to include in JSON/CSV output (e.g. the owner of the jump list artifacts)
  --cp <cp>          Code page to parse strings [default: 1252]
  --debug            Show debug information during processing
  --trace            Show trace information during processing
  --co               When true, save all console messages to <timestamp>_JLECmd_<version>_ConsoleLog.txt in the output directory (the first of --csv, --json, --html, --dumpTo given). Use --co false to turn it off
  -?, -h, --help     Show help and usage information
  --version          Show version information
```

Short options (a single letter) start with one dash, long options with two.

## Examples

```text
JLECmd.exe -f "C:\Temp\f01b4d95cf55d32a.customDestinations-ms" --mp
JLECmd.exe -f "C:\Temp\f01b4d95cf55d32a.automaticDestinations-ms" --json "D:\jsonOutput" --pretty
JLECmd.exe -d "C:\CustomDestinations" --csv "c:\temp" --html "c:\temp" -q
JLECmd.exe -d "C:\Users\e\AppData\Roaming\Microsoft\Windows\Recent" --dt "ddd yyyy MM dd HH:mm:ss.fff"
```

## Output

- **Console**: each jump list's entries (`--ld` adds more about the shortcuts in them, `--fd` all of it; `-q` leaves out
  the details).
- **CSV** (`--csv`): two files, `<timestamp>_AutomaticDestinations.csv` and `<timestamp>_CustomDestinations.csv` (with
  `--csvf`, its name with those suffixes). `TargetIDAbsolutePath` is the full path built from the shortcut's target ID
  list, with folders' real names where Windows also stores a localized one. `--user` adds a `User` column.
- **JSON** (`--json`): one file per jump list, time stamps in ISO 8601 UTC (`2024-01-02T03:04:05.6780000Z`). `--user`
  adds the user name.
- **HTML** (`--html`): one folder for the automatic and one for the custom jump lists, each with an `index.xhtml`.
- **Shortcuts** (`--dumpTo`): the shortcuts inside each jump list, saved as `.lnk` files in a folder named after the jump
  list, to process with LECmd.

Damaged or truncated jump lists keep every entry that can be read instead of failing the whole file. Unknown AppIds
show "Unknown AppId"; `--appIds` adds descriptions to the built-in list.

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Success. Damaged jump lists are reported but are not a failure |
| `1` | The tool failed: invalid arguments, a missing `-f` or `-d` path, output that can't be written, or jump lists or folders it can't read. Folders under `-d` that can't be read are reported (`Unable to read folder <path>: <reason>`) and the rest is still processed |

## What's new in 2026.10

- Uses about a fifth less memory on large sets of jump lists.
- New `--user` option: adds the user name to the CSV (a `User` column) and JSON output. JSON dates are ISO 8601 UTC.
- `--appIds` works again; unknown AppIds show "Unknown AppId" and AppId lookups ignore case.
- Damaged or truncated jump lists keep every entry that can be read; jump lists with 4096 byte sectors and large jump
  lists (about 24 MB and up) load; custom destination categories after the first get their names.
- Real folder names in `TargetIDAbsolutePath`, and all the shortcut parsing fixes listed under LECmd.
- With `--fd`, jump lists with a named property (AutoList) are no longer dropped.
- 2026.10.3: 0x2E root folder items with extension blocks after the GUID (a known folder such as Downloads) are read, and zip contents with data in bytes 4-9 (JumpList 2026.10.9, Lnk 2026.10.10, ExtensionBlocks 2026.10.8).
- 2026.10.5: `--dumpTo` on its own also saves the console log, next to the dumped lnk files.
- 2026.10.4: `--co` (on by default; `--co false` turns it off) saves all console messages to `<yyyyMMddHHmmss>_JLECmd_<version>_ConsoleLog.txt` in the output directory (the first of `--csv`, `--json`, `--html` given). `--debug` and `--trace` show the time of each message again (they showed the `--dt` date).
- Exit code 1 when the tool fails (it was always 0). Error reports include source file and line numbers.

The full list is in the [ChangeLog](https://ericzimmerman.github.io/changelog/).

## Building

```text
dotnet build JLECmd.slnx -c Release
```

Builds `net10.0` and `net472`. Jump list parsing is in the [JumpList](https://github.com/EricZimmerman/JumpList) library
([NuGet](https://www.nuget.org/packages/JumpList)), shortcut parsing in the [Lnk](https://github.com/EricZimmerman/Lnk)
library ([NuGet](https://www.nuget.org/packages/Lnk)).

## Further reading

- [Jump lists in depth: Understand the format to better understand what your tools are (or aren't) doing](https://binaryforay.blogspot.com/2016/02/jump-lists-in-depth-understand-format.html)
- [Introducing JLECmd!](https://binaryforay.blogspot.com/2016/03/introducing-jlecmd.html)
- [PECmd, LECmd, and JLECmd updated!](https://binaryforay.blogspot.com/2016/03/pecmd-lecmd-and-jlecmd-updated.html)
- [LECmd and JLECmd updated](https://binaryforay.blogspot.com/2016/04/lecmd-and-jlecmd-updated.html)
- [JLECmd v0.9.6.0 released](https://binaryforay.blogspot.com/2016/09/jlecmd-v0960-released.html)

## Special thanks

Open source development funding and support provided by the [SANS Institute](https://www.sans.org/) and
[SANS DFIR](https://www.sans.org/digital-forensics-incident-response/).

## License

[MIT](license), Copyright (c) Eric R. Zimmerman 2027.

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

JLECmd (Jump List Explorer Command line edition) is a single-project .NET console tool (`net472;net10.0`) that parses Windows jump lists (`*.automaticDestinations-ms`, `*.customDestinations-ms`) and the LNK files in them, and shows them or exports them to CSV, JSON and HTML. It is part of Eric Zimmerman's DFIR tools. It is not packable: the user builds, signs and releases it.

The parsing is in sibling NuGet packages, not here: JumpList (the containers and DestList), Lnk (the LNK files), ExtensionBlocks (shell items, extension blocks, property stores) and GuidMapping. Fixes to those formats belong in those repos, followed by a package bump here. Input may be corrupt or hostile.

## Commands

The SDK is pinned in `global.json` (.NET 10).

```bash
dotnet build -c Release                                # builds JLECmd.slnx, both net472 and net10.0
dotnet run --project JLECmd -f net10.0 -- -d path/to/dir --csv out --json out -q
dotnet run --project JLECmd -f net10.0 -- -f file.automaticDestinations-ms --fd     # full LNK details
dotnet run --project JLECmd -f net10.0 -- -d dir --dumpTo out                       # the LNKs inside, as .lnk files
```

There is no test project. CI (`.github/workflows/build.yml`) builds both targets. Dependabot opens monthly PRs for GitHub Actions and NuGet updates.

## Behavior to keep

- CSVs: `<timestamp>_AutomaticDestinations.csv` and `<timestamp>_CustomDestinations.csv` (or `--csvf` with those suffixes). Keep the columns and their order; new columns go at the end (Timeline Explorer's plugins match them by their exact header set). A `User` column is there only with `--user`.
- JSON dates are ISO 8601 UTC with a Z (`ConfigureJsonDates()`, as the other tools).
- Exit codes: 0, or 1 for bad arguments, a file that failed (`_failedFiles`) or an export that failed. `Main` returns the code (no `Environment.Exit`).
- Shell items are ExtensionBlocks' classes (since Lnk 2026.10.6): the `--ld`/`--fd` details switch on those types (`ShellBag0X31`/`0X35`/`0X36`/`0X39`/`0XB1`/`0X3A` file entries, `ShellBagCFSF`, network `0X41`-`0X47`, ...), and the target's MFT columns come from the last 0xbeef0004 of `TargetBlocks` (a 0x00 property view's `DelegateExtensionBlocks`, then its own blocks). `int.Parse` on property keys threw on named properties (`AutoList`) and dropped the whole jump list with `--fd`: keys are parsed with `TryParse`.

## Validation of a change

Build the old and the new version, run both over the jump list corpus (`Corpus/Jump lists and lnks`, outside the repo; never copy its files into it; skip its `jumplist tools` folder) with `--csv` and `--json`, and compare the CSVs as multisets of rows (ignoring `SourceAccessed`, `SourceCreated` and `SourceModified`, which reading the files can change). Explain every difference. `--dumpTo` gives the LNKs inside the jump lists for LNK-level checks (LECmd, or Lnk directly).

## Versioning

`<Version>` in `JLECmd/JLECmd.csproj`, bumped by hand. The dependency order is GuidMapping → ExtensionBlocks → Lnk → JumpList → JLECmd; JLECmd references ExtensionBlocks, Lnk and JumpList directly, so keep them in step with the versions JumpList uses.

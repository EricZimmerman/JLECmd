using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.Invocation;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Alphaleonis.Win32.Filesystem;
using Exceptionless;
using ExtensionBlocks;
using JLECmd.Properties;
using JumpList.Automatic;
using JumpList.Custom;
using Lnk;
using Lnk.ExtraData;
using Lnk.ShellItems;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using ServiceStack;
using ServiceStack.Text;
using CsvWriter = CsvHelper.CsvWriter;
#if !NET6_0_OR_GREATER
using Directory = Alphaleonis.Win32.Filesystem.Directory;
using File = Alphaleonis.Win32.Filesystem.File;
using FileInfo = Alphaleonis.Win32.Filesystem.FileInfo;
using Path = Alphaleonis.Win32.Filesystem.Path;

#else
using Directory = System.IO.Directory;
using File = System.IO.File;
using FileInfo = System.IO.FileInfo;
using Path = System.IO.Path;
#endif


namespace JLECmd
{
    internal class Program
    {
        private static readonly string _preciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

        private static List<string> _failedFiles;

        //exports that failed (CSV, JSON, HTML): the exit code is 1
        private static int _exportErrors;

        //files that could not be opened or read (a damaged file is only listed in _failedFiles)
        private static int _readErrors;
        private static string ActiveDateTimeFormat;

        private static readonly Dictionary<string, string> MacList = new();

        private static List<AutomaticDestination> _processedAutoFiles;
        private static List<CustomDestination> _processedCustomFiles;

        private static RootCommand _rootCommand;

        private static DateTimeOffset ts = DateTimeOffset.UtcNow;

        private static string Header =
            $"JLECmd version {Assembly.GetExecutingAssembly().GetName().Version.ToString(3)}" +
            "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
            "\r\nhttps://github.com/EricZimmerman/JLECmd";

        private static string Footer =
            @"Examples: JLECmd.exe -f ""C:\Temp\f01b4d95cf55d32a.customDestinations-ms"" --mp" + "\r\n\t " +
            @" JLECmd.exe -f ""C:\Temp\f01b4d95cf55d32a.automaticDestinations-ms"" --json ""D:\jsonOutput"" --pretty" +
            "\r\n\t " +
            @" JLECmd.exe -d ""C:\CustomDestinations"" --csv ""c:\temp"" --html ""c:\temp"" -q" +
            "\r\n\t " +
            @" JLECmd.exe -d ""C:\Users\e\AppData\Roaming\Microsoft\Windows\Recent"" --dt ""ddd yyyy MM dd HH:mm:ss.fff"" " +
            "\r\n\t" +
            "\r\n\t" +
            "  Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes\r\n";

        private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private static bool IsAdministrator()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return true;
            }

            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static async Task<int> Main(string[] args)
        {
            ExceptionlessClient.Default.Startup("ZqbYbvr4FIRjkpUrqCLC5N4RfKIuo9YIVmpQuOje");

            var fOpt = new Option<string>("-f")
            {
                Description = "File to process. Either this or -d is required"
            };

            var dOpt = new Option<string>("-d")
            {
                Description = "Directory to recursively process. Either this or -f is required"
            };

            var allOpt = new Option<bool>("--all")
            {
                Description =
                    "Process all files in directory vs. only files matching *.automaticDestinations-ms or *.customDestinations-ms"
            };

            var csvOpt = new Option<string>(
                "--csv")
            {
                Description =
                    "Directory to save CSV formatted results to. Be sure to include the full path in double quotes"
            };

            var csvfOpt = new Option<string>(
                "--csvf")
            {
                Description = "File name to save CSV formatted results to. When present, overrides default name"
            };
            var jsonOpt = new Option<string>(
                "--json")
            {
                Description =
                    "Directory to save JSON formatted results to. Be sure to include the full path in double quotes. Use --pretty for a more human readable layout"
            };

            var jsonfOpt = new Option<string>(
                "--jsonf")
            {
                Description = "File name to save JSON formatted results to. When present, overrides default name"
            };

            var htmlOpt = new Option<string>(
                "--html")
            {
                Description =
                    "Directory to save xhtml formatted results to. Be sure to include the full path in double quotes"
            };

            var prettyOpt = new Option<bool>("--pretty")
            {
                Description = "When exporting to json, use a more human readable layout",
                DefaultValueFactory = _ => false
            };

            var qOpt = new Option<bool>("-q")
            {
                Description =
                    "Do not dump full details about each file processed. Speeds up processing when using --json or --csv",
                DefaultValueFactory = _ => false
            };

            var ldOpt = new Option<bool>("--ld")
            {
                Description = "Include more information about lnk files",
                DefaultValueFactory = _ => false
            };

            var fdOpt = new Option<bool>("--fd")
            {
                Description =
                    "Include full information about lnk files (Alternatively, dump lnk files using --dumpTo and process with LECmd)",
                DefaultValueFactory = _ => false
            };

            var appIdOpt = new Option<string>(
                "--appIds")
            {
                Description =
                    "Path to file containing AppIDs and descriptions (appid|description format). New appIds are added to the built-in list, existing appIds will have their descriptions updated",
            };

            var dumpOpt = new Option<string>(
                "--dumpTo")
            {
                Description = "Directory to save exported lnk files",
            };

            var dtOpt = new Option<string>(
                "--dt")
            {
                Description =
                    "The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options",
                DefaultValueFactory = _ => "yyyy-MM-dd HH:mm:ss"
            };

            var mpOpt = new Option<bool>("--mp")
            {
                Description = "When true, display higher precision for timestamps",
                DefaultValueFactory = _ => false
            };

            var withDirOpt = new Option<bool>("--withDir")
            {
                Description = "When true, show contents of Directory not accounted for in DestList entries",
                DefaultValueFactory = _ => false
            };

            var cpOpt = new Option<int>("--cp")
            {
                Description = "Code page to parse strings",
                DefaultValueFactory = _ => 1252
            };

            var userOpt = new Option<string>("--user")
            {
                Description = "Username to include in JSON/CSV output (e.g. the owner of the jump list artifacts)"
            };

            var debugOpt = new Option<bool>("--debug")
            {
                Description = "Show debug information during processing",
                DefaultValueFactory = _ => false
            };
            var traceOpt = new Option<bool>("--trace")
            {
                Description = "Show trace information during processing",
                DefaultValueFactory = _ => false
            };
            var coOpt = new Option<bool>("--co")
            {
                Description = "When true, save all console messages to <timestamp>_JLECmd_<version>_ConsoleLog.txt in the output directory (the first of --csv, --json, --html given). Use --co false to turn it off",
                DefaultValueFactory = _ => true
            };

            _rootCommand = new RootCommand
            {
                fOpt,
                dOpt,
                allOpt,
                csvOpt,
                csvfOpt,
                jsonOpt,
                jsonfOpt,
                htmlOpt,
                prettyOpt,
                qOpt,
                ldOpt,
                fdOpt,
                appIdOpt,
                dumpOpt,
                dtOpt,
                mpOpt,
                withDirOpt,
                userOpt,
                cpOpt,
                debugOpt,
                traceOpt,
                coOpt
            };

            _rootCommand.Description = Header + "\r\n\r\n" + Footer;

            _rootCommand.SetAction(result => DoWork(result.GetValue(fOpt), result.GetValue(dOpt),
                result.GetValue(allOpt),
                result.GetValue(csvOpt), result.GetValue(csvfOpt), result.GetValue(jsonOpt), result.GetValue(jsonfOpt),
                result.GetValue(htmlOpt), result.GetValue(prettyOpt), result.GetValue(qOpt), result.GetValue(ldOpt),
                result.GetValue(fdOpt), result.GetValue(appIdOpt), result.GetValue(dumpOpt), result.GetValue(dtOpt),
                result.GetValue(mpOpt), result.GetValue(withDirOpt), result.GetValue(debugOpt),
                result.GetValue(traceOpt), result.GetValue(cpOpt), result.GetValue(userOpt), result.GetValue(coOpt)));

            var exitCode = await _rootCommand.Parse(args).InvokeAsync();

            Log.CloseAndFlush();

            return exitCode;
        }

        private static int DoWork(string f, string d, bool all, string csv, string csvf, string json, string jsonf,
            string html, bool pretty, bool q, bool ld, bool fd, string appIds, string dumpTo, string dt, bool mp,
            bool withDir, bool debug, bool trace, int cp, string user, bool co)
        {
            var levelSwitch = new LoggingLevelSwitch();

            JsConfig.DateHandler = DateHandler.ISO8601;
            JsConfig<DateTimeOffset>.SerializeFn = JsonDate;
            JsConfig<DateTimeOffset?>.SerializeFn = d => d.HasValue ? JsonDate(d.Value) : null;
            JsConfig<DateTime>.SerializeFn = JsonDate;
            JsConfig<DateTime?>.SerializeFn = d => d.HasValue ? JsonDate(d.Value) : null;

            ActiveDateTimeFormat = dt;

            if (mp)
            {
                ActiveDateTimeFormat = _preciseTimeFormat;
            }

            var formatter =
                new DateTimeOffsetFormatter(CultureInfo.CurrentCulture);


            var template = "{Message:lj}{NewLine}{Exception}";

            if (debug)
            {
                levelSwitch.MinimumLevel = LogEventLevel.Debug;
                template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
            }

            if (trace)
            {
                levelSwitch.MinimumLevel = LogEventLevel.Verbose;
                template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
            }

            var conf = new LoggerConfiguration()
                .WriteTo.Console(outputTemplate: template, formatProvider: formatter)
                .MinimumLevel.ControlledBy(levelSwitch);

            var consoleLog = co ? ConsoleLog.GetPath(ts, csv, json, html) : null;
            if (consoleLog != null)
            {
                conf.WriteTo.File(consoleLog, outputTemplate: ConsoleLog.Template, formatProvider: formatter);
            }

            Log.Logger = conf.CreateLogger();

            if (f.IsNullOrEmpty() && d.IsNullOrEmpty())
            {
                var aaa = new CustomHelpAction(new HelpAction());
                aaa.Invoke(_rootCommand.Parse("Either -f or -d is required. Exiting"));


                Console.WriteLine();
                return 1;
            }

            if (f.IsNullOrEmpty() == false && !File.Exists(f))
            {
                Log.Warning("File {F} not found. Exiting", f);
                Console.WriteLine();
                return 1;
            }

            if (d.IsNullOrEmpty() == false &&
                !Directory.Exists(d))
            {
                Log.Warning("Directory {D} not found. Exiting", d);
                Console.WriteLine();
                return 1;
            }


            Log.Information("{Header}", Header);
            Console.WriteLine();
            Log.Information("Command line: {Args}", string.Join(" ", Environment.GetCommandLineArgs().Skip(1)));
            Console.WriteLine();

            if (consoleLog != null)
            {
                Log.Information("Console log: {Path}", consoleLog);
                Console.WriteLine();
            }

            if (IsAdministrator() == false)
            {
                Log.Warning("Warning: Administrator privileges not found!");
                Console.WriteLine();
            }

            if (mp)
            {
                dt = _preciseTimeFormat;
            }

            _processedAutoFiles = new List<AutomaticDestination>();
            _processedCustomFiles = new List<CustomDestination>();

            _failedFiles = new List<string>();
            _readErrors = 0;


            if (appIds?.Length > 0)
            {
                if (File.Exists(appIds))
                {
                    Log.Information("Looking for AppIDs in {AppIds}", appIds);

                    var added = JumpList.JumpList.AppIdList.LoadAppListFromFile(appIds);

                    Log.Information("Loaded {Added:N0} new AppIDs from {AppIds}", added, appIds);
                    Console.WriteLine();
                }
                else
                {
                    Log.Warning("{AppIds} does not exist!", appIds);
                }
            }

            if (f?.Length > 0)
            {
                f = Path.GetFullPath(f);

                if (IsAutomaticDestinationFile(f))
                {
                    try
                    {
                        AutomaticDestination adjl;
                        adjl = ProcessAutoFile(f, q, dt, fd, ld, withDir, cp);
                        if (adjl != null)
                        {
                            _processedAutoFiles.Add(adjl);
                        }
                    }
                    catch (UnauthorizedAccessException ua)
                    {
                        Log.Error(ua,
                            "Unable to access {F}. Are you running as an administrator? Error: {Message}", f,
                            ua.Message);
                        Console.WriteLine();
                        return 1;
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Error processing jump list: {Message}", ex.Message);
                        Console.WriteLine();
                        return 1;
                    }
                }
                else
                {
                    try
                    {
                        CustomDestination cdjl;
                        cdjl = ProcessCustomFile(f, q, dt, ld, cp);
                        if (cdjl != null)
                        {
                            _processedCustomFiles.Add(cdjl);
                        }
                    }
                    catch (UnauthorizedAccessException ua)
                    {
                        Log.Error(ua,
                            "Unable to access {F}. Are you running as an administrator? Error: {Message}", f,
                            ua.Message);
                        return 1;
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex,
                            "Error processing jump list. Error: {Message}", ex.Message);
                        return 1;
                    }
                }
            }
            else
            {
                Log.Information("Looking for jump list files in {D}", d);
                Console.WriteLine();

                d = Path.GetFullPath(d);

                var jumpFiles = new List<string>();

                try
                {
#if !NET6_0_OR_GREATER
                    var filters = new DirectoryEnumerationFilters();
                    filters.InclusionFilter = fsei =>
                    {
                        var mask = ".*Destinations-ms".ToUpperInvariant();
                        if (all)
                        {
                            mask = "*";
                        }

                        if (mask == "*")
                        {
                            return true;
                        }

                        if (fsei.Extension.ToUpperInvariant() == ".AUTOMATICDESTINATIONS-MS" ||
                            fsei.Extension.ToUpperInvariant() == ".CUSTOMDESTINATIONS-MS")
                        {
                            return true;
                        }

                        return false;
                    };

                    filters.RecursionFilter = entryInfo => !entryInfo.IsMountPoint && !entryInfo.IsSymbolicLink;

                    filters.ErrorFilter = UnreadableFolder;

                    var dirEnumOptions =
                        DirectoryEnumerationOptions.Files | DirectoryEnumerationOptions.Recursive |
                        DirectoryEnumerationOptions.SkipReparsePoints |
                        DirectoryEnumerationOptions.BasicSearch;

                    var files2 = Directory.EnumerateFileSystemEntries(d, dirEnumOptions, filters);
#else
                    var mask = "*.*Destinations-ms".ToUpperInvariant();
                    if (all)
                    {
                        mask = "*";
                    }

                    //folders that can't be read are reported and counted (IgnoreInaccessible skipped them silently)
                    var files2 = FindEntries(d, mask);
#endif


                    jumpFiles.AddRange(files2);
                }
                catch (UnauthorizedAccessException ua)
                {
                    Log.Error(ua,
                        "Unable to access {D}. Error message: {Message}", d, ua.Message);
                    return 1;
                }
                catch (Exception ex)
                {
                    Log.Error(ex,
                        "Error getting jump list files in {D}. Error: {Message}", d, ex.Message);
                    return 1;
                }

                Log.Information("Found {Count:N0} files", jumpFiles.Count);
                Console.WriteLine();

                var sw = new Stopwatch();
                sw.Start();

                foreach (var file in jumpFiles)
                {
                    if (IsAutomaticDestinationFile(file))
                    {
                        AutomaticDestination adjl;
                        adjl = ProcessAutoFile(file, q, dt, fd, ld, withDir, cp);
                        if (adjl != null)
                        {
                            _processedAutoFiles.Add(adjl);
                        }
                    }
                    else
                    {
                        CustomDestination cdjl;
                        cdjl = ProcessCustomFile(file, q, dt, ld, cp);
                        if (cdjl != null)
                        {
                            _processedCustomFiles.Add(cdjl);
                        }
                    }
                }

                sw.Stop();

                if (q)
                {
                    Console.WriteLine();
                }

                Log.Information(
                    "Processed {ProcessedCount:N0} out of {Count:N0} files in {TotalSeconds:N4} seconds",
                    jumpFiles.Count - _failedFiles.Count, jumpFiles.Count, sw.Elapsed.TotalSeconds);
                if (_failedFiles.Count > 0)
                {
                    Console.WriteLine();
                    Log.Information("Failed files");
                    foreach (var failedFile in _failedFiles)
                    {
                        Log.Information("  {FailedFile}", failedFile);
                    }
                }
            }

            //export lnks if requested
            if (dumpTo?.Length > 0)
            {
                Console.WriteLine();
                Log.Information(
                    "Dumping lnk files to {DumpTo}", dumpTo);

                if (Directory.Exists(dumpTo) == false)
                {
                    Directory.CreateDirectory(dumpTo);
                }

                foreach (var processedCustomFile in _processedCustomFiles)
                foreach (var entry in processedCustomFile.Entries)
                {
                    if (entry.LnkFiles.Count == 0)
                    {
                        continue;
                    }

                    var outDir = Path.Combine(dumpTo,
                        Path.GetFileName(processedCustomFile.SourceFile));

                    if (Directory.Exists(outDir) == false)
                    {
                        Directory.CreateDirectory(outDir);
                    }

                    entry.DumpAllLnkFiles(outDir, processedCustomFile.AppId.AppId);
                }

                foreach (var automaticDestination in _processedAutoFiles)
                {
                    if (automaticDestination.DestListCount == 0 &&
                        withDir == false)
                    {
                        continue;
                    }

                    var outDir = Path.Combine(dumpTo,
                        Path.GetFileName(automaticDestination.SourceFile));

                    if (Directory.Exists(outDir) == false)
                    {
                        Directory.CreateDirectory(outDir);
                    }

                    automaticDestination.DumpAllLnkFiles(outDir);
                }
            }

            if (_processedAutoFiles.Count > 0)
            {
                ExportAuto(csv, csvf, json, jsonf, html, pretty, dt, debug, withDir, user);
            }

            if (_processedCustomFiles.Count > 0)
            {
                ExportCustom(csv, csvf, json, jsonf, html, pretty, dt, user);
            }

            //1 when a file could not be opened or read, or an export failed. Damaged files are listed under Failed files but are
            //not a failure of the run
            return _readErrors > 0 || _exportErrors > 0 ? 1 : 0;
        }

        private static void ExportCustom(string csv, string csvf, string json, string jsonf, string html, bool pretty,
            string dt, string user)
        {
            Console.WriteLine();


            try
            {
                CsvWriter csvCustom = null;
                StreamWriter swCustom = null;

                if (csv?.Length > 0)
                {
                    if (Directory.Exists(csv) == false)
                    {
                        Log.Information("{Csv} does not exist. Creating...", csv);
                        Directory.CreateDirectory(csv);
                    }


                    var outName = $"{ts:yyyyMMddHHmmss}_CustomDestinations.csv";

                    if (csvf.IsNullOrEmpty() == false)
                    {
                        outName =
                            $"{Path.GetFileNameWithoutExtension(csvf)}_CustomDestinations{Path.GetExtension(csvf)}";
                    }

                    var outFile = Path.Combine(csv, outName);


                    Log.Information(
                        "CustomDestinations CSV output will be saved to {OutFile}", outFile);

                    try
                    {
                        swCustom = new StreamWriter(outFile);
                        csvCustom = new CsvWriter(swCustom, CultureInfo.InvariantCulture);

                        //only add the User column when --user is used, so the default layout does not change
                        var mapCustom = csvCustom.Context.AutoMap<CustomCsvOut>();
                        if (user.IsNullOrEmpty())
                        {
                            mapCustom.Map(m => m.User).Ignore();
                        }

                        csvCustom.Context.RegisterClassMap(mapCustom);

                        csvCustom.WriteHeader(typeof(CustomCsvOut));
                        csvCustom.NextRecord();
                    }
                    catch (Exception ex)
                    {
                        _exportErrors++;
                        Log.Error(ex,
                            "Unable to write to {Csv}. Custom CSV export canceled. Error: {Message}", csv, ex.Message);
                    }
                }

                if (json?.Length > 0)
                {
                    if (Directory.Exists(json) == false)
                    {
                        Log.Information("{Json} does not exist. Creating...", json);
                        Directory.CreateDirectory(json);
                    }

                    Log.Information("Saving Custom json output to {Json}", json);
                }


                XmlTextWriter xml = null;

                if (html?.Length > 0)
                {
                    if (Directory.Exists(html) == false)
                    {
                        Log.Information("{Html} does not exist. Creating...", html);
                        Directory.CreateDirectory(html);
                    }


                    var outDir = Path.Combine(html,
                        $"{ts:yyyyMMddHHmmss}_JLECmd_Custom_Output_for_{html.Replace(@":\", "_").Replace(@"\", "_")}");

                    if (Directory.Exists(outDir) == false)
                    {
                        Directory.CreateDirectory(outDir);
                    }

                    var styleDir = Path.Combine(outDir, "styles");
                    if (Directory.Exists(styleDir) == false)
                    {
                        Directory.CreateDirectory(styleDir);
                    }

                    File.WriteAllText(Path.Combine(styleDir, "normalize.css"), Resources.normalize);
                    File.WriteAllText(Path.Combine(styleDir, "style.css"), Resources.style);

                    var outFile = Path.Combine(html, outDir, "index.xhtml");

                    Log.Information("Saving HTML output to {OutFile}", outFile);

                    xml = new XmlTextWriter(outFile, Encoding.UTF8)
                    {
                        Formatting = Formatting.Indented,
                        Indentation = 4
                    };

                    xml.WriteStartDocument();

                    xml.WriteProcessingInstruction("xml-stylesheet", "href=\"styles/normalize.css\"");
                    xml.WriteProcessingInstruction("xml-stylesheet", "href=\"styles/style.css\"");

                    xml.WriteStartElement("document");
                }

                foreach (var processedFile in _processedCustomFiles)
                {
                    if (json?.Length > 0)
                    {
                        SaveJsonCustom(processedFile, pretty,
                            json, jsonf, user);
                    }


                    var records = GetCustomCsvFormat(processedFile, dt, user);

                    try
                    {
                        csvCustom?.WriteRecords(records);
                    }
                    catch (Exception ex)
                    {
                        _exportErrors++;
                        Log.Error(ex,
                            "Error writing record for {SourceFile} to {Csv}. Error: {Message}",
                            processedFile.SourceFile, csv, ex.Message);
                    }


                    //XHTML
                    xml?.WriteStartElement("Container");

                    xml?.WriteStartElement("SourceFile");
                    xml?.WriteAttributeString("title", "Note: Location and name of processed Jump List");
                    xml?.WriteString(processedFile.SourceFile);
                    xml?.WriteEndElement();


                    var fs = new FileInfo(processedFile.SourceFile);
                    var ct = DateTimeOffset.FromFileTime(fs.CreationTime.ToFileTime()).ToUniversalTime();
                    var mt = DateTimeOffset.FromFileTime(fs.LastWriteTime.ToFileTime()).ToUniversalTime();
                    var at = DateTimeOffset.FromFileTime(fs.LastAccessTime.ToFileTime()).ToUniversalTime();

                    xml?.WriteElementString("SourceCreated", ct.ToString(dt));
                    xml?.WriteElementString("SourceModified",
                        mt.ToString(dt));
                    xml?.WriteElementString("SourceAccessed",
                        at.ToString(dt));

                    xml?.WriteElementString("AppId", processedFile.AppId.AppId);
                    xml?.WriteElementString("AppIdDescription", processedFile.AppId.Description);

                    var index = 0;
                    foreach (var o in records)
                    {
                        xml?.WriteStartElement("lftColumn");
                        xml?.WriteStartElement("EntryNumber_large");
                        xml?.WriteAttributeString("title", "Note: Lnk position in file");
                        xml?.WriteString(index.ToString());
                        xml?.WriteEndElement();
                        xml?.WriteEndElement();

                        xml?.WriteStartElement("rgtColumn");

                        xml?.WriteElementString("EntryName", o.EntryName);

                        xml?.WriteElementString("TargetIDAbsolutePath", o.TargetIDAbsolutePath);
                        if (o.Arguments.Length > 0)
                        {
                            xml?.WriteElementString("Arguments", o.Arguments);
                        }


                        xml?.WriteElementString("TargetCreated", o.TargetCreated);
                        xml?.WriteElementString("TargetModified", o.TargetModified);
                        xml?.WriteElementString("TargetAccessed", o.TargetAccessed);
                        xml?.WriteElementString("FileSize", o.FileSize.ToString());
                        xml?.WriteElementString("RelativePath", o.RelativePath);
                        xml?.WriteElementString("WorkingDirectory", o.WorkingDirectory);
                        xml?.WriteElementString("FileAttributes", o.FileAttributes);
                        xml?.WriteElementString("HeaderFlags", o.HeaderFlags);
                        xml?.WriteElementString("DriveType", o.DriveType);
                        xml?.WriteElementString("VolumeSerialNumber", o.VolumeSerialNumber);
                        xml?.WriteElementString("VolumeLabel", o.VolumeLabel);
                        xml?.WriteElementString("LocalPath", o.LocalPath);
                        xml?.WriteElementString("CommonPath", o.CommonPath);

                        xml?.WriteElementString("TargetMFTEntryNumber", $"{o.TargetMFTEntryNumber}");
                        xml?.WriteElementString("TargetMFTSequenceNumber", $"{o.TargetMFTSequenceNumber}");


                        xml?.WriteElementString("MachineID", o.MachineID);
                        xml?.WriteElementString("MachineMACAddress", o.MachineMACAddress);
                        xml?.WriteElementString("TrackerCreatedOn", o.TrackerCreatedOn);

                        xml?.WriteElementString("ExtraBlocksPresent", o.ExtraBlocksPresent);

                        xml?.WriteEndElement();
                        index += 1;
                    }

                    xml?.WriteEndElement();
                }


                //Close CSV stuff
                swCustom?.Flush();
                swCustom?.Close();

                //Close XML
                xml?.WriteEndElement();
                xml?.WriteEndDocument();
                xml?.Flush();
            }
            catch (Exception ex)
            {
                _exportErrors++;
                Log.Error(ex,
                    "Error exporting Custom Destinations data! Error: {Message}", ex.Message);
            }
        }

        private static void ExportAuto(string csv, string csvf, string json, string jsonf, string html, bool pretty,
            string dt, bool debug, bool wd, string user)
        {
            Console.WriteLine();

            try
            {
                CsvWriter csvAuto = null;
                StreamWriter swAuto = null;

                if (csv?.Length > 0)
                {
                    if (Directory.Exists(csv) == false)
                    {
                        Log.Information("{Csv} does not exist. Creating...", csv);
                        Directory.CreateDirectory(csv);
                    }

                    var outName = $"{ts:yyyyMMddHHmmss}_AutomaticDestinations.csv";

                    if (csvf.IsNullOrEmpty() == false)
                    {
                        outName =
                            $"{Path.GetFileNameWithoutExtension(csvf)}_AutomaticDestinations{Path.GetExtension(csvf)}";
                    }

                    var outFile = Path.Combine(csv, outName);

                    Log.Information(
                        "AutomaticDestinations CSV output will be saved to {OutFile}", outFile);

                    try
                    {
                        swAuto = new StreamWriter(outFile);
                        csvAuto = new CsvWriter(swAuto, CultureInfo.InvariantCulture);

                        //only add the User column when --user is used, so the default layout does not change
                        var mapAuto = csvAuto.Context.AutoMap<AutoCsvOut>();
                        if (user.IsNullOrEmpty())
                        {
                            mapAuto.Map(m => m.User).Ignore();
                        }

                        csvAuto.Context.RegisterClassMap(mapAuto);

                        csvAuto.WriteHeader(typeof(AutoCsvOut));
                        csvAuto.NextRecord();
                    }
                    catch (Exception ex)
                    {
                        _exportErrors++;
                        Log.Error(ex,
                            "Unable to write to {Csv}. Automatic CSV export canceled. Error: {Message}", csv,
                            ex.Message);
                    }
                }

                if (json?.Length > 0)
                {
                    if (Directory.Exists(json) == false)
                    {
                        Log.Information("{Json} does not exist. Creating...", json);
                        Directory.CreateDirectory(json);
                    }

                    Log.Information("Saving Automatic json output to {Json}", json);
                }


                XmlTextWriter xml = null;

                if (html?.Length > 0)
                {
                    if (Directory.Exists(html) == false)
                    {
                        Log.Information("{Html} does not exist. Creating...", html);
                        Directory.CreateDirectory(html);
                    }

                    var outDir = Path.Combine(html,
                        $"{ts:yyyyMMddHHmmss}_JLECmd_Automatic_Output_for_{html.Replace(@":\", "_").Replace(@"\", "_")}");

                    if (Directory.Exists(outDir) == false)
                    {
                        Directory.CreateDirectory(outDir);
                    }

                    var stylesDir = Path.Combine(outDir, "styles");
                    if (Directory.Exists(stylesDir) == false)
                    {
                        Directory.CreateDirectory(stylesDir);
                    }

                    File.WriteAllText(Path.Combine(stylesDir, "normalize.css"), Resources.normalize);
                    File.WriteAllText(Path.Combine(stylesDir, "style.css"), Resources.style);

                    var outFile = Path.Combine(html, outDir, "index.xhtml");

                    Log.Information("Saving HTML output to {OutFile}", outFile);

                    xml = new XmlTextWriter(outFile, Encoding.UTF8)
                    {
                        Formatting = Formatting.Indented,
                        Indentation = 4
                    };

                    xml.WriteStartDocument();

                    xml.WriteProcessingInstruction("xml-stylesheet", "href=\"styles/normalize.css\"");
                    xml.WriteProcessingInstruction("xml-stylesheet", "href=\"styles/style.css\"");

                    xml.WriteStartElement("document");
                }

                foreach (var processedFile in _processedAutoFiles)
                {
                    if (json?.Length > 0)
                    {
                        SaveJsonAuto(processedFile, pretty,
                            json, jsonf, user);
                    }

                    var records = GetAutoCsvFormat(processedFile, debug, dt, wd, user);

                    try
                    {
                        csvAuto?.WriteRecords(records);
                    }
                    catch (Exception ex)
                    {
                        _exportErrors++;
                        Log.Error(ex,
                            "Error writing record for {SourceFile} to {Csv}. Error: {Message}",
                            processedFile.SourceFile, csv, ex.Message);
                    }

                    var fs = new FileInfo(processedFile.SourceFile);
                    var ct = DateTimeOffset.FromFileTime(fs.CreationTime.ToFileTime()).ToUniversalTime();
                    var mt = DateTimeOffset.FromFileTime(fs.LastWriteTime.ToFileTime()).ToUniversalTime();
                    var at = DateTimeOffset.FromFileTime(fs.LastAccessTime.ToFileTime()).ToUniversalTime();

                    if (xml != null)
                    {
                        xml.WriteStartElement("Container");
                        xml.WriteElementString("SourceFile", processedFile.SourceFile);
                        xml.WriteElementString("SourceCreated",
                            ct.ToString(dt));
                        xml.WriteElementString("SourceModified",
                            mt.ToString(dt));
                        xml.WriteElementString("SourceAccessed",
                            at.ToString(dt));

                        xml.WriteElementString("AppId", processedFile.AppId.AppId);
                        xml.WriteElementString("AppIdDescription", processedFile.AppId.Description);
                        xml.WriteElementString("DestListVersion", processedFile.DestListVersion.ToString());
                        xml.WriteElementString("LastUsedEntryNumber", processedFile.LastUsedEntryNumber.ToString());


                        foreach (var o in records)
                        {
                            //XHTML
                            xml.WriteStartElement("lftColumn");

                            xml.WriteStartElement("EntryNumber_large");
                            xml.WriteAttributeString("title", "Entry number");
                            xml.WriteString(o.EntryNumber);
                            xml.WriteEndElement();

                            xml.WriteEndElement();

                            xml.WriteStartElement("rgtColumn");

                            //       xml.WriteElementString("EntryNumber", o.EntryNumber);
                            xml.WriteElementString("TargetIDAbsolutePath",
                                RemoveInvalidXmlChars(o.TargetIDAbsolutePath));

                            xml.WriteElementString("CreationTime", o.CreationTime);
                            xml.WriteElementString("LastModified", o.LastModified);
                            xml.WriteElementString("Hostname", o.Hostname);
                            xml.WriteElementString("MacAddress", o.MacAddress);
                            xml.WriteElementString("Path", o.Path);
                            xml.WriteElementString("PinStatus", o.PinStatus);
                            xml.WriteElementString("FileBirthDroid", o.FileBirthDroid);
                            xml.WriteElementString("FileDroid", o.FileDroid);
                            xml.WriteElementString("VolumeBirthDroid", o.VolumeBirthDroid);
                            xml.WriteElementString("VolumeDroid", o.VolumeDroid);


                            if (o.Arguments?.Length > 0)
                            {
                                xml.WriteElementString("Arguments", o.Arguments);
                            }

                            xml.WriteElementString("TargetCreated", o.TargetCreated);
                            xml.WriteElementString("TargetModified", o.TargetModified);
                            xml.WriteElementString("TargetAccessed", o.TargetAccessed);
                            xml.WriteElementString("InteractionCount", o.InteractionCount);

                            xml.WriteElementString("FileSize", o.FileSize.ToString());
                            xml.WriteElementString("RelativePath", o.RelativePath);
                            xml.WriteElementString("WorkingDirectory", o.WorkingDirectory);
                            xml.WriteElementString("FileAttributes", o.FileAttributes);
                            xml.WriteElementString("HeaderFlags", o.HeaderFlags);
                            xml.WriteElementString("DriveType", o.DriveType);
                            xml.WriteElementString("VolumeSerialNumber", o.VolumeSerialNumber);
                            xml.WriteElementString("VolumeLabel", o.VolumeLabel);
                            xml.WriteElementString("LocalPath", o.LocalPath);
                            xml.WriteElementString("CommonPath", o.CommonPath);


                            xml.WriteElementString("TargetMFTEntryNumber", $"{o.TargetMFTEntryNumber}");
                            xml.WriteElementString("TargetMFTSequenceNumber", $"{o.TargetMFTSequenceNumber}");

                            xml.WriteElementString("MachineID", o.MachineID);
                            xml.WriteElementString("MachineMACAddress", o.MachineMACAddress);
                            xml.WriteElementString("TrackerCreatedOn", o.TrackerCreatedOn);

                            xml.WriteElementString("ExtraBlocksPresent", o.ExtraBlocksPresent);

                            xml.WriteEndElement();
                        }

                        xml.WriteEndElement();
                    }
                }

                //Close CSV stuff
                swAuto?.Flush();
                swAuto?.Close();

                //Close XML
                xml?.WriteEndElement();
                xml?.WriteEndDocument();
                xml?.Flush();
            }
            catch (Exception ex)
            {
                _exportErrors++;
                Log.Error(ex,
                    "Error exporting Automatic Destinations data: {Message}", ex.Message);
            }
        }

        private static string RemoveInvalidXmlChars(string text)
        {
            if (text == null)
            {
                return "";
            }

            if (text.Trim().Length == 0)
            {
                return text;
            }

            var validXmlChars = text.Where(XmlConvert.IsXmlChar).ToArray();
            return new string(validXmlChars);
        }

        private static bool IsAutomaticDestinationFile(string file)
        {
            const ulong signature = 0xe11ab1a1e011cfd0;

            try
            {
                var sig = BitConverter.ToUInt64(File.ReadAllBytes(file), 0);

                return signature == sig;
            }
            catch (Exception)
            {
                // ignored
            }

            return false;
        }

        private static List<CustomCsvOut> GetCustomCsvFormat(CustomDestination cust, string dt, string user)
        {
            var csList = new List<CustomCsvOut>();

            var fs = new FileInfo(cust.SourceFile);
            var ct = DateTimeOffset.FromFileTime(fs.CreationTime.ToFileTime()).ToUniversalTime();
            var mt = DateTimeOffset.FromFileTime(fs.LastWriteTime.ToFileTime()).ToUniversalTime();
            var at = DateTimeOffset.FromFileTime(fs.LastAccessTime.ToFileTime()).ToUniversalTime();

            foreach (var entry in cust.Entries)
            foreach (var lnk in entry.LnkFiles)
            {
                var csOut = new CustomCsvOut
                {
                    User = user,
                    SourceFile = cust.SourceFile,
                    SourceCreated = ct.ToString(dt),
                    SourceModified = mt.ToString(dt),
                    SourceAccessed = at.ToString(dt),
                    AppId = cust.AppId.AppId,
                    AppIdDescription = cust.AppId.Description,
                    EntryName = entry.Name,
                    TargetCreated =
                        lnk.Header.TargetCreationDate.Year == 1601
                            ? string.Empty
                            : lnk.Header.TargetCreationDate.ToString(dt),
                    TargetModified =
                        lnk.Header.TargetModificationDate.Year == 1601
                            ? string.Empty
                            : lnk.Header.TargetModificationDate.ToString(
                                dt),
                    TargetAccessed =
                        lnk.Header.TargetLastAccessedDate.Year == 1601
                            ? string.Empty
                            : lnk.Header.TargetLastAccessedDate.ToString(
                                dt),
                    CommonPath = lnk.CommonPath,
                    VolumeLabel = lnk.VolumeInfo?.VolumeLabel,
                    VolumeSerialNumber = lnk.VolumeInfo?.VolumeSerialNumber,
                    DriveType =
                        lnk.VolumeInfo == null ? "(None)" : GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType),
                    FileAttributes = lnk.Header.FileAttributes.ToString(),
                    FileSize = lnk.Header.FileSize,
                    HeaderFlags = lnk.Header.DataFlags.ToString(),
                    LocalPath = lnk.LocalPath,
                    RelativePath = lnk.RelativePath
                };

                csOut.Arguments = string.Empty;
                if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) ==
                    Lnk.Header.DataFlag.HasArguments)
                {
                    csOut.Arguments = lnk.Arguments ?? string.Empty;
                }

                if (lnk.TargetIDs?.Count > 0)
                {
                    csOut.TargetIDAbsolutePath = GetAbsolutePathFromTargetIDs(lnk.TargetIDs);
                }

                csOut.WorkingDirectory = lnk.WorkingDirectory;

                var ebPresent = string.Empty;

                if (lnk.ExtraBlocks.Count > 0)
                {
                    var names = new List<string>();

                    foreach (var extraDataBase in lnk.ExtraBlocks)
                    {
                        names.Add(extraDataBase.GetType().Name);
                    }

                    ebPresent = string.Join(", ", names);
                }

                csOut.ExtraBlocksPresent = ebPresent;

                var tnb = lnk.ExtraBlocks.SingleOrDefault(t => t.GetType().Name.ToUpper() == "TRACKERDATABASEBLOCK");

                if (tnb != null)
                {
                    var tnbBlock = tnb as TrackerDataBaseBlock;

                    csOut.TrackerCreatedOn =
                        tnbBlock?.CreationTime.ToString(dt);

                    csOut.MachineID = tnbBlock?.MachineId;
                    csOut.MachineMACAddress = tnbBlock?.MacAddress;
                }

                if (lnk.TargetIDs?.Count > 0)
                {
                    var si = lnk.TargetIDs.Last();

                    if (TargetBlocks(si).Count > 0)
                    {
                        var eb = TargetBlocks(si).LastOrDefault(t => t is Beef0004);
                        if (eb is Beef0004)
                        {
                            var eb4 = eb as Beef0004;
                            if (eb4.MFTInformation.MFTEntryNumber != null)
                            {
                                csOut.TargetMFTEntryNumber =
                                    $"0x{eb4.MFTInformation.MFTEntryNumber.Value:X}";
                            }

                            if (eb4.MFTInformation.MFTSequenceNumber != null)
                            {
                                csOut.TargetMFTSequenceNumber =
                                    $"0x{eb4.MFTInformation.MFTSequenceNumber.Value:X}";
                            }
                        }
                    }
                }

                csList.Add(csOut);
            }

            return csList;
        }

        private static List<AutoCsvOut> GetAutoCsvFormat(AutomaticDestination auto, bool debug, string dt, bool wd, string user)
        {
            var csList = new List<AutoCsvOut>();

            var fs = new FileInfo(auto.SourceFile);
            var ct = DateTimeOffset.FromFileTime(fs.CreationTime.ToFileTime()).ToUniversalTime();
            var mt = DateTimeOffset.FromFileTime(fs.LastWriteTime.ToFileTime()).ToUniversalTime();
            var at = DateTimeOffset.FromFileTime(fs.LastAccessTime.ToFileTime()).ToUniversalTime();


            foreach (var destListEntry in auto.DestListEntries)
            {
                if (debug)
                {
                    Log.Debug("Dumping destListEntry");
                    destListEntry.PrintDump();
                }

                var csOut = new AutoCsvOut
                {
                    User = user,
                    SourceFile = auto.SourceFile,
                    SourceCreated = ct.ToString(dt),
                    SourceModified = mt.ToString(dt),
                    SourceAccessed = at.ToString(dt),
                    AppId = auto.AppId.AppId,
                    AppIdDescription = auto.AppId.Description,
                    HasSps = auto.HasSps,
                    DestListVersion = auto.DestListVersion.ToString(),
                    MRU = destListEntry.MRUPosition.ToString("F0"),
                    LastUsedEntryNumber = auto.LastUsedEntryNumber.ToString(),
                    EntryNumber = destListEntry.EntryNumber.ToString("X"),
                    CreationTime =
                        destListEntry.CreatedOn.Year == 1582
                            ? string.Empty
                            : destListEntry.CreatedOn.ToString(dt),
                    LastModified = destListEntry.LastModified.ToString(dt),
                    Hostname = destListEntry.Hostname,
                    MacAddress =
                        destListEntry.MacAddress == "00:00:00:00:00:00" ? string.Empty : destListEntry.MacAddress,
                    Path = destListEntry.Path,
                    PinStatus = destListEntry.Pinned.ToString(),
                    FileBirthDroid =
                        destListEntry.FileBirthDroid.ToString() == "00000000-0000-0000-0000-000000000000"
                            ? string.Empty
                            : destListEntry.FileBirthDroid.ToString(),
                    FileDroid =
                        destListEntry.FileDroid.ToString() == "00000000-0000-0000-0000-000000000000"
                            ? string.Empty
                            : destListEntry.FileDroid.ToString(),
                    VolumeBirthDroid =
                        destListEntry.VolumeBirthDroid.ToString() == "00000000-0000-0000-0000-000000000000"
                            ? string.Empty
                            : destListEntry.VolumeBirthDroid.ToString(),
                    VolumeDroid =
                        destListEntry.VolumeDroid.ToString() == "00000000-0000-0000-0000-000000000000"
                            ? string.Empty
                            : destListEntry.VolumeDroid.ToString(),
                    TargetCreated =
                        destListEntry.Lnk?.Header.TargetCreationDate.Year == 1601
                            ? string.Empty
                            : destListEntry.Lnk?.Header.TargetCreationDate.ToString(
                                dt),
                    TargetModified =
                        destListEntry.Lnk?.Header.TargetModificationDate.Year == 1601
                            ? string.Empty
                            : destListEntry.Lnk?.Header.TargetModificationDate.ToString(
                                dt),
                    InteractionCount = destListEntry.InteractionCount.ToString(),
                    TargetAccessed =
                        destListEntry.Lnk?.Header.TargetLastAccessedDate.Year == 1601
                            ? string.Empty
                            : destListEntry.Lnk?.Header.TargetLastAccessedDate.ToString(
                                dt),
                    CommonPath = destListEntry.Lnk?.CommonPath,
                    VolumeLabel = destListEntry.Lnk?.VolumeInfo?.VolumeLabel,
                    VolumeSerialNumber = destListEntry.Lnk?.VolumeInfo?.VolumeSerialNumber,
                    DriveType =
                        destListEntry.Lnk?.VolumeInfo == null
                            ? "(None)"
                            : GetDescriptionFromEnumValue(destListEntry.Lnk?.VolumeInfo.DriveType),
                    FileAttributes = destListEntry.Lnk?.Header.FileAttributes.ToString(),
                    FileSize = destListEntry.Lnk?.Header.FileSize ?? 0,
                    HeaderFlags = destListEntry.Lnk?.Header.DataFlags.ToString(),
                    LocalPath = destListEntry.Lnk?.LocalPath,
                    RelativePath = destListEntry.Lnk?.RelativePath,
                    Notes = string.Empty
                };

                if (debug)
                {
                    Log.Debug("CSOut values:");
                    csOut.PrintDump();
                }


                if (destListEntry.Lnk == null)
                {
                    csList.Add(csOut);
                    continue;
                }

                Log.Debug("Lnk file isn't null. Continuing");

                Log.Debug("Getting absolute path. TargetID count: {Count:N0}", destListEntry.Lnk.TargetIDs.Count);

                var target = GetAbsolutePathFromTargetIDs(destListEntry.Lnk.TargetIDs);

                Log.Debug("GetAbsolutePathFromTargetIDs Target is: {Target}", target);

                if (target.Length == 0)
                {
                    Log.Debug("Target length is 0. building alternate path");

                    if (destListEntry.Lnk.NetworkShareInfo != null)
                    {
                        target =
                            $"{destListEntry.Lnk.NetworkShareInfo.NetworkShareName}\\{destListEntry.Lnk.CommonPath}";
                    }
                    else
                    {
                        target =
                            $"{destListEntry.Lnk.LocalPath}\\{destListEntry.Lnk.CommonPath}";
                    }
                }

                csOut.TargetIDAbsolutePath = target;

                Log.Debug("Target is: {Target}", target);

                /*  if (destListEntry.Lnk.TargetIDs?.Count > 0)
                  {
                      csOut.TargetIDAbsolutePath = GetAbsolutePathFromTargetIDs(destListEntry.Lnk.TargetIDs);
                  }*/

                csOut.Arguments = string.Empty;
                if ((destListEntry.Lnk.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) ==
                    Lnk.Header.DataFlag.HasArguments)
                {
                    csOut.Arguments = destListEntry.Lnk.Arguments ?? string.Empty;
                }

                Log.Debug("csOut.Arguments is: {Arguments}", csOut.Arguments);

                csOut.WorkingDirectory = destListEntry.Lnk.WorkingDirectory;

                Log.Debug("csOut.WorkingDirectory is: {WorkingDirectory}", csOut.WorkingDirectory);

                var ebPresent = string.Empty;

                if (destListEntry.Lnk.ExtraBlocks.Count > 0)
                {
                    var names = new List<string>();

                    foreach (var extraDataBase in destListEntry.Lnk.ExtraBlocks)
                    {
                        names.Add(extraDataBase.GetType().Name);
                    }

                    ebPresent = string.Join(", ", names);
                }

                Log.Debug("csOut.ExtraBlocksPresent is: {EbPresent}", ebPresent);

                csOut.ExtraBlocksPresent = ebPresent;

                var tnb =
                    destListEntry.Lnk.ExtraBlocks.SingleOrDefault(t =>
                        t.GetType().Name.ToUpper() == "TRACKERDATABASEBLOCK");

                if (tnb != null)
                {
                    Log.Debug("Found tracker block");

                    var tnbBlock = tnb as TrackerDataBaseBlock;

                    csOut.TrackerCreatedOn =
                        tnbBlock?.CreationTime.ToString(dt);

                    csOut.MachineID = tnbBlock?.MachineId;
                    csOut.MachineMACAddress = tnbBlock?.MacAddress;
                }

                if (destListEntry.Lnk.TargetIDs?.Count > 0)
                {
                    Log.Debug("Target ID count: {Count:N0}", destListEntry.Lnk.TargetIDs.Count);

                    var si = destListEntry.Lnk.TargetIDs.Last();

                    if (TargetBlocks(si).Count > 0)
                    {
                        var eb = TargetBlocks(si).LastOrDefault(t => t is Beef0004);
                        if (eb is Beef0004)
                        {
                            var eb4 = eb as Beef0004;
                            if (eb4.MFTInformation.MFTEntryNumber != null)
                            {
                                csOut.TargetMFTEntryNumber =
                                    $"0x{eb4.MFTInformation.MFTEntryNumber.Value:X}";
                            }

                            if (eb4.MFTInformation.MFTSequenceNumber != null)
                            {
                                csOut.TargetMFTSequenceNumber =
                                    $"0x{eb4.MFTInformation.MFTSequenceNumber.Value:X}";
                            }
                        }
                    }
                }

                Log.Debug("Adding to csList");

                csList.Add(csOut);
            }

            if (wd)
            {
                foreach (var directoryEntry in auto.Directory)
                {
                    if (directoryEntry.DirectoryName.Equals("Root Entry") ||
                        directoryEntry.DirectoryName.Equals("DestList"))
                    {
                        continue;
                    }

                    if (auto.DestListEntries.Any(t => t.EntryNumber.ToString("X") == directoryEntry.DirectoryName))
                    {
                        continue;
                    }

                    //this directory entry is not in destlist

                    var f = auto.GetLnkFromDirectoryName(directoryEntry.DirectoryName);

                    if (f != null)
                    {
                        var csOut = new AutoCsvOut
                        {
                            User = user,
                            SourceFile = auto.SourceFile,
                            SourceCreated = ct.ToString(dt),
                            SourceModified = mt.ToString(dt),
                            SourceAccessed = at.ToString(dt),
                            AppId = auto.AppId.AppId,
                            AppIdDescription = auto.AppId.Description,
                            HasSps = auto.HasSps,
                            DestListVersion = auto.DestListVersion.ToString(),
                            LastUsedEntryNumber = auto.LastUsedEntryNumber.ToString(),
                            EntryNumber = directoryEntry.DirectoryName,
                            CreationTime =
                                directoryEntry.CreationTime?.Year == 1582
                                    ? string.Empty
                                    : directoryEntry.CreationTime?.ToString(
                                        dt),
                            LastModified =
                                directoryEntry.ModifiedTime?.ToString(dt),
                            Hostname = string.Empty,
                            MacAddress = string.Empty,
                            Path = string.Empty,
                            PinStatus = string.Empty,
                            FileBirthDroid = string.Empty,
                            FileDroid =
                                string.Empty,
                            VolumeBirthDroid = string.Empty,
                            VolumeDroid = string.Empty,
                            TargetCreated = f?.Header.TargetCreationDate.Year == 1601
                                ? string.Empty
                                : f?.Header.TargetCreationDate.ToString(
                                    dt),
                            TargetModified =
                                f?.Header.TargetModificationDate.Year == 1601
                                    ? string.Empty
                                    : f?.Header.TargetModificationDate.ToString(
                                        dt),
                            TargetAccessed =
                                f?.Header.TargetLastAccessedDate.Year == 1601
                                    ? string.Empty
                                    : f?.Header.TargetLastAccessedDate.ToString(
                                        dt),
                            CommonPath = f?.CommonPath,
                            VolumeLabel = f?.VolumeInfo?.VolumeLabel,
                            VolumeSerialNumber = f?.VolumeInfo?.VolumeSerialNumber,

                            DriveType =
                                f?.VolumeInfo == null
                                    ? "(None)"
                                    : GetDescriptionFromEnumValue(f?.VolumeInfo.DriveType),
                            FileAttributes = f?.Header.FileAttributes.ToString(),
                            FileSize = f?.Header.FileSize ?? 0,
                            HeaderFlags = f?.Header.DataFlags.ToString(),
                            LocalPath = f?.LocalPath,
                            RelativePath = f?.RelativePath,
                            Notes = "Found in Directory, not DestList"
                        };


                        /*  if (f.TargetIDs?.Count > 0)
                        {*/
                        var target = GetAbsolutePathFromTargetIDs(f.TargetIDs);
                        if (target.Length == 0)
                        {
                            target = $"{f.NetworkShareInfo?.NetworkShareName}\\\\{f.CommonPath}";
                        }

                        csOut.TargetIDAbsolutePath = target;
                        //}

                        csOut.Arguments = string.Empty;
                        if ((f.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) ==
                            Lnk.Header.DataFlag.HasArguments)
                        {
                            csOut.Arguments = f.Arguments ?? string.Empty;
                        }

                        csOut.WorkingDirectory = f.WorkingDirectory;

                        var ebPresent = string.Empty;

                        if (f.ExtraBlocks.Count > 0)
                        {
                            var names = new List<string>();

                            foreach (var extraDataBase in f.ExtraBlocks)
                            {
                                names.Add(extraDataBase.GetType().Name);
                            }

                            ebPresent = string.Join(", ", names);
                        }

                        csOut.ExtraBlocksPresent = ebPresent;

                        var tnb =
                            f.ExtraBlocks.SingleOrDefault(t => t.GetType().Name.ToUpper() == "TRACKERDATABASEBLOCK");

                        if (tnb != null)
                        {
                            var tnbBlock = tnb as TrackerDataBaseBlock;

                            csOut.TrackerCreatedOn =
                                tnbBlock?.CreationTime.ToString(dt);

                            csOut.MachineID = tnbBlock?.MachineId;
                            csOut.MachineMACAddress = tnbBlock?.MacAddress;
                        }

                        if (f.TargetIDs?.Count > 0)
                        {
                            var si = f.TargetIDs.Last();

                            if (TargetBlocks(si).Count > 0)
                            {
                                var eb = TargetBlocks(si).LastOrDefault(t => t is Beef0004);
                                if (eb is Beef0004)
                                {
                                    var eb4 = eb as Beef0004;
                                    if (eb4.MFTInformation.MFTEntryNumber != null)
                                    {
                                        csOut.TargetMFTEntryNumber =
                                            $"0x{eb4.MFTInformation.MFTEntryNumber.Value:X}";
                                    }

                                    if (eb4.MFTInformation.MFTSequenceNumber != null)
                                    {
                                        csOut.TargetMFTSequenceNumber =
                                            $"0x{eb4.MFTInformation.MFTSequenceNumber.Value:X}";
                                    }
                                }
                            }
                        }

                        csList.Add(csOut);
                    }
                }
            }


            return csList;
        }

        private const string JsonDateFormat = "yyyy-MM-ddTHH:mm:ss.fffffffZ";

        private static string JsonDate(DateTimeOffset d)
        {
            return d.UtcDateTime.ToString(JsonDateFormat, CultureInfo.InvariantCulture);
        }

        private static string JsonDate(DateTime d)
        {
            //timestamps without a kind are already UTC; ToUniversalTime would treat them as local time
            if (d.Kind == DateTimeKind.Unspecified)
            {
                d = DateTime.SpecifyKind(d, DateTimeKind.Utc);
            }

            return d.ToUniversalTime().ToString(JsonDateFormat, CultureInfo.InvariantCulture);
        }

        private static void DumpToJsonAuto(AutomaticDestination auto, bool pretty, string outFile, string user)
        {
            object target = user.IsNullOrEmpty() ? auto : new { User = user, JumpList = auto };
            if (pretty)
            {
                File.WriteAllText(outFile, target.Dump());
            }
            else
            {
                File.WriteAllText(outFile, target.ToJson());
            }
        }

        private static void DumpToJsonCustom(CustomDestination cust, bool pretty, string outFile, string user)
        {
            object target = user.IsNullOrEmpty() ? cust : new { User = user, JumpList = cust };
            if (pretty)
            {
                File.WriteAllText(outFile, target.Dump());
            }
            else
            {
                File.WriteAllText(outFile, target.ToJson());
            }
        }

        private static void SaveJsonCustom(CustomDestination cust, bool pretty, string outDir, string jsonf, string user)
        {
            try
            {
                if (Directory.Exists(outDir) == false)
                {
                    Directory.CreateDirectory(outDir);
                }

                var outName =
                    $"{ts:yyyyMMddHHmmss}_{Path.GetFileName(cust.SourceFile)}.json";
                var outFile = Path.Combine(outDir, outName);

                if (jsonf.IsNullOrEmpty() == false)
                {
                    outName =
                        $"{Path.GetFileNameWithoutExtension(jsonf)}_CustomDestinations{Path.GetExtension(jsonf)}";
                }

                outFile = Path.Combine(outDir, outName);

                DumpToJsonCustom(cust, pretty, outFile, user);
            }
            catch (Exception ex)
            {
                _exportErrors++;
                Log.Error(ex, "Error exporting json for {SourceFile}. Error: {Message}", cust.SourceFile, ex.Message);
            }
        }

        private static void SaveJsonAuto(AutomaticDestination auto, bool pretty, string outDir, string jsonf, string user)
        {
            try
            {
                if (Directory.Exists(outDir) == false)
                {
                    Directory.CreateDirectory(outDir);
                }

                var outName =
                    $"{ts:yyyyMMddHHmmss}_{Path.GetFileName(auto.SourceFile)}.json";
                var outFile = Path.Combine(outDir, outName);

                if (jsonf.IsNullOrEmpty() == false)
                {
                    outName =
                        $"{Path.GetFileNameWithoutExtension(jsonf)}_AutomaticDestinations{Path.GetExtension(jsonf)}";
                }

                outFile = Path.Combine(outDir, outName);

                DumpToJsonAuto(auto, pretty, outFile, user);
            }
            catch (Exception ex)
            {
                _exportErrors++;
                Log.Error(ex, "Error exporting json for {SourceFile}. Error: {Message}", auto.SourceFile, ex.Message);
            }
        }

        private static string GetDescriptionFromEnumValue(Enum value)
        {
            var attribute = value.GetType()
                .GetField(value.ToString())
                .GetCustomAttributes(typeof(DescriptionAttribute), false)
                .SingleOrDefault() as DescriptionAttribute;
            return attribute?.Description;
        }

        /// <summary>A file or folder item: its short name, modified time and extension blocks</summary>
        private static void LogFileEntry(string shortName, DateTimeOffset? modified, List<IExtensionBlock> blocks, bool blankLineAfterBlocks)
        {
            Log.Information("    Short name: {ShortName}", shortName);
            if (modified.HasValue)
            {
                Log.Information(
                    "    Modified:    {LastModificationTime}", modified);
            }
            else
            {
                Log.Information("    Modified:");
            }

            LogItemBlocks(blocks, blankLineAfterBlocks);
        }

        private static void LogItemBlocks(List<IExtensionBlock> blocks, bool blankLineAfterBlocks)
        {
            if (blocks == null || blocks.Count == 0)
            {
                return;
            }

            Log.Information("    Extension block count: {Count:N0}", blocks.Count);
            Console.WriteLine();

            var extensionNumber = 0;
            foreach (var extensionBlock in blocks)
            {
                Log.Information(
                    "    --------- Block {ExtensionNumber:N0} ({Name}) ---------", extensionNumber,
                    extensionBlock.GetType().Name);
                switch (extensionBlock)
                {
                    case Beef0004 b4:
                        Log.Information("    Long name: {LongName}", b4.LongName);
                        if (b4.LocalisedName?.Length > 0)
                        {
                            Log.Information("    Localized name: {LocalisedName}", b4.LocalisedName);
                        }

                        if (b4.CreatedOnTime.HasValue)
                        {
                            Log.Information(
                                "    Created:     {CreatedOnTime}", b4.CreatedOnTime);
                        }
                        else
                        {
                            Log.Information("    Created:");
                        }

                        if (b4.LastAccessTime.HasValue)
                        {
                            Log.Information(
                                "    Last access: {LastAccessTime}", b4.LastAccessTime);
                        }
                        else
                        {
                            Log.Information("    Last access: ");
                        }

                        if (b4.MFTInformation.MFTEntryNumber > 0)
                        {
                            Log.Information(
                                "    MFT entry/sequence #: {MftEntryNumber}/{MftSequenceNumber} (0x{MftEntryNumber2:X}/0x{MftSequenceNumber2:X})",
                                b4.MFTInformation.MFTEntryNumber, b4.MFTInformation.MFTSequenceNumber,
                                b4.MFTInformation.MFTEntryNumber, b4.MFTInformation.MFTSequenceNumber);
                        }

                        break;
                    case Beef0025 b25:
                        Log.Information(
                            "    Filetime 1: {FileTime1}, Filetime 2: {FileTime2}", b25.FileTime1.Value,
                            b25.FileTime2.Value);
                        break;
                    case Beef0003 b3:
                        Log.Information("    GUID: {Guid1} ({Guid1Folder})", b3.GUID1, b3.GUID1Folder);
                        break;
                    case Beef001a b1a:
                        Log.Information("    File document type: {FileDocumentTypeString}",
                            b1a.FileDocumentTypeString);
                        break;
                    default:
                        Log.Information("    {ExtensionBlock}", extensionBlock);
                        break;
                }

                extensionNumber += 1;
                if (blankLineAfterBlocks)
                {
                    Console.WriteLine();
                }
            }
        }

        private static void LogPropertyStore(PropertyStore store)
        {
            if (store == null || store.Sheets.Count == 0)
            {
                return;
            }

            Log.Information("  >> Property store (Format: GUID\\ID Description ==> Value)");
            var propCount = 0;

            foreach (var prop in store.Sheets)
            foreach (var propertyName in prop.PropertyNames)
            {
                propCount += 1;

                var prefix = $"{prop.GUID}\\{propertyName.Key}".PadRight(43);

                //(int.Parse threw on a named property)
                var suffix = int.TryParse(propertyName.Key, out var key)
                    ? $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, key)}".PadRight(35)
                    : string.Empty.PadRight(35);

                Log.Information("     {Prefix} {Suffix} ==> {PropertyName}", prefix, suffix,
                    propertyName.Value);
            }

            if (propCount == 0)
            {
                Log.Warning("     (Property store is empty)");
            }
        }

        /// <summary>
        ///     The target's extension blocks: a property view's (0x00) are those of the items in its Delegate ID List and then
        ///     its own, so the last 0xbeef0004 is the deepest item's (as when Lnk had its own shell item classes)
        /// </summary>
        private static List<IExtensionBlock> TargetBlocks(IShellBag item)
        {
            var own = item.ExtensionBlocks ?? new List<IExtensionBlock>();
            return item is ShellBag0X00 propertyView ? propertyView.DelegateExtensionBlocks.Concat(own).ToList() : own;
        }

        private static string GetAbsolutePathFromTargetIDs(List<IShellBag> ids)
        {
            if (ids == null)
            {
                return "(No target IDs present)";
            }

            var absPath = string.Empty;

            if (ids.Count == 0)
            {
                return absPath;
            }

            foreach (var shellBag in ids)
            {
                absPath += shellBag.Value + @"\";
            }

            absPath = absPath.Substring(0, absPath.Length - 1);

            return absPath;
        }

        private static AutomaticDestination ProcessAutoFile(string jlFile, bool q, string dt, bool fd, bool ld, bool wd,
            int codepage = 1252)
        {
            if (q == false)
            {
                Log.Information("Processing {File}", jlFile);
                Console.WriteLine();
            }

            var sw = new Stopwatch();
            sw.Start();

            var distListCountAdjust = 2;

            try
            {
                Log.Debug("Opening {File}", jlFile);

                var autoDest = JumpList.JumpList.LoadAutoJumplist(jlFile, codepage);

                Log.Debug("Opened {File}", jlFile);


                if (q == false)
                {
                    Log.Information("Source file: {SourceFile}", autoDest.SourceFile);

                    Console.WriteLine();

                    Log.Information("--- AppId information ---");
                    Log.Information("  AppID: {AppId}", autoDest.AppId.AppId);
                    Log.Information("  Description: {Description}", autoDest.AppId.Description);
                    Console.WriteLine();

                    Log.Information("--- DestList information ---");
                    Log.Information("  Expected DestList entries:  {DestListCount:N0}", autoDest.DestListCount);
                    Log.Information("  Actual DestList entries:    {DestListCount:N0}", autoDest.DestListCount);
                    Log.Information("  DestList version:           {DestListVersion}", autoDest.DestListVersion);


                    if (autoDest.DestListPropertyStore != null && autoDest.EmptyDestListPropertyStore == false)
                    {
                        distListCountAdjust = 3;
                    }

                    if (autoDest.DestListCount > 0 &&
                        autoDest.DestListCount != autoDest.Directory.Count - distListCountAdjust)
                    {
                        Console.WriteLine();
                        Log.Warning(
                            "  There are more items in the Directory ({DirectoryCount:N0}) than are contained in the DestList ({DestListCount:N0}). Use {Switch} to view/export them",
                            autoDest.Directory.Count - 2, autoDest.DestListCount, "--withDir");
                    }

                    Console.WriteLine();

                    if (autoDest.DestListPropertyStore != null)
                    {
                        Log.Information("   --- DestList Property Store information ---");
                        foreach (var property in autoDest.DestListPropertyStore.PropertyNames)
                        {
                            Log.Information("    Property {Name}: {Value}", property.Key, property.Value);
                        }
                    }

                    Console.WriteLine();


                    Log.Information("--- DestList entries ---");
                    foreach (var autoDestList in autoDest.DestListEntries)
                    {
                        Log.Information("Entry #: {EntryNumber}", autoDestList.EntryNumber);
                        Log.Information("  MRU: {MRUPosition}", autoDestList.MRUPosition);
                        Log.Information("  Path: {Path}", autoDestList.Path);
                        Log.Information("  Pinned: {Pinned}", autoDestList.Pinned);
                        Log.Information(
                            "  Created on:    {CreatedOn}", autoDestList.CreatedOn);
                        Log.Information(
                            "  Last modified: {LastModified}", autoDestList.LastModified);
                        Log.Information("  Hostname: {Hostname}", autoDestList.Hostname);
                        Log.Information(
                            "  Mac Address: {Mac}",
                            (autoDestList.MacAddress == "00:00:00:00:00:00" ? string.Empty : autoDestList.MacAddress));
                        Log.Information(
                            "  Interaction count: {InteractionCount:N0}", autoDestList.InteractionCount);

                        Console.WriteLine();
                        Log.Information("--- Lnk information ---");

                        if (fd)
                        {
                            DateTimeOffset? tc = autoDestList.Lnk.Header.TargetCreationDate.Year == 1601
                                ? null
                                : autoDestList.Lnk.Header.TargetCreationDate;
                            DateTimeOffset? tm = autoDestList.Lnk.Header.TargetModificationDate.Year == 1601
                                ? null
                                : autoDestList.Lnk.Header.TargetModificationDate;
                            DateTimeOffset? ta = autoDestList.Lnk.Header.TargetLastAccessedDate.Year == 1601
                                ? null
                                : autoDestList.Lnk.Header.TargetLastAccessedDate;


                            Log.Information("  Lnk target created:  {Tc}", tc);
                            Log.Information("  Lnk target modified: {Tm}", tm);
                            Log.Information("  Lnk target accessed: {Ta}", ta);

                            Console.WriteLine();

                            DumpLnkFile(autoDestList.Lnk, dt);
                        }
                        else if (ld)
                        {
                            DumpLnkDetail(autoDestList.Lnk);
                        }


                        if (!fd &&
                            autoDestList.Lnk?.TargetIDs.Count > 0)
                        {
                            var target = GetAbsolutePathFromTargetIDs(autoDestList.Lnk.TargetIDs);

                            if (target.Length == 0)
                            {
                                target =
                                    $"{autoDestList.Lnk.NetworkShareInfo.NetworkShareName}\\\\{autoDestList.Lnk.CommonPath}";
                            }

                            Log.Information("  Absolute path: {Target}", target);
                            Console.WriteLine();
                        }
                        else
                        {
                            Log.Warning("   (lnk file not present)");
                            Console.WriteLine();
                        }


                        Console.WriteLine();

                        if (autoDestList.Sps != null)
                        {
                            Log.Information("   --- Serialized Property Store information ---");
                            foreach (var property in autoDestList.Sps.PropertyNames)
                            {
                                Log.Information("    Property {Name}: {Value}", property.Key, property.Value);
                            }
                        }
                    }

                    if (wd)
                    {
                        Log.Information("{Dir} entries not represented by {Dest} entries", "Directory", "DestList");

                        foreach (var directoryEntry in autoDest.Directory)
                        {
                            if (directoryEntry.DirectoryName.Equals("Root Entry") ||
                                directoryEntry.DirectoryName.Equals("DestList"))
                            {
                                continue;
                            }

                            if (autoDest.DestListEntries.Any(t =>
                                    t.EntryNumber.ToString("X") == directoryEntry.DirectoryName))
                            {
                                continue;
                            }

                            //this directory entry is not in destlist
                            if (!q)
                            {
                                Log.Information("Directory Name: {DirectoryName}", directoryEntry.DirectoryName);
                            }

                            var f = autoDest.GetLnkFromDirectoryName(directoryEntry.DirectoryName);

                            if (f != null)
                            {
                                if (fd)
                                {
                                    DateTimeOffset? tc = f.Header.TargetCreationDate.Year == 1601
                                        ? null
                                        : f.Header.TargetCreationDate;
                                    DateTimeOffset? tm = f.Header.TargetModificationDate.Year == 1601
                                        ? null
                                        : f.Header.TargetModificationDate;
                                    DateTimeOffset? ta = f.Header.TargetLastAccessedDate.Year == 1601
                                        ? null
                                        : f.Header.TargetLastAccessedDate;


                                    Log.Information("  Lnk target created:  {Tc}", tc);
                                    Log.Information("  Lnk target modified: {Tm}", tm);
                                    Log.Information("  Lnk target accessed: {Ta}", ta);
                                    Console.WriteLine();

                                    DumpLnkFile(f, dt);
                                }
                                else if (ld)
                                {
                                    DumpLnkDetail(f);
                                }

                                if (!fd)
                                {
                                    Console.WriteLine();

                                    var target = GetAbsolutePathFromTargetIDs(f.TargetIDs);

                                    if (target.Length == 0)
                                    {
                                        target = $"{f.NetworkShareInfo?.NetworkShareName}\\\\{f.CommonPath}";
                                    }

                                    Log.Information("  Absolute path: {Target}", target);
                                    Console.WriteLine();
                                }
                            }
                            else
                            {
                                Log.Debug("  No lnk file found for directory entry {DirectoryName}",
                                    directoryEntry.DirectoryName);
                            }
                        }
                    }
                }

                sw.Stop();

                if (q == false)
                {
                    Console.WriteLine();
                }

                if (autoDest.DestListPropertyStore != null && autoDest.EmptyDestListPropertyStore == false)
                {
                    distListCountAdjust = 3;
                }

                if (autoDest.DestListCount > 0 &&
                    autoDest.DestListCount != autoDest.Directory.Count - distListCountAdjust)
                {
                    Console.WriteLine();
                    Log.Warning(
                        "  There are more items in the Directory ({DirectoryCount:N0}) than are contained in the DestList ({DestListCount:N0}). Use {Switch} to view/export them",
                        autoDest.Directory.Count - 2, autoDest.DestListCount, "--withDir");
                }

                if (autoDest.HasSps || autoDest.EmptyDestListPropertyStore == false)
                {
                    Log.Warning("** {Warn} **",
                        "JumpList has serialized property store(s)! View its contents via -f for details");
                    Console.WriteLine();
                }

                Log.Information(
                    "---------- Processed {SourceFile} in {TotalSeconds:N8} seconds ----------", autoDest.SourceFile,
                    sw.Elapsed.TotalSeconds);

                if (q == false)
                {
                    Console.WriteLine();
                }

                return autoDest;
            }

            catch (Exception ex)
            {
                _failedFiles.Add($"{jlFile} ==> ({ex.Message})");
                Log.Fatal(ex, "Error opening {File}. Message: {Message}", jlFile, ex.Message);
                Console.WriteLine();

                if (IsReadError(ex))
                {
                    _readErrors++;
                }
            }

            return null;
        }

        /// <summary>
        ///     True when a file could not be opened or read (access denied, locked, an IO error), as opposed to a file that
        ///     was read but is damaged
        /// </summary>
        private static bool IsReadError(Exception ex)
        {
            return ex is UnauthorizedAccessException || (ex is IOException && !(ex is EndOfStreamException));
        }

#if !NET6_0_OR_GREATER
        /// <summary>
        ///     AlphaFS ErrorFilter: a folder that can't be read is reported and counted as a read error (exit 1 at the end),
        ///     and the search goes on (it was skipped silently).
        ///     The options leave out ContinueOnException: with it, AlphaFS never calls ErrorFilter
        /// </summary>
        private static bool UnreadableFolder(int errorCode, string errorMessage, string pathProcessed)
        {
            Log.Error("Unable to read folder {Folder}: {Reason}", pathProcessed, errorMessage);
            _readErrors++;
            return true;
        }
#else
        /// <summary>
        ///     Entries under a folder, recursively, matching a pattern (case-insensitive), as
        ///     Directory.EnumerateFileSystemEntries did with IgnoreInaccessible. Folders that can't be read are skipped and kept
        ///     in Unreadable (IgnoreInaccessible skipped them silently)
        /// </summary>
        private sealed class EntryFinder : System.IO.Enumeration.FileSystemEnumerator<string>
        {
            private readonly string _pattern;

            //the subfolder whose handle is opened right after ShouldRecurseIntoEntry; an error then is about that folder
            private string _lastRecursed;

            //folder (null: the folder searched, as the base constructor can report it before this one runs) and reason
            public readonly List<(string Folder, string Reason)> Unreadable = new();

            public EntryFinder(string directory, string pattern) : base(directory,
                new System.IO.EnumerationOptions { IgnoreInaccessible = false, RecurseSubdirectories = true, AttributesToSkip = 0 })
            {
                _pattern = pattern;
            }

            protected override bool ShouldRecurseIntoEntry(ref System.IO.Enumeration.FileSystemEntry entry)
            {
                _lastRecursed = entry.ToSpecifiedFullPath();
                return true;
            }

            protected override bool ShouldIncludeEntry(ref System.IO.Enumeration.FileSystemEntry entry)
            {
                _lastRecursed = null;
                return System.IO.Enumeration.FileSystemName.MatchesSimpleExpression(_pattern, entry.FileName, true);
            }

            protected override string TransformEntry(ref System.IO.Enumeration.FileSystemEntry entry)
            {
                return entry.ToSpecifiedFullPath();
            }

            protected override bool ContinueOnError(int error)
            {
                Unreadable.Add((_lastRecursed, new Win32Exception(error).Message));
                _lastRecursed = null;
                return true;
            }
        }

        /// <summary>
        ///     The entries found; each folder that can't be read is reported and counted as a read error (exit 1 at the end)
        /// </summary>
        private static List<string> FindEntries(string directory, string pattern)
        {
            using var finder = new EntryFinder(directory, pattern);
            var entries = new List<string>();
            while (finder.MoveNext())
            {
                entries.Add(finder.Current);
            }

            foreach (var (folder, reason) in finder.Unreadable)
            {
                Log.Error("Unable to read folder {Folder}: {Reason}", folder ?? directory, reason);
                _readErrors++;
            }

            return entries;
        }
#endif

        private static void DumpLnkDetail(LnkFile lnk)
        {
            if (lnk == null)
            {
                Log.Warning("(lnk file not present)");
                return;
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasName) == Lnk.Header.DataFlag.HasName)
            {
                Log.Information("  Name: {Name}", lnk.Name);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasRelativePath) ==
                Lnk.Header.DataFlag.HasRelativePath)
            {
                Log.Information("  Relative Path: {RelativePath}", lnk.RelativePath);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasWorkingDir) ==
                Lnk.Header.DataFlag.HasWorkingDir)
            {
                Log.Information("  Working Directory: {WorkingDirectory}", lnk.WorkingDirectory);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) ==
                Lnk.Header.DataFlag.HasArguments)
            {
                Log.Information("  Arguments: {Arguments}", lnk.Arguments);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasLinkInfo) ==
                Lnk.Header.DataFlag.HasLinkInfo)
            {
                Console.WriteLine();
                Log.Information("--- Link information ---");
                Log.Information("Flags: {LocationFlags}", lnk.LocationFlags);

                if (lnk.VolumeInfo != null)
                {
                    Console.WriteLine();
                    Log.Information(">> Volume information");
                    Log.Information(
                        "  Drive type: {Desc}", GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType));
                    Log.Information("  Serial number: {VolumeSerialNumber}", lnk.VolumeInfo.VolumeSerialNumber);

                    var label = lnk.VolumeInfo.VolumeLabel.Length > 0
                        ? lnk.VolumeInfo.VolumeLabel
                        : "(No label)";

                    Log.Information("  Label: {Label}", label);
                }

                if (lnk.NetworkShareInfo != null)
                {
                    Console.WriteLine();
                    Log.Information("  Network share information");

                    if (lnk.NetworkShareInfo.DeviceName.Length > 0)
                    {
                        Log.Information("    Device name: {NetworkShareInfoDeviceName}",
                            lnk.NetworkShareInfo.DeviceName);
                    }

                    Log.Information("    Share name: {NetworkShareInfoNetworkShareName}",
                        lnk.NetworkShareInfo.NetworkShareName);

                    Log.Information(
                        "    Provider type: {NetworkShareInfoNetworkProviderType}",
                        lnk.NetworkShareInfo.NetworkProviderType);
                    Log.Information("    Share flags: {NetworkShareInfoShareFlags}", lnk.NetworkShareInfo.ShareFlags);
                    Console.WriteLine();
                }

                if (lnk.LocalPath?.Length > 0)
                {
                    Log.Information("  Local path: {LocalPath}", lnk.LocalPath);
                }

                if (lnk.CommonPath.Length > 0)
                {
                    Log.Information("  Common path: {CommonPath}", lnk.CommonPath);
                }
            }
        }

        private static void DumpLnkFile(LnkFile lnk, string dt)
        {
            if (lnk == null)
            {
                Log.Warning("(lnk file not present)");
                return;
            }

            Log.Information("--- Header ---");
            Console.WriteLine();

            DateTimeOffset? tc1 = lnk.Header.TargetCreationDate.Year == 1601
                ? null
                : lnk.Header.TargetCreationDate;
            DateTimeOffset? tm1 = lnk.Header.TargetModificationDate.Year == 1601
                ? null
                : lnk.Header.TargetModificationDate;
            DateTimeOffset? ta1 = lnk.Header.TargetLastAccessedDate.Year == 1601
                ? null
                : lnk.Header.TargetLastAccessedDate;

            Log.Information("  Target created:  {Tc1}", tc1);
            Log.Information("  Target modified: {Tm1}", tm1);
            Log.Information("  Target accessed: {Ta1}", ta1);
            Console.WriteLine();
            Log.Information("  File size: {FileSize:N0}", lnk.Header.FileSize);
            Log.Information("  Flags: {DataFlags}", lnk.Header.DataFlags);
            Log.Information("  File attributes: {FileAttributes}", lnk.Header.FileAttributes);

            if (lnk.Header.HotKey.Length > 0)
            {
                Log.Information("  Hot key: {HotKey}", lnk.Header.HotKey);
            }

            Log.Information("  Icon index: {IconIndex}", lnk.Header.IconIndex);
            Log.Information(
                "  Show window: {ShowWindow} ({Desc})", lnk.Header.ShowWindow,
                GetDescriptionFromEnumValue(lnk.Header.ShowWindow));

            Console.WriteLine();

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasName) == Lnk.Header.DataFlag.HasName)
            {
                Log.Information("Name: {Name}", lnk.Name);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasRelativePath) == Lnk.Header.DataFlag.HasRelativePath)
            {
                Log.Information("Relative Path: {RelativePath}", lnk.RelativePath);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasWorkingDir) == Lnk.Header.DataFlag.HasWorkingDir)
            {
                Log.Information("Working Directory: {WorkingDirectory}", lnk.WorkingDirectory);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) == Lnk.Header.DataFlag.HasArguments)
            {
                Log.Information("Arguments: {Arguments}", lnk.Arguments);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasIconLocation) == Lnk.Header.DataFlag.HasIconLocation)
            {
                Log.Information("Icon Location: {IconLocation}", lnk.IconLocation);
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasLinkInfo) == Lnk.Header.DataFlag.HasLinkInfo)
            {
                Console.WriteLine();
                Log.Information("--- Link information ---");
                Log.Information("Flags: {LocationFlags}", lnk.LocationFlags);

                if (lnk.VolumeInfo != null)
                {
                    Console.WriteLine();
                    Log.Information(">> Volume information");
                    Log.Information("  Drive type: {Desc}", GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType));
                    Log.Information("  Serial number: {VolumeSerialNumber}", lnk.VolumeInfo.VolumeSerialNumber);

                    var label = lnk.VolumeInfo.VolumeLabel.Length > 0
                        ? lnk.VolumeInfo.VolumeLabel
                        : "(No label)";

                    Log.Information("  Label: {Label}", label);
                }

                if (lnk.NetworkShareInfo != null)
                {
                    Console.WriteLine();
                    Log.Information("  Network share information");

                    if (lnk.NetworkShareInfo.DeviceName.Length > 0)
                    {
                        Log.Information("    Device name: {DeviceName}", lnk.NetworkShareInfo.DeviceName);
                    }

                    Log.Information("    Share name: {NetworkShareName}", lnk.NetworkShareInfo.NetworkShareName);

                    Log.Information("    Provider type: {NetworkProviderType}",
                        lnk.NetworkShareInfo.NetworkProviderType);
                    Log.Information("    Share flags: {ShareFlags}", lnk.NetworkShareInfo.ShareFlags);
                    Console.WriteLine();
                }

                if (lnk.LocalPath?.Length > 0)
                {
                    Log.Information("  Local path: {LocalPath}", lnk.LocalPath);
                }

                if (lnk.CommonPath.Length > 0)
                {
                    Log.Information("  Common path: {CommonPath}", lnk.CommonPath);
                }
            }

            if (lnk.TargetIDs.Count > 0)
            {
                Console.WriteLine();

                // var absPath = string.Empty;
                //
                // foreach (var shellBag in lnk.TargetIDs)
                // {
                //     absPath += shellBag.Value + @"\";
                // }

                Log.Information("--- Target ID information (Format: Type ==> Value) ---");
                Console.WriteLine();
                Log.Information("  Absolute path: {Abs}", GetAbsolutePathFromTargetIDs(lnk.TargetIDs));
                Console.WriteLine();

                foreach (var shellBag in lnk.TargetIDs)
                {
                    //HACK
                    //This is a total hack until i can refactor some shellbag code to clean things up

                    var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;

                    Log.Information("  -{FriendlyName} ==> {Val}", shellBag.FriendlyName, val);

                    //the shell items are ExtensionBlocks' classes (shared with ShellBags Explorer, SBECmd and the RECmd plugins)
                    switch (shellBag)
                    {
                        //files and folders: 0x32 and 0x36 (a blank line after each block), 0x31 (and 0x35, 0x39, 0xB1, read
                        //as one) and 0x3A file entries
                        case ShellBag0X32 b:
                            LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks, true);
                            break;
                        case ShellBag0X36 b:
                            LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks, true);
                            break;
                        case ShellBag0X31 b:
                            LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks, false);
                            break;
                        case ShellBag0X35 b:
                            LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks, false);
                            break;
                        case ShellBag0X39 b:
                            LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks, false);
                            break;
                        case ShellBag0XB1 b:
                            LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks, false);
                            break;
                        case ShellBag0X3A b when b.ShortName != null:
                            LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks, false);
                            break;

                        //users files folder (a 0x74 item with CFSF)
                        case ShellBagCFSF b:
                            if (b.LastModificationTime.HasValue)
                            {
                                Log.Information(
                                    "    Modified:    {LastModificationTime}", b.LastModificationTime);
                            }
                            else
                            {
                                Log.Information("    Modified:");
                            }

                            LogItemBlocks(b.ExtensionBlocks, false);
                            break;

                        case ShellBag0X00 b:
                            foreach (var store in b.PropertyStores)
                            {
                                LogPropertyStore(store);
                            }

                            break;
                        case ShellBag0X1F b:
                            LogPropertyStore(b.PropertyStore);
                            break;
                        case ShellBag0X01 b:
                            if (b.DriveLetter?.Length > 0)
                            {
                                Log.Information("  Drive letter: {DriveLetter}", b.DriveLetter);
                            }

                            break;
                        case ShellBag0X71 b:
                            if (b.PropertyStore?.Sheets.Count > 0)
                            {
                                Log.Warning(
                                    "Property stores found! Please email lnk file to {Email} so support can be added!!",
                                    "saericzimmerman@gmail.com");
                            }

                            break;

                        //the value above is all there is
                        case ShellBagDamaged:
                        case ShellBag0X2E:
                        case ShellBag0X2F:
                        case ShellBag0X2A:
                        case ShellBag0X3A:
                        case ShellBag0X41:
                        case ShellBag0X42:
                        case ShellBag0X43:
                        case ShellBag0X46:
                        case ShellBag0X47:
                        case ShellBag0X61:
                        case ShellBag0X23:
                        case ShellBag0XC3:
                        case ShellBag0X74:
                        case ShellBag0X79:
                        case ShellBag0X4C:
                        case ShellBagZipContents:
                            break;

                        default:
                            Log.Warning(">> UNMAPPED Type! Please email lnk file to {Email} so support can be added!",
                                "saericzimmerman@gmail.com");
                            Log.Warning(">>{ShellBag}", shellBag);
                            break;
                    }
                    Console.WriteLine();
                }

                Log.Information("--- End Target ID information ---");
            }


            if (lnk.ExtraBlocks.Count > 0)
            {
                Console.WriteLine();
                Log.Information("--- Extra blocks information ---");
                Console.WriteLine();

                foreach (var extraDataBase in lnk.ExtraBlocks)
                {
                    switch (extraDataBase.GetType().Name)
                    {
                        case "ConsoleDataBlock":
                            var cdb = extraDataBase as ConsoleDataBlock;
                            Log.Information(">> Console data block");
                            Log.Information("   Fill Attributes: {FillAttributes}", cdb.FillAttributes);
                            Log.Information("   Popup Attributes: {PopupFillAttributes}", cdb.PopupFillAttributes);
                            Log.Information(
                                "   Buffer Size (Width x Height): {ScreenWidthBufferSize} x {cdb.ScreenHeightBufferSize}",
                                cdb.ScreenWidthBufferSize, cdb.ScreenHeightBufferSize);
                            Log.Information(
                                "   Window Size (Width x Height): {WindowWidth} x {cdb.WindowHeight}", cdb.WindowWidth,
                                cdb.WindowHeight);
                            Log.Information("   Origin (X/Y): {WindowOriginX}/{cdb.WindowOriginY}", cdb.WindowOriginX,
                                cdb.WindowOriginY);
                            Log.Information("   Font Size: {FontSize}", cdb.FontSize);
                            Log.Information("   Is Bold: {IsBold}", cdb.IsBold);
                            Log.Information("   Face Name: {FaceName}", cdb.FaceName);
                            Log.Information("   Cursor Size: {CursorSize}", cdb.CursorSize);
                            Log.Information("   Is Full Screen: {IsFullScreen}", cdb.IsFullScreen);
                            Log.Information("   Is Quick Edit: {IsQuickEdit}", cdb.IsQuickEdit);
                            Log.Information("   Is Insert Mode: {IsInsertMode}", cdb.IsInsertMode);
                            Log.Information("   Is Auto Positioned: {IsAutoPositioned}", cdb.IsAutoPositioned);
                            Log.Information("   History Buffer Size: {HistoryBufferSize}", cdb.HistoryBufferSize);
                            Log.Information("   History Buffer Count: {HistoryBufferCount}", cdb.HistoryBufferCount);
                            Log.Information("   History Duplicates Allowed: {HistoryDuplicatesAllowed}",
                                cdb.HistoryDuplicatesAllowed);
                            Console.WriteLine();
                            break;
                        case "ConsoleFEDataBlock":
                            var cfedb = extraDataBase as ConsoleFeDataBlock;
                            Log.Information(">> Console FE data block");
                            Log.Information("   Code page: {CodePage}", cfedb.CodePage);
                            Console.WriteLine();
                            break;
                        case "DarwinDataBlock":
                            var ddb = extraDataBase as DarwinDataBlock;
                            Log.Information(">> Darwin data block");
                            Log.Information("   Application ID: {ApplicationIdentifierUnicode}",
                                ddb.ApplicationIdentifierUnicode);
                            Console.WriteLine();
                            break;
                        case "EnvironmentVariableDataBlock":
                            var evdb = extraDataBase as EnvironmentVariableDataBlock;
                            Log.Information(">> Environment variable data block");
                            Log.Information("   Environment variables: {EnvironmentVariablesUnicode}",
                                evdb.EnvironmentVariablesUnicode);
                            Console.WriteLine();
                            break;
                        case "IconEnvironmentDataBlock":
                            var iedb = extraDataBase as IconEnvironmentDataBlock;
                            Log.Information(">> Icon environment data block");
                            Log.Information("   Icon path: {IconPathUni}", iedb.IconPathUni);
                            Console.WriteLine();
                            break;
                        case "KnownFolderDataBlock":
                            var kfdb = extraDataBase as KnownFolderDataBlock;
                            Log.Information(">> Known folder data block");
                            Log.Information(
                                "   Known folder GUID: {KnownFolderId} ==> {KnownFolderName}", kfdb.KnownFolderId,
                                kfdb.KnownFolderName);
                            Console.WriteLine();
                            break;
                        case "PropertyStoreDataBlock":
                            var psdb = extraDataBase as PropertyStoreDataBlock;

                            if (psdb.PropertyStore.Sheets.Count > 0)
                            {
                                Log.Information(
                                    ">> Property store data block (Format: GUID\\ID Description ==> Value)");
                                var propCount = 0;

                                foreach (var prop in psdb.PropertyStore.Sheets)
                                foreach (var propertyName in prop.PropertyNames)
                                {
                                    propCount += 1;

                                    var prefix = $"{prop.GUID}\\{propertyName.Key}".PadRight(43);
                                    var suffix =
                                        $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
                                            .PadRight(35);

                                    Log.Information("   {Prefix} {Suffix} ==> {Value}", prefix, suffix,
                                        propertyName.Value);
                                }

                                if (propCount == 0)
                                {
                                    Log.Information("   (Property store is empty)");
                                }
                            }

                            Console.WriteLine();
                            break;
                        case "ShimDataBlock":
                            var sdb = extraDataBase as ShimDataBlock;
                            Log.Information(">> Shimcache data block");
                            Log.Information("   LayerName: {LayerName}", sdb.LayerName);
                            Console.WriteLine();
                            break;
                        case "SpecialFolderDataBlock":
                            var sfdb = extraDataBase as SpecialFolderDataBlock;
                            Log.Information(">> Special folder data block");
                            Log.Information("   Special Folder ID: {SpecialFolderId}", sfdb.SpecialFolderId);
                            Console.WriteLine();
                            break;
                        case "TrackerDataBaseBlock":
                            var tdb = extraDataBase as TrackerDataBaseBlock;
                            Log.Information(">> Tracker database block");
                            Log.Information("   Machine ID:  {MachineId}", tdb.MachineId);
                            Log.Information("   MAC Address: {MacAddress}", tdb.MacAddress);
                            Log.Information("   MAC Vendor:  {Mac}", GetVendorFromMac(tdb.MacAddress));
                            Log.Information(
                                "   Creation:    {CreationTime}", tdb.CreationTime);
                            Console.WriteLine();
                            Log.Information("   Volume Droid:       {VolumeDroid}", tdb.VolumeDroid);
                            Log.Information("   Volume Droid Birth: {VolumeDroidBirth}", tdb.VolumeDroidBirth);
                            Log.Information("   File Droid:         {FileDroid}", tdb.FileDroid);
                            Log.Information("   File Droid birth:   {FileDroidBirth}", tdb.FileDroidBirth);
                            Console.WriteLine();
                            break;
                        case "VistaAndAboveIdListDataBlock":
                            var vdb = extraDataBase as VistaAndAboveIdListDataBlock;
                            Log.Information(">> Vista and above ID List data block");

                            foreach (var shellBag in vdb.TargetIDs)
                            {
                                var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;
                                Log.Information("   {FriendlyName} ==> {Val}", shellBag.FriendlyName, val);
                            }

                            Console.WriteLine();
                            break;
                    }
                }
            }
        }

        private static string GetVendorFromMac(string macAddress)
        {
            //00-00-00	XEROX CORPORATION
            //"00:14:22:0d:94:04"

            var mac = string.Join("-", macAddress.Split(':').Take(3)).ToUpperInvariant();
            // .Replace(":", "-").ToUpper();

            var vendor = "(Unknown vendor)";

            if (MacList.ContainsKey(mac))
            {
                vendor = MacList[mac];
            }

            return vendor;
        }

        private static CustomDestination ProcessCustomFile(string jlFile, bool q, string dt, bool ld,
            int codepage = 1252)
        {
            if (q == false)
            {
                Log.Information("Processing {File}", jlFile);
                Console.WriteLine();
            }

            var sw = new Stopwatch();
            sw.Start();

            try
            {
                var customDest = JumpList.JumpList.LoadCustomJumplist(jlFile, codepage);

                if (q == false)
                {
                    Log.Error("Source file: {SourceFile}", customDest.SourceFile);

                    Console.WriteLine();

                    Log.Information("--- AppId information ---");
                    Log.Information("AppID: {AppId}, Description: {Description}", customDest.AppId.AppId,
                        customDest.AppId.Description);
                    Log.Information("--- DestList information ---");
                    Log.Information("  Entries:  {Count:N0}", customDest.Entries.Count);
                    Console.WriteLine();

                    var entryNum = 0;
                    foreach (var entry in customDest.Entries)
                    {
                        Log.Information("  Entry #: {EntryNum}, lnk count: {Count:N0} Rank: {Rank:G5}", entryNum,
                            entry.LnkFiles.Count, entry.Rank);

                        if (entry.Name.Length > 0)
                        {
                            Log.Information("   Name: {Name}", entry.Name);
                        }

                        Console.WriteLine();

                        var lnkCounter = 0;

                        foreach (var lnkFile in entry.LnkFiles)
                        {
                            DateTimeOffset? tc = lnkFile.Header.TargetCreationDate.Year == 1601
                                ? null
                                : lnkFile.Header.TargetCreationDate;
                            DateTimeOffset? tm = lnkFile.Header.TargetModificationDate.Year == 1601
                                ? null
                                : lnkFile.Header.TargetModificationDate;
                            DateTimeOffset? ta = lnkFile.Header.TargetLastAccessedDate.Year == 1601
                                ? null
                                : lnkFile.Header.TargetLastAccessedDate;


                            Log.Information("--- Lnk #{LnkCounter:N0} information ---", lnkCounter);
                            Log.Information("  Lnk target created:  {Tc}", tc);
                            Log.Information("  Lnk target modified: {Tm}", tm);
                            Log.Information("  Lnk target accessed: {Ta}", ta);

                            if (ld)
                            {
                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasName) ==
                                    Lnk.Header.DataFlag.HasName)
                                {
                                    Log.Information("  Name: {Name}", lnkFile.Name);
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasRelativePath) ==
                                    Lnk.Header.DataFlag.HasRelativePath)
                                {
                                    Log.Information("  Relative Path: {RelativePath}", lnkFile.RelativePath);
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasWorkingDir) ==
                                    Lnk.Header.DataFlag.HasWorkingDir)
                                {
                                    Log.Information("  Working Directory: {WorkingDirectory}",
                                        lnkFile.WorkingDirectory);
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) ==
                                    Lnk.Header.DataFlag.HasArguments)
                                {
                                    Log.Information("  Arguments: {Arguments}", lnkFile.Arguments);
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasLinkInfo) ==
                                    Lnk.Header.DataFlag.HasLinkInfo)
                                {
                                    Console.WriteLine();
                                    Log.Information("--- Link information ---");
                                    Console.WriteLine();
                                    Log.Information("Flags: {LocationFlags}", lnkFile.LocationFlags);

                                    if (lnkFile.VolumeInfo != null)
                                    {
                                        Console.WriteLine();
                                        Log.Information(">> Volume information");
                                        Log.Information(
                                            "  Drive type: {Desc}",
                                            GetDescriptionFromEnumValue(lnkFile.VolumeInfo.DriveType));
                                        Log.Information("  Serial number: {VolumeSerialNumber}",
                                            lnkFile.VolumeInfo.VolumeSerialNumber);

                                        var label = lnkFile.VolumeInfo.VolumeLabel.Length > 0
                                            ? lnkFile.VolumeInfo.VolumeLabel
                                            : "(No label)";

                                        Log.Information("  Label: {Label}", label);
                                    }

                                    if (lnkFile.NetworkShareInfo != null)
                                    {
                                        Console.WriteLine();
                                        Log.Information("  Network share information");

                                        if (lnkFile.NetworkShareInfo.DeviceName.Length > 0)
                                        {
                                            Log.Information("    Device name: {DeviceName}",
                                                lnkFile.NetworkShareInfo.DeviceName);
                                        }

                                        Log.Information("    Share name: {NetworkShareName}",
                                            lnkFile.NetworkShareInfo.NetworkShareName);

                                        Log.Information(
                                            "    Provider type: {NetworkProviderType}",
                                            lnkFile.NetworkShareInfo.NetworkProviderType);
                                        Log.Information("    Share flags: {ShareFlags}",
                                            lnkFile.NetworkShareInfo.ShareFlags);
                                        Console.WriteLine();
                                    }

                                    if (lnkFile.LocalPath?.Length > 0)
                                    {
                                        Log.Information("  Local path: {LocalPath}", lnkFile.LocalPath);
                                    }

                                    if (lnkFile.CommonPath.Length > 0)
                                    {
                                        Log.Information("  Common path: {CommonPath}", lnkFile.CommonPath);
                                    }
                                }
                            }

                            if (lnkFile.TargetIDs.Count > 0)
                            {
                                Console.WriteLine();

                                var absPath = string.Empty;

                                foreach (var shellBag in lnkFile.TargetIDs)
                                {
                                    absPath += shellBag.Value + @"\";
                                }

                                Log.Information("  Absolute path: {Path}",
                                    GetAbsolutePathFromTargetIDs(lnkFile.TargetIDs));
                                Console.WriteLine();
                            }


                            lnkCounter += 1;
                        }

                        Console.WriteLine();
                        entryNum += 1;
                    }
                }


                sw.Stop();

                if (q == false)
                {
                    Console.WriteLine();
                }

                Log.Information(
                    "---------- Processed {SourceFile} in {TotalSeconds:N8} seconds ----------", customDest.SourceFile,
                    sw.Elapsed.TotalSeconds);

                if (q == false)
                {
                    Console.WriteLine();
                }

                return customDest;
            }

            catch (Exception ex)
            {
                if (ex.Message.Equals("Empty custom destinations jump list"))
                {
                    Log.Warning("Error processing {File}. Message: {Message}", jlFile,
                        "Empty custom destinations jump list");
                }
                else
                {
                    _failedFiles.Add($"{jlFile} ==> ({ex.Message})");
                    Log.Fatal(ex, "Error opening {File}. Message: {Message}", jlFile, ex.Message);

                    if (IsReadError(ex))
                    {
                        _readErrors++;
                    }
                }

                Console.WriteLine();
            }

            return null;
        }

        private class CustomHelpAction : SynchronousCommandLineAction
        {
            private readonly HelpAction _defaultHelp;

            public CustomHelpAction(HelpAction action)
            {
                _defaultHelp = action;
            }

            public override int Invoke(ParseResult parseResult)
            {
                var result = _defaultHelp.Invoke(parseResult);

                Log.Warning("{Msg}", string.Join(" ", parseResult.Tokens));

                return result;
            }
        }

        class DateTimeOffsetFormatter : IFormatProvider, ICustomFormatter
        {
            private readonly IFormatProvider _innerFormatProvider;

            public DateTimeOffsetFormatter(IFormatProvider innerFormatProvider)
            {
                _innerFormatProvider = innerFormatProvider;
            }

            public string Format(string format, object arg, IFormatProvider formatProvider)
            {
                if (arg is DateTimeOffset)
                {
                    //an explicit format (the log templates' {Timestamp:HH:mm:ss.fff}) wins over --dt
                    var size = (DateTimeOffset)arg;
                    return size.ToString(string.IsNullOrEmpty(format) ? ActiveDateTimeFormat : format);
                }

                var formattable = arg as IFormattable;
                if (formattable != null)
                {
                    return formattable.ToString(format, _innerFormatProvider);
                }

                return arg.ToString();
            }

            public object GetFormat(Type formatType)
            {
                return formatType == typeof(ICustomFormatter) ? this : _innerFormatProvider.GetFormat(formatType);
            }
        }
    }

    public sealed class AutoCsvOut
    {
        public string User { get; set; }
        //jump list info
        public string SourceFile { get; set; }
        public string SourceCreated { get; set; }
        public string SourceModified { get; set; }
        public string SourceAccessed { get; set; }
        public string AppId { get; set; }
        public string AppIdDescription { get; set; }

        public bool HasSps { get; set; }

        public string DestListVersion { get; set; }
        public string LastUsedEntryNumber { get; set; }
        public string MRU { get; set; }

        //destlist entry
        public string EntryNumber { get; set; }
        public string CreationTime { get; set; }
        public string LastModified { get; set; }
        public string Hostname { get; set; }
        public string MacAddress { get; set; }
        public string Path { get; set; }
        public string InteractionCount { get; set; }
        public string PinStatus { get; set; }
        public string FileBirthDroid { get; set; }
        public string FileDroid { get; set; }
        public string VolumeBirthDroid { get; set; }
        public string VolumeDroid { get; set; }


        //lnk file info
        public string TargetCreated { get; set; }
        public string TargetModified { get; set; }
        public string TargetAccessed { get; set; }
        public uint FileSize { get; set; }
        public string RelativePath { get; set; }
        public string WorkingDirectory { get; set; }
        public string FileAttributes { get; set; }
        public string HeaderFlags { get; set; }
        public string DriveType { get; set; }
        public string VolumeSerialNumber { get; set; }
        public string VolumeLabel { get; set; }

        public string LocalPath { get; set; }
        public string CommonPath { get; set; }
        public string TargetIDAbsolutePath { get; set; }
        public string TargetMFTEntryNumber { get; set; }
        public string TargetMFTSequenceNumber { get; set; }
        public string MachineID { get; set; }
        public string MachineMACAddress { get; set; }

        public string TrackerCreatedOn { get; set; }
        public string ExtraBlocksPresent { get; set; }
        public string Arguments { get; set; }
        public string Notes { get; set; }
    }

    public sealed class CustomCsvOut
    {
        public string User { get; set; }
        //
        public string SourceFile { get; set; }
        public string SourceCreated { get; set; }
        public string SourceModified { get; set; }
        public string SourceAccessed { get; set; }
        public string AppId { get; set; }
        public string AppIdDescription { get; set; }


        public string EntryName { get; set; }

        //lnk file info

        public string TargetCreated { get; set; }
        public string TargetModified { get; set; }
        public string TargetAccessed { get; set; }
        public uint FileSize { get; set; }
        public string RelativePath { get; set; }
        public string WorkingDirectory { get; set; }
        public string FileAttributes { get; set; }
        public string HeaderFlags { get; set; }
        public string DriveType { get; set; }
        public string VolumeSerialNumber { get; set; }
        public string VolumeLabel { get; set; }
        public string LocalPath { get; set; }
        public string CommonPath { get; set; }
        public string TargetIDAbsolutePath { get; set; }
        public string TargetMFTEntryNumber { get; set; }
        public string TargetMFTSequenceNumber { get; set; }
        public string MachineID { get; set; }
        public string MachineMACAddress { get; set; }

        public string TrackerCreatedOn { get; set; }
        public string ExtraBlocksPresent { get; set; }
        public string Arguments { get; set; }
    }
}
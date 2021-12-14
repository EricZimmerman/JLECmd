using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.NamingConventionBinder;
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
using NLog;
using NLog.Config;
using NLog.Targets;
using ServiceStack;
using ServiceStack.Text;
using CsvWriter = CsvHelper.CsvWriter;
#if !NET6_0
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
using ShellBag = Lnk.ShellItems.ShellBag;
using ShellBag0X31 = Lnk.ShellItems.ShellBag0X31;


namespace JLECmd
{
    internal class Program
    {
        private static Logger _logger;

        private static readonly string _preciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

        private static List<string> _failedFiles;

        private static readonly Dictionary<string, string> MacList = new();

        private static List<AutomaticDestination> _processedAutoFiles;
        private static List<CustomDestination> _processedCustomFiles;

        private static RootCommand _rootCommand;
        
        private static string Header =
            $"JLECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
            "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
            "\r\nhttps://github.com/EricZimmerman/JLECmd";

        private static string Footer = @"Examples: JLECmd.exe -f ""C:\Temp\f01b4d95cf55d32a.customDestinations-ms"" --mp" + "\r\n\t " +
                     @" JLECmd.exe -f ""C:\Temp\f01b4d95cf55d32a.automaticDestinations-ms"" --json ""D:\jsonOutput"" --jsonpretty" +
                     "\r\n\t " +
                     @" JLECmd.exe -d ""C:\CustomDestinations"" --csv ""c:\temp"" --html ""c:\temp"" -q" +
                     "\r\n\t " +
                     @" JLECmd.exe -d ""C:\Users\e\AppData\Roaming\Microsoft\Windows\Recent"" --dt ""ddd yyyy MM dd HH:mm:ss.fff"" " +
                     "\r\n\t" +
                     "\r\n\t" +
                     "  Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes\r\n";

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

        private static async Task Main(string[] args)
        {
            ExceptionlessClient.Default.Startup("ZqbYbvr4FIRjkpUrqCLC5N4RfKIuo9YIVmpQuOje");

            SetupNLog();

            _logger = LogManager.GetCurrentClassLogger();

            _rootCommand = new RootCommand
            {
                new Option<string>(
                    "-f",
                    "File to process. Either this or -d is required"),

                new Option<string>(
                    "-d",
                    "Directory to recursively process. Either this or -f is required"),
                
                new Option<bool>(
                    "--all",
                    getDefaultValue:()=>false,
                    "Process all files in directory vs. only files matching *.automaticDestinations-ms or *.customDestinations-ms"),

                new Option<string>(
                    "--csv",
                    "Directory to save CSV formatted results to. This or --json required unless --de or --body is specified"),

                new Option<string>(
                    "--csvf",
                    "File name to save CSV formatted results to. When present, overrides default name\r\n"),
                
                new Option<string>(
                    "--json",
                    "Directory to save json representation to. Use --pretty for a more human readable layout"),

                new Option<string>(
                    "--html",
                    "Directory to save xhtml formatted results to. Be sure to include the full path in double quotes"),

                new Option<bool>(
                    "--pretty",
                    getDefaultValue:()=>false,
                    "When exporting to json, use a more human readable layout"),

                new Option<bool>(
                    "-q",
                    getDefaultValue:()=>false,
                    "Only show the filename being processed vs all output. Useful to speed up exporting to json and/or csv"),
                
                new Option<bool>(
                    "-ld",
                    getDefaultValue:()=>false,
                    "Include more information about lnk files"),
                
                new Option<bool>(
                    "-fd",
                    getDefaultValue:()=>false,
                    "Include full information about lnk files (Alternatively, dump lnk files using --dumpTo and process with LECmd)"),

                new Option<string>(
                    "--appIds",
                    "Path to file containing AppIDs and descriptions (appid|description format). New appIds are added to the built-in list, existing appIds will have their descriptions updated"),

                new Option<string>(
                    "--dumpTo",
                    "Directory to save exported lnk files"),
                
                new Option<string>(
                    "--dt",
                    getDefaultValue:()=>"yyyy-MM-dd HH:mm:ss",
                    "The custom date/time format to use when displaying timestamps. See https://goo.gl/CNVq0k for options. Default is: yyyy-MM-dd HH:mm:ss"),
                
                new Option<bool>(
                    "--mp",
                    getDefaultValue:()=>false,
                    "Display higher precision for timestamps"),
                
                new Option<bool>(
                    "--withDir",
                    getDefaultValue:()=>false,
                    "When true, show contents of Directory not accounted for in DestList entries"),
                
                new Option<bool>(
                    "--debug",
                    getDefaultValue:()=>false,
                    "Show debug information during processing"),



            };
            
            _rootCommand.Description = Header + "\r\n\r\n" + Footer;

            _rootCommand.Handler = CommandHandler.Create(DoWork);

            await _rootCommand.InvokeAsync(args);
        }

        private static void DoWork(string f, string d, bool all, string csv, string csvf, string json, string html, bool pretty, bool q, bool ld, bool fd, string appIds, string dumpTo, string dt, bool mp, bool withDir, bool debug)
        {
            if (f.IsNullOrEmpty() && d.IsNullOrEmpty())
            {
                var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);
                var hc = new HelpContext(helpBld, _rootCommand, Console.Out);

                helpBld.Write(hc);

                _logger.Warn("Either -f or -d is required. Exiting");
                return;
            }

            if (f.IsNullOrEmpty() == false &&  !File.Exists(f))
            {
                _logger.Warn($"File '{f}' not found. Exiting");
                return;
            }

            if (d.IsNullOrEmpty() == false &&
                !Directory.Exists(d))
            {
                _logger.Warn($"Directory '{d}' not found. Exiting");
                return;
            }

        
            _logger.Info(Header);
            _logger.Info("");
            _logger.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}\r\n");

            if (IsAdministrator() == false)
            {
                _logger.Fatal($"Warning: Administrator privileges not found!\r\n");
            }

            if (mp)
            {
                dt = _preciseTimeFormat;
            }

            _processedAutoFiles = new List<AutomaticDestination>();
            _processedCustomFiles = new List<CustomDestination>();

            _failedFiles = new List<string>();

            if (debug)
            {
                LogManager.Configuration.LoggingRules.First().EnableLoggingForLevel(LogLevel.Debug);
                LogManager.ReconfigExistingLoggers();
            }

            if (appIds?.Length > 0)
            {
                if (File.Exists(appIds))
                {
                    _logger.Info($"Looking for AppIDs in '{appIds}'");

                    var added =   JumpList.JumpList.AppIdList.LoadAppListFromFile(appIds);

                    _logger.Info($"Loaded {added:N0} new AppIDs from '{appIds}'\r\n");
                }
                else
                {
                    _logger.Warn($"'{appIds}' does not exist!");
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
                        adjl = ProcessAutoFile(f,q,dt,fd,ld,withDir);
                        if (adjl != null)
                        {
                            _processedAutoFiles.Add(adjl);
                        }
                    }
                    catch (UnauthorizedAccessException ua)
                    {
                        _logger.Error(
                            $"Unable to access '{f}'. Are you running as an administrator? Error: {ua.Message}");
                        return;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(
                            $"Error processing jump list. Error: {ex.Message}");
                        return;
                    }
                }
                else
                {
                    try
                    {
                        CustomDestination cdjl ;
                        cdjl = ProcessCustomFile(f,q,dt,ld);
                        if (cdjl != null)
                        {
                            _processedCustomFiles.Add(cdjl);
                        }
                    }
                    catch (UnauthorizedAccessException ua)
                    {
                        _logger.Error(
                            $"Unable to access '{f}'. Are you running as an administrator? Error: {ua.Message}");
                        return;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(
                            $"Error processing jump list. Error: {ex.Message}");
                        return;
                    }
                }
            }
            else
            {
                _logger.Info($"Looking for jump list files in '{d}'");
                _logger.Info("");

                d = Path.GetFullPath(d);

                var jumpFiles = new List<string>();
                
                try
                {
                    

                    #if !NET6_0
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

                            if (fsei.Extension.ToUpperInvariant() == ".AUTOMATICDESTINATIONS-MS" || fsei.Extension.ToUpperInvariant() == ".CUSTOMDESTINATIONS-MS")
                            {
                                return true;
                            }

                            return false;
                        };

                        filters.RecursionFilter = entryInfo => !entryInfo.IsMountPoint && !entryInfo.IsSymbolicLink;

                        filters.ErrorFilter = (errorCode, errorMessage, pathProcessed) => true;

                        var dirEnumOptions =
                            DirectoryEnumerationOptions.Files | DirectoryEnumerationOptions.Recursive |
                            DirectoryEnumerationOptions.SkipReparsePoints | DirectoryEnumerationOptions.ContinueOnException |
                            DirectoryEnumerationOptions.BasicSearch;
                        
                        var files2 = Directory.EnumerateFileSystemEntries(d, dirEnumOptions, filters);
                    #else
                        var mask = "*.*Destinations-ms".ToUpperInvariant();
                        if (all)
                        {
                            mask = "*";
                        }
                        var enumerationOptions = new EnumerationOptions
                        {
                            IgnoreInaccessible = true,
                            MatchCasing = MatchCasing.CaseInsensitive,
                            RecurseSubdirectories = true,
                            AttributesToSkip = 0
                        };
                        
                       var files2 =
                            Directory.EnumerateFileSystemEntries(d, mask,enumerationOptions);
                    #endif
                   


                  jumpFiles.AddRange(files2);



                }
                catch (UnauthorizedAccessException ua)
                {
                    _logger.Error(
                        $"Unable to access '{d}'. Error message: {ua.Message}");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        $"Error getting jump list files in '{d}'. Error: {ex.Message}");
                    return;
                }

                _logger.Info($"Found {jumpFiles.Count:N0} files");
                _logger.Info("");

                var sw = new Stopwatch();
                sw.Start();

                foreach (var file in jumpFiles)
                {
                    if (IsAutomaticDestinationFile(file))
                    {
                        AutomaticDestination adjl ;
                        adjl = ProcessAutoFile(file,q,dt,fd,ld,withDir);
                        if (adjl != null)
                        {
                            _processedAutoFiles.Add(adjl);
                        }
                    }
                    else
                    {
                        CustomDestination cdjl ;
                        cdjl = ProcessCustomFile(file,q,dt,ld);
                        if (cdjl != null)
                        {
                            _processedCustomFiles.Add(cdjl);
                        }
                    }
                }

                sw.Stop();

                if (q)
                {
                    _logger.Info("");
                }

                _logger.Info(
                    $"Processed {jumpFiles.Count - _failedFiles.Count:N0} out of {jumpFiles.Count:N0} files in {sw.Elapsed.TotalSeconds:N4} seconds");
                if (_failedFiles.Count > 0)
                {
                    _logger.Info("");
                    _logger.Warn("Failed files");
                    foreach (var failedFile in _failedFiles)
                    {
                        _logger.Info($"  {failedFile}");
                    }
                }
            }

            //export lnks if requested
            if (dumpTo?.Length > 0)
            {
                _logger.Info("");
                _logger.Warn(
                    $"Dumping lnk files to '{dumpTo}'");

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
                ExportAuto(csv,csvf,json,html,pretty,dt,debug,withDir);
            }

            if (_processedCustomFiles.Count > 0)
            {
                ExportCustom(csv,csvf,json,html,pretty,dt);
            }
        }

        private static void ExportCustom(string csv, string csvf, string json, string html, bool pretty, string dt)
        {
            _logger.Info("");

            try
            {
                CsvWriter csvCustom = null;
                StreamWriter swCustom = null;

                if (csv?.Length > 0)
                {
                    if (Directory.Exists(csv) == false)
                    {
                        _logger.Warn($"'{csv} does not exist. Creating...'");
                        Directory.CreateDirectory(csv);
                    }


                    var outName = $"{DateTimeOffset.Now:yyyyMMddHHmmss}_CustomDestinations.csv";

                    if (csvf.IsNullOrEmpty() == false)
                    {
                        outName =
                            $"{Path.GetFileNameWithoutExtension(csvf)}_CustomDestinations{Path.GetExtension(csvf)}";
                    }

                    var outFile = Path.Combine(csv, outName);
                    

                    _logger.Warn(
                        $"CustomDestinations CSV output will be saved to '{outFile}'");

                    try
                    {
                        swCustom = new StreamWriter(outFile);
                        csvCustom = new CsvWriter(swCustom,CultureInfo.InvariantCulture);
                        
                        csvCustom.WriteHeader(typeof(CustomCsvOut));
                        csvCustom.NextRecord();
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(
                            $"Unable to write to '{csv}'. Custom CSV export canceled. Error: {ex.Message}");
                    }
                }

                if (json?.Length > 0)
                {
                    if (Directory.Exists(json) == false)
                    {
                        _logger.Warn($"'{json} does not exist. Creating...'");
                        Directory.CreateDirectory(json);
                    }
                    _logger.Warn($"Saving Custom json output to '{json}'");
                }


                XmlTextWriter xml = null;

                if (html?.Length > 0)
                {
                    if (Directory.Exists(html) == false)
                    {
                        _logger.Warn($"'{html} does not exist. Creating...'");
                        Directory.CreateDirectory(html);
                    }


                    var outDir = Path.Combine(html,
                        $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_JLECmd_Custom_Output_for_{html.Replace(@":\", "_").Replace(@"\", "_")}");

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

                    _logger.Warn($"Saving HTML output to '{outFile}'");

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
                            json);
                    }


                    var records = GetCustomCsvFormat(processedFile,dt);

                    try
                    {
                        csvCustom?.WriteRecords(records);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(
                            $"Error writing record for '{processedFile.SourceFile}' to '{csv}'. Error: {ex.Message}");
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
                _logger.Error(
                    $"Error exporting Custom Destinations data! Error: {ex.Message}");
            }
        }

        private static void ExportAuto(string csv, string csvf, string json, string html, bool pretty, string dt, bool debug, bool wd)
        {
            _logger.Info("");

            try
            {
                CsvWriter csvAuto = null;
                StreamWriter swAuto = null;

                if (csv?.Length > 0)
                {
                    if (Directory.Exists(csv) == false)
                    {
                        _logger.Warn($"'{csv} does not exist. Creating...'");
                        Directory.CreateDirectory(csv);
                    }

                    var outName = $"{DateTimeOffset.Now:yyyyMMddHHmmss}_AutomaticDestinations.csv";

                    if (csvf.IsNullOrEmpty() == false)
                    {
                        outName =
                            $"{Path.GetFileNameWithoutExtension(csvf)}_AutomaticDestinations{Path.GetExtension(csvf)}";
                    }
                    
                    var outFile = Path.Combine(csv, outName);
                    
                    _logger.Warn(
                        $"AutomaticDestinations CSV output will be saved to '{outFile}'");

                    try
                    {
                        swAuto = new StreamWriter(outFile);
                        csvAuto = new CsvWriter(swAuto,CultureInfo.InvariantCulture);
                
                        csvAuto.WriteHeader(typeof(AutoCsvOut));
                        csvAuto.NextRecord();
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(
                            $"Unable to write to '{csv}'. Automatic CSV export canceled. Error: {ex.Message}");
                    }
                }

                if (json?.Length > 0)
                {
                    if (Directory.Exists(json) == false)
                    {
                        _logger.Warn($"'{json} does not exist. Creating...'");
                        Directory.CreateDirectory(json);
                    }

                    _logger.Warn($"Saving Automatic json output to '{json}'");
                }


                XmlTextWriter xml = null;

                if (html?.Length > 0)
                {
                    if (Directory.Exists(html) == false)
                    {
                        _logger.Warn($"'{html} does not exist. Creating...'");
                        Directory.CreateDirectory(html);
                    }

                    var outDir = Path.Combine(html,
                        $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_JLECmd_Automatic_Output_for_{html.Replace(@":\", "_").Replace(@"\", "_")}");

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

                    _logger.Warn($"Saving HTML output to '{outFile}'");

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
                            json);
                    }

                    var records = GetAutoCsvFormat(processedFile,debug,dt,wd);

                    try
                    {
                        csvAuto?.WriteRecords(records);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(
                            $"Error writing record for '{processedFile.SourceFile}' to '{csv}'. Error: {ex.Message}");
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
                _logger.Error(
                    $"Error exporting Automatic Destinations data! Error: {ex.Message}");
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

        private static List<CustomCsvOut> GetCustomCsvFormat(CustomDestination cust, string dt)
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

                    if (si.ExtensionBlocks?.Count > 0)
                    {
                        var eb = si.ExtensionBlocks.LastOrDefault(t => t is Beef0004);
                        if (eb is Beef0004)
                        {
                            var eb4 = eb as Beef0004;
                            if (eb4.MFTInformation.MFTEntryNumber != null)
                            {
                                csOut.TargetMFTEntryNumber =
                                    $"0x{eb4.MFTInformation.MFTEntryNumber.Value.ToString("X")}";
                            }

                            if (eb4.MFTInformation.MFTSequenceNumber != null)
                            {
                                csOut.TargetMFTSequenceNumber =
                                    $"0x{eb4.MFTInformation.MFTSequenceNumber.Value.ToString("X")}";
                            }
                        }
                    }
                }

                csList.Add(csOut);
            }

            return csList;
        }

        private static List<AutoCsvOut> GetAutoCsvFormat(AutomaticDestination auto, bool debug, string dt, bool wd)
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
                    _logger.Debug("Dumping destListEntry");
                    destListEntry.PrintDump();
                }

                var csOut = new AutoCsvOut
                {
                    SourceFile = auto.SourceFile,
                    SourceCreated = ct.ToString(dt),
                    SourceModified = mt.ToString(dt),
                    SourceAccessed = at.ToString(dt),
                    AppId = auto.AppId.AppId,
                    AppIdDescription = auto.AppId.Description,
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
                    _logger.Debug("CSOut values:");
                    csOut.PrintDump();
                }


                if (destListEntry.Lnk == null)
                {
                    csList.Add(csOut);
                    continue;
                }

                _logger.Debug("Lnk file isn't null. Continuing");

                _logger.Debug($"Getting absolute path. TargetID count: {destListEntry.Lnk.TargetIDs.Count}");

                var target = GetAbsolutePathFromTargetIDs(destListEntry.Lnk.TargetIDs);

                _logger.Debug($"GetAbsolutePathFromTargetIDs Target is: {target}");

                if (target.Length == 0)
                {
                    _logger.Debug($"Target length is 0. building alternate path");

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

                _logger.Debug($"Target is: {target}");

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

                _logger.Debug($"csOut.Arguments is: {csOut.Arguments}");

                csOut.WorkingDirectory = destListEntry.Lnk.WorkingDirectory;

                _logger.Debug($"csOut.WorkingDirectory is: {csOut.WorkingDirectory}");

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

                _logger.Debug($"csOut.ExtraBlocksPresent is: {ebPresent}");

                csOut.ExtraBlocksPresent = ebPresent;

                var tnb =
                    destListEntry.Lnk.ExtraBlocks.SingleOrDefault(
                        t => t.GetType().Name.ToUpper() == "TRACKERDATABASEBLOCK");

                if (tnb != null)
                {
                    _logger.Debug($"Found tracker block");

                    var tnbBlock = tnb as TrackerDataBaseBlock;

                    csOut.TrackerCreatedOn =
                        tnbBlock?.CreationTime.ToString(dt);

                    csOut.MachineID = tnbBlock?.MachineId;
                    csOut.MachineMACAddress = tnbBlock?.MacAddress;
                }

                if (destListEntry.Lnk.TargetIDs?.Count > 0)
                {
                    _logger.Debug($"Target ID count: {destListEntry.Lnk.TargetIDs.Count}");

                    var si = destListEntry.Lnk.TargetIDs.Last();

                    if (si.ExtensionBlocks?.Count > 0)
                    {
                        var eb = si.ExtensionBlocks.LastOrDefault(t => t is Beef0004);
                        if (eb is Beef0004)
                        {
                            var eb4 = eb as Beef0004;
                            if (eb4.MFTInformation.MFTEntryNumber != null)
                            {
                                csOut.TargetMFTEntryNumber =
                                    $"0x{eb4.MFTInformation.MFTEntryNumber.Value.ToString("X")}";
                            }

                            if (eb4.MFTInformation.MFTSequenceNumber != null)
                            {
                                csOut.TargetMFTSequenceNumber =
                                    $"0x{eb4.MFTInformation.MFTSequenceNumber.Value.ToString("X")}";
                            }
                        }
                    }
                }

                _logger.Debug($"Adding to csList");

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

                    if (auto.DestListEntries.Any(
                        t => t.EntryNumber.ToString("X") == directoryEntry.DirectoryName))
                    {
                        continue;
                    }

                    //this directory entry is not in destlist

                    var f = auto.GetLnkFromDirectoryName(directoryEntry.DirectoryName);

                    if (f != null)
                    {
                        var csOut = new AutoCsvOut
                        {
                            SourceFile = auto.SourceFile,
                            SourceCreated = ct.ToString(dt),
                            SourceModified = mt.ToString(dt),
                            SourceAccessed = at.ToString(dt),
                            AppId = auto.AppId.AppId,
                            AppIdDescription = auto.AppId.Description,
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
                            f.ExtraBlocks.SingleOrDefault(
                                t => t.GetType().Name.ToUpper() == "TRACKERDATABASEBLOCK");

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

                            if (si.ExtensionBlocks?.Count > 0)
                            {
                                var eb = si.ExtensionBlocks.LastOrDefault(t => t is Beef0004);
                                if (eb is Beef0004)
                                {
                                    var eb4 = eb as Beef0004;
                                    if (eb4.MFTInformation.MFTEntryNumber != null)
                                    {
                                        csOut.TargetMFTEntryNumber =
                                            $"0x{eb4.MFTInformation.MFTEntryNumber.Value.ToString("X")}";
                                    }

                                    if (eb4.MFTInformation.MFTSequenceNumber != null)
                                    {
                                        csOut.TargetMFTSequenceNumber =
                                            $"0x{eb4.MFTInformation.MFTSequenceNumber.Value.ToString("X")}";
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

        private static void DumpToJsonAuto(AutomaticDestination auto, bool pretty, string outFile)
        {
            if (pretty)
            {
                File.WriteAllText(outFile, auto.Dump());
            }
            else
            {
                File.WriteAllText(outFile, auto.ToJson());
            }
        }

        private static void DumpToJsonCustom(CustomDestination cust, bool pretty, string outFile)
        {
            if (pretty)
            {
                File.WriteAllText(outFile, cust.Dump());
            }
            else
            {
                File.WriteAllText(outFile, cust.ToJson());
            }
        }

        private static void SaveJsonCustom(CustomDestination cust, bool pretty, string outDir)
        {
            try
            {
                if (Directory.Exists(outDir) == false)
                {
                    Directory.CreateDirectory(outDir);
                }

                var outName =
                    $"{DateTimeOffset.UtcNow.ToString("yyyyMMddHHmmss")}_{Path.GetFileName(cust.SourceFile)}.json";
                var outFile = Path.Combine(outDir, outName);

                DumpToJsonCustom(cust, pretty, outFile);
            }
            catch (Exception ex)
            {
                _logger.Error($"Error exporting json for '{cust.SourceFile}'. Error: {ex.Message}");
            }
        }

        private static void SaveJsonAuto(AutomaticDestination auto, bool pretty, string outDir)
        {
            try
            {
                if (Directory.Exists(outDir) == false)
                {
                    Directory.CreateDirectory(outDir);
                }

                var outName =
                    $"{DateTimeOffset.UtcNow.ToString("yyyyMMddHHmmss")}_{Path.GetFileName(auto.SourceFile)}.json";
                var outFile = Path.Combine(outDir, outName);

                DumpToJsonAuto(auto, pretty, outFile);
            }
            catch (Exception ex)
            {
                _logger.Error($"Error exporting json for '{auto.SourceFile}'. Error: {ex.Message}");
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

        private static string GetAbsolutePathFromTargetIDs(List<ShellBag> ids)
        {
            if (ids == null)
            {
                return $"(No target IDs present)";
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

        private static AutomaticDestination ProcessAutoFile(string jlFile, bool q, string dt, bool fd, bool ld, bool wd)
        {
            if (q == false)
            {
                _logger.Warn($"Processing '{jlFile}'");
                _logger.Info("");
            }

            var sw = new Stopwatch();
            sw.Start();

            try
            {
                _logger.Debug($"Opening {jlFile}");

                var autoDest = JumpList.JumpList.LoadAutoJumplist(jlFile);

                _logger.Debug($"Opened {jlFile}");

                if (q == false)
                {
                    _logger.Error($"Source file: {autoDest.SourceFile}");

                    _logger.Info("");

                    _logger.Warn("--- AppId information ---");
                    _logger.Info($"  AppID: {autoDest.AppId.AppId}");
                    _logger.Info($"  Description: {autoDest.AppId.Description}");
                    _logger.Info("");

                    _logger.Warn("--- DestList information ---");
                    _logger.Info($"  Expected DestList entries:  {autoDest.DestListCount:N0}");
                    _logger.Info($"  Actual DestList entries: {autoDest.DestListCount:N0}");
                    _logger.Info($"  DestList version: {autoDest.DestListVersion}");

                    if (autoDest.DestListCount != autoDest.Directory.Count - 2)
                    {
                        _logger.Info("");
                        _logger.Fatal(
                            $"  There are more items in the Directory ({autoDest.Directory.Count - 2:N0}) than are contained in the DestList ({autoDest.DestListCount:N0}). Use --WithDir to view/export them");
                    }

                    _logger.Info("");

                    _logger.Warn("--- DestList entries ---");
                    foreach (var autoDestList in autoDest.DestListEntries)
                    {
                        _logger.Info($"Entry #: {autoDestList.EntryNumber}");
                        _logger.Info($"  MRU: {autoDestList.MRUPosition}");
                        _logger.Info($"  Path: {autoDestList.Path}");
                        _logger.Info($"  Pinned: {autoDestList.Pinned}");
                        _logger.Info(
                            $"  Created on: {autoDestList.CreatedOn.ToString(dt)}");
                        _logger.Info(
                            $"  Last modified: {autoDestList.LastModified.ToString(dt)}");
                        _logger.Info($"  Hostname: {autoDestList.Hostname}");
                        _logger.Info(
                            $"  Mac Address: {(autoDestList.MacAddress == "00:00:00:00:00:00" ? string.Empty : autoDestList.MacAddress)}");
                        _logger.Info(
                            $"  Interaction count: {autoDestList.InteractionCount:N0}");

                        _logger.Error("\r\n--- Lnk information ---\r\n");

                        if (fd)
                        {
                            var tc = autoDestList.Lnk.Header.TargetCreationDate.Year == 1601
                                ? ""
                                : autoDestList.Lnk.Header.TargetCreationDate.ToString(
                                    dt);
                            var tm = autoDestList.Lnk.Header.TargetModificationDate.Year == 1601
                                ? ""
                                : autoDestList.Lnk.Header.TargetModificationDate.ToString(
                                    dt);
                            var ta = autoDestList.Lnk.Header.TargetLastAccessedDate.Year == 1601
                                ? ""
                                : autoDestList.Lnk.Header.TargetLastAccessedDate.ToString(
                                    dt);


                            _logger.Info($"  Lnk target created: {tc}");
                            _logger.Info($"  Lnk target modified: {tm}");
                            _logger.Info($"  Lnk target accessed: {ta}");

                            _logger.Info("");

                            DumpLnkFile(autoDestList.Lnk,dt);
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

                            _logger.Info($"  Absolute path: {target}");
                            _logger.Info("");
                        }
                        else
                        {
                            _logger.Info($"  (lnk file not present)");
                            _logger.Info("");
                        }

                        _logger.Info("");
                    }

                    if (wd)
                    {
                        _logger.Fatal("Directory entries not represented by DestList entries");

                        foreach (var directoryEntry in autoDest.Directory)
                        {
                            if (directoryEntry.DirectoryName.Equals("Root Entry") ||
                                directoryEntry.DirectoryName.Equals("DestList"))
                            {
                                continue;
                            }

                            if (autoDest.DestListEntries.Any(
                                t => t.EntryNumber.ToString("X") == directoryEntry.DirectoryName))
                            {
                                continue;
                            }

                            //this directory entry is not in destlist
                            if (!q)
                            {
                                _logger.Info($"Directory Name: {directoryEntry.DirectoryName}");
                            }


                            var f = autoDest.GetLnkFromDirectoryName(directoryEntry.DirectoryName);

                            if (f != null)
                            {
                                if (fd)
                                {
                                    var tc = f.Header.TargetCreationDate.Year == 1601
                                        ? ""
                                        : f.Header.TargetCreationDate.ToString(
                                            dt);
                                    var tm = f.Header.TargetModificationDate.Year == 1601
                                        ? ""
                                        : f.Header.TargetModificationDate.ToString(
                                            dt);
                                    var ta = f.Header.TargetLastAccessedDate.Year == 1601
                                        ? ""
                                        : f.Header.TargetLastAccessedDate.ToString(
                                            dt);


                                    _logger.Info($"  Lnk target created: {tc}");
                                    _logger.Info($"  Lnk target modified: {tm}");
                                    _logger.Info($"  Lnk target accessed: {ta}");
                                    _logger.Info("");

                                    DumpLnkFile(f,dt);
                                }
                                else if (ld)
                                {
                                    DumpLnkDetail(f);
                                }

                                if (!fd)
                                {
                                    _logger.Info("");

                                    var target = GetAbsolutePathFromTargetIDs(f.TargetIDs);

                                    if (target.Length == 0)
                                    {
                                        target = $"{f.NetworkShareInfo?.NetworkShareName}\\\\{f.CommonPath}";
                                    }

                                    _logger.Info($"  Absolute path: {target}");
                                    _logger.Info("");
                                }
                            }
                            else
                            {
                                _logger.Debug(
                                    $"  No lnk file found for directory entry '{directoryEntry.DirectoryName}'");
                            }
                        }
                    }
                }

                sw.Stop();

                if (q == false)
                {
                    _logger.Info("");
                }

                if (autoDest.DestListCount != autoDest.Directory.Count - 2)
                {
                    _logger.Fatal(
                        $"** There are more items in the Directory ({autoDest.Directory.Count - 2:N0}) than are contained in the DestList ({autoDest.DestListCount:N0}). Use --WithDir to view them **");
                    _logger.Info("");
                }

                _logger.Info(
                    $"---------- Processed '{autoDest.SourceFile}' in {sw.Elapsed.TotalSeconds:N8} seconds ----------");

                if (q == false)
                {
                    _logger.Info("\r\n");
                }

                return autoDest;
            }

            catch (Exception ex)
            {
                _failedFiles.Add($"{jlFile} ==> ({ex.Message})");
                _logger.Fatal($"Error opening '{jlFile}'. Message: {ex.Message}");
                _logger.Info("");
            }

            return null;
        }

        private static void DumpLnkDetail(LnkFile lnk)
        {
            if (lnk == null)
            {
                _logger.Warn("(lnk file not present)");
                return;
            }
            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasName) == Lnk.Header.DataFlag.HasName)
            {
                _logger.Info($"  Name: {lnk.Name}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasRelativePath) ==
                Lnk.Header.DataFlag.HasRelativePath)
            {
                _logger.Info($"  Relative Path: {lnk.RelativePath}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasWorkingDir) ==
                Lnk.Header.DataFlag.HasWorkingDir)
            {
                _logger.Info($"  Working Directory: {lnk.WorkingDirectory}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) ==
                Lnk.Header.DataFlag.HasArguments)
            {
                _logger.Info($"  Arguments: {lnk.Arguments}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasLinkInfo) ==
                Lnk.Header.DataFlag.HasLinkInfo)
            {
                _logger.Info("");
                _logger.Error("--- Link information ---");
                _logger.Info($"Flags: {lnk.LocationFlags}");

                if (lnk.VolumeInfo != null)
                {
                    _logger.Info("");
                    _logger.Warn(">>Volume information");
                    _logger.Info(
                        $"  Drive type: {GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType)}");
                    _logger.Info($"  Serial number: {lnk.VolumeInfo.VolumeSerialNumber}");

                    var label = lnk.VolumeInfo.VolumeLabel.Length > 0
                        ? lnk.VolumeInfo.VolumeLabel
                        : "(No label)";

                    _logger.Info($"  Label: {label}");
                }

                if (lnk.NetworkShareInfo != null)
                {
                    _logger.Info("");
                    _logger.Warn("  Network share information");

                    if (lnk.NetworkShareInfo.DeviceName.Length > 0)
                    {
                        _logger.Info($"    Device name: {lnk.NetworkShareInfo.DeviceName}");
                    }

                    _logger.Info($"    Share name: {lnk.NetworkShareInfo.NetworkShareName}");

                    _logger.Info(
                        $"    Provider type: {lnk.NetworkShareInfo.NetworkProviderType}");
                    _logger.Info($"    Share flags: {lnk.NetworkShareInfo.ShareFlags}");
                    _logger.Info("");
                }

                if (lnk.LocalPath?.Length > 0)
                {
                    _logger.Info($"  Local path: {lnk.LocalPath}");
                }

                if (lnk.CommonPath.Length > 0)
                {
                    _logger.Info($"  Common path: {lnk.CommonPath}");
                }
            }
        }

        private static void DumpLnkFile(LnkFile lnk, string dt)
        {
            if (lnk == null)
            {
                _logger.Warn($"(lnk file not present)");
                return;
            }
            _logger.Warn("--- Header ---");

            var tc1 = lnk.Header.TargetCreationDate.Year == 1601
                ? ""
                : lnk.Header.TargetCreationDate.ToString(dt);
            var tm1 = lnk.Header.TargetModificationDate.Year == 1601
                ? ""
                : lnk.Header.TargetModificationDate.ToString(dt);
            var ta1 = lnk.Header.TargetLastAccessedDate.Year == 1601
                ? ""
                : lnk.Header.TargetLastAccessedDate.ToString(dt);

            _logger.Info($"  Target created:  {tc1}");
            _logger.Info($"  Target modified: {tm1}");
            _logger.Info($"  Target accessed: {ta1}");
            _logger.Info("");
            _logger.Info($"  File size: {lnk.Header.FileSize:N0}");
            _logger.Info($"  Flags: {lnk.Header.DataFlags}");
            _logger.Info($"  File attributes: {lnk.Header.FileAttributes}");
            
            if (lnk.Header.HotKey.Length > 0)
            {
                _logger.Info($"  Hot key: {lnk.Header.HotKey}");
            }

            _logger.Info($"  Icon index: {lnk.Header.IconIndex}");
            _logger.Info(
                $"  Show window: {lnk.Header.ShowWindow} ({GetDescriptionFromEnumValue(lnk.Header.ShowWindow)})");

            _logger.Info("");

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasName) ==Lnk.Header.DataFlag.HasName)
            {
                _logger.Info($"Name: {lnk.Name}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasRelativePath) == Lnk.Header.DataFlag.HasRelativePath)
            {
                _logger.Info($"Relative Path: {lnk.RelativePath}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasWorkingDir) == Lnk.Header.DataFlag.HasWorkingDir)
            {
                _logger.Info($"Working Directory: {lnk.WorkingDirectory}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) == Lnk.Header.DataFlag.HasArguments)
            {
                _logger.Info($"Arguments: {lnk.Arguments}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasIconLocation) == Lnk.Header.DataFlag.HasIconLocation)
            {
                _logger.Info($"Icon Location: {lnk.IconLocation}");
            }

            if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasLinkInfo) == Lnk.Header.DataFlag.HasLinkInfo)
            {
                _logger.Info("");
                _logger.Error("--- Link information ---");
                _logger.Info($"Flags: {lnk.LocationFlags}");

                if (lnk.VolumeInfo != null)
                {
                    _logger.Info("");
                    _logger.Warn(">>Volume information");
                    _logger.Info($"  Drive type: {GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType)}");
                    _logger.Info($"  Serial number: {lnk.VolumeInfo.VolumeSerialNumber}");

                    var label = lnk.VolumeInfo.VolumeLabel.Length > 0
                        ? lnk.VolumeInfo.VolumeLabel
                        : "(No label)";

                    _logger.Info($"  Label: {label}");
                }

                if (lnk.NetworkShareInfo != null)
                {
                    _logger.Info("");
                    _logger.Warn("  Network share information");

                    if (lnk.NetworkShareInfo.DeviceName.Length > 0)
                    {
                        _logger.Info($"    Device name: {lnk.NetworkShareInfo.DeviceName}");
                    }

                    _logger.Info($"    Share name: {lnk.NetworkShareInfo.NetworkShareName}");

                    _logger.Info($"    Provider type: {lnk.NetworkShareInfo.NetworkProviderType}");
                    _logger.Info($"    Share flags: {lnk.NetworkShareInfo.ShareFlags}");
                    _logger.Info("");
                }

                if (lnk.LocalPath?.Length > 0)
                {
                    _logger.Info($"  Local path: {lnk.LocalPath}");
                }

                if (lnk.CommonPath.Length > 0)
                {
                    _logger.Info($"  Common path: {lnk.CommonPath}");
                }
            }

            if (lnk.TargetIDs.Count > 0)
            {
                _logger.Info("");

                // var absPath = string.Empty;
                //
                // foreach (var shellBag in lnk.TargetIDs)
                // {
                //     absPath += shellBag.Value + @"\";
                // }

                _logger.Error("--- Target ID information (Format: Type ==> Value) ---");
                _logger.Info("");
                _logger.Info($"  Absolute path: {GetAbsolutePathFromTargetIDs(lnk.TargetIDs)}");
                _logger.Info("");

                foreach (var shellBag in lnk.TargetIDs)
                {
                    //HACK
                    //This is a total hack until i can refactor some shellbag code to clean things up

                    var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;

                    _logger.Info($"  -{shellBag.FriendlyName} ==> {val}");

                    switch (shellBag.GetType().Name.ToUpper())
                    {
                        case "SHELLBAG0X36":
                        case "SHELLBAG0X32":
                            var b32 = shellBag as ShellBag0X32;

                            _logger.Info($"    Short name: {b32.ShortName}");
                            if (b32.LastModificationTime.HasValue)
                            {
                                _logger.Info(
                                    $"    Modified: {b32.LastModificationTime.Value.ToString(dt)}");
                            }
                            else
                            {
                                _logger.Info($"    Modified:");
                            }


                            var extensionNumber32 = 0;
                            if (b32.ExtensionBlocks.Count > 0)
                            {
                                _logger.Info($"    Extension block count: {b32.ExtensionBlocks.Count:N0}");
                                _logger.Info("");
                                foreach (var extensionBlock in b32.ExtensionBlocks)
                                {
                                    _logger.Info(
                                        $"    --------- Block {extensionNumber32:N0} ({extensionBlock.GetType().Name}) ---------");
                                    if (extensionBlock is Beef0004)
                                    {
                                        var b4 = extensionBlock as Beef0004;

                                        _logger.Info($"    Long name: {b4.LongName}");
                                        if (b4.LocalisedName.Length > 0)
                                        {
                                            _logger.Info($"    Localized name: {b4.LocalisedName}");
                                        }

                                        if (b4.CreatedOnTime.HasValue)
                                        {
                                            _logger.Info(
                                                $"    Created: {b4.CreatedOnTime.Value.ToString(dt)}");
                                        }
                                        else
                                        {
                                            _logger.Info($"    Created:");
                                        }

                                        if (b4.LastAccessTime.HasValue)
                                        {
                                            _logger.Info(
                                                $"    Last access: {b4.LastAccessTime.Value.ToString(dt)}");
                                        }
                                        else
                                        {
                                            _logger.Info($"    Last access: ");
                                        }

                                        if (b4.MFTInformation.MFTEntryNumber > 0)
                                        {
                                            _logger.Info(
                                                $"    MFT entry/sequence #: {b4.MFTInformation.MFTEntryNumber}/{b4.MFTInformation.MFTSequenceNumber} (0x{b4.MFTInformation.MFTEntryNumber:X}/0x{b4.MFTInformation.MFTSequenceNumber:X})");
                                        }
                                    }
                                    else if (extensionBlock is Beef0025)
                                    {
                                        var b25 = extensionBlock as Beef0025;
                                        _logger.Info(
                                            $"    Filetime 1: {b25.FileTime1.Value.ToString(dt)}, Filetime 2: {b25.FileTime2.Value.ToString(dt)}");
                                    }
                                    else if (extensionBlock is Beef0003)
                                    {
                                        var b3 = extensionBlock as Beef0003;
                                        _logger.Info($"    GUID: {b3.GUID1} ({b3.GUID1Folder})");
                                    }
                                    else if (extensionBlock is Beef001a)
                                    {
                                        var b3 = extensionBlock as Beef001a;
                                        _logger.Info($"    File document type: {b3.FileDocumentTypeString}");
                                    }
                                    else
                                    {
                                        _logger.Info($"    {extensionBlock}");
                                    }

                                    extensionNumber32 += 1;
                                }
                            }

                            break;
                        case "SHELLBAG0X31":

                            var b3x = shellBag as ShellBag0X31;

                            _logger.Info($"    Short name: {b3x.ShortName}");
                            if (b3x.LastModificationTime.HasValue)
                            {
                                _logger.Info(
                                    $"    Modified: {b3x.LastModificationTime.Value.ToString(dt)}");
                            }
                            else
                            {
                                _logger.Info($"    Modified:");
                            }

                            var extensionNumber = 0;
                            if (b3x.ExtensionBlocks.Count > 0)
                            {
                                _logger.Info($"    Extension block count: {b3x.ExtensionBlocks.Count:N0}");
                                _logger.Info("");
                                foreach (var extensionBlock in b3x.ExtensionBlocks)
                                {
                                    _logger.Info(
                                        $"    --------- Block {extensionNumber:N0} ({extensionBlock.GetType().Name}) ---------");
                                    if (extensionBlock is Beef0004)
                                    {
                                        var b4 = extensionBlock as Beef0004;

                                        _logger.Info($"    Long name: {b4.LongName}");
                                        if (b4.LocalisedName.Length > 0)
                                        {
                                            _logger.Info($"    Localized name: {b4.LocalisedName}");
                                        }

                                        if (b4.CreatedOnTime.HasValue)
                                        {
                                            _logger.Info(
                                                $"    Created: {b4.CreatedOnTime.Value.ToString(dt)}");
                                        }
                                        else
                                        {
                                            _logger.Info($"    Created:");
                                        }

                                        if (b4.LastAccessTime.HasValue)
                                        {
                                            _logger.Info(
                                                $"    Last access: {b4.LastAccessTime.Value.ToString(dt)}");
                                        }
                                        else
                                        {
                                            _logger.Info($"    Last access: ");
                                        }

                                        if (b4.MFTInformation.MFTEntryNumber > 0)
                                        {
                                            _logger.Info(
                                                $"    MFT entry/sequence #: {b4.MFTInformation.MFTEntryNumber}/{b4.MFTInformation.MFTSequenceNumber} (0x{b4.MFTInformation.MFTEntryNumber:X}/0x{b4.MFTInformation.MFTSequenceNumber:X})");
                                        }
                                    }
                                    else if (extensionBlock is Beef0025)
                                    {
                                        var b25 = extensionBlock as Beef0025;
                                        _logger.Info(
                                            $"    Filetime 1: {b25.FileTime1.Value.ToString(dt)}, Filetime 2: {b25.FileTime2.Value.ToString(dt)}");
                                    }
                                    else if (extensionBlock is Beef0003)
                                    {
                                        var b3 = extensionBlock as Beef0003;
                                        _logger.Info($"    GUID: {b3.GUID1} ({b3.GUID1Folder})");
                                    }
                                    else if (extensionBlock is Beef001a)
                                    {
                                        var b3 = extensionBlock as Beef001a;
                                        _logger.Info($"    File document type: {b3.FileDocumentTypeString}");
                                    }
                                    else
                                    {
                                        _logger.Info($"    {extensionBlock}");
                                    }

                                    extensionNumber += 1;
                                }
                            }
                            break;

                        case "SHELLBAG0X00":
                            var b00 = shellBag as ShellBag0X00;

                            if (b00.PropertyStore.Sheets.Count > 0)
                            {
                                _logger.Warn("  >> Property store (Format: GUID\\ID Description ==> Value)");
                                var propCount = 0;

                                foreach (var prop in b00.PropertyStore.Sheets)
                                foreach (var propertyName in prop.PropertyNames)
                                {
                                    propCount += 1;

                                    var prefix = $"{prop.GUID}\\{propertyName.Key}".PadRight(43);

                                    var suffix =
                                        $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
                                            .PadRight(35);

                                    _logger.Info($"     {prefix} {suffix} ==> {propertyName.Value}");
                                }

                                if (propCount == 0)
                                {
                                    _logger.Warn("     (Property store is empty)");
                                }
                            }

                            break;
                        case "SHELLBAG0X01":
                            var baaaa1f = shellBag as ShellBag0X01;
                            if (baaaa1f.DriveLetter.Length > 0)
                            {
                                _logger.Info($"  Drive letter: {baaaa1f.DriveLetter}");
                            }
                            break;
                        case "SHELLBAG0X1F":

                            var b1f = shellBag as ShellBag0X1F;

                            if (b1f.PropertyStore.Sheets.Count > 0)
                            {
                                _logger.Warn("  >> Property store (Format: GUID\\ID Description ==> Value)");
                                var propCount = 0;

                                foreach (var prop in b1f.PropertyStore.Sheets)
                                foreach (var propertyName in prop.PropertyNames)
                                {
                                    propCount += 1;

                                    var prefix = $"{prop.GUID}\\{propertyName.Key}".PadRight(43);

                                    var suffix =
                                        $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
                                            .PadRight(35);

                                    _logger.Info($"     {prefix} {suffix} ==> {propertyName.Value}");
                                }

                                if (propCount == 0)
                                {
                                    _logger.Warn("     (Property store is empty)");
                                }
                            }

                            break;
                        case "SHELLBAG0X2E":
                            break;
                        case "SHELLBAG0X2F":
                            var b2f = shellBag as ShellBag0X2F;

                            break;
                        case "SHELLBAG0X40":
                            break;
                        case "SHELLBAG0X61":

                            break;
                        case "SHELLBAG0X71":
                            var b71 = shellBag as ShellBag0X71;
                            if (b71.PropertyStore?.Sheets.Count > 0)
                            {
                                _logger.Fatal(
                                    "Property stores found! Please email lnk file to saericzimmerman@gmail.com so support can be added!!");
                            }

                            break;
                        case "SHELLBAG0X74":
                            var b74 = shellBag as ShellBag0X74;

                            if (b74.LastModificationTime.HasValue)
                            {
                                _logger.Info(
                                    $"    Modified: {b74.LastModificationTime.Value.ToString(dt)}");
                            }
                            else
                            {
                                _logger.Info($"    Modified:");
                            }

                            var extensionNumber74 = 0;
                            if (b74.ExtensionBlocks.Count > 0)
                            {
                                _logger.Info($"    Extension block count: {b74.ExtensionBlocks.Count:N0}");
                                _logger.Info("");
                                foreach (var extensionBlock in b74.ExtensionBlocks)
                                {
                                    _logger.Info(
                                        $"    --------- Block {extensionNumber74:N0} ({extensionBlock.GetType().Name}) ---------");
                                    if (extensionBlock is Beef0004)
                                    {
                                        var b4 = extensionBlock as Beef0004;

                                        _logger.Info($"    Long name: {b4.LongName}");
                                        if (b4.LocalisedName.Length > 0)
                                        {
                                            _logger.Info($"    Localized name: {b4.LocalisedName}");
                                        }

                                        if (b4.CreatedOnTime.HasValue)
                                        {
                                            _logger.Info(
                                                $"    Created: {b4.CreatedOnTime.Value.ToString(dt)}");
                                        }
                                        else
                                        {
                                            _logger.Info($"    Created:");
                                        }

                                        if (b4.LastAccessTime.HasValue)
                                        {
                                            _logger.Info(
                                                $"    Last access: {b4.LastAccessTime.Value.ToString(dt)}");
                                        }
                                        else
                                        {
                                            _logger.Info($"    Last access: ");
                                        }
                                        if (b4.MFTInformation.MFTEntryNumber > 0)
                                        {
                                            _logger.Info(
                                                $"    MFT entry/sequence #: {b4.MFTInformation.MFTEntryNumber}/{b4.MFTInformation.MFTSequenceNumber} (0x{b4.MFTInformation.MFTEntryNumber:X}/0x{b4.MFTInformation.MFTSequenceNumber:X})");
                                        }
                                    }
                                    else if (extensionBlock is Beef0025)
                                    {
                                        var b25 = extensionBlock as Beef0025;
                                        _logger.Info(
                                            $"    Filetime 1: {b25.FileTime1.Value.ToString(dt)}, Filetime 2: {b25.FileTime2.Value.ToString(dt)}");
                                    }
                                    else if (extensionBlock is Beef0003)
                                    {
                                        var b3 = extensionBlock as Beef0003;
                                        _logger.Info($"    GUID: {b3.GUID1} ({b3.GUID1Folder})");
                                    }
                                    else if (extensionBlock is Beef001a)
                                    {
                                        var b3 = extensionBlock as Beef001a;
                                        _logger.Info($"    File document type: {b3.FileDocumentTypeString}");
                                    }
                                    else
                                    {
                                        _logger.Info($"    {extensionBlock}");
                                    }

                                    extensionNumber74 += 1;
                                }
                            }
                            break;
                        case "SHELLBAG0XC3":
                            break;
                        case "SHELLBAGZIPCONTENTS":
                            break;
                        default:
                            _logger.Fatal(
                                $">> UNMAPPED Type! Please email lnk file to saericzimmerman@gmail.com so support can be added!");
                            _logger.Fatal($">>{shellBag}");
                            break;
                    }

                    _logger.Info("");
                }
                _logger.Error("--- End Target ID information ---");
            }


            if (lnk.ExtraBlocks.Count > 0)
            {
                _logger.Info("");
                _logger.Error("--- Extra blocks information ---");
                _logger.Info("");

                foreach (var extraDataBase in lnk.ExtraBlocks)
                {
                    switch (extraDataBase.GetType().Name)
                    {
                        case "ConsoleDataBlock":
                            var cdb = extraDataBase as ConsoleDataBlock;
                            _logger.Warn(">> Console data block");
                            _logger.Info($"   Fill Attributes: {cdb.FillAttributes}");
                            _logger.Info($"   Popup Attributes: {cdb.PopupFillAttributes}");
                            _logger.Info(
                                $"   Buffer Size (Width x Height): {cdb.ScreenWidthBufferSize} x {cdb.ScreenHeightBufferSize}");
                            _logger.Info(
                                $"   Window Size (Width x Height): {cdb.WindowWidth} x {cdb.WindowHeight}");
                            _logger.Info($"   Origin (X/Y): {cdb.WindowOriginX}/{cdb.WindowOriginY}");
                            _logger.Info($"   Font Size: {cdb.FontSize}");
                            _logger.Info($"   Is Bold: {cdb.IsBold}");
                            _logger.Info($"   Face Name: {cdb.FaceName}");
                            _logger.Info($"   Cursor Size: {cdb.CursorSize}");
                            _logger.Info($"   Is Full Screen: {cdb.IsFullScreen}");
                            _logger.Info($"   Is Quick Edit: {cdb.IsQuickEdit}");
                            _logger.Info($"   Is Insert Mode: {cdb.IsInsertMode}");
                            _logger.Info($"   Is Auto Positioned: {cdb.IsAutoPositioned}");
                            _logger.Info($"   History Buffer Size: {cdb.HistoryBufferSize}");
                            _logger.Info($"   History Buffer Count: {cdb.HistoryBufferCount}");
                            _logger.Info($"   History Duplicates Allowed: {cdb.HistoryDuplicatesAllowed}");
                            _logger.Info("");
                            break;
                        case "ConsoleFEDataBlock":
                            var cfedb = extraDataBase as ConsoleFeDataBlock;
                            _logger.Warn(">> Console FE data block");
                            _logger.Info($"   Code page: {cfedb.CodePage}");
                            _logger.Info("");
                            break;
                        case "DarwinDataBlock":
                            var ddb = extraDataBase as DarwinDataBlock;
                            _logger.Warn(">> Darwin data block");
                            _logger.Info($"   Application ID: {ddb.ApplicationIdentifierUnicode}");
                            _logger.Info("");
                            break;
                        case "EnvironmentVariableDataBlock":
                            var evdb = extraDataBase as EnvironmentVariableDataBlock;
                            _logger.Warn(">> Environment variable data block");
                            _logger.Info($"   Environment variables: {evdb.EnvironmentVariablesUnicode}");
                            _logger.Info("");
                            break;
                        case "IconEnvironmentDataBlock":
                            var iedb = extraDataBase as IconEnvironmentDataBlock;
                            _logger.Warn(">> Icon environment data block");
                            _logger.Info($"   Icon path: {iedb.IconPathUni}");
                            _logger.Info("");
                            break;
                        case "KnownFolderDataBlock":
                            var kfdb = extraDataBase as KnownFolderDataBlock;
                            _logger.Warn(">> Known folder data block");
                            _logger.Info(
                                $"   Known folder GUID: {kfdb.KnownFolderId} ==> {kfdb.KnownFolderName}");
                            _logger.Info("");
                            break;
                        case "PropertyStoreDataBlock":
                            var psdb = extraDataBase as PropertyStoreDataBlock;

                            if (psdb.PropertyStore.Sheets.Count > 0)
                            {
                                _logger.Warn(
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

                                    _logger.Info($"   {prefix} {suffix} ==> {propertyName.Value}");
                                }

                                if (propCount == 0)
                                {
                                    _logger.Warn("   (Property store is empty)");
                                }
                            }
                            _logger.Info("");
                            break;
                        case "ShimDataBlock":
                            var sdb = extraDataBase as ShimDataBlock;
                            _logger.Warn(">> Shimcache data block");
                            _logger.Info($"   LayerName: {sdb.LayerName}");
                            _logger.Info("");
                            break;
                        case "SpecialFolderDataBlock":
                            var sfdb = extraDataBase as SpecialFolderDataBlock;
                            _logger.Warn(">> Special folder data block");
                            _logger.Info($"   Special Folder ID: {sfdb.SpecialFolderId}");
                            _logger.Info("");
                            break;
                        case "TrackerDataBaseBlock":
                            var tdb = extraDataBase as TrackerDataBaseBlock;
                            _logger.Warn(">> Tracker database block");
                            _logger.Info($"   Machine ID: {tdb.MachineId}");
                            _logger.Info($"   MAC Address: {tdb.MacAddress}");
                            _logger.Info($"   MAC Vendor: {GetVendorFromMac(tdb.MacAddress)}");
                            _logger.Info(
                                $"   Creation: {tdb.CreationTime.ToString(dt)}");
                            _logger.Info("");
                            _logger.Info($"   Volume Droid: {tdb.VolumeDroid}");
                            _logger.Info($"   Volume Droid Birth: {tdb.VolumeDroidBirth}");
                            _logger.Info($"   File Droid: {tdb.FileDroid}");
                            _logger.Info($"   File Droid birth: {tdb.FileDroidBirth}");
                            _logger.Info("");
                            break;
                        case "VistaAndAboveIdListDataBlock":
                            var vdb = extraDataBase as VistaAndAboveIdListDataBlock;
                            _logger.Warn(">> Vista and above ID List data block");

                            foreach (var shellBag in vdb.TargetIDs)
                            {
                                var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;
                                _logger.Info($"   {shellBag.FriendlyName} ==> {val}");
                            }

                            _logger.Info("");
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

        private static CustomDestination ProcessCustomFile(string jlFile, bool q, string dt, bool ld)
        {
            if (q == false)
            {
                _logger.Warn($"Processing '{jlFile}'");
                _logger.Info("");
            }

            var sw = new Stopwatch();
            sw.Start();

            try
            {
                var customDest = JumpList.JumpList.LoadCustomJumplist(jlFile);

                if (q == false)
                {
                    _logger.Error($"Source file: {customDest.SourceFile}");

                    _logger.Info("");

                    _logger.Warn("--- AppId information ---");
                    _logger.Warn($"AppID: {customDest.AppId.AppId}, Description: {customDest.AppId.Description}");
                    _logger.Warn("--- DestList information ---");
                    _logger.Info($"  Entries:  {customDest.Entries.Count:N0}");
                    _logger.Info("");

                    var entryNum = 0;
                    foreach (var entry in customDest.Entries)
                    {
                        _logger.Warn($"  Entry #: {entryNum}, lnk count: {entry.LnkFiles.Count:N0} Rank: {entry.Rank}");

                        if (entry.Name.Length > 0)
                        {
                            _logger.Info($"   Name: {entry.Name}");
                        }

                        _logger.Info("");

                        var lnkCounter = 0;

                        foreach (var lnkFile in entry.LnkFiles)
                        {
                            var tc = lnkFile.Header.TargetCreationDate.Year == 1601
                                ? ""
                                : lnkFile.Header.TargetCreationDate.ToString(
                                    dt);
                            var tm = lnkFile.Header.TargetModificationDate.Year == 1601
                                ? ""
                                : lnkFile.Header.TargetModificationDate.ToString(
                                    dt);
                            var ta = lnkFile.Header.TargetLastAccessedDate.Year == 1601
                                ? ""
                                : lnkFile.Header.TargetLastAccessedDate.ToString(
                                    dt);


                            _logger.Warn($"--- Lnk #{lnkCounter:N0} information ---");
                            _logger.Info($"  Lnk target created: {tc}");
                            _logger.Info($"  Lnk target modified: {tm}");
                            _logger.Info($"  Lnk target accessed: {ta}");

                            if (ld)
                            {
                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasName) == Lnk.Header.DataFlag.HasName)
                                {
                                    _logger.Info($"  Name: {lnkFile.Name}");
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasRelativePath) ==
                                    Lnk.Header.DataFlag.HasRelativePath)
                                {
                                    _logger.Info($"  Relative Path: {lnkFile.RelativePath}");
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasWorkingDir) ==
                                    Lnk.Header.DataFlag.HasWorkingDir)
                                {
                                    _logger.Info($"  Working Directory: {lnkFile.WorkingDirectory}");
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) ==
                                    Lnk.Header.DataFlag.HasArguments)
                                {
                                    _logger.Info($"  Arguments: {lnkFile.Arguments}");
                                }

                                if ((lnkFile.Header.DataFlags & Lnk.Header.DataFlag.HasLinkInfo) ==
                                    Lnk.Header.DataFlag.HasLinkInfo)
                                {
                                    _logger.Info("");
                                    _logger.Error("--- Link information ---");
                                    _logger.Info($"Flags: {lnkFile.LocationFlags}");

                                    if (lnkFile.VolumeInfo != null)
                                    {
                                        _logger.Info("");
                                        _logger.Warn(">>Volume information");
                                        _logger.Info(
                                            $"  Drive type: {GetDescriptionFromEnumValue(lnkFile.VolumeInfo.DriveType)}");
                                        _logger.Info($"  Serial number: {lnkFile.VolumeInfo.VolumeSerialNumber}");

                                        var label = lnkFile.VolumeInfo.VolumeLabel.Length > 0
                                            ? lnkFile.VolumeInfo.VolumeLabel
                                            : "(No label)";

                                        _logger.Info($"  Label: {label}");
                                    }

                                    if (lnkFile.NetworkShareInfo != null)
                                    {
                                        _logger.Info("");
                                        _logger.Warn("  Network share information");

                                        if (lnkFile.NetworkShareInfo.DeviceName.Length > 0)
                                        {
                                            _logger.Info($"    Device name: {lnkFile.NetworkShareInfo.DeviceName}");
                                        }

                                        _logger.Info($"    Share name: {lnkFile.NetworkShareInfo.NetworkShareName}");

                                        _logger.Info(
                                            $"    Provider type: {lnkFile.NetworkShareInfo.NetworkProviderType}");
                                        _logger.Info($"    Share flags: {lnkFile.NetworkShareInfo.ShareFlags}");
                                        _logger.Info("");
                                    }

                                    if (lnkFile.LocalPath?.Length > 0)
                                    {
                                        _logger.Info($"  Local path: {lnkFile.LocalPath}");
                                    }

                                    if (lnkFile.CommonPath.Length > 0)
                                    {
                                        _logger.Info($"  Common path: {lnkFile.CommonPath}");
                                    }
                                }
                            }

                            if (lnkFile.TargetIDs.Count > 0)
                            {
                                _logger.Info("");

                                var absPath = string.Empty;

                                foreach (var shellBag in lnkFile.TargetIDs)
                                {
                                    absPath += shellBag.Value + @"\";
                                }

                                _logger.Info($"  Absolute path: {GetAbsolutePathFromTargetIDs(lnkFile.TargetIDs)}");
                                _logger.Info("");
                            }


                            lnkCounter += 1;
                        }
                        _logger.Info("");
                        entryNum += 1;
                    }
                }


                sw.Stop();

                if (q == false)
                {
                    _logger.Info("");
                }

                _logger.Info(
                    $"---------- Processed '{customDest.SourceFile}' in {sw.Elapsed.TotalSeconds:N8} seconds ----------");

                if (q == false)
                {
                    _logger.Info("\r\n");
                }

                return customDest;
            }

            catch (Exception ex)
            {
                _failedFiles.Add($"{jlFile} ==> ({ex.Message})");
                _logger.Fatal($"Error opening '{jlFile}'. Message: {ex.Message}");
                _logger.Info("");
            }

            return null;
        }

        private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private static void SetupNLog()
        {
            if (File.Exists( Path.Combine(BaseDirectory,"Nlog.config")))
            {
                return;
            }
            var config = new LoggingConfiguration();
            var loglevel = LogLevel.Info;

            var layout = @"${message}";

            var consoleTarget = new ColoredConsoleTarget();

            config.AddTarget("console", consoleTarget);

            consoleTarget.Layout = layout;

            var rule1 = new LoggingRule("*", loglevel, consoleTarget);
            config.LoggingRules.Add(rule1);

            LogManager.Configuration = config;
        }
    }

    public sealed class AutoCsvOut
    {
        //jump list info
        public string SourceFile { get; set; }
        public string SourceCreated { get; set; }
        public string SourceModified { get; set; }
        public string SourceAccessed { get; set; }
        public string AppId { get; set; }
        public string AppIdDescription { get; set; }

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
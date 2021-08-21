# JLECmd

## Command Line Interface

    JLECmd version 1.4.0.0
    
    Author: Eric Zimmerman (saericzimmerman@gmail.com)
    https://github.com/EricZimmerman/JLECmd
    
            d               Directory to recursively process. Either this or -f is required
            f               File to process. Either this or -d is required
            q               Only show the filename being processed vs all output. Useful to speed up exporting to json and/or csv. Default is FALSE
    
            all             Process all files in directory vs. only files matching *.automaticDestinations-ms or *.customDestinations-ms. Default is FALSE
    
            csv             Directory to save CSV formatted results to. Be sure to include the full path in double quotes
            csvf            File name to save CSV formatted results to. When present, overrides default name
    
            html            Directory to save xhtml formatted results to. Be sure to include the full path in double quotes
            json            Directory to save json representation to. Use --pretty for a more human readable layout
            pretty          When exporting to json, use a more human readable layout. Default is FALSE
    
            ld              Include more information about lnk files. Default is FALSE
            fd              Include full information about lnk files (Alternatively, dump lnk files using --dumpTo and process with LECmd). Default is FALSE
    
            appIds          Path to file containing AppIDs and descriptions (appid|description format). New appIds are added to the built-in list, existing appIds will have their descriptions updated
            dumpTo          Directory to save exported lnk files
            withDir         When true, show contents of Directory not accounted for in DestList entries
            Debug           Debug mode
    
            dt              The custom date/time format to use when displaying timestamps. See https://goo.gl/CNVq0k for options. Default is: yyyy-MM-dd HH:mm:ss
            mp              Display higher precision for timestamps. Default is FALSE
    
    Examples: JLECmd.exe -f "C:\Temp\f01b4d95cf55d32a.customDestinations-ms" --mp
              JLECmd.exe -f "C:\Temp\f01b4d95cf55d32a.automaticDestinations-ms" --json "D:\jsonOutput" --jsonpretty
              JLECmd.exe -d "C:\CustomDestinations" --csv "c:\temp" --html "c:\temp" -q
              JLECmd.exe -d "C:\Users\e\AppData\Roaming\Microsoft\Windows\Recent" --dt "ddd yyyy MM dd HH:mm:ss.fff"
    
              Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes  

## Documentation

Automatic and Custom Destinations jump list parser with Windows 10/11 support.

[Jump lists in depth: Understand the format to better understand what your tools are (or aren't) doing](https://binaryforay.blogspot.com/2016/02/jump-lists-in-depth-understand-format.html)
[Introducing JLECmd!](https://binaryforay.blogspot.com/2016/03/introducing-jlecmd.html)
[PECmd, LECmd, and JLECmd updated!](https://binaryforay.blogspot.com/2016/03/pecmd-lecmd-and-jlecmd-updated.html)
[LECmd and JLECmd updated](https://binaryforay.blogspot.com/2016/04/lecmd-and-jlecmd-updated.html)
[JLECmd v0.9.6.0 released](https://binaryforay.blogspot.com/2016/09/jlecmd-v0960-released.html)

# Download Eric Zimmerman's Tools

All of Eric Zimmerman's tools can be downloaded [here](https://ericzimmerman.github.io/#!index.md). Use the [Get-ZimmermanTools](https://f001.backblazeb2.com/file/EricZimmermanTools/Get-ZimmermanTools.zip) PowerShell script to automate the download and updating of the EZ Tools suite. Additionally, you can automate each of these tools using [KAPE](https://www.kroll.com/en/services/cyber-risk/incident-response-litigation-support/kroll-artifact-parser-extractor-kape)!

# Special Thanks

Open Source Development funding and support provided by the following contributors: [SANS Institute](http://sans.org/) and [SANS DFIR](http://dfir.sans.org/).

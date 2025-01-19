![NanoXLSX](https://raw.githubusercontent.com/rabanti-github/PicoXLSX/refs/heads/master/Documentation/icons/PicoXLSX.png)

# PicoXLSX

![nuget](https://img.shields.io/nuget/v/picoXLSX.svg?maxAge=86400)
![NuGet Downloads](https://img.shields.io/nuget/dt/PicoXLSX)
![GitHub License](https://img.shields.io/github/license/rabanti-github/PicoXLSX)
[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX.svg?type=shield)](https://app.fossa.io/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX?ref=badge_shield)

 PicoXLSX is a small .NET library written in C#, to create Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

* **Minimum of dependencies** (\*
* No need for an installation of Microsoft Office
* No need for Office interop libraries
* No need for 3rd party libraries
* No need for an installation of the Microsoft Open Office XML SDK (OOXML)

**Please have a look at the successor library [NanoXLSX](https://github.com/rabanti-github/NanoXLSX) for reader support.**

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch) 
See the **[Change Log](https://github.com/rabanti-github/PicoXLSX/blob/master/Changelog.md)** for recent updates.

## What's new in version 3.x

* Copy functions for worksheets
* Several additional checks, exception handling and updated documentation

Note: Most changes came from the rewritten [NanoXLSX](https://github.com/rabanti-github/NanoXLSX) library v2.0. Unit testing was also introduced there.
Therefore, the change list in PicoXLSX is not as long as in NanoXLSX, since many these changes are dealing with reader functionality. 

## Roadmap
Version 3.x of PicoXLSX was completely overhauled along with NanoXLSX v3.x.
However, v3.x it is not planned as a LTS version. The upcoming v4.x is supposed to introduce some important functions, like in-line cell formatting, better formula handling and additional worksheet features.
Furthermore, it is planned to introduce more modern OOXML features like the SHA256 implementation of worksheet passwords.
One of the main aspects of this upcoming version is the retirement of the original code base in favor of a facade, using NanoXLSX as single dependency. This will reduce the maintenance effort dramatically.


## Requirements

PicoXLSX was created with .NET version 4.5. Newer versions like 4.6 are working and tested. Furthermore, .NET Standard 2.0 is supported since v2.9. Older versions of.NET like 3.5 and 4.0 may also work with minor changes. Some functions introduced in .NET 4.5 were used and must be adapted in this case. 

### .NET 4.5 or newer

*)The only requirement to compile the library besides .NET (v4.5 or newer) is the assembly **WindowsBase**, as well as **System.IO.Compression**. These assemblies are **standard components in all Microsoft Windows systems** (except Windows RT systems). If your IDE of choice supports referencing assemblies from the Global Assembly Cache (**GAC**) of Windows, select WindowsBase and Compression from there. If you want so select the DLLs manually and Microsoft Visual Studio is installed on your system, the DLL of WindowsBase can be found most likely under "c:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll", as well as System.IO.Compression under "c:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.IO.Compression.dll". Otherwise you find them in the GAC, under "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\WindowsBase" and "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.IO.Compression"

The NuGet package **does not require dependencies**

### .NET Standard

.NET Standard v2.0 resolves the dependency System.IO.Compression automatically, using NuGet and does not rely anymore on WindowsBase in the development environment. In contrast to the .NET >=4.5 version, **no manually added dependencies necessary** (as assembly references) to compile the library.

Please note that the demo project of the .NET Standard version will not work in Visual Studio 2017. To get the build working, unload the demo project of the .NET Standard version.

### Documentation project

If you want to compile the documentation project (folder: Documentation; project file: shfbproj), you need also the **[Sandcastle Help File Builder (SHFB)](https://github.com/EWSoftware/SHFB)**. It is also freely available. But you don't need the documentation project to build the NanoXLSX library.

The .NET version of the documentation may vary, based on the installation. If v4.5 is not available, upgrade to target to a newer version, like v4.6

## Installation

### Using NuGet

By Package Manager (PM):

```sh
Install-Package PicoXLSX
```

By .NET CLI:

```sh
dotnet add package PicoXLSX
```

### As DLL

Simply place the PicoXLSX DLL into your .NET project and add a reference to it. Please keep in mind that the .NET version of your solution must match with the runtime version of the PicoXLSX DLL (currently compiled with 4.5 and .NET Standard 2.0).

### As source files

Place all .CS files from the PicoXLSX source folder into your project. You can place them into a sub-folder if you wish. The files contains definitions for workbooks, worksheets, cells, styles, meta-data, low level methods and exceptions. In case of the .NET >=4.5 version, the necessary dependencies have to be referenced as well.

## Usage

### Quick Start (shortened syntax)

```c#
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");         // Create new workbook with a worksheet called Sheet1
 workbook.WS.Value("Some Data");                                        // Add cell A1
 workbook.WS.Formula("=A1");                                            // Add formula to cell B1
 workbook.WS.Down();                                                    // Go to row 2
 workbook.WS.Value(DateTime.Now, Style.BasicStyles.Bold);               // Add formatted value to cell A2
 workbook.Save();                                                       // Save the workbook as myWorkbook.xlsx
```

### Quick Start (regular syntax)

```c#
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");         // Create new workbook with a worksheet called Sheet1
 workbook.CurrentWorksheet.AddNextCell("Some Data");                    // Add cell A1
 workbook.CurrentWorksheet.AddNextCell(42);                             // Add cell B1
 workbook.CurrentWorksheet.GoToNextRow();                               // Go to row 2
 workbook.CurrentWorksheet.AddNextCell(DateTime.Now);                   // Add cell A2
 workbook.Save();                                                       // Save the workbook as myWorkbook.xlsx
```

## Further References

See the full **API-Documentation** at: [https://rabanti-github.github.io/PicoXLSX/](https://rabanti-github.github.io/PicoXLSX/).

The [Demo project](https://github.com/rabanti-github/PicoXLSX/tree/master/Demo) contains 17 simple use cases. You can find also the full documentation in the [Documentation-Folder](https://github.com/rabanti-github/PicoXLSX/tree/master/docs) (html files or single chm file) or as C# documentation in the particular .CS files.

See also: [Getting started in the Wiki](https://github.com/rabanti-github/PicoXLSX/wiki/Getting-started)

## License

[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX.svg?type=large)](https://app.fossa.io/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX?ref=badge_large)

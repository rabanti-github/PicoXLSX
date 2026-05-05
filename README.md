![PicoXLSX](https://raw.githubusercontent.com/rabanti-github/PicoXLSX/refs/heads/master/PicoXLSX/PicoXLSX.png)

# PicoXLSX

![nuget](https://img.shields.io/nuget/v/picoXLSX.svg?maxAge=86400)
![NuGet Downloads](https://img.shields.io/nuget/dt/PicoXLSX)
![GitHub License](https://img.shields.io/github/license/rabanti-github/PicoXLSX)
[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX.svg?type=shield)](https://app.fossa.io/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX?ref=badge_shield)

 PicoXLSX is a small .NET library written in C#, to create Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

* :white_check_mark: **Minimum of dependencies** (\*
* :x: No need for an installation of Microsoft Office
* :x: No need for Office interop libraries
* :x: No need for proprietary 3rd party libraries
* :x: No need for an installation of the Microsoft Open Office XML SDK (OOXML)

:arrow_right: Please have a look at the **codebase library [NanoXLSX](https://github.com/rabanti-github/NanoXLSX)** for the full source code of PicoXLSX and NanoXLSX.

:globe_with_meridians: Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

:page_facing_up: See the **[Change Log](https://github.com/rabanti-github/PicoXLSX/blob/master/Changelog.md)** for recent updates.

## :package: Modules

PicoXLSX v4 is split into modular NuGet packages:

| Module | Status | Description |
|--------|--------|-------------|
| **[NanoXLSX.Core](https://www.nuget.org/packages/NanoXLSX.Core)** | :green_circle: Mandatory, Bundled | Core library with workbooks, worksheets, cells, styles. No external dependencies |
| **[NanoXLSX.Writer](https://www.nuget.org/packages/NanoXLSX.Writer)** |  :large_blue_circle: Optional, Bundled  | Extension methods to write/save XLSX files. Depends on Core |
| **[NanoXLSX.Formatting](https://www.nuget.org/packages/NanoXLSX.Formatting)** | :large_blue_circle: Optional, Bundled | In-line cell formatting (rich text). [External repo](https://github.com/rabanti-github/NanoXLSX.Formatting). Depends on Core |
| **[PicoXLSX](https://www.nuget.org/packages/PicoXLSX)** | :star: Meta-Package | Bundles all of the above. **Recommended for most users** |
| **[NanoXLSX.Reader](https://www.nuget.org/packages/NanoXLSX.Reader)** | :white_circle: Optional, Not bundled | Extension methods to read/load XLSX files. Depends on Core |

> **Note:** All bundled modules are included when you install the `PicoXLSX` meta-package. Optional, non-bundled modules will extend the functionality of PicoXLSX

For advanced scenarios, you can install only the specific packages you need (e.g. `NanoXLSX.Core` + `NanoXLSX.Reader` for read-only applications).


## :sparkles: What's new in version 4.x

PicoXLSX v4 is a major release with significant architectural changes:

* **Modular architecture** - Split into separate NuGet packages (Core, Reader, Writer, Formatting) with a plugin system
* **New Color system** - Unified `Color` class supporting RGB, ARGB, indexed, theme and system colors
* **Redesigned Font and Fill** - Font properties now use proper enums; Fill supports flexible color definitions with tint
* **PascalCase naming** - All enums and constants follow C# naming conventions
* **Immutable value types** - `Address` and `Range` structs are now immutable
* **In-line formatting** - Rich text cell formatting via the NanoXLSX.Formatting module
* **Utils reorganization** - `Utils` class split into `DataUtils`, `ParserUtils`, `Validators`

:warning: **Breaking changes from v3.x** - There are breaking changes between PicoXLSX v3.4.5 and v4.0.0, mostly related to namespace changes and renamed enum values. See the **[Migration Guide](MigrationGuide.md)** for detailed upgrade instructions.

## :world_map: Roadmap

PicoXLSX v4.x (NanoXLSX v3.x) is planned as the **long-term supported version**. Possible future enhancements include:

* :lock: Modern password handling (e.g. SHA-256 for worksheet protection)
* :art: Auto-formatting capabilities
* :1234: Formula assistant for easier formula creation
* :paintbrush: Modern Style builder API
* :speech_balloon: Support for cell comments
* :framed_picture: Embedded images and charts
* :rocket: Performance optimizations


## :robot: For AI Agents
For AI agents and LLM tooling, a machine-readable [`llms.txt`](llms.txt) is available.
It lists all packages, installation commands, API documentation, source repositories, and a quick-start code snippet.

## :gear: Requirements

The library is currently on compatibility level with .NET version 4.5 and .NET Standard 2.0. Newer versions should of course work as well. Older versions, like .NET 3.5 have only limited support, since newer language features were used.

### .NET 4.5 or newer

\*)The only requirement to compile the library besides .NET (v4.5 or newer) is the assembly **WindowsBase**, as well as **System.IO.Compression**. These assemblies are **standard components in all Microsoft Windows systems** (except Windows RT systems). If your IDE of choice supports referencing assemblies from the Global Assembly Cache (**GAC**) of Windows, select WindowsBase and Compression from there. If you want so select the DLLs manually and Microsoft Visual Studio is installed on your system, the DLL of WindowsBase can be found most likely under "c:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll", as well as System.IO.Compression under "c:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.IO.Compression.dll". Otherwise you find them in the GAC, under "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\WindowsBase" and "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.IO.Compression"

The NuGet package **does not require dependencies**

### .NET Standard

.NET Standard v2.0 resolves the dependency System.IO.Compression automatically, using NuGet and does not rely anymore on WindowsBase in the development environment. In contrast to the .NET >=4.5 version, **no manually added dependencies necessary** (as assembly references) to compile the library.

## :hammer_and_wrench: Development and Testing

The full source code of PicoXLSX is available in the [NanoXLSX repository](https://github.com/rabanti-github/NanoXLSX). There are also thousands of unit tests available, ensuring a very high code coverage and security against unwanted side effects on changes.

## :inbox_tray: Installation

### Using NuGet (recommended)

By package Manager (PM):

```sh
Install-Package PicoXLSX
```

By .NET CLI:

```sh
dotnet add package PicoXLSX
```

:information_source: **Note**: Other methods like adding DLLs or source files directly into your project are technically still possible, but **not recommended** anymore. Use dependency management, whenever possible

## :bulb: Usage

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

## :link: Further References

See the full **API-Documentation** at: [https://rabanti-github.github.io/PicoXLSX/](https://rabanti-github.github.io/PicoXLSX/).

The **[Demo Project](https://github.com/rabanti-github/NanoXLSX.Demo)** contains 25 examples covering various use cases. The demo project is maintained in a separate repository.
See the section **[PicoXLSX](https://github.com/rabanti-github/NanoXLSX.Demo/tree/main/PicoXLSX)** for the specific examples related to PicoXLSX.

See also: [Getting started in the Wiki](https://github.com/rabanti-github/PicoXLSX/wiki/Getting-started)

Hint: You will find most certainly any function, and the way how to use it, in the [Unit Test Project of nanoXLSX](https://github.com/rabanti-github/NanoXLSX/tree/master/NanoXlsx%20Test)

## :balance_scale: License

PicoXLSX is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

This library claims to be free of any dependencies on proprietary software or libraries.
[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX.svg?type=large)](https://app.fossa.io/projects/git%2Bgithub.com%2Frabanti-github%2FPicoXLSX?ref=badge_large)

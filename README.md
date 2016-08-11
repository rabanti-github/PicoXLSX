# PicoXLSX
PicoXLSX is a small .NET / C# library to create XLSX files (Microsoft Excel 2007 or newer) in an easy and native way

* No need for an installation of Microsoft Office
* No need for Office interop libraries
* No need for 3rd party libraries
* No need for an installation of the Microsoft Open Office XML SDK (OOXML)

# Requirements
PicoXLSX was created with .NET version 4.5. But older versions like 3.5 and 4.0 may also work with minor or no changes. However, this was not tested yet.
The only requirement to compile the library besides .NET is the assembly **WindowsBase**. This assembly is a standard component in all Microsoft Windows systems (except Windows RT systems) and can be found most likely under "c:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll", according this [MSDN Blog entry](http://blogs.msdn.com/b/dmahugh/archive/2006/12/14/finding-windowsbase-dll.aspx).<br><br>
If you want to compile the documentation project (folder: Documentation; project file: shfbproj), you need also the **[Sandcastle Help File Builder (SHFB)](https://github.com/EWSoftware/SHFB)**. It is also freely available. But you don't need the documentation project to build the PicoXLSX library.

# Installation
## As DLL
Simply place the PicoXLSX DLL into your .NET project and make a reference (in VS or SharpDevelop) to it
## As source files
Place all .CS files from the PicoXLSX source folder into your project. You can place them into a sub-folder if you wish. The files contains definitions for workbooks, worksheets, cells, styles, meta-data, low level methods and exceptions.

# Usage
The [Demo project](https://github.com/rabanti-github/PicoXLSX/tree/master/Demo) contains nine simple use cases. You can find also the full documentation in the [Documentation-Folder](https://github.com/rabanti-github/PicoXLSX/tree/master/Documentation) or as C# documentation in the .CS files.<br>
See also: [Getting started in the Wiki](https://github.com/rabanti-github/PicoXLSX/wiki/Getting-started)

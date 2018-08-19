# Change Log

## v2.5.1

---
Release Date: **19.08.2018**

- Fixed a bug in the Font style class
- Fixed typos


## v2.5.0

---
Release Date: **02.07.2018**

- Added address types (no fixed rows and columns, fixed rows, fixed columns, fixed rows and columns; Useful in formulas)
- Added new CellDirection Disabled, if the addresses of the cells are defined manually (AddNextCell will override the current cell in this case)
- Altered Demo 3 to demonstrate disabling of automatic cell addressing
- Extended Demo 1 to demonstrate the new address types
- Minor, internal changes


## v2.4.0
---
Release Date: **07.06.2018**

- Added style appending (builder / method chaining)
- Added new basic styles ColorizedText, ColorizedBackground and Font as functions
- Added a new constructor for Workbooks without file name to handle stream-only workbooks more logical
- Added the functions HasCell, GetLastColumnNumber and GetLastRowNumber in the Worksheet class
- Fixed a bug when overriding a worksheet name with sanitizing
- Added new demo for the introduced style features
- Internal optimizations and fixes


## v2.3.2
---
Release Date: **30.05.2018**

- Fixed a bug in the processing of column widths. Bug fix provided by Johan Lindvall
- Added numeric data types byte, sbyte, decimal, uint, ulong and short and ushort (proposal by Johan Lindvall)
- Changed the behavior of cell type casting. User defined cell types will now only be overwritten if the type is DEFAULT (proposal by Johan Lindvall)


## v2.3.1
---
Release Date: **12.03.2018**

**Note**: Due to some refactoring (see below) in this version, changes of existing code may be necessary. However, most introduced changes are on a rather low level and probably only used internally although publicly accessible

- Renamed the method addStyleComponent  in the class Workbook to AddStyleComponent, to follow conventions
- Renamed the Properties RowAddress and ColumnAddress to RowNumber and ColumnNumber in the class Cell for clarity
- Renamed the methods GetCurrentColumnAddress, GetCurrentRowAddress, SetCurrentColumnAddress and SetCurrentRowAddress in the class Worksheet to GetCurrentColumnNumber, GetCurrentRowNumber, SetCurrentColumnNumber and SetCurrentRowNumber for clarity
- Renamed the constants MIN_ROW_ADDRESS, MAX_ROW_ADDRESS, MIN_COLUMN_ADDRESS, MAX_COLUMN_ADDRESS in the class Worksheet to MIN_ROW_NUMBER, MAX_ROW_NUMBER, MIN_COLUMN_NUMBER, MAX_COLUMN_NUMBER for clarity
- Many optimizations
- Fixed typos
- Extensive documentation update


## v2.3.0
---
Release Date: **09.02.2018**

- Added most important formulas as static method calls in the sub-class Cell.BasicFormulas (round, floor, ceil, min, max, average, median, sum, vlookup)
- Removed overloaded methods to add cells as type Cell. This can be done now with the overloading of the type object (no code changes necessary)
- Added new constructors for Address and Range
- Added demo for the new static formula methods
- Fixed a bug that kept the stream of the saved file open as long as the program was running 
- Minor bug fixes
- Documentation update

## v2.2.0
---
Release Date: **10.12.2017**

- Added Shortener class (WS) in workbook for quicker writing of data / formulas
- Updated descriptions
- Added full NuGet support

## v2.1.1
---
Release Date: **07.12.2017**

- Documentation Update
- Fixed version number of the assembly
- Preparation for NuGet release

## v2.1.0
---
Release Date: **03.12.2017**

- Pushed back to .NET 4.5 due to platform compatibility reasons (multi platform architecture is planned)
- Added SaveToStream method in Workbook class
- Added demo for the new stream save method
- Changed log to MD format

## v2.0.0
---
Release Date: **01.11.2017**

**Note**: This major version is not compatible with code of v1.x. However, porting is feasible with moderate effort
- Complete replacement of style handling
- Added an option to add styles with the cell values in one step
- Added a sanitizing function for worksheet names (with auto-sanitizing when adding a worksheet as option)
- Changed specific exception to general exceptions (e.g. StyleException, FormatException or WorksheetException)
- Added function to retrieve cell values easier
- Added functions to get the current column or row number
- Many internal optimizations
- Added more documentation
- Added new functionality to the demos

## v1.6.3
---
Release Date: **24.08.2017**

- Added further null checks
- Minor optimizations
- Fixed typos

## v1.6.2
---
Release Date: **12.08.2017**

- fixed a bug in the function to remove merged cells (Worksheet class)
- Fixed typos

## v1.6.1
---
Release Date: **08.08.2017**

**Note**: Due to a (now fixed) typo in a public parameter name, it is possible that some function calls on existing code must be fixed too (just renaming).
- Fixed typos (parameters and text)
- Minor optimization
- Complete reformatting of code (alphabetical order & regions)
- HTML documentation moved to folder 'docs' to provide an automatic API documentation on the hosting platform

## v1.6.0 
---
Release Date: **07.04.2017**

**Note**: Using this version of the library with old code can cause compatibility issues due to the simplification of some methods (see below).
- Simplified all style assignment methods. Referencing of the workbook is not necessary anymore (can cause compatibility issues with existing code; just remove the workbook references)
- Removed SetCellAddress Method. Replaced by Getters and Setters
- Due to the impossibility of overloading getters and setters, two getters and setters "CellAddress"(string-based) and "CellAddress2" (number based) are introduced
- Fixed a bug in the handling of style assignment
- Additional checks in the assignment methods for columns and rows
- Minor changes (code and documentation)

## v1.5.6 
---
Release Date: **01.04.2017**

- Fixed a bug induced by non-Gregorian calendars (e.g Minguo, Heisei period, Hebrew) on the host system
- Code cleanup
- Minor bug fixes
- Fixed typos

## v1.5.5 
---
Release Date: **24.03.2017**

- Fixed a Out-of-Memory bug when saving very big files
- Improved the performance of the save() method (reduction of processing time from minutes to second when handling big amount of data)
- Added a debug and release version of the executable
- Added some testing utils in the demo project (tests are currently commented out)

## v1.5.4 
---
Release Date: **20.03.2017**

- Extended the sanitizing of allowed XML characters according the XML specifications to avoid errors with illegal characters in passed strings
- Updated project settings of the documentation solution
- Fixed typos

## v1.5.3 
---
Release Date: **17.11.2016**

- Fixed general bug in the handling of the sharedStrings table. Please update
- Passed null values to cells are now interpreted as empty values. Caused an exception until now

## v1.5.2 
---
Release Date: **15.11.2016**

- Fixed a bug in the sharedStrings table

## v1.5.1 
---
Release Date: **16.08.2016**

- Fixed a bug in the cell type resolution / formatting assignment

## v1.5.0 
---
Release Date: **12.08.2016**

**Note**: Using this version of the library with old code can cause compatibility issues due to the removal of some methods (see below).
- Removed all overloaded methods with various input values for adding cells. Object is sufficient
- Added sharedStrings table to manage strings more efficient (Excel standard)
- Pushed solution to .NET 4.6.1 (no changes necessary)
- Changed demos according to removed overloaded methods (List&lt;string&gt; is now List&lt;object&gt;)
- Added support for long (64bit) data type
- Fixed a bug in the type recognition of cells
  

## v1.4.0
---
Release Date: **11.08.2016**

- Added support for Cell selection
- Added support for worksheet selection
- Removed XML namespace 'x' as prefix in OOXML output. No use for this at the moment
- Removed newlines from OOXML output. No relevance for parser
- Added further demo for the new features

## v1.3.1
---
Release Date: **18.01.2016**

- Fixed a bug in the auto filter section
- Code cleanup
- Fixed some documentation issues

## v1.3.0
---
Release Date: **17.01.2016**

- Added support for auto filter (columns)
- Added support for hiding columns and rows
- Added new Column class (sub-class of Worksheet) to manage column based properties more efficiently
- Removed unused Exception UnsupportedDataTypeException
- Added more documentation (exceptions are now better defined)
- Minor bug fixes + typos 
- Added further demo for the new features

## v1.2.4
---
Release Date: **08.11.2015**

- Fixed a bug in the meta data section


## v1.2.3
---
Release Date: **02.11.2015**

- Added support for protecting workbooks
- Minor bug fixes

## v1.2.2
---
Release Date: **01.11.2015**

- Added support to protect worksheets with a password
- Minor bug fixes
- Fixed some code formatting issues
- Fixed issue in documentation project to include the private LowLevel class (and sub classes)

## v1.2.1
---
Release Date: **31.10.2015**

- Fixed typos (in parameter names)

## v1.2.0
---
Release Date: **29.10.2015**

- Added support for merging cells
- Added support for Protecting worksheets (no support for passwords yet)
- Minor bug fixes
- Fixed typos
- Added more documentation
- Added further demo for the new features

## v1.1.2
---
Release Date: **12.10.2015**

- Added a method to generate random style names using a Crypto Service Provider. Fixed a problem of identical style names due to too fast processing when using a standard RNG
- Minor bug fixes
- Fixed typos

## v1.1.1
---
Release Date: **06.10.2015**

- Minor bug fixes

## v1.1.0
---
Release Date: **29.09.2015**

- Added extensive support for styling
- Added support for meta data (title, subject etc.)
- Added support for cell width and cell height
- Fixed many spelling errors
- Added new methods to add cells as objects with automatic casting
- Added more documentation
- Added post build macros to VS project for easier deployment
- Many bug fixes and optimizations

**Note**: Styling is not complete and fully tested yet. Several additions and changes are possible with the next version

## v.1.0.1 
---
Release Date: **22.08.2015**

- Fixed uncritical / silent casting exception

## v1.0.0
---
Release Date: **21.08.2015**

- Initial release

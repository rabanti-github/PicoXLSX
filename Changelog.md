# Change Log

## v3.0.2

---
Release Date: **05.10.2022**

- Minor adaptions
- Code formatting and maintenance

## v3.0.1

---
Release Date: **01.10.2022**

- Fixed a bug in the functions to write custom number formats
- Fixed behavior of empty cells and added re-evaluation if values are set by the Value property
- Fixed a bug in the functions to write font values (styles)
- Updated documentation

Note: 
- When defining a custom number format, now the CustomFormatCode property must always be defined as well, since an empty value leads to an invalid Workbook 
- When a cell is now created (by constructor) with the type EMPTY, any passed value will be discarded in this cell

## v3.0.0

---
Release Date: **03.09.2022 - Major Release**

Note: some of the mentioned changes may be already implemented in v2.x as preview functionality

### Workbook and Shortener

- Added a list of MRU colors that can be defined in the Workbook class (methods AddMruColor, ClearMruColors)
- Added an exposed property for the workbook protection password hash (will be filled when loading a workbook)
- Added the method SetSelectedWorksheet by name in the Workbook class
- Added two methods GetWorksheet by name or index in the Workbook class
- Added the methods CopyWorksheetIntoThis and CopyWorksheetTo with several overloads in the Workbook class
- Added the function RemoveWorksheet by index with the option of resetting the current worksheet, in the Workbook class
- Added the function SetCurrentWorksheet by index in the Workbook class
- Added the function SetSelectedWorksheet by name in the Workbook class
- Added a Shortener-Class constructor with a workbook reference
- The shortener functions Down and Right have now an option to keep row and column positions
- Added two shortener functions Up and Left
- Made several style assigning methods deprecated in the Workbook class (will be removed in future versions)

### Worksheet

- Added an exposed property for the worksheet protection password hash (will be filled when loading a workbook)
- Added the methods GetFirstDataColumnNumber, GetFirstDataColumnNumber, GetFirstDataRowNumber, GetFirstRowNumber, GetLastDataColumnNumber, GetFirstCellAddress, GetFirstDataCellAddress, GetLastDataColumnNumber, GetLastDataRowNumber, GetLastRowNumber, GetLastCellAddress,  GetLastCellAddress and GetLastDataCellAddress
- Added the methods GetRow and GetColumns by address string or index
- Added the method Copy to copy a worksheet (deep copy)
- Added a constructor with only the worksheet name as parameter
- Added and option in GoToNextColumn and GoToNextRow to either keep the current row or column
- Added the methods RemoveRowHeight and RemoveAllowedActionOnSheetProtection
- Renamed columnAddress and rowAddress to columnNumber and rowNumber in the AddNextCell, AddCellFormula and RemoveCell methods
- Added several validations for worksheet data

### Cells, Rows and Columns

- In Cell, the address can now have reference modifiers ($)
- The worksheet reference in the Cell constructor was removed. Assigning to a worksheet is now managed automatically by the worksheet when adding a cell
- Added a property CellAddressType in Cell
- Cells can now have null as value, interpreted as empty
- Added a new overloaded function ResolveCellCoordinate to resolve the address type as well
- Added ValidateColumnNumber and ValidateRowNumber in Cell
- In Address, the constructor with string and address type now only needs a string, since reference modifiers ($) are resolved automatically
- Address objects are now comparable
- Implemented better address validation
- Range start and end addresses are swapped automatically, if reversed

### Styles

- Font has now an enum of possible underline values (e.g. instead of a bool)
- CellXf supports now indentation
- A new, internal style repository was introduced to streamline the style management
- Color (RGB) values are now validated (Fill class has a function ValidateColor)
- Style components have now more appropriate default values
- MRU colors are now not collected from defined style colors but from the MRU list in the workbook object
- The ToString function of Styles and all sub parts will now give a complete outline of all elements
- Fixed several issues with style comparison
- Several style default values were introduced as constants

### Formulas

- Added uint as possible formula value. Valid types are int, uint, long, ulong, float, double, byte, sbyte, decimal, short and ushort
- Added several validity checks


### Misc
- Added several constants for boundary dates in the LowLevel class
- Added several functions for pane splitting in the LowLevel class
- Exposed the (legacy) password generation method in the LowLevel class
- Updated documentation among the whole project
- Exceptions have no sub-tiles anymore
- Overhauled the whole writer
- Removed lot of dead code for better maintenance

## v2.11.5

---
Release Date: **06.08.2022**

- Fixed a bug when setting a workbook protection password

## v2.11.4

---
Release Date: **27.03.2022**

- Fixed a follow-up issue on finding first/last cell addresses on explicitly defined, empty cells

## v2.11.3

---
Release Date: **20.03.2022**

- Fixed a regression bug, caused by changes of v2.11.2

## v2.11.2

---
Release Date: **10.03.2022**

- Added functions to determine the first cell address, column number or row number of a worksheet
- Adapted internal style handling
- Adapted the internal building of XML documents
- Fixed a bug in the handling of border colors

## v2.11.1

---
Release Date: **12.09.2021**

- Added methods to remove styles from cell ranges
- Added missing methods to remove row height definitions, to reset column definitions and to remove actions for the worksheet protection
- Introduced general methods to validate row and column numbers (Cell class)
- Added checks for worksheet IDs and column widths
- Introduced (internal) style handling by a repository, in favor of the old style manager
- Fixed a bug in the handling of white spaces and newlines in string values
- Fixed a bug when hiding worksheets (Note: It is not possible anymore to remove all worksheets from a workbook, or to set a hidden one as active. This would lead to an invalid Excel file)
- Fixed a bug in the hash code generation in the Style class
- Fixed date and time handling in OADates
- Fixed implementation of GetLastRowNumber and GetLastColumnNumber (Worksheet)
- Fixed handling of cell coordinate resolution (accepts now addresses with fixed rows and columns)
- Removed worksheet reference from the Cell object
- Adapted Exceptions
- Improved the handling of worksheet removal
- Improved shortener and added option to keep the row number when jumping to the next column (same for next row)
- Documentation update
- Many optimizations and minor bug fixes

Note: This patch release contains many fixes, optimizations and features as preview of the next planned minor release v2.12 or mayor release v3.0.
These features are already available due to a recently reported bug and its fix, where pending changes were already published upstream (dev channel).


## v2.11.0

---
Release Date: **10.07.2020**

- Added functions to split (and freeze) a worksheet horizontally and vertically into panes
- Added a property to set the visibility of a workbook
- Added a property to set the visibility of worksheets
- Added two examples in the demo for the introduced split, freeze and visibility functionalities
- Added the possibility to define column widths and row height even if there are no cells defined
- Fixed the internal representation of column widths and row heights
- Minor code maintenance

Note: The column widths and row heights may change slightly with this release, since now the actual (internal) width and height is applied when setting a non-standard column width or row height

## v2.10.0

---
Release Date: **06.06.2020**

- Added functions to determine the last row, column or cell with data
- Fixed documentation formatting issues
- Updated readme and documentation

## v2.9.0

---
Release Date: **18.04.2021**

- Introduced library version for .NET Standard 2.0 (and assigned demos)
- Updated project structure (two projects for .NET >=4.5 and two for .NET Standard 2.0)
- Added function SetStyle in the Worksheet class
- Added demo for the new SetStyle function
- Changed behavior of empty cells. They are now not string but implicit numeric cells
- Added new function ResolveEnclosedAddresses in Cell.Range class
- Added new function GetAddressScope in Cell class
- Fixed the validation of cell addresses (single cell)
- Introduced several generalizations of Lists

Thanks to the following people for their contributions in NanoXLSX, that are based on the above changes:

- Shobb for the introduction of IReadOnlyList (generalizations)
- John Lenz for the port to .NET Standard
- Ned Marinov for the proposal of the new SetStyle function

## v2.8.1

---
Release Date: **10.12.2020**

- Formal update for NuGet. Fixed wrong readme

Note: No release will be published for this version, only a Nuget package

## v2.8.0

---
Release Date: **10.12.2020**

- Added indentation property of horizontal text alignment (CellXF) as style 
- Added example in demo for text indentation
- Code Cleanup

## v2.7.0

---
Release Date: **30.08.2020**

- Added new data type TIME, represented by TimeSpan objects
- Added time (TimeSpan) examples to the demos
- Added a check to ensure dates are not set beyond 9999-12-31 (limitation of OAdate)
- Updated documentation
- Fixed some code formatting issues

## v2.6.6

---
Release Date: **19.07.2020**

- Fixed a bug in the method AddNextCellFormula (Fix provided by Thiago Souza)

## v2.6.5

---
Release Date: **11.01.2020**

- Fixed a potential bug when parsing numbers (using certain locales)
- Formal changes

## v2.6.4

---
Release Date: **20.05.2019**

- Fixed a bug in the handling of streams (streams can be left open now)
- Updated stream demo
- Code Cleanup
- Removed executable folder, since executables are available through releases, compilation or NuGet

## v2.6.3

---
Release Date: **08.12.2018**

- Improved the performance of adding stylized cells by factor 10 to 100

## v2.6.2

---
Release Date: **04.11.2018**

- Fixed a bug in the style handling of merged cells. Bug fix provided by David Courtel

## v2.6.1

---
Release Date: **06.10.2018**

- Fixed a bug in the demo for the async handling
- Removed redundant code

## v2.6.0

---
Release Date: **04.10.2018**

- Added asynchronous methods SaveAsync, SaveAsAsync and SaveAsStreamAsync
- Added a new constructor in the Cell class with the address as string
- Added a new example for the introduced async methods
- Minor bug fixes and optimizations
- Fixed typos

## v2.5.1

---
Release Date: **19.08.2018**

- Fixed a bug in the Font style class
- Fixed typos

## v2.5.0

---
Release Date: **02.07.2018**

- Added address types (no fixed rows and columns, fixed rows, fixed columns, fixed rows and columns; Useful in formulas)
- Added new option CellDirection Disabled, if the addresses of the cells are defined manually (AddNextCell will override the current cell in this case)
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

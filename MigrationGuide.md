# Migration Guide v3.x to v4.0.0

## Introduction

PicoXLSX is now a subset of [NanoXLSX](https://github.com/rabanti-github/NanoXLSX) and therefore underwent some major changes. Since the used dependencies of NanoXLSX (v3.0.0) was also changed quite drastically between v2.x and 3.x (improved usability, consistency and modularization), these changes are reflecting now in PicoXLSX v4.x. 
There are some breaking changes, that are summarized in this migration guide.

## General changes

1. The namespace `PicoXLSX` was changed to `NanoXLSX`, to reflect the integration of PicoXLSX into NanoXLSX. Therefore, all occurrences of `PicoXLSX` in the code have to be replaced by `NanoXLSX`.
2. All sub-classes, like `Cell.Address`, `Style.Font`, etc. are now independent classes. Therefore, all occurrences of these classes have to be adapted in the code, e.g. instead of `Style.Font font = new Font();`, `Font font = new Font();` is used. Additionally, some namespaces, like `NanoXLSX.Styles`, have to be added as usings.
3. All constanst and enum values were renamed to follow the C# naming conventions (PascalCase for public members). Therefore, all occurrences of these constants and enum values have to be adapted in the code.

Important to note is, that PicoXLSX can now be extended by adding further NanoXLSX package references. For instance, if `NanoXLSX.Reader` is added as NuGet reference to PicoXLSX, it gets reader functionalities and has basically the same scope as NanoXLSX (Meta-Package).

The following guide lists all necessary changes, as well as changed behaviors of properties and methods.

---

## Core classes

### Workbook

  - The method `Workbook.AddStyle(Style)` was completely removed, after marked as obsolete in version 2.x. Styles should be added directly to cells or ranges. 
  - The method `Workbook.AddStyleComponent(Style, AbstractStyle)` was completely removed, after marked as obsolete in version 2.x. Styles should be modified directly on cells, e.g. `workbook.CurrentWorksheet.Cells["A1"].CellStyle.CurrentFont.Bold = true;` or `workbook.CurrentWorksheet.Cells["A1"].CellStyle.Append(fontStyle)`. 
  - The methods `Workbook.RemoveStyle(Style)`, `Workbook.RemoveStyle(Style, bool)`, `Workbook.RemoveStyle(string)` and `Workbook.RemoveStyle(string, bool)` were completely removed, after marked as obsolete in version 2.x. Styles should be removed directly from cells (e.g. `workbook.CurrentWorksheet.Cells["A1"].RemoveStyle()`.

---

### Worksheet

  - The method `Worksheet.SetSelectedCells(Range)`(and overloads) was replaced by three methods: `Worksheet.AddSelectedCell(Range)`, `Worksheet.AddSelectedCells(String)` (address or range), `Worksheet.AddSelectedCells(Address)` and `Worksheet.AddSelectedCells(Address, Address)`.
  - The method `Worksheet.RemoveSelectedCells()` was replaced by `Worksheet.ClearSelectedCells()`. 
  - The methods `Worksheet.RemoveSelectedCells(Range)`, `Worksheet.RemoveSelectedCells(String)` (address or range), `Worksheet.RemoveSelectedCells(Address)` and `Worksheet.RemoveSelectedCells(Address, Address)` were introduced to remove specific selected cells or ranges from the selection.
  - The property `Worksheet.SelectedCells` was changed from type `Range` to `List<Range>`, to allow multiple selected ranges in a worksheet.
  - The property `Workshet.SelectedCellRanges` was replaced by `Worksheet.SelectedCells`.
  - The property `Workshet.SelectedCellRange` was replaced by `Worksheet.SelectedCells`, where the value would be the only entry in the list.
  - The property `Worksheet.SheetProtectionPassword` was changed from type `string` to the interface `IPassword` (namespace `NanoXLSX.Interfaces`), by default implemented by `LegacyPassword`. `IPassword` contains several methods to get or set a password, and to  get or set its hash value. The property is instantiated by default on every worksheet. The property has to be adapted in the code, e.g. `string plainTextPassword = worksheet.SheetProtectionPassword.GetPassword()`. 
  - The property `Worksheet.SheetProtectionPasswordHash` was removed, as the password handling was changed. The property can now be found in the property: `Worksheet.SheetProtectionPassword.PasswordHash`). 
  
  - The following property and method behaviors have changed:

| Property/Method               | Old Behavior                | New Behavior             |
|-------------------------------|-----------------------------|--------------------------|
| `Worksheet.SelectedCells`     | List was null by default    | List is empty by default |
| `Worksheet.SelectedCells`     | Every added range was present | If ranges are overlapping, or even enclosed in other ranges, the ranges are automatically recalculated, so that every address only occurs in one particular range |
| `Worksheet.SetSelectedCells(string)` | Method accepted ranges or addresses with fixed cell ranges (e.g. "$A$1:$R$1") | Method transforms addresses or ranges automatically in neutral addresses or ranges (e.g. "A1:R1") |
| `Worksheet.SheetProtectionPassword`  | Was null (string) by default         | Is instantiated by default with a structured object of the type `LegacyPassword`  |

- The public constant values of the `Worksheet` class were renamed, according to the following overview:

| Old Constant                | New Constant             |
|-----------------------------|---------------------------
| `MAX_WORKSHEET_NAME_LENGTH` | `MaxWorksheetNameLength` |
| `DEFAULT_COLUMN_WIDTH`      | `DefaultWorksheetColumnWidth`     |
| `DEFAULT_ROW_HEIGHT`        | `DefaultWorksheetRowHeight`       |
| `MAX_COLUMN_NUMBER`         | `MaxColumnNumber`        |
| `MIN_COLUMN_NUMBER`         | `MinColumnNumber`        |
| `MIN_COLUMN_WIDTH`          | `MinColumnWidth`         |
| `MIN_ROW_HEIGHT`            | `MinRowHeight`           |
| `MAX_COLUMN_WIDTH`          | `MaxColumnWidth`         |
| `MAX_ROW_NUMBER`            | `MaxRowNumber`           |
| `MIN_ROW_NUMBER`            | `MinRowNumber`           |
| `MAX_ROW_HEIGHT`            | `MaxRowHeight`           |
| `AUTO_ZOOM_FACTOR`          | `AutoZoomFactor`         |
| `MIN_ZOOM_FACTOR`           | `MinZoomFactor`          |
| `MAX_ZOOM_FACTOR`           | `MaxZoomFactor`          |

- The enum values of `Worksheet.SheetProtectionValue` were renamed, according to the following overview:

| Old Enum Value                              | New Enum Value                  | Remarks |
| ------------------------------------------- | ------------------------------- | ------- |
| `SheetProtectionValue.objects`              | `SheetProtectionValue.Objects`             |         |
| `SheetProtectionValue.scenarios`            | `SheetProtectionValue.Scenarios`           |         |
| `SheetProtectionValue.formatCells`          | `SheetProtectionValue.FormatCells`         |         |
| `SheetProtectionValue.formatColumns`        | `SheetProtectionValue.FormatColumns`       |         |
| `SheetProtectionValue.formatRows`           | `SheetProtectionValue.FormatRows`          |         |
| `SheetProtectionValue.insertColumns`        | `SheetProtectionValue.InsertColumns`       |         |
| `SheetProtectionValue.insertRows`           | `SheetProtectionValue.InsertRows`          |         |
| `SheetProtectionValue.insertHyperlinks`     | `SheetProtectionValue.InsertHyperlinks`    |         |
| `SheetProtectionValue.deleteColumns`        | `SheetProtectionValue.DeleteColumns`       |         |
| `SheetProtectionValue.deleteRows`           | `SheetProtectionValue.DeleteRows`          |         |
| `SheetProtectionValue.selectLockedCells`    | `SheetProtectionValue.SelectLockedCells`   |         |
| `SheetProtectionValue.sort`                 | `SheetProtectionValue.Sort`                |         |
| `SheetProtectionValue.autoFilter`           | `SheetProtectionValue.AutoFilter`          |         |
| `SheetProtectionValue.pivotTables`          | `SheetProtectionValue.PivotTables`         |         |
| `SheetProtectionValue.selectUnlockedCells`  | `SheetProtectionValue.SelectUnlockedCells` |         |

- The enum values of `Worksheet.WorksheetPane` were renamed, according to the following overview:

| Old Enum Value                | New Enum Value                  | Remarks |
| ----------------------------- | ------------------------------- | ------- |
| `WorksheetPane.bottomRight`   | `WorksheetPane.BottomRight`     |         |
| `WorksheetPane.topRight`	    | `WorksheetPane.TopRight`        |         |
| `WorksheetPane.bottomLeft`    | `WorksheetPane.BottomLeft`      |         |
| `WorksheetPane.topLeft`       | `WorksheetPane.TopLeft`         |         |

- The enum values of `Worksheet.SheetViewType` were renamed, according to the following overview:

| Old Enum Value                   | New Enum Value                  | Remarks |
| -------------------------------- | ------------------------------- | ------- |
| `SheetViewType.normal`           | `SheetViewType.Normal`          |         |
| `SheetViewType.PageBreakPreview` | `SheetViewType.PageBreakPreview`|         |
| `SheetViewType.PageLayout`       | `SheetViewType.PageLayout`      |         |

---

### Cell

- All sub-classes, like `Cell.Address`, are now independent classes. Therefore, all occurrences of these classes have to be adapted in the code, e.g. instead of `Cell.Address address = new Cell.Address();`, `Address address = new Address();` is used. Additionally, some namespaces, like `NanoXLSX` have to be added (see particular sections).
- The enum values of `Cell.CellType` were renamed, according to the following overview:

| Old Enum Value      | New Enum Value         | Remarks  |
|---------------------|------------------------|----------|
| `CellType.STRING`   | `CellType.String`      |          |
| `CellType.NUMBER`   | `CellType.Number`      |          |
| `CellType.DATE`     | `CellType.Date`        |          |
| `CellType.TIME`     | `CellType.Time`        |          |
| `CellType.BOOL`     | `CellType.Bool`        |          |
| `CellType.FORMULA`  | `CellType.Formula`     |          |
| `CellType.EMPTY`    | `CellType.Empty`       |          |
| `CellType.DEFAULT`  | `CellType.Default`     |          |

---

### Address (struct)

- The class `Cell.Address` is now an independent struct `Address` in the namespace `NanoXLSX`. Therefore, all occurrences of this class have to be adapted in the code, e.g. instead of `Cell.Address address = new Cell.Address();`, `Address address = new Address();` is used. Additionally, the namespace `NanoXLSX` has to be added.
- The property `Row` is now read-only (immutable). To change the property, a new Address object has to be created
- The property `Column` is now read-only (immutable). To change the property, a new Address object has to be created
- The property `Type` is now read-only (immutable). To change the property, a new Address object has to be created

---

### Range (struct)

- The class `Cell.Range` is now an independent struct `Range` in the namespace `NanoXLSX`. Therefore, all occurrences of this class have to be adapted in the code, e.g. instead of `Cell.Range range = new Cell.Range();`, `Range range = new Range();` is used. Additionally, the namespace `NanoXLSX` has to be added.
- The property `StartAddress` is now read-only (immutable). To change the property, a new Range object has to be created
- The property `EndAddress` is now read-only (immutable). To change the property, a new Range object has to be created
 
---

### BasicFormulas

The class `Cell.BasicFormulas` was moved to an own class in the namespace `NanoXLSX`. The class name has to be changed in the code.

---

### Styles 

Styles were undergoing several changes in version 4.0.0, to improve usability and consistency.
Especially the `Font` class was completely redesigned, according to the Excel specifications.
Furthermore, a lot of constants were renamed to follow the C# naming conventions.

- General in any Style class: All (s)RGB values are automatically validated and cast to upper case. If valid hex values are used, no actions are necessary. If existing code uses invalid hex values, these have to be adapted.

### BasicStyles

The class `Style.BasicStyles` was moved to an own class in the namespace `NanoXLSX.Styles`. The class name has to be changed in the code.

---

#### Font

- The class `Style.Font` was moved to an own class in the namespace `NanoXLSX.Styles`. The class name has to be changed in the code.
- The public constant values of the `Font` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|------------------------|----------|
| `DEFAULT_MAJOR_FONT`     | `DefaultMajorFont`     |          |
| `DEFAULT_MINOR_FONT`     | `DefaultMinorFont`     |          |
| `DEFAULT_FONT_NAME`      | `DefaultFontName`      |          |
| `DEFAULT_FONT_SCHEME`    | `DefaultFontScheme`    |          |
| `MIN_FONT_SIZE`          | `MinFontSize`          |          |
| `MAX_FONT_SIZE`          | `MaxFontSize`          |          |
| `DEFAULT_FONT_SIZE`      | `DefaultFontSize`      |          |
| `DEFAULT_FONT_FAMILY`    | `DefaultFontSize`      | The type was changed from `int` to the enum `Font.FontFamilyValue` |
| `DEFAULT_VERTICAL_ALIGN` | `DefaultVerticalAlign` | The type was changed from enum `Font.VerticalAlignValue` to `Font.VerticalTextAlignValue` |

- The enum values of `Font.SchemeValue` were renamed, according to the following overview:

| Old Enum Value                | New Enum Value              | Remarks |
| ----------------------------- | --------------------------- | ------- |
| `SchemeValue.major`           | `SchemeValue.Major`         |         |
| `SchemeValue.minor`           | `SchemeValue.Minor`         |         |
| `SchemeValue.none`            | `SchemeValue.None`          |         |

- The enum values of `Font.VerticalTextAlignValue` were renamed, according to the following overview:

| Old Enum Value                      | New Enum Value              | Remarks |
| ----------------------------------- | --------------------------- | ------- |
| `VerticalTextAlignValue.subscript`  | `VerticalTextAlignValue.Subscript`    |         |
| `VerticalTextAlignValue.superscript`| `VerticalTextAlignValue.Superscript`  |         |
| `VerticalTextAlignValue.none`       | `VerticalTextAlignValue.None`         |         |

- The enum values of `Font.UnderlineValue` were renamed, according to the following overview:

| Old Enum Value                      | New Enum Value              | Remarks |
| ----------------------------------- | --------------------------- | ------- |
| `UnderlineValue.u_single`           | `UnderlineValue.Single`     |         |
| `UnderlineValue.u_double_`          | `UnderlineValue.Double`     |         |
| `UnderlineValue.singleAccounting`   | `UnderlineValue.SingleAccounting`  |         |
| `UnderlineValue.doubleAccounting`   | `UnderlineValue.DoubleAccounting`  |         |
| `UnderlineValue.none`               | `UnderlineValue.None`       |         |

- The property `Font.Family` was changed from type `string` to the enum `Font.FontFamilyValue`.The value has to be replaced by one of the following available values:
```cs
NotApplicable, Roman, Swiss, Modern, Script, Decorative, Reserved1, Reserved2, Reserved3, Reserved4, Reserved5, Reserved6, Reserved7, Reserved8, Reserved9
// Mostly used: Roman, Swiss, Modern, Script, Decorative
```

- The property `Font.Charset` was changed from type `string` to the enum `Font.CharsetValue`. The initialization default value is `CharsetValue.Default` The value has to be replaced by one of the following available values:
```cs
ApplicationDefined, ANSI, Default, Symbols, Mac, ShiftJIS, Hangul, Johab, GBK, Big5, Greek, Turkish, Vietnamese, Hebrew, Arabic, Baltic, Cyrillic, Thai, EasternEuropean, OEM
// ApplicableDefined is usually ignored, and Default may be used instead
```

- The property `Font.ColorScheme` was changed from type `int` to the enum `Theme.ColorSchemeElement`. The value has to be replaced by one of the available values (See **Theme section** ). The initialization default value is `Theme.ColorSchemeElement.light1`.
- The property `Font.VerticalAlign` was changed from type `Font.VerticalAlignValue` to the enum `Font.VerticalTextAlignValue`. Only the enum name has to be changed (see below):
- The enum `Font.VerticalAlignValue` was renamed to `Font.VerticalTextAlignValue`. The available values remain unchanged

#### Border

- The class `Style.Border` was moved to an own class in the namespace `NanoXLSX.Styles`. The class name has to be changed in the code.
- The public constant values of the `Border` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|------------------------|----------|
| `DEFAULT_BORDER_STYLE`   | `DefaultBorderStyle`   |          |
| `DEFAULT_COLOR`          | `DefaultBorderColor`   |          |

- The enum values of `Border.StyleValue` were renamed, according to the following overview:

| Old Enum Value                | New Enum Value              | Remarks |
| ----------------------------- | --------------------------- | ------- |
| `StyleValue.hair`             | `CellType.Hair`             |         |
| `StyleValue.dotted`           | `CellType.Dotted`           |         |
| `StyleValue.dashDotDot`       | `CellType.DashDotDot`       |         |
| `StyleValue.dashDot`          | `CellType.DashDot`          |         |
| `StyleValue.dashed`           | `CellType.Dashed`           |         |
| `StyleValue.thin`             | `CellType.Thin`             |         |
| `StyleValue.mediumDashDotDot` | `CellType.MediumDashDotDot` |         |
| `StyleValue.slantDashDot`     | `CellType.SlantDashDot`     |         |
| `StyleValue.mediumDashDot`    | `CellType.MediumDashDot`    |         |
| `StyleValue.mediumDashed`     | `CellType.MediumDashed`     |         |
| `StyleValue.medium`           | `CellType.Medium`           |         |
| `StyleValue.thick`            | `CellType.Thick`            |         |
| `StyleValue.s_double`         | `CellType.Double`           |         |

#### Fill

- The class `Style.Fill` was moved to an own class in the namespace `NanoXLSX.Styles`. The class name has to be changed in the code.
- The public constant values of the `Fill` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|------------------------|----------|
| `DEFAULT_COLOR`          | `DefaultColor`         |          |
| `DEFAULT_INDEXED_COLOR`  | `DefaultIndexedColor`  |          |
| `DEFAULT_PATTERN_FILL`   | `DefaultPatternFill`   |          |

- The enum values of `Fill.FillType` were renamed, according to the following overview:

| Old Enum Value          | New Enum Value              | Remarks |
| ----------------------- | --------------------------- | ------- |
| `FillType.patternColor` | `FillType.PatternColor` |         |
| `FillType.fillColor`    | `FillType.FillColor`    |         |

- The enum values of `Fill.PatternValue` were renamed, according to the following overview:

| Old Enum Value           | New Enum Value             | Remarks |
| ------------------------ | -------------------------- | ------- |
| `PatternValue.solid`     | `PatternValue.Solid`       |         |
| `PatternValue.darkGray`  | `PatternValue.DarkGray`    |         |
| `PatternValue.mediumGray`| `PatternValue.MediumGray`  |         |
| `PatternValue.lightGray` | `PatternValue.LightGray`   |         |
| `PatternValue.gray0625`  | `PatternValue.Gray0625`    |         |
| `PatternValue.gray125`   | `PatternValue.Gray125`     |         |
| `PatternValue.none`      | `PatternValue.None`        |         |

- The static method `Fill.ValidateColr(string,bool, bool)` was moved to the utils class `Validators.ValidateColr(string,bool, bool)` in namespace `NanoXLSX.Utils`. The class has to be changed in the code, but the method signature remains unchanged.

#### CellXf

- The class `Style.CellXf` was moved to an own class in the namespace `NanoXLSX.Styles`. The class name has to be changed in the code.
- The public constant values of the `CellXf` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|------------------------|----------|
| `DEFAULT_HORIZONTAL_ALIGNMENT` | `DefaultHorizontalAlignment`|          |
| `DEFAULT_ALIGNMENT`      | `DefaultAlignment`     |          |
| `DEFAULT_TEXT_DIRECTION` | `DefaultTextDirection` |          |
| `DEFAULT_VERTICAL_ALIGNMENT`   | `DefaultVerticalAlignment`  |          |

- The enum values of `CellXf.HorizontalAlignValue` were renamed, according to the following overview:

| Old Enum Value                | New Enum Value              | Remarks |
| ------------------------------- | --------------------------- | ------- |
| `HorizontalAlignValue.left`     | `HorizontalAlignValue.Left`      |         |
| `HorizontalAlignValue.center`   | `HorizontalAlignValue.Center`    |         |
| `HorizontalAlignValue.right`    | `HorizontalAlignValue.Right`     |         |
| `HorizontalAlignValue.fill`     | `HorizontalAlignValue.Fill`      |         |
| `HorizontalAlignValue.justify`  | `HorizontalAlignValue.Justify`   |         |
| `HorizontalAlignValue.general`  | `HorizontalAlignValue.General`   |         |
| `HorizontalAlignValue.centerContinuous` | `HorizontalAlignValue.CenterContinuous` |         |
| `HorizontalAlignValue.distributed` | `HorizontalAlignValue.Distributed` |         |
| `HorizontalAlignValue.none`     | `HorizontalAlignValue.None` |         |

- The enum values of `CellXf.TextBreakValue` were renamed, according to the following overview:

| Old Enum Value                | New Enum Value              | Remarks |
| ----------------------------- | --------------------------- | ------- |
| `TextBreakValue.wrapText`     | `TextBreakValue.WrapText`   |         |
| `TextBreakValue.shrinkToFit`  | `TextBreakValue.ShrinkToFit`|         |
| `TextBreakValue.none`         | `TextBreakValue.None`       |         |

- The enum values of `CellXf.TextDirectionValue` were renamed, according to the following overview:

| Old Enum Value                  | New Enum Value              | Remarks |
| ------------------------------- | --------------------------- | ------- |
| `TextDirectionValue.horizontal` | `TextDirectionValue.Horizontal` |         |
| `TextDirectionValue.vertical`   | `TextDirectionValue.Vertical`   |         |

- The enum values of `CellXf.VerticalAlignValue` were renamed, according to the following overview:

| Old Enum Value                  | New Enum Value              | Remarks |
| ------------------------------- | --------------------------- | ------- |
| `VerticalAlignValue.bottom`     | `VerticalAlignValue.Horizontal` |         |
| `VerticalAlignValue.top`        | `VerticalAlignValue.Vertical`   |         |
| `VerticalAlignValue.center`     | `VerticalAlignValue.Horizontal` |         |
| `VerticalAlignValue.justify`    | `VerticalAlignValue.Justify`    |         |
| `VerticalAlignValue.distributed`| `VerticalAlignValue.Distributed`|         |
| `VerticalAlignValue.none`       | `VerticalAlignValue.None`       |         |

#### NumberFormat

- The class `Style.NumberFormat` was moved to an own class in the namespace `NanoXLSX.Styles`. The class name has to be changed in the code.
- The public constant values of the `NumberFormat` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|------------------------|----------|
| `CUSTOMFORMAT_START_NUMBER` | `CustomFormatStartNumber`|          |
| `DEFAULT_NUMBER`         | `DefaultNumber`        |          |

- The enum values of `NumberFormat.FormatRange` were renamed, according to the following overview:

| Old Enum Value           | New Enum Value         | Remarks  |
|--------------------------|------------------------|----------|
| `FormatRange.defined_format` | `FormatRange.DefinedFormat`   |          |
| `FormatRange.custom_format_` | `FormatRange.CustomFormat`    |          |
| `FormatRange.invalied`   | `FormatRange.Inavlid`  |          |
| `FormatRange.undefined`  | `FormatRange.Undefined`|          |

- The enum values of `NumberFormat.NumberFormat` were renamed, according to the following overview:

| Old Enum Value           | New Enum Value          | Remarks                   |
| ------------------------ | ----------------------- | ------------------------- |
| `NumberFormat.format_1`  | `NumberFormat.Format1`  |                           |
| `NumberFormat.format_2`  | `NumberFormat.Format2`  |                           |
| `NumberFormat.format_3`  | `NumberFormat.Format3`  |                           |
| `NumberFormat.format_4`  | `NumberFormat.Format4`  |                           |
| `NumberFormat.format_5`  | `NumberFormat.Format5`  |                           |
| `NumberFormat.format_6`  | `NumberFormat.Format6`  |                           |
| `NumberFormat.format_7`  | `NumberFormat.Format7`  |                           |
| `NumberFormat.format_8`  | `NumberFormat.Format8`  |                           |
| `NumberFormat.format_9`  | `NumberFormat.Format9`  |                           |
| `NumberFormat.format_10` | `NumberFormat.Format10` |                           |
| `NumberFormat.format_11` | `NumberFormat.Format11` |                           |
| `NumberFormat.format_12` | `NumberFormat.Format12` |                           |
| `NumberFormat.format_13` | `NumberFormat.Format13` |                           |
| `NumberFormat.format_14` | `NumberFormat.Format14` |                           |
| `NumberFormat.format_15` | `NumberFormat.Format15` |                           |
| `NumberFormat.format_16` | `NumberFormat.Format16` |                           |
| `NumberFormat.format_17` | `NumberFormat.Format17` |                           |
| `NumberFormat.format_18` | `NumberFormat.Format18` |                           |
| `NumberFormat.format_19` | `NumberFormat.Format19` |                           |
| `NumberFormat.format_20` | `NumberFormat.Format20` |                           |
| `NumberFormat.format_21` | `NumberFormat.Format21` |                           |
| `NumberFormat.format_22` | `NumberFormat.Format22` |                           |
| `NumberFormat.format_37` | `NumberFormat.Format37` |                           |
| `NumberFormat.format_38` | `NumberFormat.Format38` |                           |
| `NumberFormat.format_39` | `NumberFormat.Format39` |                           |
| `NumberFormat.format_40` | `NumberFormat.Format40` |                           |
| `NumberFormat.format_45` | `NumberFormat.Format45` |                           |
| `NumberFormat.format_46` | `NumberFormat.Format46` |                           |
| `NumberFormat.format_47` | `NumberFormat.Format47` |                           |
| `NumberFormat.format_48` | `NumberFormat.Format48` |                           |
| `NumberFormat.format_49` | `NumberFormat.Format49` |                           |
| `NumberFormat.custom`    | `NumberFormat.Custom`   | Custom formats (ID ≥ 164) |
| `NumberFormat.none`      | `NumberFormat.None`     |                           |

---

### Theme
The `Theme` class was introduced with NanoXLSX v3.0.0 It represents the theme of a workbook, which contains several color schemes and font schemes.
The class can mostly be ignored unless specific stylings are required.
Theme may be references ind Styles, especially in Fonts.
- The enum `Theme.ColorSchemeElement` was introduced to represent the color scheme elements of a theme. The available values are:
```cs
 dark1, light1, dark2, light2, accent1, accent2, accent3, accent4, accent5, accent6, hyperlink, followedHyperlink
```

---

## Common

All changes related to common, mostly static functions

### LowLevel

The class `LowLevel` was copletely removed. All functionalities are now in particular utils or writer classes. Please see the documentation of NanoXLSX if you want to alter writer functionalities, since these are now modularized.

### Utils

- Several utils methods were removed from its original classes and replaced by several specific utils classes in the namespace `NanoXLSX.Utils`. The class name has to be adapted, according to the following method overview:

| Utils Method            | Old Class  | New Utils Class      | Remarks        |
|-------------------------|----------------------|------------------------|
| `GetOADateTimeString`   | `LowLevel` |`DataUtils`          | No changes of the signature |
| `GetOATimeString`       | `LowLevel` |`DataUtils`          | No changes of the signature |
| `GetInternalColumnWidth`| `LowLevel` | `DataUtils`          | No changes of the signature |
| `GetInternalRowHeight`  | `LowLevel` | `DataUtils`          | No changes of the signature |
| `GetInternalPaneSplitWidth` | `LowLevel` | `DataUtils`      | No changes of the signature |
| `GetInternalPaneSplitHeight`| `LowLevel` | `DataUtils`      | No changes of the signature |
| `GeneratePasswordHash`  | `LowLevel` | `NanoXLSX.LegacyPassword` - new method name: `GenerateLegacyPasswordHash(string)` | No longer an utils method |

- The public constant values of the former `Utils` class were moved to specific utils classes in the name space `NanoXLSX.Utils.Constants`. The class names have to be adapted, according to the following overview:

| Old Constant                | New Class and Constant       | Remarks        |
|-----------------------------|------------------------------|------------------------|
| `Utils.MIN_OA_DATE_VALUE`   | `DataUtils.MinOADateValue`   |
| `Utils.MAX_OA_DATE_VALUE`   | `DataUtils.MaxOADateValue`   |
| `Utils.FIRST_ALLOWED_EXCEL_DATE` | `DataUtils.FirstAllowedExcelDate`|
| `Utils.LAST_ALLOWED_EXCEL_DATE`  | `DataUtils.LastAllowedExcelDate` |
| `Utils.INVARIANT_CULTURE`   | `DataUtils.InvariantCulture` |
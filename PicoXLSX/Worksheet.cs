﻿/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace PicoXLSX
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using static PicoXLSX.Cell;

    /// <summary>
    /// Class representing a worksheet of a workbook
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// Threshold, using when floats are compared
        /// </summary>
        private const float FLOAT_THRESHOLD = 0.0001f;

        /// <summary>
        /// Maximum number of characters a worksheet name can have
        /// </summary>
        public static readonly int MAX_WORKSHEET_NAME_LENGTH = 31;

        /// <summary>
        /// Default column width as constant
        /// </summary>
        public const float DEFAULT_COLUMN_WIDTH = 10f;

        /// <summary>
        /// Default row height as constant
        /// </summary>
        public const float DEFAULT_ROW_HEIGHT = 15f;

        /// <summary>
        /// Maximum column number (zero-based) as constant
        /// </summary>
        public const int MAX_COLUMN_NUMBER = 16383;

        /// <summary>
        /// Minimum column number (zero-based) as constant
        /// </summary>
        public const int MIN_COLUMN_NUMBER = 0;

        /// <summary>
        /// Minimum column width as constant
        /// </summary>
        public const float MIN_COLUMN_WIDTH = 0f;

        /// <summary>
        /// Minimum row height as constant
        /// </summary>
        public const float MIN_ROW_HEIGHT = 0f;

        /// <summary>
        /// Maximum column width as constant
        /// </summary>
        public const float MAX_COLUMN_WIDTH = 255f;

        /// <summary>
        /// Maximum row number (zero-based) as constant
        /// </summary>
        public const int MAX_ROW_NUMBER = 1048575;

        /// <summary>
        /// Minimum row number (zero-based) as constant
        /// </summary>
        public const int MIN_ROW_NUMBER = 0;

        /// <summary>
        /// Maximum row height as constant
        /// </summary>
        public const float MAX_ROW_HEIGHT = 409.5f;

        /// <summary>
        /// Enum to define the direction when using AddNextCell method
        /// </summary>
        public enum CellDirection
        {
            /// <summary>The next cell will be on the same row (A1,B1,C1...)</summary>
            ColumnToColumn,
            /// <summary>The next cell will be on the same column (A1,A2,A3...)</summary>
            RowToRow,
            /// <summary>The address of the next cell will be not changed when adding a cell (for manual definition of cell addresses)</summary>
            Disabled
        }

        /// <summary>
        /// Enum to define the possible protection types when protecting a worksheet
        /// </summary>
        public enum SheetProtectionValue
        {
            // sheet, // Is always on 1 if protected
            /// <summary>If selected, the user can edit objects if the worksheets is protected</summary>
            objects,
            /// <summary>If selected, the user can edit scenarios if the worksheets is protected</summary>
            scenarios,
            /// <summary>If selected, the user can format cells if the worksheets is protected</summary>
            formatCells,
            /// <summary>If selected, the user can format columns if the worksheets is protected</summary>
            formatColumns,
            /// <summary>If selected, the user can format rows if the worksheets is protected</summary>
            formatRows,
            /// <summary>If selected, the user can insert columns if the worksheets is protected</summary>
            insertColumns,
            /// <summary>If selected, the user can insert rows if the worksheets is protected</summary>
            insertRows,
            /// <summary>If selected, the user can insert hyper links if the worksheets is protected</summary>
            insertHyperlinks,
            /// <summary>If selected, the user can delete columns if the worksheets is protected</summary>
            deleteColumns,
            /// <summary>If selected, the user can delete rows if the worksheets is protected</summary>
            deleteRows,
            /// <summary>If selected, the user can select locked cells if the worksheets is protected</summary>
            selectLockedCells,
            /// <summary>If selected, the user can sort cells if the worksheets is protected</summary>
            sort,
            /// <summary>If selected, the user can use auto filters if the worksheets is protected</summary>
            autoFilter,
            /// <summary>If selected, the user can use pivot tables if the worksheets is protected</summary>
            pivotTables,
            /// <summary>If selected, the user can select unlocked cells if the worksheets is protected</summary>
            selectUnlockedCells
        }

        /// <summary>
        /// Enum to define the pane position or active pane in a slip worksheet
        /// </summary>
        public enum WorksheetPane
        {
            /// <summary>The pane is located in the bottom right of the split worksheet</summary>
            bottomRight,
            /// <summary>The pane is located in the top right of the split worksheet</summary>
            topRight,
            /// <summary>The pane is located in the bottom left of the split worksheet</summary>
            bottomLeft,
            /// <summary>The pane is located in the top left of the split worksheet</summary>
            topLeft
        }

        /// <summary>
        /// Defines the activeStyle
        /// </summary>
        private Style activeStyle;

        /// <summary>
        /// Defines the autoFilterRange
        /// </summary>
        private Nullable<Cell.Range> autoFilterRange;

        /// <summary>
        /// Defines the cells
        /// </summary>
        private readonly Dictionary<string, Cell> cells;

        /// <summary>
        /// Defines the columns
        /// </summary>
        private readonly Dictionary<int, Column> columns;

        /// <summary>
        /// Defines the sheetName
        /// </summary>
        private string sheetName;

        /// <summary>
        /// Defines the currentRowNumber
        /// </summary>
        private int currentRowNumber;

        /// <summary>
        /// Defines the currentColumnNumber
        /// </summary>
        private int currentColumnNumber;

        /// <summary>
        /// Defines the defaultRowHeight
        /// </summary>
        private float defaultRowHeight;

        /// <summary>
        /// Defines the defaultColumnWidth
        /// </summary>
        private float defaultColumnWidth;

        /// <summary>
        /// Defines the rowHeights
        /// </summary>
        private readonly Dictionary<int, float> rowHeights;

        /// <summary>
        /// Defines the hiddenRows
        /// </summary>
        private readonly Dictionary<int, bool> hiddenRows;

        /// <summary>
        /// Defines the mergedCells
        /// </summary>
        private readonly Dictionary<string, Cell.Range> mergedCells;

        /// <summary>
        /// Defines the sheetProtectionValues
        /// </summary>
        private readonly List<SheetProtectionValue> sheetProtectionValues;

        /// <summary>
        /// Defines the useActiveStyle
        /// </summary>
        private bool useActiveStyle;

        /// <summary>
        /// Defines the hidden
        /// </summary>
        private bool hidden;

        /// <summary>
        /// Defines the workbookReference
        /// </summary>
        private Workbook workbookReference;

        /// <summary>
        /// Defines the sheetProtectionPassword
        /// </summary>
        private string sheetProtectionPassword = null;

        /// <summary>
        /// Defines the sheetProtectionPasswordHash
        /// </summary>
        private string sheetProtectionPasswordHash = null;

        /// <summary>
        /// Defines the selectedCells
        /// </summary>
        private List<Range> selectedCells;

        /// <summary>
        /// Defines the freezeSplitPanes
        /// </summary>
        private bool? freezeSplitPanes;

        /// <summary>
        /// Defines the paneSplitLeftWidth
        /// </summary>
        private float? paneSplitLeftWidth;

        /// <summary>
        /// Defines the paneSplitTopHeight
        /// </summary>
        private float? paneSplitTopHeight;

        /// <summary>
        /// Defines the paneSplitTopLeftCell
        /// </summary>
        private Cell.Address? paneSplitTopLeftCell;

        /// <summary>
        /// Defines the paneSplitAddress
        /// </summary>
        private Cell.Address? paneSplitAddress;

        /// <summary>
        /// Defines the activePane
        /// </summary>
        private WorksheetPane? activePane;

        /// <summary>
        /// Defines the sheetID
        /// </summary>
        private int sheetID;

        /// <summary>
        /// Gets the range of the auto-filter. Wrapped to Nullable to provide null as value. If null, no auto-filter is applied
        /// </summary>
        public Cell.Range? AutoFilterRange
        {
            get { return autoFilterRange; }
        }

        /// <summary>
        /// Gets the cells of the worksheet as dictionary with the cell address as key and the cell object as value
        /// </summary>
        public Dictionary<string, Cell> Cells
        {
            get { return cells; }
        }

        /// <summary>
        /// Gets all columns with non-standard properties, like auto filter applied or a special width as dictionary with the zero-based column index as key and the column object as value
        /// </summary>
        public Dictionary<int, Column> Columns
        {
            get { return columns; }
        }

        /// <summary>
        /// Gets or sets the direction when using AddNextCell method
        /// </summary>
        public CellDirection CurrentCellDirection { get; set; }

        /// <summary>
        /// Gets or sets the default column width
        /// </summary>
        public float DefaultColumnWidth
        {
            get { return defaultColumnWidth; }
            set
            {
                if (value < MIN_COLUMN_WIDTH || value > MAX_COLUMN_WIDTH)
                {
                    throw new RangeException("OutOfRangeException", "The passed default column width is out of range (" + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + ")");
                }
                defaultColumnWidth = value;
            }
        }

        /// <summary>
        /// Gets or sets the default Row height
        /// </summary>
        public float DefaultRowHeight
        {
            get { return defaultRowHeight; }
            set
            {
                if (value < MIN_ROW_HEIGHT || value > MAX_ROW_HEIGHT)
                {
                    throw new RangeException("OutOfRangeException", "The passed default row height is out of range (" + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + ")");
                }
                defaultRowHeight = value;
            }
        }

        /// <summary>
        /// Gets the hidden rows as dictionary with the zero-based row number as key and a boolean as value. True indicates hidden, false visible.
        /// </summary>
        public Dictionary<int, bool> HiddenRows
        {
            get { return hiddenRows; }
        }

        /// <summary>
        /// Gets the merged cells (only references) as dictionary with the cell address as key and the range object as value
        /// </summary>
        public Dictionary<string, Cell.Range> MergedCells
        {
            get { return mergedCells; }
        }

        /// <summary>
        /// Gets defined row heights as dictionary with the zero-based row number as key and the height (float from 0 to 409.5) as value
        /// </summary>
        public Dictionary<int, float> RowHeights
        {
            get { return rowHeights; }
        }

        /// <summary>
        /// Returns either null (if no cells are selected), or the first defined range of selected cells
        /// </summary>
        /// <remarks>Use <see cref="SelectedCellRanges"/> to get all defined ranges</remarks>
        [Obsolete("This method is a deprecated subset of the function SelectedCellRanges. SelectedCellRanges will get this function name in a future version. Therefore, the type will change")]
        public Range? SelectedCells
        {
            get
            {
                if (selectedCells.Count == 0)
                {
                    return null;
                }
                else
                {
                    return selectedCells[0];
                }
            }
        }

        /// <summary>
        /// Gets all ranges of selected cells of this worksheet. An empty list is returned if no cells are selected
        /// </summary>
        public List<Range> SelectedCellRanges
        {
            get { return selectedCells; }
        }

        /// <summary>
        /// Gets or sets the internal ID of the worksheet
        /// </summary>
        public int SheetID
        {
            get => sheetID;
            set
            {
                if (value < 1)
                {
                    throw new FormatException("The ID " + value + " is invalid. Worksheet IDs must be >0");
                }
                sheetID = value;
            }
        }

        /// <summary>
        /// Gets or sets the name of the worksheet
        /// </summary>
        public string SheetName
        {
            get { return sheetName; }
            set { SetSheetName(value); }
        }

        /// <summary>
        /// Gets the password used for sheet protection. See <see cref="SetSheetProtectionPassword"/> to set the password
        /// </summary>
        public string SheetProtectionPassword
        {
            get { return sheetProtectionPassword; }
        }

        /// <summary>
        /// gets the encrypted hash of the password, defined with <see cref="SheetProtectionPassword"/>. The value will be null, if no password is defined
        /// </summary>
        public string SheetProtectionPasswordHash
        {
            get { return sheetProtectionPasswordHash; }
        }

        /// <summary>
        /// Gets the list of SheetProtectionValues. These values define the allowed actions if the worksheet is protected
        /// </summary>
        public List<SheetProtectionValue> SheetProtectionValues
        {
            get { return sheetProtectionValues; }
        }

        /// <summary>
        /// Gets or sets whether the worksheet is protected. If true, protection is enabled
        /// </summary>
        public bool UseSheetProtection { get; set; }

        /// <summary>
        /// Gets or sets the Reference to the parent Workbook
        /// </summary>
        public Workbook WorkbookReference
        {
            get { return workbookReference; }
            set
            {
                workbookReference = value;
                if (value != null)
                {
                    workbookReference.ValidateWorksheets();
                }
            }
        }

        /// <summary>
        /// Gets or sets whether the worksheet is hidden. If true, the worksheet is not listed as tab in the workbook's worksheet selection<br/>
        /// If the worksheet is not part of a workbook, or the only one in the workbook, an exception will be thrown.<br/>
        /// If the worksheet is the selected one, and attempted to set hidden, an exception will be thrown. Define another selected worksheet prior to this call, in this case.
        /// </summary>
        public bool Hidden
        {
            get { return hidden; }
            set
            {
                hidden = value;
                if (value && workbookReference != null)
                {
                    workbookReference.ValidateWorksheets();
                }
            }
        }

        /// <summary>
        /// Gets the height of the upper, horizontal split pane, measured from the top of the window.<br/>
        /// The value is nullable. If null, no horizontal split of the worksheet is applied.<br/>
        /// The value is only applicable to split the worksheet into panes, but not to freeze them.<br/>
        /// See also: <see cref="PaneSplitAddress"/>
        /// </summary>
        public float? PaneSplitTopHeight
        {
            get { return paneSplitTopHeight; }
        }

        /// <summary>
        /// Gets the width of the left, vertical split pane, measured from the left of the window.<br/>
        /// The value is nullable. If null, no vertical split of the worksheet is applied<br/>
        /// The value is only applicable to split the worksheet into panes, but not to freeze them.<br/>
        /// See also: <see cref="PaneSplitAddress"/>
        /// </summary>
        public float? PaneSplitLeftWidth
        {
            get { return paneSplitLeftWidth; }
        }

        /// <summary>
        /// Gets the FreezeSplitPanes
        /// Gets whether split panes are frozen.<br/>
        /// The value is nullable. If null, no freezing is applied. This property also does not apply if <see cref="PaneSplitAddress"/> is null
        /// </summary>
        public bool? FreezeSplitPanes
        {
            get { return freezeSplitPanes; }
        }

        /// <summary>
        /// Gets the Top Left cell address of the bottom right pane if applicable and splitting is applied.<br/>
        /// The column is only relevant for vertical split, whereas the row component is only relevant for a horizontal split.<br/>
        /// The value is nullable. If null, no splitting was defined.
        /// </summary>
        public Cell.Address? PaneSplitTopLeftCell
        {
            get { return paneSplitTopLeftCell; }
        }

        /// <summary>
        /// Gets the split address for frozen panes or if pane split was defined in number of columns and / or rows.<br/> 
        /// For vertical splits, only the column component is considered. For horizontal splits, only the row component is considered.<br/>
        /// The value is nullable. If null, no frozen panes or split by columns / rows are applied to the worksheet. 
        /// However, splitting can still be applied, if the value is defined in characters.<br/>
        /// See also: <see cref="PaneSplitLeftWidth"/> and <see cref="PaneSplitTopHeight"/> for splitting in characters (without freezing)
        /// </summary>
        public Cell.Address? PaneSplitAddress
        {
            get { return paneSplitAddress; }
        }

        /// <summary>
        /// Gets the active Pane is splitting is applied.<br/>
        /// The value is nullable. If null, no splitting was defined
        /// </summary>
        public WorksheetPane? ActivePane
        {
            get { return activePane; }
        }

        /// <summary>
        /// Gets the active Style of the worksheet. If null, no style is defined as active
        /// </summary>
        public Style ActiveStyle
        {
            get { return activeStyle; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Worksheet"/> class
        /// </summary>
        public Worksheet()
        {
            CurrentCellDirection = CellDirection.ColumnToColumn;
            cells = new Dictionary<string, Cell>();
            currentRowNumber = 0;
            currentColumnNumber = 0;
            defaultColumnWidth = DEFAULT_COLUMN_WIDTH;
            defaultRowHeight = DEFAULT_ROW_HEIGHT;
            rowHeights = new Dictionary<int, float>();
            mergedCells = new Dictionary<string, Cell.Range>();
            sheetProtectionValues = new List<SheetProtectionValue>();
            hiddenRows = new Dictionary<int, bool>();
            columns = new Dictionary<int, Column>();
            selectedCells = new List<Range>();
            activeStyle = null;
            workbookReference = null;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Worksheet"/> class
        /// </summary>
        /// <param name="name">The name<see cref="string"/>.</param>
        public Worksheet(string name)
            : this()
        {
            SetSheetName(name);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Worksheet"/> class
        /// </summary>
        /// <param name="name">Name of the worksheet.</param>
        /// <param name="id">ID of the worksheet (for internal use).</param>
        /// <param name="reference">Reference to the parent Workbook.</param>
        public Worksheet(string name, int id, Workbook reference)
            : this()
        {
            SetSheetName(name);
            SheetID = id;
            workbookReference = reference;
        }

        /// <summary>
        /// Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert.</param>
        public void AddNextCell(object value)
        {
            AddNextCell(CastValue(value, currentColumnNumber, currentRowNumber), true, null);
        }

        /// <summary>
        /// Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert.</param>
        /// <param name="style">Style object to apply on this cell.</param>
        public void AddNextCell(object value, Style style)
        {
            AddNextCell(CastValue(value, currentColumnNumber, currentRowNumber), true, style);
        }

        /// <summary>
        /// Method to insert a generic cell to the next cell position
        /// </summary>
        /// <param name="cell">Cell object to insert.</param>
        /// <param name="incremental">If true, the address value (row or column) will be incremented, otherwise not.</param>
        /// <param name="style">If not null, the defined style will be applied to the cell, otherwise no style or the default style will be applied.</param>
        private void AddNextCell(Cell cell, bool incremental, Style style)
        {
            // date and time styles are already defined by the passed cell object
            if (style != null || (activeStyle != null && useActiveStyle))
            {

                if (cell.CellStyle == null && useActiveStyle)
                {
                    cell.SetStyle(activeStyle);
                }
                else if (cell.CellStyle == null && style != null)
                {
                    cell.SetStyle(style);
                }
                else if (cell.CellStyle != null && useActiveStyle)
                {
                    Style mixedStyle = (Style)cell.CellStyle.Copy();
                    mixedStyle.Append(activeStyle);
                    cell.SetStyle(mixedStyle);
                }
                else if (cell.CellStyle != null && style != null)
                {
                    Style mixedStyle = (Style)cell.CellStyle.Copy();
                    mixedStyle.Append(style);
                    cell.SetStyle(mixedStyle);
                }
            }
            string address = cell.CellAddress;
            if (cells.ContainsKey(address))
            {
                cells[address] = cell;
            }
            else
            {
                cells.Add(address, cell);
            }
            if (incremental)
            {
                if (CurrentCellDirection == CellDirection.ColumnToColumn)
                {
                    currentColumnNumber++;
                }
                else if (CurrentCellDirection == CellDirection.RowToRow)
                {
                    currentRowNumber++;
                }
                // else = disabled
            }
            else
            {
                if (CurrentCellDirection == CellDirection.ColumnToColumn)
                {
                    currentColumnNumber = cell.ColumnNumber + 1;
                    currentRowNumber = cell.RowNumber;
                }
                else if (CurrentCellDirection == CellDirection.RowToRow)
                {
                    currentColumnNumber = cell.ColumnNumber;
                    currentRowNumber = cell.RowNumber + 1;
                }
                // else = Disabled
            }
        }

        /// <summary>
        /// Method to cast a value or align an object of the type Cell to the context of the worksheet
        /// </summary>
        /// <param name="value">Unspecified value or object of the type Cell.</param>
        /// <param name="column">Column index.</param>
        /// <param name="row">Row index.</param>
        /// <returns>Cell object.</returns>
        private Cell CastValue(object value, int column, int row)
        {
            Cell c;
            if (value != null && value.GetType() == typeof(Cell))
            {
                c = (Cell)value;
                c.CellAddress2 = new Cell.Address(column, row);
            }
            else
            {
                c = new Cell(value, Cell.CellType.DEFAULT, column, row);
            }
            return c;
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert.</param>
        /// <param name="columnNumber">Column number (zero based).</param>
        /// <param name="rowNumber">Row number (zero based).</param>
        public void AddCell(object value, int columnNumber, int rowNumber)
        {
            AddNextCell(CastValue(value, columnNumber, rowNumber), false, null);
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert.</param>
        /// <param name="columnNumber">Column number (zero based).</param>
        /// <param name="rowNumber">Row number (zero based).</param>
        /// <param name="style">Style to apply on the cell.</param>
        public void AddCell(object value, int columnNumber, int rowNumber, Style style)
        {
            AddNextCell(CastValue(value, columnNumber, rowNumber), false, style);
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert.</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576.</param>
        public void AddCell(object value, string address)
        {
            int column, row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddCell(value, column, row);
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert.</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576.</param>
        /// <param name="style">Style to apply on the cell.</param>
        public void AddCell(object value, string address, Style style)
        {
            int column, row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddCell(value, column, row, style);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert.</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576.</param>
        public void AddCellFormula(string formula, string address)
        {
            int column, row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            Cell c = new Cell(formula, Cell.CellType.FORMULA, column, row);
            AddNextCell(c, false, null);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert.</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576.</param>
        /// <param name="style">Style to apply on the cell.</param>
        public void AddCellFormula(string formula, string address, Style style)
        {
            int column, row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            Cell c = new Cell(formula, Cell.CellType.FORMULA, column, row);
            AddNextCell(c, false, style);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert.</param>
        /// <param name="columnNumber">Column number (zero based).</param>
        /// <param name="rowNumber">Row number (zero based).</param>
        public void AddCellFormula(string formula, int columnNumber, int rowNumber)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
            AddNextCell(c, false, null);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert.</param>
        /// <param name="columnNumber">Column number (zero based).</param>
        /// <param name="rowNumber">Row number (zero based).</param>
        /// <param name="style">Style to apply on the cell.</param>
        public void AddCellFormula(string formula, int columnNumber, int rowNumber, Style style)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
            AddNextCell(c, false, style);
        }

        /// <summary>
        /// Adds a formula as string to the next cell position
        /// </summary>
        /// <param name="formula">Formula to insert.</param>
        public void AddNextCellFormula(string formula)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, currentColumnNumber, currentRowNumber);
            AddNextCell(c, true, null);
        }

        /// <summary>
        /// Adds a formula as string to the next cell position
        /// </summary>
        /// <param name="formula">Formula to insert.</param>
        /// <param name="style">Style to apply on the cell.</param>
        public void AddNextCellFormula(string formula, Style style)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, currentColumnNumber, currentRowNumber);
            AddNextCell(c, true, style);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert.</param>
        /// <param name="startAddress">Start address.</param>
        /// <param name="endAddress">End address.</param>
        public void AddCellRange(IReadOnlyList<object> values, Cell.Address startAddress, Cell.Address endAddress)
        {
            AddCellRangeInternal(values, startAddress, endAddress, null);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert.</param>
        /// <param name="startAddress">Start address.</param>
        /// <param name="endAddress">End address.</param>
        /// <param name="style">Style to apply on the all cells of the range.</param>
        public void AddCellRange(IReadOnlyList<object> values, Cell.Address startAddress, Cell.Address endAddress, Style style)
        {
            AddCellRangeInternal(values, startAddress, endAddress, style);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert.</param>
        /// <param name="cellRange">Cell range as string in the format like A1:D1 or X10:X22.</param>
        public void AddCellRange(IReadOnlyList<object> values, string cellRange)
        {
            Cell.Range range = Cell.ResolveCellRange(cellRange);
            AddCellRangeInternal(values, range.StartAddress, range.EndAddress, null);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert.</param>
        /// <param name="cellRange">Cell range as string in the format like A1:D1 or X10:X22.</param>
        /// <param name="style">Style to apply on the all cells of the range.</param>
        public void AddCellRange(IReadOnlyList<object> values, string cellRange, Style style)
        {
            Cell.Range range = Cell.ResolveCellRange(cellRange);
            AddCellRangeInternal(values, range.StartAddress, range.EndAddress, style);
        }

        /// <summary>
        /// Internal function to add a generic list of value to the defined cell range
        /// </summary>
        /// <typeparam name="T">Data type of the generic value list.</typeparam>
        /// <param name="values">List of values.</param>
        /// <param name="startAddress">Start address.</param>
        /// <param name="endAddress">End address.</param>
        /// <param name="style">Style to apply on the all cells of the range.</param>
        private void AddCellRangeInternal<T>(IReadOnlyList<T> values, Cell.Address startAddress, Cell.Address endAddress, Style style)
        {
            List<Cell.Address> addresses = Cell.GetCellRange(startAddress, endAddress) as List<Cell.Address>;
            if (values.Count != addresses.Count)
            {
                throw new RangeException("OutOfRangeException", "The number of passed values (" + values.Count + ") differs from the number of cells within the range (" + addresses.Count + ")");
            }
            List<Cell> list = Cell.ConvertArray(values) as List<Cell>;
            int len = values.Count;
            for (int i = 0; i < len; i++)
            {
                list[i].RowNumber = addresses[i].Row;
                list[i].ColumnNumber = addresses[i].Column;
                AddNextCell(list[i], false, style);
            }
        }

        /// <summary>
        /// Removes a previous inserted cell at the defined address
        /// </summary>
        /// <param name="columnNumber">Column number (zero based).</param>
        /// <param name="rowNumber">Row number (zero based).</param>
        /// <returns>Returns true if the cell could be removed (existed), otherwise false (did not exist).</returns>
        public bool RemoveCell(int columnNumber, int rowNumber)
        {
            string address = Cell.ResolveCellAddress(columnNumber, rowNumber);
            return cells.Remove(address);
        }

        /// <summary>
        /// Removes a previous inserted cell at the defined address
        /// </summary>
        /// <param name="address">Cell address in the format A1 - XFD1048576.</param>
        /// <returns>Returns true if the cell could be removed (existed), otherwise false (did not exist).</returns>
        public bool RemoveCell(string address)
        {
            int row, column;
            Cell.ResolveCellCoordinate(address, out column, out row);
            return RemoveCell(column, row);
        }

        /// <summary>
        /// Sets the passed style on the passed cell range. If cells are already existing, the style will be added or replaced
        /// </summary>
        /// <param name="cellRange">Cell range to apply the style.</param>
        /// <param name="style">Style to apply.</param>
        public void SetStyle(Cell.Range cellRange, Style style)
        {
            IReadOnlyList<Cell.Address> addresses = cellRange.ResolveEnclosedAddresses();
            foreach (Cell.Address address in addresses)
            {
                string key = address.GetAddress();
                if (this.cells.ContainsKey(key))
                {
                    if (style == null)
                    {
                        cells[key].RemoveStyle();
                    }
                    else
                    {
                        cells[key].SetStyle(style);
                    }
                }
                else
                {
                    if (style != null)
                    {
                        AddCell(null, address.Column, address.Row, style);
                    }
                }
            }
        }

        /// <summary>
        /// Sets the passed style on the passed cell range, derived from a start and end address. If cells are already existing, the style will be added or replaced
        /// Sets the passed style on the passed cell range, derived from a start and end address. If cells are already existing, the style will be added or replaced
        /// </summary>
        /// <param name="startAddress">Start address of the cell range.</param>
        /// <param name="endAddress">End address of the cell range.</param>
        /// <param name="style">Style to apply or null to clear the range.</param>
        public void SetStyle(Cell.Address startAddress, Cell.Address endAddress, Style style)
        {
            SetStyle(new Cell.Range(startAddress, endAddress), style);
        }

        /// <summary>
        /// Sets the passed style on the passed (singular) cell address. If the cell is already existing, the style will be added or replaced
        /// Sets the passed style on the passed (singular) cell address. If the cell is already existing, the style will be added or replaced
        /// </summary>
        /// <param name="address">Cell address to apply the style.</param>
        /// <param name="style">Style to apply or null to clear the range.</param>
        public void SetStyle(Cell.Address address, Style style)
        {
            SetStyle(address, address, style);
        }

        /// <summary>
        /// Sets the passed style on the passed address expression. Such an expression may be a single cell or a cell range
        /// Sets the passed style on the passed address expression. Such an expression may be a single cell or a cell range
        /// Sets the passed style on the passed address expression. Such an expression may be a single cell or a cell range
        /// </summary>
        /// <param name="addressExpression">Expression of a cell address or range of addresses.</param>
        /// <param name="style">Style to apply or null to clear the range.</param>
        public void SetStyle(string addressExpression, Style style)
        {
            Cell.AddressScope scope = Cell.GetAddressScope(addressExpression);
            if (scope == Cell.AddressScope.SingleAddress)
            {
                Cell.Address address = new Cell.Address(addressExpression);
                SetStyle(address, style);
            }
            else if (scope == Cell.AddressScope.Range)
            {
                Cell.Range range = new Cell.Range(addressExpression);
                SetStyle(range, style);
            }
            else
            {
                throw new FormatException("The passed address'" + addressExpression + "' is neither a cell address, nor a range");
            }
        }

        /// <summary>
        /// Method to add allowed actions if the worksheet is protected. If one or more values are added, UseSheetProtection will be set to true
        /// </summary>
        /// <param name="typeOfProtection">Allowed action on the worksheet or cells.</param>
        public void AddAllowedActionOnSheetProtection(SheetProtectionValue typeOfProtection)
        {
            if (!sheetProtectionValues.Contains(typeOfProtection))
            {
                if (typeOfProtection == SheetProtectionValue.selectLockedCells && !sheetProtectionValues.Contains(SheetProtectionValue.selectUnlockedCells))
                {
                    sheetProtectionValues.Add(SheetProtectionValue.selectUnlockedCells);
                }
                sheetProtectionValues.Add(typeOfProtection);
                UseSheetProtection = true;
            }
        }

        /// <summary>
        /// Sets the defined column as hidden
        /// </summary>
        /// <param name="columnNumber">Column number to hide on the worksheet.</param>
        public void AddHiddenColumn(int columnNumber)
        {
            SetColumnHiddenState(columnNumber, true);
        }

        /// <summary>
        /// Sets the defined column as hidden
        /// </summary>
        /// <param name="columnAddress">Column address to hide on the worksheet.</param>
        public void AddHiddenColumn(string columnAddress)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnHiddenState(columnNumber, true);
        }

        /// <summary>
        /// Sets the defined row as hidden
        /// </summary>
        /// <param name="rowNumber">Row number to hide on the worksheet.</param>
        public void AddHiddenRow(int rowNumber)
        {
            SetRowHiddenState(rowNumber, true);
        }

        /// <summary>
        /// Clears the active style of the worksheet. All later added calls will contain no style unless another active style is set
        /// </summary>
        public void ClearActiveStyle()
        {
            useActiveStyle = false;
            activeStyle = null;
        }

        /// <summary>
        /// Gets the cell of the specified address
        /// </summary>
        /// <param name="address">Address of the cell.</param>
        /// <returns>Cell object.</returns>
        public Cell GetCell(Cell.Address address)
        {
            if (!cells.ContainsKey(address.GetAddress()))
            {
                throw new WorksheetException("The cell with the address " + address.GetAddress() + " does not exist in this worksheet");
            }
            return cells[address.GetAddress()];
        }

        /// <summary>
        /// Gets the cell of the specified column and row number (zero-based)
        /// </summary>
        /// <param name="columnNumber">Column number of the cell.</param>
        /// <param name="rowNumber">Row number of the cell.</param>
        /// <returns>Cell object.</returns>
        public Cell GetCell(int columnNumber, int rowNumber)
        {
            return GetCell(new Cell.Address(columnNumber, rowNumber));
        }

        /// <summary>
        /// Gets whether the specified address exists in the worksheet. Existing means that a value was stored at the address
        /// </summary>
        /// <param name="address">Address to check.</param>
        /// <returns>The <see cref="bool"/>.</returns>
        public bool HasCell(Cell.Address address)
        {
            return cells.ContainsKey(address.GetAddress());
        }

        /// <summary>
        /// Gets whether the specified address exists in the worksheet. Existing means that a value was stored at the address
        /// </summary>
        /// <param name="columnNumber">Column number of the cell to check (zero-based).</param>
        /// <param name="rowNumber">Row number of the cell to check (zero-based).</param>
        /// <returns>The <see cref="bool"/>.</returns>
        public bool HasCell(int columnNumber, int rowNumber)
        {
            return HasCell(new Cell.Address(columnNumber, rowNumber));
        }

        /// <summary>
        /// Resets the defined column, if existing. The corresponding instance will be removed from <see cref="Columns"/>
        /// </summary>
        /// <param name="columnNumber">Column number to reset (zero-based).</param>
        public void ResetColumn(int columnNumber)
        {
            if (columns.ContainsKey(columnNumber) && !columns[columnNumber].HasAutoFilter) // AutoFilters cannot have gaps 
            {
                columns.Remove(columnNumber);
            }
            else if (columns.ContainsKey(columnNumber))
            {
                columns[columnNumber].IsHidden = false;
                columns[columnNumber].Width = DEFAULT_COLUMN_WIDTH;
            }
        }

        /// <summary>
        /// Gets the first existing column number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of an empty worksheet, -1 will be returned.</returns>
        public int GetFirstColumnNumber()
        {
            return GetBoundaryNumber(false, true);
        }

        /// <summary>
        /// Gets the first existing column number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of an empty worksheet, -1 will be returned.</returns>
        public int GetFirstDataColumnNumber()
        {
            return GetBoundaryDataNumber(false, true, true);
        }

        /// <summary>
        /// Gets the first existing row number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of an empty worksheet, -1 will be returned.</returns>
        public int GetFirstRowNumber()
        {
            return GetBoundaryNumber(true, true);
        }

        /// <summary>
        /// Gets the first existing row number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of an empty worksheet, -1 will be returned.</returns>
        public int GetFirstDataRowNumber()
        {
            return GetBoundaryDataNumber(true, true, true);
        }

        /// <summary>
        /// Gets the last existing column number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of an empty worksheet, -1 will be returned.</returns>
        public int GetLastColumnNumber()
        {
            return GetBoundaryNumber(false, false);
        }

        /// <summary>
        /// Gets the last existing column number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. in case of an empty worksheet, -1 will be returned.</returns>
        public int GetLastDataColumnNumber()
        {
            return GetBoundaryDataNumber(false, false, true);
        }

        /// <summary>
        /// Gets the last existing row number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of an empty worksheet, -1 will be returned.</returns>
        public int GetLastRowNumber()
        {
            return GetBoundaryNumber(true, false);
        }

        /// <summary>
        /// Gets the last existing row number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. in case of an empty worksheet, -1 will be returned.</returns>
        public int GetLastDataRowNumber()
        {
            return GetBoundaryDataNumber(true, false, true);
        }

        /// <summary>
        /// Gets the last existing cell in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned.</returns>
        public Cell.Address? GetLastCellAddress()
        {
            int lastRow = GetLastRowNumber();
            int lastColumn = GetLastColumnNumber();
            if (lastRow < 0 || lastColumn < 0)
            {
                return null;
            }
            return new Cell.Address(lastColumn, lastRow);
        }

        /// <summary>
        /// Gets the last existing cell with data in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned.</returns>
        public Cell.Address? GetLastDataCellAddress()
        {
            int lastRow = GetLastDataRowNumber();
            int lastColumn = GetLastDataColumnNumber();
            if (lastRow < 0 || lastColumn < 0)
            {
                return null;
            }
            return new Cell.Address(lastColumn, lastRow);
        }

        /// <summary>
        /// Gets the first existing cell in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned.</returns>
        public Cell.Address? GetFirstCellAddress()
        {
            int firstRow = GetFirstRowNumber();
            int firstColumn = GetFirstColumnNumber();
            if (firstRow < 0 || firstColumn < 0)
            {
                return null;
            }
            return new Cell.Address(firstColumn, firstRow);
        }

        /// <summary>
        /// Gets the first existing cell with data in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned.</returns>
        public Cell.Address? GetFirstDataCellAddress()
        {
            int firstRow = GetFirstDataRowNumber();
            int firstColumn = GetLastDataColumnNumber();
            if (firstRow < 0 || firstColumn < 0)
            {
                return null;
            }
            return new Cell.Address(firstColumn, firstRow);
        }

        /// <summary>
        /// Gets either the minimum or maximum row or column number, considering only calls with data
        /// </summary>
        /// <param name="row">If true, the min or max row is returned, otherwise the column.</param>
        /// <param name="min">If true, the min value of the row or column is defined, otherwise the max value.</param>
        /// <param name="ignoreEmpty">If true, empty cell values are ignored, otherwise considered without checking the content.</param>
        /// <returns>Min or max number, or -1 if not defined.</returns>
        private int GetBoundaryDataNumber(bool row, bool min, bool ignoreEmpty)
        {
            if (cells.Count == 0)
            {
                return -1;
            }
            if (!ignoreEmpty)
            {
                if (row && min)
                {
                    return cells.Min(x => x.Value.RowNumber);
                }
                else if (row)
                {
                    return cells.Max(x => x.Value.RowNumber);
                }
                else if (min)
                {
                    return cells.Min(x => x.Value.ColumnNumber);
                }
                else
                {
                    return cells.Max(x => x.Value.ColumnNumber);
                }
            }
            List<Cell> nonEmptyCells = cells.Values.Where(x => x.Value != null).ToList();
            if (nonEmptyCells.Count == 0)
            {
                return -1;
            }
            if (row && min)
            {
                return nonEmptyCells.Where(x => x.Value.ToString() != string.Empty).Min(x => x.RowNumber);
            }
            else if (row)
            {
                return nonEmptyCells.Where(x => x.Value.ToString() != string.Empty).Max(x => x.RowNumber);
            }
            else if (min)
            {
                return nonEmptyCells.Where(x => x.Value.ToString() != string.Empty).Max(x => x.ColumnNumber);
            }
            else
            {
                return nonEmptyCells.Where(x => x.Value.ToString() != string.Empty).Min(x => x.ColumnNumber);
            }
        }

        /// <summary>
        /// Gets either the minimum or maximum row or column number, considering all available data
        /// </summary>
        /// <param name="row">If true, the min or max row is returned, otherwise the column.</param>
        /// <param name="min">If true, the min value of the row or column is defined, otherwise the max value.</param>
        /// <returns>Min or max number, or -1 if not defined.</returns>
        private int GetBoundaryNumber(bool row, bool min)
        {
            int cellBoundary = GetBoundaryDataNumber(row, min, false);
            if (row)
            {
                int heightBoundary = -1;
                if (rowHeights.Count > 0)
                {
                    heightBoundary = min ? RowHeights.Min(x => x.Key) : RowHeights.Max(x => x.Key);
                }
                int hiddenBoundary = -1;
                if (hiddenRows.Count > 0)
                {
                    hiddenBoundary = min ? HiddenRows.Min(x => x.Key) : HiddenRows.Max(x => x.Key);
                }
                return min ? GetMinRow(cellBoundary, heightBoundary, hiddenBoundary) : GetMaxRow(cellBoundary, heightBoundary, hiddenBoundary);
            }
            else
            {
                int columnDefBoundary = -1;
                if (columns.Count > 0)
                {
                    columnDefBoundary = min ? Columns.Min(x => x.Key) : Columns.Max(x => x.Key);
                }
                if (min)
                {
                    return cellBoundary > 0 && cellBoundary < columnDefBoundary ? cellBoundary : columnDefBoundary;
                }
                else
                {
                    return cellBoundary > 0 && cellBoundary > columnDefBoundary ? cellBoundary : columnDefBoundary;
                }
            }
        }

        /// <summary>
        /// Gets the maximum row coordinate either from cell data, height definitions or hidden rows
        /// </summary>
        /// <param name="cellBoundary">Row number of max cell data.</param>
        /// <param name="heightBoundary">Row number of max defined row height.</param>
        /// <param name="hiddenBoundary">Row number of max defined hidden row.</param>
        /// <returns>Max row number or -1 if nothing valid defined.</returns>
        private int GetMaxRow(int cellBoundary, int heightBoundary, int hiddenBoundary)
        {
            int highest = -1;
            if (cellBoundary >= 0)
            {
                highest = cellBoundary;
            }
            if (heightBoundary >= 0 && heightBoundary > highest)
            {
                highest = heightBoundary;
            }
            if (hiddenBoundary >= 0 && hiddenBoundary > highest)
            {
                highest = hiddenBoundary;
            }
            return highest;
        }

        /// <summary>
        /// Gets the minimum row coordinate either from cell data, height definitions or hidden rows
        /// </summary>
        /// <param name="cellBoundary">Row number of min cell data.</param>
        /// <param name="heightBoundary">Row number of min defined row height.</param>
        /// <param name="hiddenBoundary">Row number of min defined hidden row.</param>
        /// <returns>Min row number or -1 if nothing valid defined.</returns>
        private int GetMinRow(int cellBoundary, int heightBoundary, int hiddenBoundary)
        {
            int lowest = int.MaxValue;
            if (cellBoundary >= 0)
            {
                lowest = cellBoundary;
            }
            if (heightBoundary >= 0 && heightBoundary < lowest)
            {
                lowest = heightBoundary;
            }
            if (hiddenBoundary >= 0 && hiddenBoundary < lowest)
            {
                lowest = hiddenBoundary;
            }
            return lowest == int.MaxValue ? -1 : lowest;
        }

        /// <summary>
        /// Gets the current column number (zero based)
        /// </summary>
        /// <returns>Column number (zero-based).</returns>
        public int GetCurrentColumnNumber()
        {
            return currentColumnNumber;
        }

        /// <summary>
        /// Gets the current row number (zero based)
        /// </summary>
        /// <returns>Row number (zero-based).</returns>
        public int GetCurrentRowNumber()
        {
            return currentRowNumber;
        }

        /// <summary>
        /// Moves the current position to the next column
        /// </summary>
        public void GoToNextColumn()
        {
            currentColumnNumber++;
            currentRowNumber = 0;
            Cell.ValidateColumnNumber(currentColumnNumber);
        }

        /// <summary>
        /// Moves the current position to the next column with the number of cells to move
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to move.</param>
        /// <param name="keepRowPosition">If true, the row position is preserved, otherwise set to 0.</param>
        public void GoToNextColumn(int numberOfColumns, bool keepRowPosition = false)
        {
            currentColumnNumber += numberOfColumns;
            if (!keepRowPosition)
            {
                currentRowNumber = 0;
            }
            Cell.ValidateColumnNumber(currentColumnNumber);
        }

        /// <summary>
        /// Moves the current position to the next row (use for a new line)
        /// </summary>
        public void GoToNextRow()
        {
            currentRowNumber++;
            currentColumnNumber = 0;
            Cell.ValidateRowNumber(currentRowNumber);
        }

        /// <summary>
        /// Moves the current position to the next row with the number of cells to move (use for a new line)
        /// </summary>
        /// <param name="numberOfRows">Number of rows to move.</param>
        /// <param name="keepColumnPosition">If true, the column position is preserved, otherwise set to 0.</param>
        public void GoToNextRow(int numberOfRows, bool keepColumnPosition = false)
        {
            currentRowNumber += numberOfRows;
            if (!keepColumnPosition)
            {
                currentColumnNumber = 0;
            }
            Cell.ValidateRowNumber(currentRowNumber);
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="cellRange">Range to merge.</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12').</returns>
        public string MergeCells(Cell.Range cellRange)
        {
            return MergeCells(cellRange.StartAddress, cellRange.EndAddress);
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="cellRange">Range to merge (e.g. 'A1:B12').</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12').</returns>
        public string MergeCells(string cellRange)
        {
            Cell.Range range = Cell.ResolveCellRange(cellRange);
            return MergeCells(range.StartAddress, range.EndAddress);
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="startAddress">Start address of the merged cell range.</param>
        /// <param name="endAddress">End address of the merged cell range.</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12').</returns>
        public string MergeCells(Cell.Address startAddress, Cell.Address endAddress)
        {
            string key = startAddress + ":" + endAddress;
            Cell.Range value = new Cell.Range(startAddress, endAddress);
            IReadOnlyList<Cell.Address> enclosedAddress = value.ResolveEnclosedAddresses();
            foreach (KeyValuePair<string, Cell.Range> item in mergedCells)
            {
                if (item.Value.ResolveEnclosedAddresses().Intersect(enclosedAddress).ToList().Count > 0)
                {
                    throw new RangeException("ConflictingRangeException", "The passed range: " + value.ToString() + " contains cells that are already in the defined merge range: " + item.Key);
                }
            }
            mergedCells.Add(key, value);
            return key;
        }

        /// <summary>
        /// Method to recalculate the auto filter (columns) of this worksheet. This is an internal method. There is no need to use it
        /// </summary>
        internal void RecalculateAutoFilter()
        {
            if (autoFilterRange == null) { return; }
            int start = autoFilterRange.Value.StartAddress.Column;
            int end = autoFilterRange.Value.EndAddress.Column;
            int endRow = 0;
            foreach (KeyValuePair<string, Cell> item in Cells)
            {
                if (item.Value.ColumnNumber < start || item.Value.ColumnNumber > end) { continue; }
                if (item.Value.RowNumber > endRow) { endRow = item.Value.RowNumber; }
            }
            Column c;
            for (int i = start; i <= end; i++)
            {
                if (!columns.ContainsKey(i))
                {
                    c = new Column(i);
                    c.HasAutoFilter = true;
                    columns.Add(i, c);
                }
                else
                {
                    columns[i].HasAutoFilter = true;
                }
            }
            Cell.Range temp = new Cell.Range();
            temp.StartAddress = new Cell.Address(start, 0);
            temp.EndAddress = new Cell.Address(end, endRow);
            autoFilterRange = temp;
        }

        /// <summary>
        /// Method to recalculate the collection of columns of this worksheet. This is an internal method. There is no need to use it
        /// </summary>
        internal void RecalculateColumns()
        {
            List<int> columnsToDelete = new List<int>();
            foreach (KeyValuePair<int, Column> col in columns)
            {
                if (!col.Value.HasAutoFilter && !col.Value.IsHidden && Math.Abs(col.Value.Width - DEFAULT_COLUMN_WIDTH) <= FLOAT_THRESHOLD)
                {
                    columnsToDelete.Add(col.Key);
                }
                if (!col.Value.HasAutoFilter && !col.Value.IsHidden && Math.Abs(col.Value.Width - DEFAULT_COLUMN_WIDTH) <= FLOAT_THRESHOLD)
                {
                    columnsToDelete.Add(col.Key);
                }
            }
            foreach (int index in columnsToDelete)
            {
                columns.Remove(index);
            }
        }

        /// <summary>
        /// Method to resolve all merged cells of the worksheet. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.<br/>
        /// This is an internal method. There is no need to use it
        /// </summary>
        internal void ResolveMergedCells()
        {
            Style mergeStyle = Style.BasicStyles.MergeCellStyle;
            Cell cell;
            foreach (KeyValuePair<string, Cell.Range> range in MergedCells)
            {
                int pos = 0;
                List<Cell.Address> addresses = Cell.GetCellRange(range.Value.StartAddress, range.Value.EndAddress) as List<Cell.Address>;
                foreach (Cell.Address address in addresses)
                {
                    if (!Cells.ContainsKey(address.GetAddress()))
                    {
                        cell = new Cell();
                        cell.DataType = Cell.CellType.EMPTY;
                        cell.RowNumber = address.Row;
                        cell.ColumnNumber = address.Column;
                        AddCell(cell, cell.ColumnNumber, cell.RowNumber);
                    }
                    else
                    {
                        cell = Cells[address.GetAddress()];
                    }
                    if (pos != 0)
                    {
                        if (cell.CellStyle == null)
                        {
                            cell.SetStyle(mergeStyle);
                        }
                        else
                        {
                            Style mixedMergeStyle = cell.CellStyle;
                            // TODO: There should be a better possibility to identify particular style elements that deviates
                            mixedMergeStyle.CurrentCellXf.ForceApplyAlignment = mergeStyle.CurrentCellXf.ForceApplyAlignment;
                            cell.SetStyle(mixedMergeStyle);
                        }
                    }
                    pos++;
                }
            }
        }

        /// <summary>
        /// Removes auto filters from the worksheet
        /// </summary>
        public void RemoveAutoFilter()
        {
            autoFilterRange = null;
        }

        /// <summary>
        /// Sets a previously defined, hidden column as visible again
        /// </summary>
        /// <param name="columnNumber">Column number to make visible again.</param>
        public void RemoveHiddenColumn(int columnNumber)
        {
            SetColumnHiddenState(columnNumber, false);
        }

        /// <summary>
        /// Sets a previously defined, hidden column as visible again
        /// </summary>
        /// <param name="columnAddress">Column address to make visible again.</param>
        public void RemoveHiddenColumn(string columnAddress)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnHiddenState(columnNumber, false);
        }

        /// <summary>
        /// Sets a previously defined, hidden row as visible again
        /// </summary>
        /// <param name="rowNumber">Row number to hide on the worksheet.</param>
        public void RemoveHiddenRow(int rowNumber)
        {
            SetRowHiddenState(rowNumber, false);
        }

        /// <summary>
        /// Removes the defined merged cell range
        /// </summary>
        /// <param name="range">Cell range to remove the merging.</param>
        public void RemoveMergedCells(string range)
        {
            if (range != null)
            {
                range = range.ToUpper();
            }
            if (range == null || !mergedCells.ContainsKey(range))
            {
                throw new RangeException("UnknownRangeException", "The cell range " + range + " was not found in the list of merged cell ranges");
            }

            List<Cell.Address> addresses = Cell.GetCellRange(range) as List<Cell.Address>;
            foreach (Cell.Address address in addresses)
            {
                if (cells.ContainsKey(address.GetAddress()))
                {
                    Cell cell = cells[address.ToString()];
                    if (Style.BasicStyles.MergeCellStyle.Equals(cell.CellStyle))
                    {
                        cell.RemoveStyle();
                    }
                    cell.ResolveCellType(); // resets the type
                }
            }
            mergedCells.Remove(range);
        }

        /// <summary>
        /// Removes the cell selection of this worksheet
        /// </summary>
        public void RemoveSelectedCells()
        {
            selectedCells.Clear();
        }

        /// <summary>
        /// Removes the defined, non-standard row height
        /// </summary>
        /// <param name="rowNumber">Row number (zero-based).</param>
        public void RemoveRowHeight(int rowNumber)
        {
            if (rowHeights.ContainsKey(rowNumber))
            {
                rowHeights.Remove(rowNumber);
            }
        }

        /// <summary>
        /// Removes an allowed action on the current worksheet or its cells
        /// </summary>
        /// <param name="value">Allowed action on the worksheet or cells.</param>
        public void RemoveAllowedActionOnSheetProtection(SheetProtectionValue value)
        {
            if (sheetProtectionValues.Contains(value))
            {
                sheetProtectionValues.Remove(value);
            }
        }

        /// <summary>
        /// Sets the active style of the worksheet. This style will be assigned to all later added cells
        /// </summary>
        /// <param name="style">Style to set as active style.</param>
        public void SetActiveStyle(Style style)
        {
            if (style == null)
            {
                useActiveStyle = false;
            }
            else
            {
                useActiveStyle = true;
            }
            activeStyle = style;
        }

        /// <summary>
        /// Sets the column auto filter within the defined column range
        /// </summary>
        /// <param name="startColumn">Column number with the first appearance of an auto filter drop down.</param>
        /// <param name="endColumn">Column number with the last appearance of an auto filter drop down.</param>
        public void SetAutoFilter(int startColumn, int endColumn)
        {
            string start = Cell.ResolveCellAddress(startColumn, 0);
            string end = Cell.ResolveCellAddress(endColumn, 0);
            if (endColumn < startColumn)
            {
                SetAutoFilter(end + ":" + start);
            }
            else
            {
                SetAutoFilter(start + ":" + end);
            }
        }

        /// <summary>
        /// Sets the column auto filter within the defined column range
        /// </summary>
        /// <param name="range">Range to apply auto filter on. The range could be 'A1:C10' for instance. The end row will be recalculated automatically when saving the file.</param>
        public void SetAutoFilter(string range)
        {
            autoFilterRange = Cell.ResolveCellRange(range);
            RecalculateAutoFilter();
            RecalculateColumns();
        }

        /// <summary>
        /// Sets the defined column as hidden or visible
        /// </summary>
        /// <param name="columnNumber">Column number to hide on the worksheet.</param>
        /// <param name="state">If true, the column will be hidden, otherwise be visible.</param>
        private void SetColumnHiddenState(int columnNumber, bool state)
        {
            Cell.ValidateColumnNumber(columnNumber);
            if (columns.ContainsKey(columnNumber))
            {
                columns[columnNumber].IsHidden = state;
            }
            else if (state)
            {
                Column c = new Column(columnNumber);
                c.IsHidden = true;
                columns.Add(columnNumber, c);
            }
            if (!columns[columnNumber].IsHidden && Math.Abs(columns[columnNumber].Width - DEFAULT_COLUMN_WIDTH) <= FLOAT_THRESHOLD && !columns[columnNumber].HasAutoFilter)
            {
                columns.Remove(columnNumber);
            }
        }

        /// <summary>
        /// Sets the width of the passed column address
        /// </summary>
        /// <param name="columnAddress">Column address (A - XFD).</param>
        /// <param name="width">Width from 0 to 255.0.</param>
        public void SetColumnWidth(string columnAddress, float width)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnWidth(columnNumber, width);
        }

        /// <summary>
        /// Sets the width of the passed column number (zero-based)
        /// </summary>
        /// <param name="columnNumber">Column number (zero-based, from 0 to 16383).</param>
        /// <param name="width">Width from 0 to 255.0.</param>
        public void SetColumnWidth(int columnNumber, float width)
        {
            Cell.ValidateColumnNumber(columnNumber);
            if (width < MIN_COLUMN_WIDTH || width > MAX_COLUMN_WIDTH)
            {
                throw new RangeException("OutOfRangeException", "The column width (" + width + ") is out of range. Range is from " + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + " (chars).");
            }
            if (columns.ContainsKey(columnNumber))
            {
                columns[columnNumber].Width = width;
            }
            else
            {
                Column c = new Column(columnNumber);
                c.Width = width;
                columns.Add(columnNumber, c);
            }
        }

        /// <summary>
        /// Set the current cell address
        /// </summary>
        /// <param name="columnNumber">Column number (zero based).</param>
        /// <param name="rowNumber">Row number (zero based).</param>
        public void SetCurrentCellAddress(int columnNumber, int rowNumber)
        {
            SetCurrentColumnNumber(columnNumber);
            SetCurrentRowNumber(rowNumber);
        }

        /// <summary>
        /// Set the current cell address
        /// </summary>
        /// <param name="address">Cell address in the format A1 - XFD1048576.</param>
        public void SetCurrentCellAddress(string address)
        {
            int row, column;
            Cell.ResolveCellCoordinate(address, out column, out row);
            SetCurrentCellAddress(column, row);
        }

        /// <summary>
        /// Sets the current column number (zero based)
        /// </summary>
        /// <param name="columnNumber">Column number (zero based).</param>
        public void SetCurrentColumnNumber(int columnNumber)
        {
            Cell.ValidateColumnNumber(columnNumber);
            currentColumnNumber = columnNumber;
        }

        /// <summary>
        /// Sets the current row number (zero based)
        /// </summary>
        /// <param name="rowNumber">Row number (zero based).</param>
        public void SetCurrentRowNumber(int rowNumber)
        {
            Cell.ValidateRowNumber(rowNumber);
            currentRowNumber = rowNumber;
        }

        /// <summary>
        /// Sets a single range of selected cells on this worksheet. All existing ranges will be removed
        /// </summary>
        /// <param name="range">Range to set as single cell range for selected cells</param>
        [Obsolete("This method is a deprecated subset of the function AddSelectedCells. It will be removed in a future version")]
        public void SetSelectedCells(Cell.Range range)
        {
            RemoveSelectedCells();
            AddSelectedCells(range);
        }

        /// <summary>
        /// Sets the selected cells on this worksheet
        /// </summary>
        /// <param name="startAddress">Start address of the range to set as single cell range for selected cells</param>
        /// <param name="endAddress">End address of the range to set as single cell range for selected cells</param>
        [Obsolete("This method is a deprecated subset of the function AddSelectedCells. It will be removed in a future version")]
        public void SetSelectedCells(Cell.Address startAddress, Cell.Address endAddress)
        {
            SetSelectedCells(new Range(startAddress, endAddress));
        }

        /// <summary>
        /// Sets a single range of selected cells on this worksheet. All existing ranges will be removed. Null will remove all selected cells
        /// </summary>
        /// <param name="range">Range as string to set as single cell range for selected cells, or null to remove the selected cells</param>
        [Obsolete("This method is a deprecated subset of the function AddSelectedCells. It will be removed in a future version")]
        public void SetSelectedCells(string range)
        {
            if (range == null)
            {
                selectedCells.Clear();
                return;
            }
            else
            {
                SetSelectedCells(new Range(range));
            }
        }

        /// <summary>
        /// Adds a range to the selected cells on this worksheet
        /// </summary>
        /// <param name="range">Cell range to be added as selected cells</param>
        public void AddSelectedCells(Range range)
        {
            selectedCells.Add(range);
        }

        /// <summary>
        /// Adds a range to the selected cells on this worksheet
        /// </summary>
        /// <param name="startAddress">Start address of the range to add</param>
        /// <param name="endAddress">End address of the range to add</param>
        public void AddSelectedCells(Address startAddress, Address endAddress)
        {
            selectedCells.Add(new Range(startAddress, endAddress));
        }

        /// <summary>
        /// Adds a range to the selected cells on this worksheet. Null or empty as value will be ignored
        /// </summary>
        /// <param name="range">Cell range to add as selected cells</param>
        public void AddSelectedCells(string range)
        {
            if (range != null)
            {
                selectedCells.Add(Cell.ResolveCellRange(range));
            }
        }

        /// <summary>
        /// Sets or removes the password for worksheet protection. If set, UseSheetProtection will be also set to true
        /// </summary>
        /// <param name="password">Password (UTF-8) to protect the worksheet. If the password is null or empty, no password will be used.</param>
        public void SetSheetProtectionPassword(string password)
        {
            if (string.IsNullOrEmpty(password))
            {
                sheetProtectionPassword = null;
                sheetProtectionPasswordHash = null;
                UseSheetProtection = false;
            }
            else
            {
                sheetProtectionPassword = password;
                sheetProtectionPasswordHash = LowLevel.GeneratePasswordHash(password);
                UseSheetProtection = true;
            }
        }

        /// <summary>
        /// Sets the height of the passed row number (zero-based)
        /// </summary>
        /// <param name="rowNumber">Row number (zero-based, 0 to 1048575).</param>
        /// <param name="height">Height from 0 to 409.5.</param>
        public void SetRowHeight(int rowNumber, float height)
        {
            Cell.ValidateRowNumber(rowNumber);
            if (height < MIN_ROW_HEIGHT || height > MAX_ROW_HEIGHT)
            {
                throw new RangeException("OutOfRangeException", "The row height (" + height + ") is out of range. Range is from " + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + " (equals 546px).");
            }
            if (rowHeights.ContainsKey(rowNumber))
            {
                rowHeights[rowNumber] = height;
            }
            else
            {
                rowHeights.Add(rowNumber, height);
            }
        }

        /// <summary>
        /// Sets the defined row as hidden or visible
        /// </summary>
        /// <param name="rowNumber">Row number to make visible again.</param>
        /// <param name="state">If true, the row will be hidden, otherwise visible.</param>
        private void SetRowHiddenState(int rowNumber, bool state)
        {
            Cell.ValidateRowNumber(rowNumber);
            if (hiddenRows.ContainsKey(rowNumber))
            {
                if (state)
                {
                    hiddenRows[rowNumber] = true;
                }
                else
                {
                    hiddenRows.Remove(rowNumber);
                }
            }
            else if (state)
            {
                hiddenRows.Add(rowNumber, true);
            }
        }

        /// <summary>
        /// Validates and sets the worksheet name
        /// </summary>
        /// <param name="name">Name to set.</param>
        public void SetSheetName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new FormatException("the worksheet name must be between 1 and " + MAX_WORKSHEET_NAME_LENGTH + " characters");
            }
            if (name.Length > MAX_WORKSHEET_NAME_LENGTH)
            {
                throw new FormatException("the worksheet name must be between 1 and " + MAX_WORKSHEET_NAME_LENGTH + " characters");
            }
            Regex regex = new Regex(@"[\[\]\*\?/\\]");
            Match match = regex.Match(name);
            if (match.Captures.Count > 0)
            {
                throw new FormatException(@"the worksheet name must not contain the characters [  ]  * ? / \ ");
            }
            sheetName = name;
        }

        /// <summary>
        /// Sets the name of the worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet.</param>
        /// <param name="sanitize">If true, the filename will be sanitized automatically according to the specifications of Excel.</param>
        public void SetSheetName(string name, bool sanitize)
        {
            if (sanitize)
            {
                sheetName = ""; // Empty name (temporary) to prevent conflicts during sanitizing
                sheetName = SanitizeWorksheetName(name, workbookReference);
            }
            else
            {
                SetSheetName(name);
            }
        }

        /// <summary>
        /// Sets the horizontal split of the worksheet into two panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="topPaneHeight">Height (similar to row height) from top of the worksheet to the split line in characters.</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the row component is important in a horizontal split.</param>
        /// <param name="activePane">Active pane in the split window.</param>
        public void SetHorizontalSplit(float topPaneHeight, Cell.Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(null, topPaneHeight, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the horizontal split of the worksheet into two panes. The measurement in rows can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfRowsFromTop">Number of rows from top of the worksheet to the split line. The particular row heights are considered.</param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable.</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the row component is important in a horizontal split.</param>
        /// <param name="activePane">Active pane in the split window.</param>
        public void SetHorizontalSplit(int numberOfRowsFromTop, bool freeze, Cell.Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(null, numberOfRowsFromTop, freeze, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the vertical split of the worksheet into two panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="leftPaneWidth">Width (similar to column width) from left of the worksheet to the split line in characters.</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the column component is important in a vertical split.</param>
        /// <param name="activePane">Active pane in the split window.</param>
        public void SetVerticalSplit(float leftPaneWidth, Cell.Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(leftPaneWidth, null, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the vertical split of the worksheet into two panes. The measurement in columns can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfColumnsFromLeft">Number of columns from left of the worksheet to the split line. The particular column widths are considered.</param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable.</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the column component is important in a vertical split.</param>
        /// <param name="activePane">Active pane in the split window.</param>
        public void SetVerticalSplit(int numberOfColumnsFromLeft, bool freeze, Cell.Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(numberOfColumnsFromLeft, null, freeze, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the horizontal and vertical split of the worksheet into four panes. The measurement in rows and columns can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfColumnsFromLeft">The numberOfColumnsFromLeft<see cref="int?"/>.</param>
        /// <param name="numberOfRowsFromTop">The numberOfRowsFromTop<see cref="int?"/>.</param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable.</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable).</param>
        /// <param name="activePane">Active pane in the split window.</param>
        public void SetSplit(int? numberOfColumnsFromLeft, int? numberOfRowsFromTop, bool freeze, Cell.Address topLeftCell, WorksheetPane activePane)
        {
            if (freeze)
            {
                if (numberOfColumnsFromLeft != null && topLeftCell.Column < numberOfColumnsFromLeft.Value)
                {
                    throw new WorksheetException("The column number " + topLeftCell.Column +
                        " is not valid for a frozen, vertical split with the split pane column number " + numberOfColumnsFromLeft.Value);
                }
                if (numberOfRowsFromTop != null && topLeftCell.Row < numberOfRowsFromTop.Value)
                {
                    throw new WorksheetException("The row number " + topLeftCell.Row +
                        " is not valid for a frozen, horizontal split height the split pane row number " + numberOfRowsFromTop.Value);
                }
            }
            this.paneSplitLeftWidth = null;
            this.paneSplitTopHeight = null;
            this.freezeSplitPanes = freeze;
            int row = numberOfRowsFromTop != null ? numberOfRowsFromTop.Value : 0;
            int column = numberOfColumnsFromLeft != null ? numberOfColumnsFromLeft.Value : 0;
            this.paneSplitAddress = new Cell.Address(column, row);
            this.paneSplitTopLeftCell = topLeftCell;
            this.activePane = activePane;
        }

        /// <summary>
        /// Sets the horizontal and vertical split of the worksheet into four panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="leftPaneWidth">The leftPaneWidth<see cref="float?"/>.</param>
        /// <param name="topPaneHeight">The topPaneHeight<see cref="float?"/>.</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable).</param>
        /// <param name="activePane">Active pane in the split window.</param>
        public void SetSplit(float? leftPaneWidth, float? topPaneHeight, Cell.Address topLeftCell, WorksheetPane activePane)
        {
            this.paneSplitLeftWidth = leftPaneWidth;
            this.paneSplitTopHeight = topPaneHeight;
            this.freezeSplitPanes = null;
            this.paneSplitAddress = null;
            this.paneSplitTopLeftCell = topLeftCell;
            this.activePane = activePane;
        }

        /// <summary>
        /// Resets splitting of the worksheet into panes, as well as their freezing
        /// </summary>
        public void ResetSplit()
        {
            this.paneSplitLeftWidth = null;
            this.paneSplitTopHeight = null;
            this.freezeSplitPanes = null;
            this.paneSplitAddress = null;
            this.paneSplitTopLeftCell = null;
            this.activePane = null;
        }

        /// <summary>
        /// Creates a (dereferenced) deep copy of this worksheet
        /// </summary>
        /// <returns>The <see cref="Worksheet"/>.</returns>
        public Worksheet Copy()
        {
            Worksheet copy = new Worksheet();
            foreach (KeyValuePair<String, Cell> cell in this.cells)
            {
                copy.AddCell(cell.Value.Copy(), cell.Key);
            }
            copy.activePane = this.activePane;
            copy.activeStyle = this.activeStyle;
            if (this.autoFilterRange.HasValue)
            {
                copy.autoFilterRange = this.autoFilterRange.Value.Copy();
            }
            foreach (KeyValuePair<int, Column> column in this.columns)
            {
                copy.columns.Add(column.Key, column.Value.Copy());
            }
            copy.CurrentCellDirection = this.CurrentCellDirection;
            copy.currentColumnNumber = this.currentColumnNumber;
            copy.currentRowNumber = this.currentRowNumber;
            copy.defaultColumnWidth = this.defaultColumnWidth;
            copy.defaultRowHeight = this.defaultRowHeight;
            copy.freezeSplitPanes = this.freezeSplitPanes;
            copy.hidden = this.hidden;
            foreach (KeyValuePair<int, bool> row in this.hiddenRows)
            {
                copy.hiddenRows.Add(row.Key, row.Value);
            }
            foreach (KeyValuePair<string, Cell.Range> cell in this.mergedCells)
            {
                copy.mergedCells.Add(cell.Key, cell.Value.Copy());
            }
            if (this.paneSplitAddress.HasValue)
            {
                copy.paneSplitAddress = this.paneSplitAddress.Value.Copy();
            }
            copy.paneSplitLeftWidth = this.paneSplitLeftWidth;
            copy.paneSplitTopHeight = this.paneSplitTopHeight;
            if (this.paneSplitTopLeftCell.HasValue)
            {
                copy.paneSplitTopLeftCell = this.paneSplitTopLeftCell.Value.Copy();
            }
            foreach (KeyValuePair<int, float> row in this.rowHeights)
            {
                copy.rowHeights.Add(row.Key, row.Value);
            }
            if (this.selectedCells.Count > 0)
            {
                foreach (Range selectedCellRange in this.selectedCells)
                {
                    copy.AddSelectedCells(selectedCellRange.Copy());
                }
            }
            copy.sheetProtectionPassword = this.sheetProtectionPassword;
            copy.sheetProtectionPasswordHash = this.sheetProtectionPasswordHash;
            foreach (SheetProtectionValue value in this.sheetProtectionValues)
            {
                copy.sheetProtectionValues.Add(value);
            }
            copy.useActiveStyle = this.useActiveStyle;
            copy.UseSheetProtection = this.UseSheetProtection;
            return copy;
        }

        /// <summary>
        /// Sanitizes a worksheet name
        /// </summary>
        /// <param name="input">Name to sanitize.</param>
        /// <param name="workbook">Workbook reference.</param>
        /// <returns>Name of the sanitized worksheet.</returns>
        public static string SanitizeWorksheetName(string input, Workbook workbook)
        {
            if (input == null) { input = "Sheet1"; }
            int len = input.Length;
            if (len > 31) { len = 31; }
            else if (len == 0)
            {
                input = "Sheet1";
            }
            StringBuilder sb = new StringBuilder(31);
            char c;
            for (int i = 0; i < len; i++)
            {
                c = input[i];
                if (c == '[' || c == ']' || c == '*' || c == '?' || c == '\\' || c == '/')
                { sb.Append('_'); }
                else
                { sb.Append(c); }
            }
            return GetUnusedWorksheetName(sb.ToString(), workbook);
        }

        /// <summary>
        /// Determines the next unused worksheet name in the passed workbook
        /// </summary>
        /// <param name="name">Original name to start the check.</param>
        /// <param name="workbook">Workbook to look for existing worksheets.</param>
        /// <returns>Not yet used worksheet name.</returns>
        private static string GetUnusedWorksheetName(string name, Workbook workbook)
        {
            if (workbook == null)
            {
                throw new WorksheetException("The workbook reference is null");
            }
            if (!WorksheetExists(name, workbook))
            { return name; }
            Regex regex = new Regex(@"^(.*?)(\d{1,31})$");
            Match match = regex.Match(name);
            string prefix = name;
            int number = 1;
            if (match.Groups.Count > 1)
            {
                prefix = match.Groups[1].Value;
                int.TryParse(match.Groups[2].Value, out number);
                // if this failed, the start number is 0 (parsed number was >max. int32)
            }
            while (true)
            {
                string numberString = number.ToString("G", CultureInfo.InvariantCulture);
                if (numberString.Length + prefix.Length > MAX_WORKSHEET_NAME_LENGTH)
                {
                    int endIndex = prefix.Length - (numberString.Length + prefix.Length - MAX_WORKSHEET_NAME_LENGTH);
                    prefix = prefix.Substring(0, endIndex);
                }
                string newName = prefix + numberString;
                if (!WorksheetExists(newName, workbook))
                { return newName; }
                number++;
            }
        }

        /// <summary>
        /// Checks whether a worksheet with the given name exists
        /// </summary>
        /// <param name="name">Name to check.</param>
        /// <param name="workbook">Workbook reference.</param>
        /// <returns>True if the name exits, otherwise false.</returns>
        private static bool WorksheetExists(string name, Workbook workbook)
        {
            if (workbook == null)
            {
                throw new WorksheetException("The workbook reference is null");
            }
            int len = workbook.Worksheets.Count;
            for (int i = 0; i < len; i++)
            {
                if (workbook.Worksheets[i].SheetName == name)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Class representing a column of a worksheet
        /// </summary>
        public class Column
        {
            /// <summary>
            /// Defines the number
            /// </summary>
            private int number;

            /// <summary>
            /// Defines the columnAddress
            /// </summary>
            private string columnAddress;

            /// <summary>
            /// Defines the width
            /// </summary>
            private float width;

            /// <summary>
            /// Gets or sets the ColumnAddress
            /// Column address (A to XFD)
            /// </summary>
            public string ColumnAddress
            {
                get { return columnAddress; }
                set
                {
                    if (string.IsNullOrEmpty(value))
                    {
                        throw new RangeException("A general range exception occurred", "The passed address was null or empty");
                    }
                    number = Cell.ResolveColumn(value);
                    columnAddress = value.ToUpper();
                }
            }

            /// <summary>
            /// Gets or sets a value indicating whether HasAutoFilter
            /// If true, the column has auto filter applied, otherwise not
            /// </summary>
            public bool HasAutoFilter { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether IsHidden
            /// If true, the column is hidden, otherwise visible
            /// </summary>
            public bool IsHidden { get; set; }

            /// <summary>
            /// Gets or sets the Number
            /// Column number (0 to 16383)
            /// </summary>
            public int Number
            {
                get { return number; }
                set
                {
                    columnAddress = Cell.ResolveColumnAddress(value);
                    number = value;
                }
            }

            /// <summary>
            /// Gets or sets the Width
            /// Width of the column
            /// </summary>
            public float Width
            {
                get { return width; }
                set
                {
                    if (value < Worksheet.MIN_COLUMN_WIDTH || value > Worksheet.MAX_COLUMN_WIDTH)
                    {
                        throw new RangeException("A general range exception occurred", "The passed column width is out of range (" + Worksheet.MIN_COLUMN_WIDTH + " to " + Worksheet.MAX_COLUMN_WIDTH + ")");
                    }
                    width = value;
                }
            }

            /// <summary>
            /// Prevents a default instance of the <see cref="Column"/> class from being created
            /// </summary>
            private Column()
            {
                Width = DEFAULT_COLUMN_WIDTH;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Column"/> class
            /// </summary>
            /// <param name="columnCoordinate">Column number (zero-based, 0 to 16383).</param>
            public Column(int columnCoordinate) : this()
            {
                Number = columnCoordinate;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Column"/> class
            /// </summary>
            /// <param name="columnAddress">Column address (A to XFD).</param>
            public Column(string columnAddress) : this()
            {
                ColumnAddress = columnAddress;
            }

            /// <summary>
            /// Creates a deep copy of this column
            /// </summary>
            /// <returns>Copy of this column.</returns>
            internal Column Copy()
            {
                Column copy = new Column();
                copy.IsHidden = this.IsHidden;
                copy.Width = this.width;
                copy.HasAutoFilter = this.HasAutoFilter;
                copy.columnAddress = this.columnAddress;
                copy.number = this.number;
                return copy;
            }
        }
    }
}

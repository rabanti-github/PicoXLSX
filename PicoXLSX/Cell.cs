/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace PicoXLSX
{
    /// <summary>
    /// Class representing a cell of a worksheet
    /// </summary>
    public class Cell : IComparable<Cell>
    {

#region enums
        /// <summary>
        /// Enum defines the basic data types of a cell
        /// </summary>
        public enum CellType
        {
            /// <summary>Type for single characters and strings</summary>
            STRING,
            /// <summary>Type for all numeric types (long, integer and float and double)</summary>
            NUMBER,
            /// <summary>Type for dates and times (Note: Dates before 1900-01-01 are not allowed)</summary>
            DATE,
            /// <summary>Type for boolean</summary>
            BOOL,
            /// <summary>Type for Formulas (The cell will be handled differently)</summary>
            FORMULA,
            /// <summary>Type for empty cells. This type is only used for merged cells (all cells except the first of the cell range)</summary>
            EMPTY,
            /// <summary>Default Type, not specified</summary>
            DEFAULT
        }
#endregion

#region privateFileds
        private Style cellStyle;
        private int columnAddress;
        private int rowAddress;
#endregion

#region properties

        /// <summary>
        /// Gets or sets the combined cell Address as string in the format A1 - XFD1048576
        /// </summary>
        public string CellAddress
        {
            get { return Cell.ResolveCellAddress(this.ColumnAddress, this.RowAddress); }
            set { Cell.ResolveCellCoordinate(value, out this.columnAddress, out this.rowAddress); }
        }

        /// <summary>Gets or sets the combined cell Address as Address object</summary>
        public Address CellAddress2
        {
            get { return new Address(this.ColumnAddress, this.RowAddress); }
            set
            {
                this.ColumnAddress = value.Column;
                this.RowAddress = value.Row;
            }
        }

        /// <summary>
        /// Assigned style of the cell
        /// </summary>
        public Style CellStyle
        {
            get { return cellStyle; }
        }

        /// <summary>Number of the column (zero-based)</summary>  
        /// <exception cref="RangeException">Throws a RangeException if the column address is out of range</exception>
        public int ColumnAddress
        {
            get { return columnAddress; }
            set
            {
                if (value < Worksheet.MIN_COLUMN_ADDRESS || value > Worksheet.MAX_COLUMN_ADDRESS)
                {
                    throw new RangeException("OutOfRangeException","The passed column number (" + value.ToString() + ") is out of range. Range is from " + Worksheet.MIN_COLUMN_ADDRESS.ToString() + " to " + Worksheet.MAX_COLUMN_ADDRESS.ToString() + " (" + (Worksheet.MAX_COLUMN_ADDRESS + 1).ToString() + " rows).");
                }
                columnAddress = value;
            }
        }

        /// <summary>Type of the cell</summary>
        public CellType DataType { get; set; }


        /// <summary>Number of the row (zero-based)</summary>
        /// <exception cref="RangeException">Throws a RangeException if the row address is out of range</exception>
        public int RowAddress
        {
            get { return rowAddress; }
            set
            {
                if (value < Worksheet.MIN_ROW_ADDRESS || value > Worksheet.MAX_ROW_ADDRESS)
                {
                    throw new RangeException("OutOfRangeException", "The passed row number (" + value.ToString() + ") is out of range. Range is from " + Worksheet.MIN_ROW_ADDRESS.ToString() + " to " + Worksheet.MAX_ROW_ADDRESS.ToString() + " (" + (Worksheet.MAX_ROW_ADDRESS + 1).ToString() + " rows).");
                }
                rowAddress = value;
            }
        }

        /// <summary>Value of the cell (generic object type)</summary>
        public object Value { get; set; }

        /// <summary>
        /// Gets or sets the parent worksheet reference
        /// </summary>
        public Worksheet WorksheetReference { get; set; }

#endregion

#region constructors
        /// <summary>Default constructor</summary>
        public Cell()
        {
            this.WorksheetReference = null;
        }

        /// <summary>
        /// Constructor with value and cell type
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        public Cell(object value, CellType type)
        {
            this.Value = value;
            this.DataType = type;
            ResolveCellType();
        }

        /// <summary>
        /// Constructor with value, cell type, row address and column address
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        /// <param name="column">Column address of the cell (zero-based)</param>
        /// <param name="row">Row address of the cell (zero-based)</param>
        /// <param name="reference">Referenced worksheet which contains the cell</param>
        public Cell(object value, CellType type, int column, int row, Worksheet reference) : this(value, type)
        {
            this.ColumnAddress = column;
            this.RowAddress = row;
            this.WorksheetReference = reference;
            if (type == CellType.DEFAULT)
            {
            	ResolveCellType();
            }
        }
#endregion

#region methods
        /// <summary>
        /// Implemented CompareTo method
        /// </summary>
        /// <param name="other">Object to compare</param>
        /// <returns>0 if values are the same, -1 if this object is smaller, 1 if it is bigger</returns>
        public int CompareTo(Cell other)
        {
            if (this.RowAddress == other.RowAddress)
            {
                return this.ColumnAddress.CompareTo(other.ColumnAddress);
            }
            else
            {
                return this.RowAddress.CompareTo(other.RowAddress);
            }
        }

        /// <summary>
        /// Removes the assigned style from the cell
        /// </summary>
        /// <exception cref="StyleException">Throws an StyleException if the style cannot be referenced</exception>
        public void RemoveStyle()
        {
            if (this.WorksheetReference == null)
            {
                throw new StyleException("UndefinedStyleException", "No worksheet reference was defined while trying to remove a style from a cell");
            }
            if (this.WorksheetReference.WorkbookReference == null)
            {
                throw new StyleException("UndefinedStyleException", "No workbook reference was defined on the worksheet while trying to remove a style from a cell");
            }
            if (this.cellStyle != null)
            {
                string styleName = this.cellStyle.Name;
                this.cellStyle = null;
                this.WorksheetReference.WorkbookReference.RemoveStyle(styleName, true);
            }
        }

        /// <summary>
        /// Method resets the Cell type and tries to find the actual type. This is used if a Cell was created with the CellType DEFAULT. CellTypes FORMULA and EMPTY will skip this method
        /// </summary>
        public void ResolveCellType()
        {
            if (this.Value == null)
            {
                this.DataType = CellType.EMPTY;
                this.Value = "";
                return;
            }
            if (this.DataType == CellType.FORMULA || this.DataType == CellType.EMPTY) { return; }
            Type t = this.Value.GetType();
            if (t == typeof(int)) { this.DataType = CellType.NUMBER; }
            else if (t == typeof(float)) { this.DataType = CellType.NUMBER; }
            else if (t == typeof(double)) { this.DataType = CellType.NUMBER; }
            else if (t == typeof(long)) { this.DataType = CellType.NUMBER; }
            else if (t == typeof(bool)) { this.DataType = CellType.BOOL; }
            else if (t == typeof(DateTime)) { this.DataType = CellType.DATE; }
            else { this.DataType = CellType.STRING; } // Default
        }

        /// <summary>
        /// Sets the lock state of the cell
        /// </summary>
        /// <param name="isLocked">If true, the cell will be locked if the worksheet is protected</param>
        /// <param name="isHidden">If true, the value of the cell will be invisible if the worksheet is protected</param>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the style used to lock cells cannot be referenced</exception>
        /// <remarks>The listed exception should never happen because the mentioned style is internally generated</remarks>
        public void SetCellLockedState(bool isLocked, bool isHidden)
        {
            Style lockStyle;
            if (this.cellStyle == null)
            {
                lockStyle = new Style();
            }
            else
            {
                lockStyle = this.cellStyle.CopyStyle();
            }
            lockStyle.CurrentCellXf.Locked = isLocked;
            lockStyle.CurrentCellXf.Hidden = isHidden;
            this.SetStyle(lockStyle);
        }

        /// <summary>
        /// Sets the style of the cell
        /// </summary>
        /// <param name="style">Style to assign</param>
        /// <returns>If the passed style already exists in the workbook, the existing one will be returned, otherwise the passed one</returns>
        /// <exception cref="StyleException">Throws an StyleException if the style cannot be referenced or no style was defined</exception>
        public Style SetStyle(Style style)
        {
            if (this.WorksheetReference == null)
            {
                throw new StyleException("UndefinedStyleException", "No worksheet reference was defined while trying to set a style to a cell");
            }
            if (this.WorksheetReference.WorkbookReference == null)
            {
                throw new StyleException("UndefinedStyleException", "No workbook reference was defined on the worksheet while trying to set a style to a cell");
            }
            if (style == null)
            {
                throw new StyleException("UndefinedStyleException", "No style to assign was defined");
            }
            Style s = this.WorksheetReference.WorkbookReference.AddStyle(style);
            this.cellStyle = s;
            return s;
        }
#endregion

#region staticMethods
        /// <summary>
        /// Converts a List of supported objects into a list of cells
        /// </summary>
        /// <typeparam name="T">Generic data type</typeparam>
        /// <param name="list">List of generic objects</param>
        /// <returns>List of cells</returns>
        public static List<Cell> ConvertArray<T>(List<T> list)
        {
            List<Cell> output = new List<Cell>();
            Cell c;
            object o;
            Type t;
            foreach (T item in list)
            {
                o = (object)item;
                t = item.GetType();

                if (t == typeof(int))
                {
                    c = new Cell((int)o, CellType.NUMBER);
                }
                else if (t == typeof(float))
                {
                    c = new Cell((float)o, CellType.NUMBER);
                }
                else if (t == typeof(double))
                {
                    c = new Cell((double)o, CellType.NUMBER);
                }
                else if (t == typeof(long))
                {
                    c = new Cell((long)o, CellType.NUMBER);
                }
                else if (t == typeof(bool))
                {
                    c = new Cell((bool)o, CellType.BOOL);
                }
                else if (t == typeof(DateTime))
                {
                    c = new Cell((DateTime)o, CellType.DATE);
                }
                else if (t == typeof(string))
                {
                    c = new Cell((string)o, CellType.STRING);
                }
                else // Default = unspecified object
                {
                    c = new Cell((string)o, CellType.DEFAULT);
                }
                output.Add(c);
            }
            return output;
        }

        /// <summary>
        /// Gets a list of cell addresses from a cell range (format A1:B3 or AAD556:AAD1000)
        /// </summary>
        /// <param name="range">Range to process</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="FormatException">Throws a FormatException if a part of the passed range is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception>
        public static List<Address> GetCellRange(string range)
        {
            Range range2 = ResolveCellRange(range);
            return GetCellRange(range2.StartAddress, range2.EndAddress);
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range
        /// </summary>
        /// <param name="startAddress">Start address as string in the format A1 - XFD1048576</param>
        /// <param name="endAddress">End address as string in the format A1 - XFD1048576</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="FormatException">Throws a FormatException if a part of the passed range is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception> 
        public static List<Address> GetCellRange(string startAddress, string endAddress)
        {
            Address start = ResolveCellCoordinate(startAddress);
            Address end = ResolveCellCoordinate(endAddress);
            return GetCellRange(start, end);
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range
        /// </summary>
        /// <param name="startColumn">Start column (zero based)</param>
        /// <param name="startRow">Start row (zero based)</param>
        /// <param name="endColumn">End column (zero based)</param>
        /// <param name="endRow">End row (zero based)</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="RangeException">Throws an RangeException if the value of one passed address parts is out of range (A-XFD and 1 to 1048576) </exception>
        public static List<Address> GetCellRange(int startColumn, int startRow, int endColumn, int endRow)
        {
            Address start = new Address(startColumn, startRow);
            Address end = new Address(endColumn, endRow);
            return GetCellRange(start, end);
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range
        /// </summary>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="FormatException">Throws a FormatException if a part of the passed addresses is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the value of one passed address is out of range (A-XFD and 1 to 1048576) </exception>
        public static List<Address> GetCellRange(Address startAddress, Address endAddress)
        {
            int startColumn, endColumn, startRow, endRow;
            if (startAddress.Column < endAddress.Column)
            {
                startColumn = startAddress.Column;
                endColumn = endAddress.Column;
            }
            else
            {
                startColumn = endAddress.Column;
                endColumn = startAddress.Column;
            }
            if (startAddress.Row < endAddress.Row)
            {
                startRow = startAddress.Row;
                endRow = endAddress.Row;
            }
            else
            {
                startRow = endAddress.Row;
                endRow = startAddress.Row;
            }
            List<Address> output = new List<Address>();
            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startColumn; j <= endColumn; j++)
                {
                    output.Add(new Address(j, i));
                }
            }
            return output;
        }

        /// <summary>
        /// Gets the address of a cell by the column and row number (zero based)
        /// </summary>
        /// <param name="column">Column address of the cell (zero-based)</param>
        /// <param name="row">Row address of the cell (zero-based)</param>
        /// <exception cref="RangeException">Throws an RangeException if the start or end address was out of range</exception>
        /// <returns>Cell Address as string in the format A1 - XFD1048576</returns>
        public static string ResolveCellAddress(int column, int row)
        {
            if (column > Worksheet.MAX_COLUMN_ADDRESS || column < Worksheet.MIN_COLUMN_ADDRESS)
            {
                throw new RangeException("OutOfRangeException", "The column number (" + column.ToString() + ") is out of range. Range is from " + Worksheet.MIN_COLUMN_ADDRESS.ToString() + " to " + Worksheet.MAX_COLUMN_ADDRESS.ToString() + " (" + (Worksheet.MAX_COLUMN_ADDRESS + 1).ToString() + " columns).");
            }
            return ResolveColumnAddress(column) + (row + 1).ToString();
        }

        /// <summary>
        /// Gets the column and row number (zero based) of a cell by the address
        /// </summary>
        /// <param name="address">Address as string in the format A1 - XFD1048576</param>
        /// <returns>Struct with row and column</returns>
        /// <exception cref="FormatException">Throws a FormatException if the passed address is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the value of the passed address is out of range (A-XFD and 1 to 1048576) </exception>
        public static Address ResolveCellCoordinate(string address)
        {
            int row, column;
            ResolveCellCoordinate(address, out column, out row);
            return new Address(column, row);
        }

        /// <summary>
        /// Gets the column and row number (zero based) of a cell by the address
        /// </summary>
        /// <param name="address">Address as string in the format A1 - XFD1048576</param>
        /// <param name="column">Column address of the cell (zero-based) as out parameter</param>
        /// <param name="row">Row address of the cell (zero-based) as out parameter</param>
        /// <exception cref="FormatException">Throws a FormatException if the range address was malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the row or column address was out of range</exception>
        public static void ResolveCellCoordinate(string address, out int column, out int row)
        {
            if (string.IsNullOrEmpty(address))
            {
                throw new FormatException("The cell address is null or empty and could not be resolved");
            }
            address = address.ToUpper();
            Regex rx = new Regex("([A-Z]{1,3})([0-9]{1,7})");
            Match mx = rx.Match(address);
            if (mx.Groups.Count != 3)
            {
                throw new FormatException("The format of the cell address (" + address + ") is malformed");
            }
            int digits = int.Parse(mx.Groups[2].Value);
            column = ResolveColumn(mx.Groups[1].Value);
            row = digits - 1;
            if (row > Worksheet.MAX_ROW_ADDRESS || row < Worksheet.MIN_ROW_ADDRESS)
            {
                throw new RangeException("OutOfRangeException", "The row number (" + row.ToString() + ") is out of range. Range is from " + Worksheet.MIN_ROW_ADDRESS.ToString() + " to " + Worksheet.MAX_ROW_ADDRESS.ToString() + " (" + (Worksheet.MAX_ROW_ADDRESS + 1).ToString() + " rows).");
            }
            if (column > Worksheet.MAX_COLUMN_ADDRESS || column < Worksheet.MIN_COLUMN_ADDRESS)
            {
                throw new RangeException("OutOfRangeException", "The column number (" + column.ToString() + ") is out of range. Range is from " + Worksheet.MIN_COLUMN_ADDRESS.ToString() + " to " + Worksheet.MAX_COLUMN_ADDRESS.ToString() + " (" + (Worksheet.MAX_COLUMN_ADDRESS + 1).ToString() + " columns).");
            }
        }

        /// <summary>
        /// Resolves a cell range from the format like A1:B3 or AAD556:AAD1000
        /// </summary>
        /// <param name="range">Range to process</param>
        /// <returns>Range object</returns>
        /// <exception cref="FormatException">Throws a FormatException if the start or end address was malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception>
        public static Range ResolveCellRange(string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                throw new FormatException("The cell range is null or empty and could not be resolved");
            }
            string[] split = range.Split(':');
            if (split.Length != 2)
            {
                throw new FormatException("The cell range (" + range + ") is malformed and could not be resolved");
            }
            return new Range(ResolveCellCoordinate(split[0]), ResolveCellCoordinate(split[1]));
        }

        /// <summary>
        /// Gets the column number from the column address (A - XFD)
        /// </summary>
        /// <param name="columnAddress">Column address (A - XFD)</param>
        /// <returns>Column number (zero-based)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed address was out of range</exception>
        public static int ResolveColumn(string columnAddress)
        {
            int temp;
            int result = 0;
            int multiplicator = 1;
            for (int i = columnAddress.Length - 1; i >= 0; i--)
            {
                temp = (int)columnAddress[i];
                temp = temp - 64;
                result = result + (temp * multiplicator);
                multiplicator = multiplicator * 26;
            }
            if (result - 1 > Worksheet.MAX_COLUMN_ADDRESS || result - 1 < Worksheet.MIN_COLUMN_ADDRESS)
            {
                throw new RangeException("OutOfRangeException","The column number (" + (result - 1).ToString() + ") is out of range. Range is from " + Worksheet.MIN_COLUMN_ADDRESS.ToString() + " to " + Worksheet.MAX_COLUMN_ADDRESS.ToString() + " (" + (Worksheet.MAX_COLUMN_ADDRESS + 1).ToString() + " columns).");
            }
            return result - 1;
        }

        /// <summary>
        /// Gets the column address (A - XFD)
        /// </summary>
        /// <param name="columnNumber">Column number (zero-based)</param>
        /// <returns>Column address (A - XFD)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed column number was out of range</exception>
        public static string ResolveColumnAddress(int columnNumber)
        {
            if (columnNumber > Worksheet.MAX_COLUMN_ADDRESS || columnNumber < Worksheet.MIN_COLUMN_ADDRESS)
            {
                throw new RangeException("OutOfRangeException", "The column number (" + columnNumber.ToString() + ") is out of range. Range is from " + Worksheet.MIN_COLUMN_ADDRESS.ToString() + " to " + Worksheet.MAX_COLUMN_ADDRESS.ToString() + " (" + (Worksheet.MAX_COLUMN_ADDRESS + 1).ToString() + " columns).");
            }
            // A - XFD
            int j = 0;
            int k = 0;
            int l = 0;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i <= columnNumber; i++)
            {
                if (j > 25)
                {
                    k++;
                    j = 0;
                }
                if (k > 25)
                {
                    l++;
                    k = 0;
                }
                j++;
            }
            if (l > 0) { sb.Append((char)(l + 64)); }
            if (k > 0) { sb.Append((char)(k + 64)); }
            sb.Append((char)(j + 64));
            return sb.ToString();
        }
#endregion

#region subClasses

        /// <summary>
        /// Struct representing the cell address as column and row (zero based)
        /// </summary>
        public struct Address
        {
            /// <summary>
            /// Column number (zero based)
            /// </summary>
            public int Column;
            /// <summary>
            /// Row number (zero based)
            /// </summary>
            public int Row;
            
            /// <summary>
            /// Constructor with arguments
            /// </summary>
            /// <param name="column">Column number (zero based)</param>
            /// <param name="row">Row number (zero based)</param>
            public Address(int column, int row)
            {
                Column = column;
                Row = row;
            }

            /// <summary>
            /// Returns the combined Address
            /// </summary>
            /// <returns>Address as string in the format A1 - XFD1048576</returns>
            public string GetAddress()
            {
                return ResolveCellAddress(Column, Row);
            }

            /// <summary>
            /// Overwritten ToString method
            /// </summary>
            /// <returns>Returns the cell address (e.g. 'A15')</returns>
            public override string ToString()
            {
                return GetAddress();
            }
            
        }

        /// <summary>
        /// Struct representing a cell range with a start and end address
        /// </summary>
        public struct Range
        {
            /// <summary>
            /// End address of the range
            /// </summary>
            public Address EndAddress;
            /// <summary>
            /// Start address of the range
            /// </summary>
            public Address StartAddress;

            /// <summary>
            /// Constructor with arguments
            /// </summary>
            /// <param name="start">Start address of the range</param>
            /// <param name="end">End address of the range</param>
            public Range(Address start, Address end)
            {
                StartAddress = start;
                EndAddress = end;
            }
            /// <summary>
            /// Overwritten ToString method
            /// </summary>
            /// <returns>Returns the range (e.g. 'A1:B12')</returns>
            public override string ToString()
            {
                return StartAddress.ToString() + ":" + EndAddress.ToString();
            }

        }

#endregion

    }
}

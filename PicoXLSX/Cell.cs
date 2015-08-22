/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
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
    /// Class representing a Cell of a worksheet
    /// </summary>
    public class Cell : IComparable<Cell>
    {
        /// <summary>
        /// Enum defines the basic data types of a cell
        /// </summary>
        public enum CellType
        {
            /// <summary>Type for single characters and strings</summary>
            STRING,
            /// <summary>Type for all numeric types (integers and floats, respectively doubles)</summary>
            NUMBER,
            /// <summary>Type for dates and times</summary>
            DATE,
            /// <summary>Type for boolean</summary>
            BOOL,
            /// <summary>Type for Formulas (The cell will be handled differently)</summary>
            FORMULA,
            /// <summary>Default Type, not specified</summary>
            DEFAULT
        }

        /// <summary>Number of the row (zerobased)</summary>
        public int RowAddress { get; set; }
        /// <summary>Number of the column (zero-based)</summary>
        public int ColumnAddress { get; set; }
        /// <summary>Value of the cell (generic object type)</summary>
        public object Value { get; set; }
        /// <summary>Type of the cell</summary>
        public CellType Fieldtype { get; set; }

        /// <summary>Combined cell address as struct (read-only)</summary>
        public Address CellAddress
        {
            get { return new Address(this.ColumnAddress, this.RowAddress); }
        }

        /// <summary>Default constructor</summary>
        public Cell()
        {

        }

        /// <summary>
        /// Constructor with value and cell type
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        public Cell(object value, CellType type)
        {
            this.Value = value;
            this.Fieldtype = type;
        }

        /// <summary>
        /// Constructor with value, cell type, row address and column address
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        /// <param name="column">Column address of the cell (zerobased)</param>
        /// <param name="row">Row address of the cell (zerobased)</param>
        public Cell(object value, CellType type, int column, int row)
        {
            this.Value = value;
            this.Fieldtype = type;
            this.ColumnAddress = column;
            this.RowAddress = row;
        }

        /// <summary>
        /// Gets the cell Address as string in the format A1 - XFD16384
        /// </summary>
        /// <returns>Cell address</returns>
        public string GetCellAddress()
        {
            return Cell.ResolveCellAddress(this.ColumnAddress, this.RowAddress);
        }

        /// <summary>
        /// Implemented CompareTo method
        /// </summary>
        /// <param name="other">object to compare</param>
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
        /// Convertrs a List of supported objects into a list of celss
        /// </summary>
        /// <typeparam name="T">Generic data type</typeparam>
        /// <param name="list">List of generic objects</param>
        /// <exception cref="UnsupportedDataTypeException">Throws a UnsupportedDataTypeException if an usupported value was passed</exception>
        /// <returns>List of cells</returns>
        public static List<Cell> ConvertArray<T>(List<T> list)
        {
            List<Cell> output = new List<Cell>();
            Cell c;
            object o;
            foreach(T item in list)
            {
                o = (object)item;
                Type t = typeof(T);

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
                else
                {
                    throw new UnsupportedDataTypeException("The data type '" + typeof(T).ToString() + "' is not supported");
                }
                output.Add(c);
            }
            return output;
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range (format A1:B3 or AAD556:AAD1000)
        /// </summary>
        /// <param name="range">Range to process</param>
        /// <returns>List of cell addresses</returns>
        public static List<Address> GetCellRange(string range)
        {
            Address start, end;
            ResolveCellRange(range, out start, out end);
            return GetCellRange(start, end);
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range
        /// </summary>
        /// <param name="startAddress">Start address as string in the format A1 - XFD16384</param>
        /// <param name="endAddress">End address as string in the format A1 - XFD16384</param>
        /// <returns>List of cell addresses</returns>
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
        /// Resolves a cell range from the format like  A1:B3 or AAD556:AAD1000
        /// </summary>
        /// <param name="range">Range to process</param>
        /// <param name="startAddress">Start address as out parameter</param>
        /// <param name="endAddress">End address as out parameter</param>
        /// <exception cref="FormatException">Throws a FormatException if the start or end address was malformed</exception>
        public static void ResolveCellRange(string range, out Address startAddress, out Address endAddress)
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
            startAddress = ResolveCellCoordinate(split[0]);
            endAddress = ResolveCellCoordinate(split[1]);
        }

        /// <summary>
        /// Gets the address of a cell by the column and row number (zero based)
        /// </summary>
        /// <param name="column">Column address of the cell (zerobased)</param>
        /// <param name="row">Row address of the cell (zerobased)</param>
        /// <exception cref="OutOfRangeException">Throws a OutOfRangeException if the start or end address was out of range</exception>
        /// <returns>Cell Address as string in the format A1 - XFD16384</returns>
        public static string ResolveCellAddress(int column, int row)
        {
            if (row >= 1048576 || row < 0)
            {
                throw new OutOfRangeException("The row number (" + row.ToString() + ") is out of range. Range is from 0 to 1048575 (1048576 rows).");
            }
            if (column >= 16384 || column < 0)
            {
                throw new OutOfRangeException("The column number (" + column.ToString() + ") is out of range. Range is from 0 to 16383 (16384 columns).");
            }
            // A - XFD
            int j = 0;
            int k = 0;
            int l = 0;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i <= column; i++)
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
            sb.Append((row + 1).ToString());
            return (sb.ToString());
        }

        /// <summary>
        /// Gets the column and row number (zero based) of a cell by the address
        /// </summary>
        /// <param name="address">Address as string in the format A1 - XFD16384</param>
        /// <param name="column">Column address of the cell (zerobased) as out parameter</param>
        /// <param name="row">Row address of the cell (zerobased) as out parameter</param>
        /// <exception cref="FormatException">Throws a FormatException if the range address was malformed</exception>
        /// <exception cref="OutOfRangeException">Throws a OutOfRangeException if the start or end address was out of range</exception>
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
            string chars = mx.Groups[1].Value;
            int digits = int.Parse(mx.Groups[2].Value);
            int temp;
            int result = 0;
            int multiplicator = 1;
            for (int i = chars.Length - 1; i >= 0; i--)
            {
                temp = (int)chars[i];
                temp = temp - 64;
                result = result + (temp * multiplicator);
                multiplicator = multiplicator * 26;
            }
            column = result - 1;
            row = digits - 1;
            if (row >= 1048576 || row < 0)
            {
                throw new OutOfRangeException("The row number (" + row.ToString() + ") is out of range. Range is from 0 to 1048575 (1048576 rows).");
            }
            if (column >= 16384 || column < 0)
            {
                throw new OutOfRangeException("The column number (" + column.ToString() + ") is out of range. Range is from 0 to 16383 (16384 columns).");
            }
        }

        /// <summary>
        /// Gets the column and row number (zero based) of a cell by the address
        /// </summary>
        /// <param name="address">Address as string in the format A1 - XFD16384</param>
        /// <returns>Strucht with row and column</returns>
        public static Address ResolveCellCoordinate(string address)
        {
            int row, column;
            ResolveCellCoordinate(address, out column, out row);
            return new Address(column, row);
        }

        /// <summary>
        /// Struct represening the cell address as column and row (zero based)
        /// </summary>
        public struct Address
        {
            /// <summary>
            /// Row number (zero based)
            /// </summary>
            public int Row;
            /// <summary>
            /// Column number (zero based)
            /// </summary>
            public int Column;

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
            /// <returns>Address as string in the format A1 - XFD16384</returns>
            public string GetAddress()
            {
                return ResolveCellAddress(Column, Row);
            }

            public override string ToString()
            {
                return GetAddress();
            }
            
        }

    }
}

/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using PicoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using PicoXLSX.Styles;
using FormatException = PicoXLSX.Exceptions.FormatException;

namespace PicoXLSX
{
        /// <summary>
        /// Struct representing the cell address as column and row (zero based)
        /// </summary>
        public struct Address
        {
        /// <summary>
        /// Enum for the referencing style of the address
        /// </summary>
        public enum AddressType
        {
            /// <summary>Default behavior (e.g. 'C3')</summary>
            Default,
            /// <summary>Row of the address is fixed (e.g. 'C$3')</summary>
            FixedRow,
            /// <summary>Column of the address is fixed (e.g. '$C3')</summary>
            FixedColumn,
            /// <summary>Row and column of the address is fixed (e.g. '$C$3')</summary>
            FixedRowAndColumn
        }

        /// <summary>
        /// Column number (zero based)
        /// </summary>
        public int Column;
            /// <summary>
            /// Row number (zero based)
            /// </summary>
            public int Row;

            /// <summary>
            /// Referencing type of the address
            /// </summary>
            public AddressType Type;

            /// <summary>
            /// Constructor with row and column as arguments
            /// </summary>
            /// <param name="column">Column number (zero based)</param>
            /// <param name="row">Row number (zero based)</param>
            /// <param name="type">Optional referencing type of the address</param>
            public Address(int column, int row, AddressType type = AddressType.Default)
            {
                Column = column;
                Row = row;
                Type = type;
            }

            /// <summary>
            /// Constructor with address as string
            /// </summary>
            /// <param name="address">Address string (e.g. 'A1:B12')</param>
            /// <param name="type">Optional referencing type of the address</param>
            public Address(string address, AddressType type = AddressType.Default)
            {
                Type = type;
                Cell.ResolveCellCoordinate(address, out Column, out Row);
            }

            /// <summary>
            /// Returns the combined Address
            /// </summary>
            /// <returns>Address as string in the format A1 - XFD1048576</returns>
            public string GetAddress()
            {
                return Cell.ResolveCellAddress(Column, Row, Type);
            }

            /// <summary>
            /// Gets the column address (A - XFD)
            /// </summary>
            /// <returns>Column address as letter(s)</returns>
            public string GetColumn()
            {
                return PicoXLSX.Column.ResolveColumnAddress(Column);
            }

            /// <summary>
            /// Overwritten ToString method
            /// </summary>
            /// <returns>Returns the cell address (e.g. 'A15')</returns>
            public override string ToString()
            {
                return GetAddress();
            }

            /// <summary>
            /// Compares two addresses whether they are equal
            /// </summary>
            /// <param name="o"> Other address</param>
            /// <returns>True if equal</returns>
            public bool Equals(Address o)
            {
                if (Row == o.Row && Column == o.Column) { return true; }
                else { return false; }
            }

        #region statiMethods

        /// <summary>
        /// Gets a list of addresses from a cell range (format A1:B3 or AAD556:AAD1000)
        /// </summary>
        /// <param name="range">Range to process</param>
        /// <returns>List of addresses</returns>
        /// <exception cref="FormatException">Throws a FormatException if a part of the passed range is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception>
        public static List<Address> GetAddressRange(string range)
        {
            Range range2 = Range.ResolveCellRange(range);
            return GetAddressRange(range2.StartAddress, range2.EndAddress);
        }

        /// <summary>
        /// Get a list of addresses from a cell range
        /// </summary>
        /// <param name="startAddress">Start address as string in the format A1 - XFD1048576</param>
        /// <param name="endAddress">End address as string in the format A1 - XFD1048576</param>
        /// <returns>List of addresses</returns>
        /// <exception cref="FormatException">Throws a FormatException if a part of the passed range is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception> 
        public static List<Address> GetAddressRange(string startAddress, string endAddress)
        {
            Address start = Cell.ResolveCellCoordinate(startAddress);
            Address end = Cell.ResolveCellCoordinate(endAddress);
            return GetAddressRange(start, end);
        }

        /// <summary>
        /// Get a list of addresses from a cell range
        /// </summary>
        /// <param name="startColumn">Start column (zero based)</param>
        /// <param name="startRow">Start row (zero based)</param>
        /// <param name="endColumn">End column (zero based)</param>
        /// <param name="endRow">End row (zero based)</param>
        /// <returns>List of addresses</returns>
        /// <exception cref="RangeException">Throws an RangeException if the value of one passed address parts is out of range (A-XFD and 1 to 1048576) </exception>
        public static List<Address> GetAddressRange(int startColumn, int startRow, int endColumn, int endRow)
        {
            Address start = new Address(startColumn, startRow);
            Address end = new Address(endColumn, endRow);
            return GetAddressRange(start, end);
        }

        /// <summary>
        /// Get a list of addresses from a cell range
        /// </summary>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// <returns>List of addresses</returns>
        /// <exception cref="FormatException">Throws a FormatException if a part of the passed addresses is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the value of one passed address is out of range (A-XFD and 1 to 1048576) </exception>
        public static List<Address> GetAddressRange(Address startAddress, Address endAddress)
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
        #endregion
    }

}

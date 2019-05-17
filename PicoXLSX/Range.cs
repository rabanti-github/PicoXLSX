/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using PicoXLSX.Exceptions;
using PicoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using FormatException = PicoXLSX.Exceptions.FormatException;

namespace PicoXLSX
{
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
            /// Constructor with addresses as arguments
            /// </summary>
            /// <param name="start">Start address of the range</param>
            /// <param name="end">End address of the range</param>
            public Range(Address start, Address end)
            {
                StartAddress = start;
                EndAddress = end;
            }

            /// <summary>
            /// Constructor with a range string as argument
            /// </summary>
            /// <param name="range">Address range (e.g. 'A1:B12')</param>
            public Range(string range)
            {
                Range r = ResolveCellRange(range);
                StartAddress = r.StartAddress;
                EndAddress = r.EndAddress;
            }

            /// <summary>
            /// Overwritten ToString method
            /// </summary>
            /// <returns>Returns the range (e.g. 'A1:B12')</returns>
            public override string ToString()
            {
                return StartAddress.ToString() + ":" + EndAddress.ToString();
            }

        #region staticMethods
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
            return new Range(Cell.ResolveCellCoordinate(split[0]), Cell.ResolveCellCoordinate(split[1]));
        }
        #endregion

    }
}

using PicoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PicoXLSX
{
    /// <summary>
    /// Class representing a column of a worksheet
    /// </summary>
    public class Column
    {
        #region constants
        /// <summary>
        /// Default column width as constant
        /// </summary>
        public const float DEFAULT_COLUMN_WIDTH = 10f;
        /// <summary>
        /// Minimum column width as constant
        /// </summary>
        public const float MIN_COLUMN_WIDTH = 0f;
        /// <summary>
        /// Maximum column width as constant
        /// </summary>
        public const float MAX_COLUMN_WIDTH = 255f;
        /// <summary>
        /// Minimum column number (zero-based) as constant
        /// </summary>
        public const int MIN_COLUMN_NUMBER = 0;
        /// <summary>
        /// Maximum column number (zero-based) as constant
        /// </summary>
        public const int MAX_COLUMN_NUMBER = 16383;

        #endregion

        private int number;
        private string columnAddress;

        /// <summary>
        /// Column address (A to XFD)
        /// </summary>
        public string ColumnAddress
        {
            get { return columnAddress; }
            set
            {
                number = Column.ResolveColumn(value);
                columnAddress = value;
            }
        }

        /// <summary>
        /// If true, the column has auto filter applied, otherwise not
        /// </summary>
        public bool HasAutoFilter { get; set; }
        /// <summary>
        /// If true, the column is hidden, otherwise visible
        /// </summary>
        public bool IsHidden { get; set; }

        /// <summary>
        /// Column number (0 to 16383)
        /// </summary>
        public int Number
        {
            get { return number; }
            set
            {
                columnAddress = Column.ResolveColumnAddress(value);
                number = value;
            }
        }

        /// <summary>
        /// Width of the column
        /// </summary>
        public float Width { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public Column()
        {
            Width = DEFAULT_COLUMN_WIDTH;
        }

        /// <summary>
        /// Constructor with column number
        /// </summary>
        /// <param name="columnCoordinate">Column number (zero-based, 0 to 16383)</param>
        public Column(int columnCoordinate) : this()
        {
            Number = columnCoordinate;
        }

        /// <summary>
        /// Constructor with column address
        /// </summary>
        /// <param name="columnAddress">Column address (A to XFD)</param>
        public Column(string columnAddress) : this()
        {
            ColumnAddress = columnAddress;
        }

        #region staticMethods
        /// <summary>
        /// Gets the column number from the column address (A - XFD)
        /// </summary>
        /// <param name="columnAddress">Column address (A - XFD)</param>
        /// <returns>Column number (zero-based)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed address was out of range</exception>
        public static int ResolveColumn(string columnAddress)
        {
            int chr;
            int result = 0;
            int multiplier = 1;
            for (int i = columnAddress.Length - 1; i >= 0; i--)
            {
                chr = columnAddress[i];
                chr = chr - 64;
                result = result + (chr * multiplier);
                multiplier = multiplier * 26;
            }
            if (result - 1 > Column.MAX_COLUMN_NUMBER || result - 1 < Column.MIN_COLUMN_NUMBER)
            {
                throw new RangeException("OutOfRangeException", "The column number (" + (result - 1).ToString() + ") is out of range. Range is from " + Column.MIN_COLUMN_NUMBER.ToString() + " to " + Column.MAX_COLUMN_NUMBER.ToString() + " (" + (Column.MAX_COLUMN_NUMBER + 1).ToString() + " columns).");
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
            if (columnNumber > Column.MAX_COLUMN_NUMBER || columnNumber < Column.MIN_COLUMN_NUMBER)
            {
                throw new RangeException("OutOfRangeException", "The column number (" + columnNumber.ToString() + ") is out of range. Range is from " + Column.MIN_COLUMN_NUMBER.ToString() + " to " + Column.MAX_COLUMN_NUMBER.ToString() + " (" + (Column.MAX_COLUMN_NUMBER + 1).ToString() + " columns).");
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
    }
}

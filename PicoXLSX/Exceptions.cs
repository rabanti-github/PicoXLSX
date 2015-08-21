/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PicoXLSX
{
    /// <summary>
    /// Class for exceptions regarding out-of-range incidents
    /// </summary>
    public class OutOfRangeException : Exception
    {
        public OutOfRangeException()
        { }
        public OutOfRangeException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding format error incidents
    /// </summary>
    public class FormatException : Exception
    {
        public FormatException()
        { }
        public FormatException(string message)
            : base(message)
        { }
        public FormatException(string message, Exception inner)
            : base(message, inner)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding stream or save error incidents
    /// </summary>
    public class IOException : Exception
    {
        public IOException()
        { }
        public IOException(string message)
            : base(message)
        { }
        public IOException(string message, Exception inner)
            : base(message, inner)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding an already existing worksheet (same name)
    /// </summary>
    public class WorksheetNameAlreadxExistsException : Exception
    {
        public WorksheetNameAlreadxExistsException()
        { }
        public WorksheetNameAlreadxExistsException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding an unknown worksheet (name not found)
    /// </summary>
    public class UnknownWorksheetException : Exception
    {
        public UnknownWorksheetException()
        { }
        public UnknownWorksheetException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding an unsuppored data type
    /// </summary>
    public class UnsupportedDataTypeException : Exception
    {
        public UnsupportedDataTypeException()
        { }
        public UnsupportedDataTypeException(string message)
            : base(message)
        { }
    }

}

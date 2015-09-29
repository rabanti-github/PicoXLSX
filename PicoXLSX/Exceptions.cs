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
        /// <summary>
        /// Default constructor
        /// </summary>
        public OutOfRangeException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public OutOfRangeException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding format error incidents
    /// </summary>
    public class FormatException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public FormatException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public FormatException(string message)
            : base(message)
        { }
        /// <summary>
        /// Constructor with passed message and inner exception
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="inner">Inner exception</param>
        public FormatException(string message, Exception inner)
            : base(message, inner)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding stream or save error incidents
    /// </summary>
    public class IOException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public IOException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public IOException(string message)
            : base(message)
        { }
        /// <summary>
        /// Constructor with passed message and inner exception
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="inner">Inner exception</param>
        public IOException(string message, Exception inner)
            : base(message, inner)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding an already existing worksheet (same name)
    /// </summary>
    public class WorksheetNameAlreadxExistsException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public WorksheetNameAlreadxExistsException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public WorksheetNameAlreadxExistsException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding an unknown worksheet (name not found)
    /// </summary>
    public class UnknownWorksheetException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public UnknownWorksheetException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public UnknownWorksheetException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding an unsupported data type
    /// </summary>
    public class UnsupportedDataTypeException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public UnsupportedDataTypeException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public UnsupportedDataTypeException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding undefined Styles
    /// </summary>
    public class UndefinedStyleException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public UndefinedStyleException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public UndefinedStyleException(string message)
            : base(message)
        { }
    }

}

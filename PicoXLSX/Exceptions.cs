/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace PicoXLSX
{
    /// <summary>
    /// Class for exceptions regarding range incidents (e.g. out-of-range)
    /// </summary>
    [Serializable]
    public class RangeException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public RangeException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public RangeException(string title, string message)
            : base( message)
        {}
    }

    /// <summary>
    /// Class for exceptions regarding format error incidents
    /// </summary>
    [Serializable]
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
        /// <param name="title">Title of the exception</param>
        public FormatException(string title, string message, Exception inner)
            : base(message, inner)
        {}
    }

    /// <summary>
    /// Class for exceptions regarding stream or save error incidents
    /// </summary>
    [Serializable]
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
    /// Class for exceptions regarding worksheet incidents
    /// </summary>
    [Serializable]
    public class WorksheetException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public WorksheetException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public WorksheetException(string message)
            : base(message)
        { }
    }

    /// <summary>
    /// Class for exceptions regarding Style incidents
    /// </summary>
    [Serializable]
    public class StyleException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public StyleException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public StyleException(string title, string message)
            : base(message)
        {  }
    }

}

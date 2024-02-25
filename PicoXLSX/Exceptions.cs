/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace PicoXLSX
{
    using System;

    /// <summary>
    /// Class for exceptions regarding range incidents (e.g. out-of-range)
    /// </summary>
    [Serializable]
    public class RangeException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RangeException"/> class
        /// </summary>
        public RangeException() : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeException"/> class
        /// </summary>
        /// <param name="title">The title<see cref="string"/>.</param>
        /// <param name="message">Message of the exception.</param>
        public RangeException(string title, string message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// Class for exceptions regarding format error incidents
    /// </summary>
    [Serializable]
    public class FormatException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FormatException"/> class
        /// </summary>
        public FormatException() : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormatException"/> class
        /// </summary>
        /// <param name="message">Message of the exception.</param>
        public FormatException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormatException"/> class
        /// </summary>
        /// <param name="title">Title of the exception.</param>
        /// <param name="message">Message of the exception.</param>
        /// <param name="inner">Inner exception.</param>
        public FormatException(string title, string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    /// <summary>
    /// Class for exceptions regarding stream or save error incidents
    /// </summary>
    [Serializable]
    public class IOException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="IOException"/> class
        /// </summary>
        public IOException() : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="IOException"/> class
        /// </summary>
        /// <param name="message">Message of the exception.</param>
        public IOException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="IOException"/> class
        /// </summary>
        /// <param name="message">Message of the exception.</param>
        /// <param name="inner">Inner exception.</param>
        public IOException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    /// <summary>
    /// Class for exceptions regarding worksheet incidents
    /// </summary>
    [Serializable]
    public class WorksheetException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetException"/> class
        /// </summary>
        public WorksheetException() : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetException"/> class
        /// </summary>
        /// <param name="message">Message of the exception.</param>
        public WorksheetException(string message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// Class for exceptions regarding Style incidents
    /// </summary>
    [Serializable]
    public class StyleException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StyleException"/> class
        /// </summary>
        public StyleException() : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="StyleException"/> class
        /// </summary>
        /// <param name="title">The title<see cref="string"/>.</param>
        /// <param name="message">Message of the exception.</param>
        public StyleException(string title, string message)
            : base(message)
        {
        }
    }
}

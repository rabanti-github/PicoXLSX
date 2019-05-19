/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using PicoXLSX.Core;
using PicoXLSX.Exceptions;
using PicoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using FormatException = PicoXLSX.Exceptions.FormatException;
using IOException = PicoXLSX.Exceptions.IOException;

namespace PicoXLSX
{
        /// <summary>
        /// Class to manage XML document paths
        /// </summary>
        public class DocumentPath
        {
            /// <summary>
            /// File name of the document
            /// </summary>
            public string Filename { get; set; }
            /// <summary>
            /// Path of the document
            /// </summary>
            public string Path { get; set; }

            /// <summary>
            /// Default constructor
            /// </summary>
            public DocumentPath()
            {
            }
            /// <summary>
            /// Constructor with defined file name and path
            /// </summary>
            /// <param name="filename">File name of the document</param>
            /// <param name="path">Path of the document</param>
            public DocumentPath(string filename, string path)
            {
                Filename = filename;
                Path = path;
            }

            /// <summary>
            /// Method to return the full path of the document
            /// </summary>
            /// <returns>Full path</returns>
            public string GetFullPath()
            {
                if (Path == null) { return Filename; }
                if (Path == "") { return Filename; }
                if (Path[Path.Length - 1] == System.IO.Path.AltDirectorySeparatorChar || Path[Path.Length - 1] == System.IO.Path.DirectorySeparatorChar)
                {
                    return System.IO.Path.AltDirectorySeparatorChar.ToString() + Path + Filename;
                }
                else
                {
                    return System.IO.Path.AltDirectorySeparatorChar.ToString() + Path + System.IO.Path.AltDirectorySeparatorChar.ToString() + Filename;
                }
            }

        }
}

/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PicoXLSX
{
    /// <summary>
    /// Class representing the meta data of a workbook
    /// </summary>
    public class Metadata
    {

        private string applicationVersion;


        /// <summary>
        /// If true, custom defined colors (in styles) will be added as recent colors (MRU)
        /// </summary>
        public bool UseColorMRU { get; set; }
        /// <summary>
        /// Title of the workbook
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// Subject of the workbook
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// Creator of the workbook. Add more than one creator by using the semicolon (;) between the authors
        /// </summary>
        public string Creator { get; set; }
        /// <summary>
        /// Keyword for the workbook. Separate the keywords with semicolons (;)
        /// </summary>
        public string Keywords { get; set; }
        /// <summary>
        /// Application which created the workbook. Default is PicoXLSX
        /// </summary>
        public string Application { get; set; }
        /// <summary>
        /// Version of the creation application. Default is the library version of PicoXLSX. Use the format xxxxx.yyyyy (e.g. 1.0 or 55.9875) in case of a custom value.
        /// </summary>
        public string ApplicationVersion
        {
            get { return applicationVersion; }
            set
            { 
                applicationVersion = value;
                CheckVersion();
            }
        }
        /// <summary>
        /// Description of the document or comment about it
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Category of the document. There are no predefined values or restrictions about the content of this field
        /// </summary>
        public string Category { get; set; }
        /// <summary>
        /// Status of the document. There are no predefined values or restrictions about the content of this field
        /// </summary>
        public string ContentStatus { get; set; }
        /// <summary>
        /// Responsible manager of the document. This value is for organizational purpose.
        /// </summary>
        public string Manager { get; set; }
        /// <summary>
        /// Company owning of the document. This value is for organizational purpose. Add more than one manager by using the semicolon (;) between the words
        /// </summary>
        public string Company { get; set; }
        /// <summary>
        /// Hyper-link base of the document.
        /// </summary>
        public string HyperlinkBase { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public Metadata()
        {
            this.UseColorMRU = false;
            this.Application = "PicoXLSX";
            //this.ApplicationVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Version vi = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.ApplicationVersion = ParseVersion(vi.Major, vi.Minor, vi.Revision, vi.Build);
        }

        /// <summary>
        /// Method to parse a common version (major.minor.revision.build) into the compatible format (major.minor). The minimum value is 0.0 and the maximum value is 99999.99999<br></br>
        /// The minor, revision and build number are joined if possible. If the number is to long, the additional characters will be removed from the right side down to five characters (e.g. 785563 will be 78556)
        /// </summary>
        /// <param name="major">Major number from 0 to 99999</param>
        /// <param name="minor">Minor number</param>
        /// <param name="build">Build number</param>
        /// <param name="revision">Revision number</param>
        /// <returns>Formated version number (e.g. 1.0 or 55.987)</returns>
        /// <exception cref="FormatException">Throws a FormatException if the major number is to long or one of the numbers is negative</exception>
        public static string ParseVersion(int major, int minor, int build, int revision)
        {
            if (major < 0 || minor < 0 || build < 0 || revision < 0)
            {
                throw new FormatException("The format of the passed version is wrong. No negative number allowed.");
            }
            if (major > 99999)
            {
                throw new FormatException("The major number may not be bigger than 99999. The passed value is " + major.ToString());
            }
            CultureInfo culture = CultureInfo.CreateSpecificCulture("en-US");
            string leftPart = major.ToString("G", culture);
            string rightPart = minor.ToString("G", culture) + build.ToString("G", culture) + revision.ToString("G", culture);
            rightPart = rightPart.TrimEnd('0');
            if (rightPart == "") { rightPart = "0"; }
            else
            {
                if (rightPart.Length > 5)
                {
                    rightPart = rightPart.Substring(0, 5);
                }
            }
            return leftPart + "." + rightPart;
        }

        /// <summary>
        /// Checks the format of the passed version string
        /// </summary>
        /// /// <exception cref="FormatException">Throws a FormatException if the version string is malformed</exception>
        private void CheckVersion()
        {
            if (string.IsNullOrEmpty(this.applicationVersion)) { return; }
            string[] split = this.applicationVersion.Split('.');
            bool state = true;
            if (split.Length != 2) { state = false; }
            else
            {
                if (split[1].Length < 1 || split[1].Length > 5) { state = false; }
                if (split[0].Length < 1 || split[0].Length > 5) { state = false; }
            }
            if (state == false)
            {
                throw new FormatException("The format of the version in the meta data is wrong (" + this.applicationVersion + "). Should be in the format and a range from '0.0' to '99999.99999'");
            }
        }

    }
}

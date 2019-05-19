using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PicoXLSX.Exceptions;
using FormatException = PicoXLSX.Exceptions.FormatException;

namespace PicoXLSX.Core
{
    public static class Utils
    {
        /// <summary>
        /// Method to escape XML characters between two XML tags
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        /// <remarks>Note: The XML specs allow characters up to the character value of 0x10FFFF. However, the C# char range is only up to 0xFFFF. PicoXLSX will neglect all values above this level in the sanitizing check. Illegal characters like 0x1 will be replaced with a white space (0x20)</remarks>
        public static string EscapeXmlChars(string input)
        {
            if (input == null) { return ""; }
            int len = input.Length;
            List<int> illegalCharacters = new List<int>(len);
            List<byte> characterTypes = new List<byte>(len);
            int i;
            for (i = 0; i < len; i++)
            {
                if ((input[i] < 0x9) || (input[i] > 0xA && input[i] < 0xD) || (input[i] > 0xD && input[i] < 0x20) || (input[i] > 0xD7FF && input[i] < 0xE000) || (input[i] > 0xFFFD))
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(0);
                    continue;
                } // Note: XML specs allow characters up to 0x10FFFF. However, the C# char range is only up to 0xFFFF; Higher values are neglected here 
                if (input[i] == 0x3C) // <
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(1);
                }
                else if (input[i] == 0x3E) // >
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(2);
                }
                else if (input[i] == 0x26) // &
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(3);
                }
            }
            if (illegalCharacters.Count == 0)
            {
                return input;
            }

            StringBuilder sb = new StringBuilder(len);
            int lastIndex = 0;
            len = illegalCharacters.Count;
            for (i = 0; i < len; i++)
            {
                sb.Append(input.Substring(lastIndex, illegalCharacters[i] - lastIndex));
                if (characterTypes[i] == 0)
                {
                    sb.Append(' '); // Whitespace as fall back on illegal character
                }
                else if (characterTypes[i] == 1) // replace <
                {
                    sb.Append("&lt;");
                }
                else if (characterTypes[i] == 2) // replace >
                {
                    sb.Append("&gt;");
                }
                else if (characterTypes[i] == 3) // replace &
                {
                    sb.Append("&amp;");
                }
                lastIndex = illegalCharacters[i] + 1;
            }
            sb.Append(input.Substring(lastIndex));
            return sb.ToString();
        }

        /// <summary>
        /// Method to escape XML characters in an XML attribute
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        public static string EscapeXmlAttributeChars(string input)
        {
            input = EscapeXmlChars(input); // Sanitize string from illegal characters beside quotes
            input = input.Replace("\"", "&quot;");
            return input;
        }

        /// <summary>
        /// Method to generate an Excel internal password hash to protect workbooks or worksheets<br></br>This method is derived from the c++ implementation by Kohei Yoshida (<a href="http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)
        /// </summary>
        /// <remarks>WARNING! Do not use this method to encrypt 'real' passwords or data outside from PicoXLSX. This is only a minor security feature. Use a proper cryptography method instead.</remarks>
        /// <param name="password">Password string in UTF-8 to encrypt</param>
        /// <returns>16 bit hash as hex string</returns>
        public static string GeneratePasswordHash(string password)
        {
            if (string.IsNullOrEmpty(password)) { return string.Empty; }
            int passwordLength = password.Length;
            int passwordHash = 0;
            char character;
            for (int i = passwordLength; i > 0; i--)
            {
                character = password[i - 1];
                passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
                passwordHash ^= character;
            }
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= (0x8000 | ('N' << 8) | 'K');
            passwordHash ^= passwordLength;
            return passwordHash.ToString("X");
        }

        /// <summary>
        /// Method to convert a date or date and time into the internal Excel time format (OAdate)
        /// </summary>
        /// <param name="date">Date to process</param>
        /// <param name="culture">CultureInfo for proper formatting of the decimal point</param>
        /// <returns>Date or date and time as Number</returns>
        /// <exception cref="FormatException">Throws a FormatException if the passed date cannot be translated to the OADate format</exception>
        /// <remarks>OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException</remarks>
        public static string GetOADateTimeString(DateTime date, CultureInfo culture)
        {
            try
            {
                double d = date.ToOADate();
                if (d < 0)
                {
                    throw new FormatException("The date is not in a valid range for Excel. Dates before 1900-01-01 are not allowed.");
                }
                return d.ToString("G", culture); //worksheet.DefaultRowHeight.ToString("G", culture) 
            }
            catch (Exception e)
            {
                throw new FormatException("ConversionException", "The date could not be transformed into Excel format (OADate).", e);
            }
        }

    }
}

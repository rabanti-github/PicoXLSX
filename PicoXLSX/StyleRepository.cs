/*
 * PicoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace PicoXLSX
{
    using System.Collections.Generic;

    /// <summary>
    /// Class to manage all styles at runtime, before writing XLSX files. The main purpose is deduplication and decoupling of styles from workbooks at runtime
    /// </summary>
    public class StyleRepository
    {
        /// <summary>
        /// Defines the lockObject
        /// </summary>
        private readonly object lockObject = new object();

        /// <summary>
        /// Defines the instance
        /// </summary>
        private static StyleRepository instance;

        /// <summary>
        /// Gets the singleton instance of the repository
        /// </summary>
        public static StyleRepository Instance
        {
            get
            {
                instance = instance ?? new StyleRepository();
                return instance;
            }
        }

        /// <summary>
        /// Defines the styles
        /// </summary>
        private Dictionary<int, Style> styles;

        /// <summary>
        /// Gets the currently managed styles of the repository
        /// </summary>
        public Dictionary<int, Style> Styles { get => styles; }

        /// <summary>
        /// Prevents a default instance of the <see cref="StyleRepository"/> class from being created
        /// </summary>
        private StyleRepository()
        {
            styles = new Dictionary<int, Style>();
        }

        /// <summary>
        /// Adds a style to the repository and returns the actual reference
        /// </summary>
        /// <param name="style">Style to add.</param>
        /// <returns>Reference from the repository. If the style to add already existed, the existing object is returned, otherwise the newly added one.</returns>
        public Style AddStyle(Style style)
        {
            lock (lockObject)
            {
                if (style == null)
                {
                    return null;
                }
                int hashCode = style.GetHashCode();
                if (!styles.ContainsKey(hashCode))
                {
                    styles.Add(hashCode, style);
                }
                return styles[hashCode];
            }
        }

        /// <summary>
        /// Empties the static repository
        /// </summary>
        public void FlushStyles()
        {
            styles.Clear();
        }
    }
}

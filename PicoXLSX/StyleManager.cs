/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
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
    /// Class representing a style manager to maintain all styles and its components of a workbook
    /// </summary>
    public class StyleManager
    {
        #region constants
        public const string BORDERPREFIX = "borders@";
        public const string CELLXFPREFIX = "/cellXf@";
        public const string FILLPREFIX = "/fill@";
        public const string FONTPREFIX = "/font@";
        public const string NUMBERFORMATPREFIX = "/numberFormat@";
        public const string STYLEPREFIX = "style=";
        #endregion

        #region privateFields
        private List<AbstractStyle> borders;
        private List<AbstractStyle> cellXfs;
        private List<AbstractStyle> fills;
        private List<AbstractStyle> fonts;
        private List<AbstractStyle> numberFormats;
        private List<AbstractStyle> styles;
        private List<string> styleNames;
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public StyleManager()
        {
            this.borders = new List<AbstractStyle>();
            this.cellXfs = new List<AbstractStyle>();
            this.fills = new List<AbstractStyle>();
            this.fonts = new List<AbstractStyle>();
            this.numberFormats = new List<AbstractStyle>();
            this.styles = new List<AbstractStyle>();
            this.styleNames = new List<string>();
        }
        #endregion

        #region methods

        /// <summary>
        /// Gets a component by its hash
        /// </summary>
        /// <param name="list">List to check</param>
        /// <param name="hash">Hash of the component</param>
        /// <returns>Determined component. If not found, null will be returned</returns>
        private AbstractStyle GetComponentByHash(ref List<AbstractStyle> list, string hash)
        {
            int len = list.Count;
            for (int i = 0; i < len; i++)
            {
                if (list[i].Hash == hash)
                {
                    return list[i];
                }
            }
            return null;
        }

        /// <summary>
        /// Gets a border by its hash
        /// </summary>
        /// <param name="hash">Hash of the border</param>
        /// <returns>Determined border</returns>
        /// <exception cref="StyleException">Throws a StyleException if the border was not found in the style manager</exception>
        public Style.Border GetBorderByHash(String hash)
        {
            AbstractStyle component = GetComponentByHash(ref this.borders, hash);
            if (component == null)
            {
                throw new StyleException("MissingReferenceException", "The style component with the hash '" + hash + "' was not found");
            }
            return (Style.Border)component;
        }

        /// <summary>
        /// Gets all borders of the style manager
        /// </summary>
        /// <returns>Array of borders</returns>
        public Style.Border[] GetBorders()
        {
            return Array.ConvertAll(this.borders.ToArray(), x => (Style.Border)x);
        }

        /// <summary>
        /// Gets the number of borders in the style manager
        /// </summary>
        /// <returns>Number of stored borders</returns>
        public int GetBorderStyleNumber()
        {
            return this.borders.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a cellXf by its hash
        /// </summary>
        /// <param name="hash">Hash of the cellXf</param>
        /// <returns>Determined cellXf</returns>
        /// <exception cref="StyleException">Throws a StyleException if the cellXf was not found in the style manager</exception>
        public Style.CellXf GetCellXfByHash(String hash)
        {
            AbstractStyle component = GetComponentByHash(ref this.cellXfs, hash);
            if (component == null)
            {
                throw new StyleException("MissingReferenceException", "The style component with the hash '" + hash + "' was not found");
            }
            return (Style.CellXf)component;
        }

        /// <summary>
        /// Gets all cellXfs of the style manager
        /// </summary>
        /// <returns>Array of cellXfs</returns>
        public Style.CellXf[] GetCellXfs()
        {
            return Array.ConvertAll(this.cellXfs.ToArray(), x => (Style.CellXf)x);
        }

        /// <summary>
        /// Gets the number of cellXfs in the style manager
        /// </summary>
        /// <returns>Number of stored cellXfs</returns>
        public int GetCellXfStyleNumber()
        {
            return this.cellXfs.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a fill by its hash
        /// </summary>
        /// <param name="hash">Hash of the fill</param>
        /// <returns>Determined fill</returns>
        /// <exception cref="StyleException">Throws a StyleException if the fill was not found in the style manager</exception>
        public Style.Fill GetFillByHash(String hash)
        {
            AbstractStyle component = GetComponentByHash(ref this.fills, hash);
            if (component == null)
            {
                throw new StyleException("MissingReferenceException", "The style component with the hash '" + hash + "' was not found");
            }
            return (Style.Fill)component;
        }

        /// <summary>
        /// Gets all fills of the style manager
        /// </summary>
        /// <returns>Array of fills</returns>
        public Style.Fill[] GetFills()
        {
            return Array.ConvertAll(this.fills.ToArray(), x => (Style.Fill)x);
        }

        /// <summary>
        /// Gets the number of fills in the style manager
        /// </summary>
        /// <returns>Number of stored fills</returns>
        public int GetFillStyleNumber()
        {
            return this.fills.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a font by its hash
        /// </summary>
        /// <param name="hash">Hash of the font</param>
        /// <returns>Determined font</returns>
        /// <exception cref="StyleException">Throws a StyleException if the font was not found in the style manager</exception>
        public Style.Font GetFontByHash(String hash)
        {
            AbstractStyle component = GetComponentByHash(ref this.fonts, hash);
            if (component == null)
            {
                throw new StyleException("MissingReferenceException", "The style component with the hash '" + hash + "' was not found");
            }
            return (Style.Font)component;
        }

        /// <summary>
        /// Gets all fonts of the style manager
        /// </summary>
        /// <returns>Array of fonts</returns>
        public Style.Font[] GetFonts()
        {
            return Array.ConvertAll(this.fonts.ToArray(), x => (Style.Font)x);
        }

        /// <summary>
        /// Gets the number of fonts in the style manager
        /// </summary>
        /// <returns>Number of stored fonts</returns>
        public int GetFontStyleNumber()
        {
            return this.fonts.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a numberFormat by its hash
        /// </summary>
        /// <param name="hash">Hash of the numberFormat</param>
        /// <returns>Determined numberFormat</returns>
        /// <exception cref="StyleException">Throws a StyleException if the numberFormat was not found in the style manager</exception>
        public Style.NumberFormat GetNumberFormatByHash(String hash)
        {
            AbstractStyle component = GetComponentByHash(ref this.numberFormats, hash);
            if (component == null)
            {
                throw new StyleException("MissingReferenceException", "The style component with the hash '" + hash + "' was not found");
            }
            return (Style.NumberFormat)component;
        }

        /// <summary>
        /// Gets all numberFormats of the style manager
        /// </summary>
        /// <returns>Array of numberFormats</returns>
        public Style.NumberFormat[] GetNumberFormats()
        {
            return Array.ConvertAll(this.numberFormats.ToArray(), x => (Style.NumberFormat)x);
        }

        /// <summary>
        /// Gets the number of numberFormats in the style manager
        /// </summary>
        /// <returns>Number of stored numberFormats</returns>
        public int GetNumberFormatStyleNumber()
        {
            return this.numberFormats.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a style by its name
        /// </summary>
        /// <param name="name">Name of the style</param>
        /// <returns>Determined style</returns>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found in the style manager</exception>
        public Style GetStyleByName(string name)
        {
            int len = this.styles.Count;
            for (int i = 0; i < len; i++)
            {
                if (((Style)this.styles[i]).Name == name)
                {
                    return (Style)this.styles[i];
                }
            }
            throw new StyleException("MissingReferenceException", "The style with the name '" + name + "' was not found");
        }

        /// <summary>
        /// Gets a style by its hash
        /// </summary>
        /// <param name="hash">Hash of the style</param>
        /// <returns>Determined style</returns>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found in the style manager</exception>
        public Style GetStyleByHash(String hash)
        {
            AbstractStyle component = GetComponentByHash(ref this.styles, hash);
            if (component == null)
            {
                throw new StyleException("MissingReferenceException", "The style component with the hash '" + hash + "' was not found");
            }
            return (Style)component;
        }

        /// <summary>
        /// Gets all styles of the style manager
        /// </summary>
        /// <returns>Array of styles</returns>
        public Style[] GetStyles()
        {
            return Array.ConvertAll(this.styles.ToArray(), x => (Style)x);
        }

        /// <summary>
        /// Gets the number of styles in the style manager
        /// </summary>
        /// <returns>Number of stored styles</returns>
        public int GetStyleNumber()
        {
            return this.styles.Count;
        }

        /* ****************************** */


        /// <summary>
        /// Adds a style component to the manager
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Added or determined style in the manager</returns>
        public Style AddStyle(Style style)
        {
            string hash = AddStyleComponent(style);
            return (Style)this.GetComponentByHash(ref this.styles, hash);
        }

        /// <summary>
        /// Adds a style component to the manager with an ID
        /// </summary>
        /// <param name="style">Component to add</param>
        /// <param name="id">Id of the component</param>
        /// <returns>Hash of the added or determined component</returns>
        private string AddStyleComponent(AbstractStyle style, int? id)
        {
            style.InternalID = id;
            return AddStyleComponent(style);
        }

        /// <summary>
        /// Adds a style component to the manager
        /// </summary>
        /// <param name="style">Component to add</param>
        /// <returns>Hash of the added or determined component</returns>
        private string AddStyleComponent(AbstractStyle style)
        {
            string hash = style.Hash;
            if (style.GetType() == typeof(Style.Border))
            {
                if (this.GetComponentByHash(ref this.borders, hash) == null) { this.borders.Add(style); }
                Reorganize(ref borders);
            }
            else if (style.GetType() == typeof(Style.CellXf))
            {
                if (this.GetComponentByHash(ref this.cellXfs, hash) == null) { this.cellXfs.Add(style); }
                Reorganize(ref cellXfs);
            }
            else if (style.GetType() == typeof(Style.Fill))
            {
                if (this.GetComponentByHash(ref this.fills, hash) == null) { this.fills.Add(style); }
                Reorganize(ref fills);
            }
            else if (style.GetType() == typeof(Style.Font))
            {
                if (this.GetComponentByHash(ref this.fonts, hash) == null) { this.fonts.Add(style); }
                Reorganize(ref fonts);
            }
            else if (style.GetType() == typeof(Style.NumberFormat))
            {
                if (this.GetComponentByHash(ref this.numberFormats, hash) == null) { this.numberFormats.Add(style); }
                Reorganize(ref numberFormats);
            }
            else if (style.GetType() == typeof(Style))
            {
                Style s = (Style)style;
                if (this.styleNames.Contains(s.Name) == true)
                {
                    throw new StyleException("StyleArleadyExistsException", "The style with the name '" + s.Name + "' already exists");
                }
                if (this.GetComponentByHash(ref this.styles, hash) == null)
                {
                    int? id;
                    if (s.InternalID.HasValue == false)
                    {
                        id = int.MaxValue;
                        s.InternalID = id;
                    }
                    else
                    {
                        id = s.InternalID.Value;
                    }
                    string temp = this.AddStyleComponent(s.BorderStyle, id);
                    s.BorderStyle = (Style.Border)this.GetComponentByHash(ref this.borders, temp);
                    temp = this.AddStyleComponent(s.CellXfStyle, id);
                    s.CellXfStyle = (Style.CellXf)this.GetComponentByHash(ref this.cellXfs, temp);
                    temp = this.AddStyleComponent(s.FillStyle, id);
                    s.FillStyle = (Style.Fill)this.GetComponentByHash(ref this.fills, temp);
                    temp = this.AddStyleComponent(s.FontStyle, id);
                    s.FontStyle = (Style.Font)this.GetComponentByHash(ref this.fonts, temp);
                    temp = this.AddStyleComponent(s.NumberFormatStyle, id);
                    s.NumberFormatStyle = (Style.NumberFormat)this.GetComponentByHash(ref this.numberFormats, temp);
                    this.styles.Add(s);
                }
                Reorganize(ref this.styles);
                hash = s.CalculateHash();
            }
            return hash;
        }

        /// <summary>
        /// Removes a style and all its components from the style manager
        /// </summary>
        /// <param name="styleName">Name of the style to remove</param>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found in the style manager</exception>
        public void RemoveStyle(string styleName)
        {
            string hash = null;
            bool match = false;
            int len = this.styles.Count;
            int index = -1;
            for (int i = 0; i < len; i++)
            {
                if (((Style)this.styles[i]).Name == styleName)
                {
                    match = true;
                    hash = ((Style)this.styles[i]).Hash;
                    index = i;
                    break;
                }
            }
            if (match == false)
            {
                throw new StyleException("MissingReferenceException", "The style with the name '" + styleName + "' was not found in the style manager");
            }
            this.styles.RemoveAt(index);
            CleanupStyleComponents();
        }

        /// <summary>
        /// Method to reorganize / reorder a list of style components
        /// </summary>
        /// <param name="list">List to reorganize as reference</param>
        private void Reorganize(ref List<AbstractStyle> list)
        {
            int len = list.Count;
            list.Sort();
            int id = 0;
            for (int i = 0; i < len; i++)
            {
                list[i].InternalID = id;
                id++;
            }
        }

        /// <summary>
        /// Method to cleanup style components in the style manager
        /// </summary>
        private void CleanupStyleComponents()
        {
            Style.Border border;
            Style.CellXf cellXf;
            Style.Fill fill;
            Style.Font font;
            Style.NumberFormat numberFormat;
            int len = this.borders.Count;
            int i;
            for (i = len; i >= 0; i--)
            {
                border = (Style.Border)this.borders[i];
                if (IsUsedByStyle(border) == false) { this.borders.RemoveAt(i); }
            }
            len = this.cellXfs.Count;
            for (i = len; i >= 0; i--)
            {
                cellXf = (Style.CellXf)this.cellXfs[i];
                if (IsUsedByStyle(cellXf) == false) { this.cellXfs.RemoveAt(i); }
            }
            len = this.fills.Count;
            for (i = len; i >= 0; i--)
            {
                fill = (Style.Fill)this.fills[i];
                if (IsUsedByStyle(fill) == false) { this.fills.RemoveAt(i); }
            }
            len = this.fonts.Count;
            for (i = len; i >= 0; i--)
            {
                font = (Style.Font)this.fonts[i];
                if (IsUsedByStyle(font) == false) { this.fonts.RemoveAt(i); }
            }
            len = this.numberFormats.Count;
            for (i = len; i >= 0; i--)
            {
                numberFormat = (Style.NumberFormat)this.numberFormats[i];
                if (IsUsedByStyle(numberFormat) == false) { this.numberFormats.RemoveAt(i); }
            }
        }
   
        /// <summary>
        /// Checks whether a style component in the style manager is used by a style
        /// </summary>
        /// <param name="component">Component to check</param>
        /// <returns>If true, the component is in use</returns>
        private bool IsUsedByStyle(AbstractStyle component)
        {
            Style s;
            bool match = false;
            String hash = component.Hash;
            int len = this.styles.Count;
            for(int i = 0; i < len; i++)
            {
                s = (Style)this.styles[i];
                if (component.GetType() == typeof(Style.Border)) { if (s.BorderStyle.Hash == hash) { match = true; break; } }
                else if (component.GetType() == typeof(Style.CellXf)) { if (s.CellXfStyle.Hash == hash) { match = true; break; } }
                if (component.GetType() == typeof(Style.Fill)) { if (s.FillStyle.Hash == hash) { match = true; break; } }
                if (component.GetType() == typeof(Style.Font)) { if (s.FontStyle.Hash == hash) { match = true; break; } }
                if (component.GetType() == typeof(Style.NumberFormat)) { if (s.NumberFormatStyle.Hash == hash) { match = true; break; } }
            }
            return match;
        }



        #endregion
    }

}

﻿/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;

namespace PicoXLSX
{
    /// <summary>
    /// Class representing a style manager to maintain all styles and its components of a workbook
    /// </summary>
    public class StyleManager
    {
        #region privateFields
        private List<AbstractStyle> borders;
        private List<AbstractStyle> cellXfs;
        private List<AbstractStyle> fills;
        private List<AbstractStyle> fonts;
        private List<AbstractStyle> numberFormats;
        private List<AbstractStyle> styles;
        private readonly List<string> styleNames;
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public StyleManager()
        {
            borders = new List<AbstractStyle>();
            cellXfs = new List<AbstractStyle>();
            fills = new List<AbstractStyle>();
            fonts = new List<AbstractStyle>();
            numberFormats = new List<AbstractStyle>();
            styles = new List<AbstractStyle>();
            styleNames = new List<string>();
        }
        #endregion

        #region methods

        /// <summary>
        /// Gets a component by its hash
        /// </summary>
        /// <param name="list">List to check</param>
        /// <param name="hash">Hash of the component</param>
        /// <returns>Determined component. If not found, null will be returned</returns>
        private AbstractStyle GetComponentByHash(ref List<AbstractStyle> list, int hash)
        {
            int len = list.Count;
            for (int i = 0; i < len; i++)
            {
                if (list[i].GetHashCode() == hash)
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
        public Style.Border GetBorderByHash(int hash)
        {
            AbstractStyle component = GetComponentByHash(ref borders, hash);
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
            return Array.ConvertAll(borders.ToArray(), x => (Style.Border)x);
        }

        /// <summary>
        /// Gets the number of borders in the style manager
        /// </summary>
        /// <returns>Number of stored borders</returns>
        public int GetBorderStyleNumber()
        {
            return borders.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a cellXf by its hash
        /// </summary>
        /// <param name="hash">Hash of the cellXf</param>
        /// <returns>Determined cellXf</returns>
        /// <exception cref="StyleException">Throws a StyleException if the cellXf was not found in the style manager</exception>
        public Style.CellXf GetCellXfByHash(int hash)
        {
            AbstractStyle component = GetComponentByHash(ref cellXfs, hash);
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
            return Array.ConvertAll(cellXfs.ToArray(), x => (Style.CellXf)x);
        }

        /// <summary>
        /// Gets the number of cellXfs in the style manager
        /// </summary>
        /// <returns>Number of stored cellXfs</returns>
        public int GetCellXfStyleNumber()
        {
            return cellXfs.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a fill by its hash
        /// </summary>
        /// <param name="hash">Hash of the fill</param>
        /// <returns>Determined fill</returns>
        /// <exception cref="StyleException">Throws a StyleException if the fill was not found in the style manager</exception>
        public Style.Fill GetFillByHash(int hash)
        {
            AbstractStyle component = GetComponentByHash(ref fills, hash);
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
            return Array.ConvertAll(fills.ToArray(), x => (Style.Fill)x);
        }

        /// <summary>
        /// Gets the number of fills in the style manager
        /// </summary>
        /// <returns>Number of stored fills</returns>
        public int GetFillStyleNumber()
        {
            return fills.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a font by its hash
        /// </summary>
        /// <param name="hash">Hash of the font</param>
        /// <returns>Determined font</returns>
        /// <exception cref="StyleException">Throws a StyleException if the font was not found in the style manager</exception>
        public Style.Font GetFontByHash(int hash)
        {
            AbstractStyle component = GetComponentByHash(ref fonts, hash);
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
            return Array.ConvertAll(fonts.ToArray(), x => (Style.Font)x);
        }

        /// <summary>
        /// Gets the number of fonts in the style manager
        /// </summary>
        /// <returns>Number of stored fonts</returns>
        public int GetFontStyleNumber()
        {
            return fonts.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets a numberFormat by its hash
        /// </summary>
        /// <param name="hash">Hash of the numberFormat</param>
        /// <returns>Determined numberFormat</returns>
        /// <exception cref="StyleException">Throws a StyleException if the numberFormat was not found in the style manager</exception>
        public Style.NumberFormat GetNumberFormatByHash(int hash)
        {
            AbstractStyle component = GetComponentByHash(ref numberFormats, hash);
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
            return Array.ConvertAll(numberFormats.ToArray(), x => (Style.NumberFormat)x);
        }

        /// <summary>
        /// Gets the number of numberFormats in the style manager
        /// </summary>
        /// <returns>Number of stored numberFormats</returns>
        public int GetNumberFormatStyleNumber()
        {
            return numberFormats.Count;
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
            int len = styles.Count;
            for (int i = 0; i < len; i++)
            {
                if (((Style)styles[i]).Name == name)
                {
                    return (Style)styles[i];
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
        public Style GetStyleByHash(int hash)
        {
            AbstractStyle component = GetComponentByHash(ref styles, hash);
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
            return Array.ConvertAll(styles.ToArray(), x => (Style)x);
        }

        /// <summary>
        /// Gets the number of styles in the style manager
        /// </summary>
        /// <returns>Number of stored styles</returns>
        public int GetStyleNumber()
        {
            return styles.Count;
        }

        /* ****************************** */


        /// <summary>
        /// Adds a style component to the manager
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Added or determined style in the manager</returns>
        public Style AddStyle(Style style)
        {
            int hash = AddStyleComponent(style);
            return (Style)GetComponentByHash(ref styles, hash);
        }

        /// <summary>
        /// Adds a style component to the manager with an ID
        /// </summary>
        /// <param name="style">Component to add</param>
        /// <param name="id">Id of the component</param>
        /// <returns>Hash of the added or determined component</returns>
        private int AddStyleComponent(AbstractStyle style, int? id)
        {
            style.InternalID = id;
            return AddStyleComponent(style);
        }

        /// <summary>
        /// Adds a style component to the manager
        /// </summary>
        /// <param name="style">Component to add</param>
        /// <returns>Hash of the added or determined component</returns>
        private int AddStyleComponent(AbstractStyle style)
        {
            int hash = style.GetHashCode();
            if (style.GetType() == typeof(Style.Border))
            {
                if (GetComponentByHash(ref borders, hash) == null) { borders.Add(style); }
                Reorganize(ref borders);
            }
            else if (style.GetType() == typeof(Style.CellXf))
            {
                if (GetComponentByHash(ref cellXfs, hash) == null) { cellXfs.Add(style); }
                Reorganize(ref cellXfs);
            }
            else if (style.GetType() == typeof(Style.Fill))
            {
                if (GetComponentByHash(ref fills, hash) == null) { fills.Add(style); }
                Reorganize(ref fills);
            }
            else if (style.GetType() == typeof(Style.Font))
            {
                if (GetComponentByHash(ref fonts, hash) == null) { fonts.Add(style); }
                Reorganize(ref fonts);
            }
            else if (style.GetType() == typeof(Style.NumberFormat))
            {
                if (GetComponentByHash(ref numberFormats, hash) == null) { numberFormats.Add(style); }
                Reorganize(ref numberFormats);
            }
            else if (style.GetType() == typeof(Style))
            {
                Style s = (Style)style;
                if (styleNames.Contains(s.Name))
                {
                    throw new StyleException("StyleAlreadyExistsException", "The style with the name '" + s.Name + "' already exists");
                }
                if (GetComponentByHash(ref styles, hash) == null)
                {
                    int? id;
                    if (!s.InternalID.HasValue)
                    {
                        id = int.MaxValue;
                        s.InternalID = id;
                    }
                    else
                    {
                        id = s.InternalID.Value;
                    }
                    int temp = AddStyleComponent(s.CurrentBorder, id);
                    s.CurrentBorder = (Style.Border)GetComponentByHash(ref borders, temp);
                    temp = AddStyleComponent(s.CurrentCellXf, id);
                    s.CurrentCellXf = (Style.CellXf)GetComponentByHash(ref cellXfs, temp);
                    temp = AddStyleComponent(s.CurrentFill, id);
                    s.CurrentFill = (Style.Fill)GetComponentByHash(ref fills, temp);
                    temp = AddStyleComponent(s.CurrentFont, id);
                    s.CurrentFont = (Style.Font)GetComponentByHash(ref fonts, temp);
                    temp = AddStyleComponent(s.CurrentNumberFormat, id);
                    s.CurrentNumberFormat = (Style.NumberFormat)GetComponentByHash(ref numberFormats, temp);
                    styles.Add(s);
                }
                Reorganize(ref styles);
                hash = s.GetHashCode();
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
            bool match = false;
            int len = styles.Count;
            int index = -1;
            for (int i = 0; i < len; i++)
            {
                if (((Style)styles[i]).Name == styleName)
                {
                    match = true;
                    index = i;
                    break;
                }
            }
            if (!match)
            {
                throw new StyleException("MissingReferenceException", "The style with the name '" + styleName + "' was not found in the style manager");
            }
            styles.RemoveAt(index);
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
            int len = borders.Count - 1;
            int i;
            for (i = len; i >= 0; i--)
            {
                border = (Style.Border)borders[i];
                if (!IsUsedByStyle(border)) { borders.RemoveAt(i); }
            }
            len = cellXfs.Count - 1;
            for (i = len; i >= 0; i--)
            {
                cellXf = (Style.CellXf)cellXfs[i];
                if (!IsUsedByStyle(cellXf)) { cellXfs.RemoveAt(i); }
            }
            len = fills.Count - 1;
            for (i = len; i >= 0; i--)
            {
                fill = (Style.Fill)fills[i];
                if (!IsUsedByStyle(fill)) { fills.RemoveAt(i); }
            }
            len = fonts.Count - 1;
            for (i = len; i >= 0; i--)
            {
                font = (Style.Font)fonts[i];
                if (!IsUsedByStyle(font)) { fonts.RemoveAt(i); }
            }
            len = numberFormats.Count - 1;
            for (i = len; i >= 0; i--)
            {
                numberFormat = (Style.NumberFormat)numberFormats[i];
                if (!IsUsedByStyle(numberFormat)) { numberFormats.RemoveAt(i); }
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
            int hash = component.GetHashCode();
            int len = styles.Count;
            for (int i = 0; i < len; i++)
            {
                s = (Style)styles[i];
                if (component.GetType() == typeof(Style.Border)) { if (s.CurrentBorder.GetHashCode() == hash) { match = true; break; } }
                else if (component.GetType() == typeof(Style.CellXf) && s.CurrentCellXf.GetHashCode() == hash)
                {
                     match = true; break; 
                }
                if (component.GetType() == typeof(Style.Fill) && s.CurrentFill.GetHashCode() == hash)
                { 
                     match = true; break;
                }
                if (component.GetType() == typeof(Style.Font) && s.CurrentFont.GetHashCode() == hash)
                {
                    match = true; break;
                }
                if (component.GetType() == typeof(Style.NumberFormat) && s.CurrentNumberFormat.GetHashCode() == hash)
                { 
                     match = true; break;
                }
            }
            return match;
        }



        #endregion
    }

}

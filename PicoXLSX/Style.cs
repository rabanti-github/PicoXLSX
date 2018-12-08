/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace PicoXLSX
{
    /// <summary>
    /// Class representing a Style with sub classes within a style sheet. An instance of this class is only a container for the different sub-classes. These sub-classes contain the actual styling information.
    /// </summary>
    public class Style : AbstractStyle
    {
        #region privateFields
        private string name;
        private bool internalStyle;
        private bool styleNameDefined = false;
        private StyleManager styleManagerReference = null;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the current Border object of the style
        /// </summary>
        [AppendAttribute(NestedProperty = true)]
        public Border CurrentBorder { get; set; }
        /// <summary>
        /// Gets or sets the  current CellXf object of the style
        /// </summary>
        [AppendAttribute(NestedProperty = true)]
        public CellXf CurrentCellXf { get; set; }
        /// <summary>
        /// Gets or sets the current Fill object of the style
        /// </summary>
        [AppendAttribute(NestedProperty = true)]
        public Fill CurrentFill { get; set; }
        /// <summary>
        /// Gets or sets the  current Font object of the style
        /// </summary>
        [AppendAttribute(NestedProperty = true)]
        public Font CurrentFont { get; set; }
        /// <summary>
        /// Gets or sets the  current NumberFormat object of the style
        /// </summary>
        [AppendAttribute(NestedProperty = true)]
        public NumberFormat CurrentNumberFormat { get; set; }
        /// <summary>
        /// Gets or sets the name of the style. If not defined, the automatically calculated hash will be used as name
        /// </summary>
        [AppendAttribute(Ignore = true)]
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                styleNameDefined = true;
            }
        }

        /// <summary>
        /// Sets the reference of the style manager
        /// </summary>
        [AppendAttribute(Ignore = true)]
        public StyleManager StyleManagerReference
        {
            set
            {
                styleManagerReference = value;
                ReorganizeStyle();
            }
        }

        /// <summary>
        /// Gets whether the style is system internal. Such styles are not meant to be altered
        /// </summary>
        [AppendAttribute(Ignore = true)]
        public bool IsInternalStyle
        {
            get { return internalStyle; }
        }

        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Style()
        {
            CurrentBorder = new Border();
            CurrentCellXf = new CellXf();
            CurrentFill = new Fill();
            CurrentFont = new Font();
            CurrentNumberFormat = new NumberFormat();
            styleNameDefined = false;
            name = this.GetHashCode().ToString();
        }

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="name">Name of the style</param>
        public Style(string name)
        {
            CurrentBorder = new Border();
            CurrentCellXf = new CellXf();
            CurrentFill = new Fill();
            CurrentFont = new Font();
            CurrentNumberFormat = new NumberFormat();
            styleNameDefined = false;
            this.name = name;
        }

        /// <summary>
        /// Constructor with parameters (internal use)
        /// </summary>
        /// <param name="name">Name of the style</param>
        /// <param name="forcedOrder">Number of the style for sorting purpose. Style will be placed to this position (internal use only)</param>
        /// <param name="internalStyle">If true, the style is marked as internal</param>
        public Style(string name, int forcedOrder, bool internalStyle)
        {
            CurrentBorder = new Border();
            CurrentCellXf = new CellXf();
            CurrentFill = new Fill();
            CurrentFont = new Font();
            CurrentNumberFormat = new NumberFormat();
            this.name = name;
            InternalID = forcedOrder;
            this.internalStyle = internalStyle;
            styleNameDefined = true;
        }
        #endregion

        #region methods

        /// <summary>
        /// Appends the specified style parts to the current one. The parts can be instances of sub-classes like Border or CellXf or a Style instance. Only the altered properties of the specified style or style part that differs from a new / untouched style instance will be appended. This enables method chaining
        /// </summary>
        /// <param name="styleToAppend">The style to append or a sub-class of Style</param>
        /// <returns>Current style with appended style parts</returns>
        public Style Append(AbstractStyle styleToAppend)
        {
            if (styleToAppend.GetType() == typeof(Border))
            {
                CurrentBorder.CopyProperties<Border>((Border)styleToAppend, new Border());
            }
            else if (styleToAppend.GetType() == typeof(CellXf))
            {
                CurrentCellXf.CopyProperties<CellXf>((CellXf)styleToAppend, new CellXf());
            }
            else if (styleToAppend.GetType() == typeof(Fill))
            {
                CurrentFill.CopyProperties<Fill>((Fill)styleToAppend, new Fill());
            }
            else if (styleToAppend.GetType() == typeof(Font))
            {
                CurrentFont.CopyProperties<Font>((Font)styleToAppend, new Font());
            }
            else if (styleToAppend.GetType() == typeof(NumberFormat))
            {
                CurrentNumberFormat.CopyProperties<NumberFormat>((NumberFormat)styleToAppend, new NumberFormat());
            }
            else if (styleToAppend.GetType() == typeof(Style))
            {
                CurrentBorder.CopyProperties<Border>(((Style)styleToAppend).CurrentBorder, new Border());
                CurrentCellXf.CopyProperties<CellXf>(((Style)styleToAppend).CurrentCellXf, new CellXf());
                CurrentFill.CopyProperties<Fill>(((Style)styleToAppend).CurrentFill, new Fill());
                CurrentFont.CopyProperties<Font>(((Style)styleToAppend).CurrentFont, new Font());
                CurrentNumberFormat.CopyProperties<NumberFormat>(((Style)styleToAppend).CurrentNumberFormat, new NumberFormat());
            }
            return this;
        }

        /// <summary>
        /// Method to reorganize / synchronize the components of this style
        /// </summary>
        private void ReorganizeStyle()
        {
            if (styleManagerReference == null) { return; }

            Style newStyle = styleManagerReference.AddStyle(this);
            CurrentBorder = newStyle.CurrentBorder;
            CurrentCellXf = newStyle.CurrentCellXf;
            CurrentFill = newStyle.CurrentFill;
            CurrentFont = newStyle.CurrentFont;
            CurrentNumberFormat = newStyle.CurrentNumberFormat;

            if (styleNameDefined == false)
            {
                name = this.GetHashCode().ToString();
            }
        }

        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class instance</returns>
        public override string ToString()
        {
            return InternalID.ToString() + "->" + this.GetHashCode();
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        /// <exception cref="StyleException">MissingReferenceException - The hash of the style could not be created because one or more components are missing as references</exception>
        public override int GetHashCode()
        {
            if (CurrentBorder == null || CurrentCellXf == null || CurrentFill == null || CurrentFont == null || CurrentNumberFormat == null)
            {
                throw new StyleException("MissingReferenceException", "The hash of the style could not be created because one or more components are missing as references");
            }

            int p = 241;
            int r = 1;
            r *= p + this.CurrentBorder.GetHashCode();
            r *= p + this.CurrentCellXf.GetHashCode();
            r *= p + this.CurrentFill.GetHashCode();
            r *= p + this.CurrentFont.GetHashCode();
            r *= p + this.CurrentNumberFormat.GetHashCode();
            return r;
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            if (CurrentBorder == null || CurrentCellXf == null || CurrentFill == null || CurrentFont == null || CurrentNumberFormat == null)
            {
                throw new StyleException("MissingReferenceException", "The style could not be copied because one or more components are missing as references");
            }
            Style copy = new Style();
            copy.CurrentBorder = CurrentBorder.CopyBorder();
            copy.CurrentCellXf = CurrentCellXf.CopyCellXf();
            copy.CurrentFill = CurrentFill.CopyFill();
            copy.CurrentFont = CurrentFont.CopyFont();
            copy.CurrentNumberFormat = CurrentNumberFormat.CopyNumberFormat();
            return copy;
        }

        /// <summary>
        /// Method to copy the current object to a new one with casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public Style CopyStyle()
        {
            return (Style)Copy();
        }

        #endregion

        /*  ************************************************************************************  */
        #region border
        /// <summary>
        /// Class representing a Border entry. The Border entry is used to define frames and cell borders
        /// </summary>
        public class Border : AbstractStyle
        {
            #region enums
            /// <summary>
            /// Enum for the border style
            /// </summary>
            public enum StyleValue
            {
                /// <summary>no border</summary>
                none,
                /// <summary>hair border</summary>
                hair,
                /// <summary>dotted border</summary>
                dotted,
                /// <summary>dashed border with double-dots</summary>
                dashDotDot,
                /// <summary>dash-dotted border</summary>
                dashDot,
                /// <summary>dashed border</summary>
                dashed,
                /// <summary>thin border</summary>
                thin,
                /// <summary>medium-dashed border with double-dots</summary>
                mediumDashDotDot,
                /// <summary>slant dash-dotted border</summary>
                slantDashDot,
                /// <summary>medium dash-dotted border</summary>
                mediumDashDot,
                /// <summary>medium dashed border</summary>
                mediumDashed,
                /// <summary>medium border</summary>
                medium,
                /// <summary>thick border</summary>
                thick,
                /// <summary>double border</summary>
                s_double,
            }
            #endregion

            #region properties
            /// <summary>
            /// Gets or sets the color code of the bottom border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string BottomColor { get; set; }
            /// <summary>
            /// Gets or sets the  style of bottom cell border
            /// </summary>
            public StyleValue BottomStyle { get; set; }
            /// <summary>
            /// Gets or sets the color code of the diagonal lines. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string DiagonalColor { get; set; }
            /// <summary>
            /// Gets or sets whether the downwards diagonal line is used. If true, the line is used
            /// </summary>
            public bool DiagonalDown { get; set; }
            /// <summary>
            /// Gets or sets whether the upwards diagonal line is used. If true, the line is used
            /// </summary>
            public bool DiagonalUp { get; set; }
            /// <summary>
            /// Gets or sets the style of the diagonal lines
            /// </summary>
            public StyleValue DiagonalStyle { get; set; }
            /// <summary>
            /// Gets or sets the color code of the left border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string LeftColor { get; set; }
            /// <summary>
            /// Gets or sets the style of left cell border
            /// </summary>
            public StyleValue LeftStyle { get; set; }
            /// <summary>
            /// Gets or sets the color code of the right border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string RightColor { get; set; }
            /// <summary>
            /// Gets or sets the style of right cell border
            /// </summary>
            public StyleValue RightStyle { get; set; }
            /// <summary>
            /// Gets or sets the color code of the top border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string TopColor { get; set; }
            /// <summary>
            /// Gets or sets the style of top cell border
            /// </summary>
            public StyleValue TopStyle { get; set; }
            #endregion

            #region constructors
            /// <summary>
            /// Default constructor
            /// </summary>
            public Border()
            {
                BottomColor = string.Empty;
                TopColor = string.Empty;
                LeftColor = string.Empty;
                RightColor = string.Empty;
                DiagonalColor = string.Empty;
                LeftStyle = StyleValue.none;
                RightStyle = StyleValue.none;
                TopStyle = StyleValue.none;
                BottomStyle = StyleValue.none;
                DiagonalStyle = StyleValue.none;
                DiagonalDown = false;
                DiagonalUp = false;
            }
            #endregion

            #region methods
            /// <summary>
            /// Returns a hash code for this instance.
            /// </summary>
            /// <returns>
            /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
            /// </returns>
            public override int GetHashCode()
            {
                int p = 271;
                int r = 1;
                r *= p + (int)this.BottomStyle;
                r *= p + (int)this.DiagonalStyle;
                r *= p + (int)this.TopStyle;
                r *= p + (int)this.LeftStyle;
                r *= p + (int)this.RightStyle;
                r *= p + this.BottomColor.GetHashCode();
                r *= p + this.DiagonalColor.GetHashCode();
                r *= p + this.TopColor.GetHashCode();
                r *= p + this.LeftColor.GetHashCode();
                r *= p + this.RightColor.GetHashCode();
                r *= p + (this.DiagonalDown ? 0 : 1);
                r *= p + (this.DiagonalUp ? 0 : 1);
                return r;
            }

            /// <summary>
            /// Method to copy the current object to a new one without casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public override AbstractStyle Copy()
            {
                Border copy = new Border();
                copy.BottomColor = BottomColor;
                copy.BottomStyle = BottomStyle;
                copy.DiagonalColor = DiagonalColor;
                copy.DiagonalDown = DiagonalDown;
                copy.DiagonalStyle = DiagonalStyle;
                copy.DiagonalUp = DiagonalUp;
                copy.LeftColor = LeftColor;
                copy.LeftStyle = LeftStyle;
                copy.RightColor = RightColor;
                copy.RightStyle = RightStyle;
                copy.TopColor = TopColor;
                copy.TopStyle = TopStyle;
                return copy;
            }

            /// <summary>
            /// Method to copy the current object to a new one with casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public Border CopyBorder()
            {
                return (Border)Copy();
            }

            /// <summary>
            /// Override toString method
            /// </summary>
            /// <returns>String of a class</returns>
            public override string ToString()
            {
                return "Border:" + this.GetHashCode();
            }

            /// <summary>
            /// Method to determine whether the object has no values but the default values (means: is empty and must not be processed)
            /// </summary>
            /// <returns>True if empty, otherwise false</returns>
            public bool IsEmpty()
            {
                bool state = true;
                if (BottomColor != string.Empty) { state = false; }
                if (TopColor != string.Empty) { state = false; }
                if (LeftColor != string.Empty) { state = false; }
                if (RightColor != string.Empty) { state = false; }
                if (DiagonalColor != string.Empty) { state = false; }
                if (LeftStyle != StyleValue.none) { state = false; }
                if (RightStyle != StyleValue.none) { state = false; }
                if (TopStyle != StyleValue.none) { state = false; }
                if (BottomStyle != StyleValue.none) { state = false; }
                if (DiagonalStyle != StyleValue.none) { state = false; }
                if (DiagonalDown != false) { state = false; }
                if (DiagonalUp != false) { state = false; }
                return state;
            }
            #endregion

            #region staticMethods
            /// <summary>
            /// Gets the border style name from the enum
            /// </summary>
            /// <param name="style">Enum to process</param>
            /// <returns>The valid value of the border style as String</returns>
            public static string GetStyleName(StyleValue style)
            {
                string output = "";
                switch (style)
                {
                    case StyleValue.none:
                        output = "";
                        break;
                    case StyleValue.hair:
                        break;
                    case StyleValue.dotted:
                        output = "dotted";
                        break;
                    case StyleValue.dashDotDot:
                        output = "dashDotDot";
                        break;
                    case StyleValue.dashDot:
                        output = "dashDot";
                        break;
                    case StyleValue.dashed:
                        output = "dashed";
                        break;
                    case StyleValue.thin:
                        output = "thin";
                        break;
                    case StyleValue.mediumDashDotDot:
                        output = "mediumDashDotDot";
                        break;
                    case StyleValue.slantDashDot:
                        output = "slantDashDot";
                        break;
                    case StyleValue.mediumDashDot:
                        output = "mediumDashDot";
                        break;
                    case StyleValue.mediumDashed:
                        output = "mediumDashed";
                        break;
                    case StyleValue.medium:
                        output = "medium";
                        break;
                    case StyleValue.thick:
                        output = "thick";
                        break;
                    case StyleValue.s_double:
                        output = "double";
                        break;
                    default:
                        output = "";
                        break;
                }
                return output;
            }
            #endregion

        }
        #endregion

        #region cellXf
        /// <summary>
        /// Class representing an XF entry. The XF entry is used to make reference to other style instances like Border or Fill and for the positioning of the cell content
        /// </summary>
        public class CellXf : AbstractStyle
        {
            #region enums
            /// <summary>
            /// Enum for the horizontal alignment of a cell 
            /// </summary>
            public enum HorizontalAlignValue
            {
                /// <summary>Content will be aligned left</summary>
                left,
                /// <summary>Content will be aligned in the center</summary>
                center,
                /// <summary>Content will be aligned right</summary>
                right,
                /// <summary>Content will fill up the cell</summary>
                fill,
                /// <summary>justify alignment</summary>
                justify,
                /// <summary>General alignment</summary>
                general,
                /// <summary>Center continuous alignment</summary>
                centerContinuous,
                /// <summary>Distributed alignment</summary>
                distributed,
                /// <summary>No alignment. The alignment will not be used in a style</summary>
                none,
            }

            /// <summary>
            /// Enum for text break options
            /// </summary>
            public enum TextBreakValue
            {
                /// <summary>Word wrap is active</summary>
                wrapText,
                /// <summary>Text will be resized to fit the cell</summary>
                shrinkToFit,
                /// <summary>Text will overflow in cell</summary>
                none,
            }

            /// <summary>
            /// Enum for the general text alignment direction
            /// </summary>
            public enum TextDirectionValue
            {
                /// <summary>Text direction is horizontal (default)</summary>
                horizontal,
                /// <summary>Text direction is vertical</summary>
                vertical,
            }

            /// <summary>
            /// Enum for the vertical alignment of a cell 
            /// </summary>
            public enum VerticalAlignValue
            {
                /// <summary>Content will be aligned on the bottom (default)</summary>
                bottom,
                /// <summary>Content will be aligned on the top</summary>
                top,
                /// <summary>Content will be aligned in the center</summary>
                center,
                /// <summary>justify alignment</summary>
                justify,
                /// <summary>Distributed alignment</summary>
                distributed,
                /// <summary>No alignment. The alignment will not be used in a style</summary>
                none,
            }
            #endregion

            #region privateFields
            private int textRotation;
            private TextDirectionValue textDirection;
            #endregion

            #region properties
            /// <summary>
            /// Gets or sets whether the applyAlignment property (used to merge cells) will be defined in the XF entry of the style. If true, applyAlignment will be defined
            /// </summary>
            public bool ForceApplyAlignment { get; set; }
            /// <summary>
            /// Gets or sets whether the hidden property (used for protection or hiding of cells) will be defined in the XF entry of the style. If true, hidden will be defined
            /// </summary>
            public bool Hidden { get; set; }
            /// <summary>
            /// Gets or sets the horizontal alignment of the style
            /// </summary>
            public HorizontalAlignValue HorizontalAlign { get; set; }
            /// <summary>
            /// Gets or sets whether the locked property (used for locking / protection of cells or worksheets) will be defined in the XF entry of the style. If true, locked will be defined
            /// </summary>
            public bool Locked { get; set; }
            /// <summary>
            /// Gets or sets the text break options of the style
            /// </summary>
            public TextBreakValue Alignment { get; set; }
            /// <summary>
            /// Gets or sets the direction of the text within the cell
            /// </summary>
            public TextDirectionValue TextDirection
            {
                get { return textDirection; }
                set
                {
                    textDirection = value;
                    CalculateInternalRotation();
                }
            }
            /// <summary>
            /// Gets or sets the text rotation in degrees (from +90 to -90)
            /// </summary>
            public int TextRotation
            {
                get { return textRotation; }
                set
                {
                    textRotation = value;
                    TextDirection = TextDirectionValue.horizontal;
                    CalculateInternalRotation();
                }
            }
            /// <summary>
            /// Gets or sets the vertical alignment of the style
            /// </summary>
            public VerticalAlignValue VerticalAlign { get; set; }
            #endregion

            #region constructors
            /// <summary>
            /// Default constructor
            /// </summary>
            public CellXf()
            {
                HorizontalAlign = HorizontalAlignValue.none;
                Alignment = TextBreakValue.none;
                textDirection = TextDirectionValue.horizontal;
                VerticalAlign = VerticalAlignValue.none;
                textRotation = 0;
            }
            #endregion

            #region methods
            /// <summary>
            /// Method to calculate the internal text rotation. The text direction and rotation are handled internally by the text rotation value
            /// </summary>
            /// <returns>Returns the valid rotation in degrees for internal uses (LowLevel)</returns>
            /// <exception cref="FormatException">Throws a FormatException if the rotation angle (-90 to 90) is out of range</exception>
            public int CalculateInternalRotation()
            {
                if (textRotation < -90 || textRotation > 90)
                {
                    throw new FormatException("The rotation value (" + textRotation.ToString() + "°) is out of range. Range is form -90° to +90°");
                }
                if (textDirection == TextDirectionValue.vertical)
                {
                    return 255;
                }
                else
                {
                    if (textRotation >= 0)
                    {
                        return textRotation;
                    }
                    else
                    {
                        return (90 - textRotation);
                    }
                }
            }

            /// <summary>
            /// Override toString method
            /// </summary>
            /// <returns>String of a class instance</returns>
            public override string ToString()
            {
                return "StyleXF:" + this.GetHashCode();
            }

            /// <summary>
            /// Returns a hash code for this instance.
            /// </summary>
            /// <returns>
            /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
            /// </returns>
            public override int GetHashCode()
            {
                int p = 269;
                int r = 1;
                r *= p + (int)this.HorizontalAlign;
                r *= p + (int)this.VerticalAlign;
                r *= p + (int)this.Alignment;
                r *= p + (int)this.TextDirection;
                r *= p + this.TextRotation;
                r *= p + (this.ForceApplyAlignment ? 0 : 1);
                r *= p + (this.Locked ? 0 : 1);
                r *= p + (this.Hidden ? 0 : 1);
                return r;
            }

            /// <summary>
            /// Method to copy the current object to a new one without casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public override AbstractStyle Copy()
            {
                CellXf copy = new CellXf();
                copy.HorizontalAlign = HorizontalAlign;
                copy.Alignment = Alignment;
                copy.TextDirection = TextDirection;
                copy.TextRotation = TextRotation;
                copy.VerticalAlign = VerticalAlign;
                copy.ForceApplyAlignment = ForceApplyAlignment;
                copy.Locked = Locked;
                copy.Hidden = Hidden;
                return copy;
            }

            /// <summary>
            /// Method to copy the current object to a new one with casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public CellXf CopyCellXf()
            {
                return (CellXf)Copy();
            }


            #endregion

        }
        #endregion

        #region fill
        /// <summary>
        /// Class representing a Fill (background) entry. The Fill entry is used to define background colors and fill patterns
        /// </summary>
        public class Fill : AbstractStyle
        {
            #region constants
            /// <summary>
            /// Default Color (foreground or background) as constant
            /// </summary>
            public const string DEFAULTCOLOR = "FF000000";
            #endregion

            #region enums
            /// <summary>
            /// Enum for the type of the color
            /// </summary>
            public enum FillType
            {
                /// <summary>Color defines a pattern color </summary>
                patternColor,
                /// <summary>Color defines a solid fill color </summary>
                fillColor,
            }
            /// <summary>
            /// Enum for the pattern values
            /// </summary>
            public enum PatternValue
            {
                /// <summary>No pattern (default)</summary>
                none,
                /// <summary>Solid fill (for colors)</summary>
                solid,
                /// <summary>Dark gray fill</summary>
                darkGray,
                /// <summary>Medium gray fill</summary>
                mediumGray,
                /// <summary>Light gray fill</summary>
                lightGray,
                /// <summary>6.25% gray fill</summary>
                gray0625,
                /// <summary>12.5% gray fill</summary>
                gray125,
            }
            #endregion

            #region properties
            /// <summary>
            /// Gets or sets the background color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string BackgroundColor { get; set; }
            /// <summary>
            /// Gets or sets the foreground color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string ForegroundColor { get; set; }
            /// <summary>
            /// Gets or sets the indexed color (Default is 64)
            /// </summary>
            public int IndexedColor { get; set; }
            /// <summary>
            /// Gets or sets the pattern type of the fill (Default is none)
            /// </summary>
            public PatternValue PatternFill { get; set; }
            #endregion

            #region constructors
            /// <summary>
            /// Default constructor
            /// </summary>
            public Fill()
            {
                IndexedColor = 64;
                PatternFill = PatternValue.none;
                ForegroundColor = DEFAULTCOLOR;
                BackgroundColor = DEFAULTCOLOR;
            }
            /// <summary>
            /// Constructor with foreground and background color
            /// </summary>
            /// <param name="foreground">Foreground color of the fill</param>
            /// <param name="background">Background color of the fill</param>
            public Fill(string foreground, string background)
            {
                BackgroundColor = background;
                ForegroundColor = foreground;
                IndexedColor = 64;
                PatternFill = PatternValue.solid;
            }

            /// <summary>
            /// Constructor with color value and fill type
            /// </summary>
            /// <param name="value">Color value</param>
            /// <param name="filltype">Fill type (fill or pattern)</param>
            public Fill(string value, FillType filltype)
            {
                if (filltype == FillType.fillColor)
                {
                    BackgroundColor = value;
                    ForegroundColor = DEFAULTCOLOR;
                }
                else
                {
                    BackgroundColor = DEFAULTCOLOR;
                    ForegroundColor = value;
                }
                IndexedColor = 64;
                PatternFill = PatternValue.solid;
            }
            #endregion

            #region methods
           
            /// <summary>
            /// Override toString method
            /// </summary>
            /// <returns>String of a class</returns>
            public override string ToString()
            {
                return "Fill:" + this.GetHashCode();
            }

            /// <summary>
            /// Method to copy the current object to a new one without casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public override AbstractStyle Copy()
            {
                Fill copy = new Fill();
                copy.BackgroundColor = BackgroundColor;
                copy.ForegroundColor = ForegroundColor;
                copy.IndexedColor = IndexedColor;
                copy.PatternFill = PatternFill;
                return copy;
            }

            /// <summary>
            /// Returns a hash code for this instance.
            /// </summary>
            /// <returns>
            /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
            /// </returns>
            public override int GetHashCode()
            {
                int p = 263;
                int r = 1;
                r *= p + this.IndexedColor;
                r *= p + (int)this.PatternFill;
                r *= p + this.ForegroundColor.GetHashCode();
                r *= p + this.BackgroundColor.GetHashCode();
                return r;
            }

            /// <summary>
            /// Method to copy the current object to a new one with casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public Fill CopyFill()
            {
                return (Fill)Copy();
            }

            /// <summary>
            /// Sets the color and the depending fill type
            /// </summary>
            /// <param name="value">color value</param>
            /// <param name="filltype">fill type (fill or pattern)</param>
            public void SetColor(string value, FillType filltype)
            {
                if (filltype == FillType.fillColor)
                {
                    ForegroundColor = value;
                    BackgroundColor = DEFAULTCOLOR;
                }
                else
                {
                    ForegroundColor = DEFAULTCOLOR;
                    BackgroundColor = value;
                }
                PatternFill = PatternValue.solid;
            }
            #endregion

            #region staticMethods
            /// <summary>
            /// Gets the pattern name from the enum
            /// </summary>
            /// <param name="pattern">Enum to process</param>
            /// <returns>The valid value of the pattern as String</returns>
            public static string GetPatternName(PatternValue pattern)
            {
                string output;
                switch (pattern)
                {
                    case PatternValue.none:
                        output = "none";
                        break;
                    case PatternValue.solid:
                        output = "solid";
                        break;
                    case PatternValue.darkGray:
                        output = "darkGray";
                        break;
                    case PatternValue.mediumGray:
                        output = "mediumGray";
                        break;
                    case PatternValue.lightGray:
                        output = "lightGray";
                        break;
                    case PatternValue.gray0625:
                        output = "gray0625";
                        break;
                    case PatternValue.gray125:
                        output = "gray125";
                        break;
                    default:
                        output = "none";
                        break;
                }
                return output;
            }
            #endregion

        }

        #endregion

        #region font
        /// <summary>
        /// Class representing a Font entry. The Font entry is used to define text formatting
        /// </summary>
        public class Font : AbstractStyle
        {
            #region constants
            /// <summary>
            /// Default font family as constant
            /// </summary>
            public const string DEFAULTFONT = "Calibri";
            #endregion

            #region enums
            /// <summary>
            /// Enum for the font scheme
            /// </summary>
            public enum SchemeValue
            {
                /// <summary>Font scheme is major</summary>
                major,
                /// <summary>Font scheme is minor (default)</summary>
                minor,
                /// <summary>No Font scheme is used</summary>
                none,
            }
            /// <summary>
            /// Enum for the vertical alignment of the text from base line
            /// </summary>
            public enum VerticalAlignValue
            {
                // baseline, // Maybe not used in Excel
                /// <summary>Text will be rendered as subscript</summary>
                subscript,
                /// <summary>Text will be rendered as superscript</summary>
                superscript,
                /// <summary>Text will be rendered normal</summary>
                none,
            }
            #endregion

            #region privateFields
            private int size;
            #endregion

            #region properties
            /// <summary>
            /// Gets or sets whether the font is bold. If true, the font is declared as bold
            /// </summary>
            public bool Bold { get; set; }
            /// <summary>
            /// Gets or sets the char set of the Font (Default is empty)
            /// </summary>
            public string Charset { get; set; }
            /// <summary>
            /// Gets or sets the font color theme (Default is 1)
            /// </summary>
            public int ColorTheme { get; set; }
            /// <summary>
            /// Gets or sets the font color (default is empty)
            /// </summary>
            public string ColorValue { get; set; }
            /// <summary>
            /// Gets or sets whether the font has a double underline. If true, the font is declared with a double underline
            /// </summary>
            public bool DoubleUnderline { get; set; }
            /// <summary>
            ///  Gets or sets the font family (Default is 2)
            /// </summary>
            public string Family { get; set; }
            /// <summary>
            /// Gets whether the font is equals the default font
            /// </summary>
            [AppendAttribute(Ignore = true)]
            public bool IsDefaultFont
            {
                get
                {
                    Font temp = new Font();
                    return Equals(temp);
                }
            }
            /// <summary>
            /// Gets or sets whether the font is italic. If true, the font is declared as italic
            /// </summary>
            public bool Italic { get; set; }
            /// <summary>
            /// Gets or sets the font name (Default is Calibri)
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// Gets or sets the font scheme (Default is minor)
            /// </summary>
            public SchemeValue Scheme { get; set; }
            /// <summary>
            /// Gets or sets the font size. Valid range is from 8 to 75
            /// </summary>
            public int Size
            {
                get { return size; }
                set
                {
                    if (value < 8) { size = 8; }
                    else if (value > 75) { size = 72; }
                    else { size = value; }
                }
            }
            /// <summary>
            /// Gets or sets whether the font is struck through. If true, the font is declared as strike-through
            /// </summary>
            public bool Strike { get; set; }
            /// <summary>
            /// Gets or sets whether the font is underlined. If true, the font is declared as underlined
            /// </summary>
            public bool Underline { get; set; }
            /// <summary>
            /// Gets or sets the alignment of the font (Default is none)
            /// </summary>
            public VerticalAlignValue VerticalAlign { get; set; }
            #endregion

            #region constructors
            /// <summary>
            /// Default constructor
            /// </summary>
            public Font()
            {
                size = 11;
                Name = DEFAULTFONT;
                Family = "2";
                ColorTheme = 1;
                ColorValue = string.Empty;
                Charset = string.Empty;
                Scheme = SchemeValue.minor;
                VerticalAlign = VerticalAlignValue.none;
            }
            #endregion

            #region methods            
            /// <summary>
            /// Override toString method
            /// </summary>
            /// <returns>String of a class</returns>
            public override string ToString()
            {
                return "Font:" + this.GetHashCode();
            }

            /// <summary>
            /// Method to copy the current object to a new one without casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public override AbstractStyle Copy()
            {
                Font copy = new Font();
                copy.Bold = Bold;
                copy.Charset = Charset;
                copy.ColorTheme = ColorTheme;
                copy.ColorValue = ColorValue;
                copy.VerticalAlign = VerticalAlign;
                copy.DoubleUnderline = DoubleUnderline;
                copy.Family = Family;
                copy.Italic = Italic;
                copy.Name = Name;
                copy.Scheme = Scheme;
                copy.Size = Size;
                copy.Strike = Strike;
                copy.Underline = Underline;
                return copy;
            }

            /// <summary>
            /// Returns a hash code for this instance.
            /// </summary>
            /// <returns>
            /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
            /// </returns>
            public override int GetHashCode()
            {
                int p = 257;
                int r = 1;
                r *= p + (this.Bold ? 0 : 1);
                r *= p + (this.Italic ? 0 : 1);
                r *= p + (this.Underline ? 0 : 1);
                r *= p + (this.DoubleUnderline ? 0 : 1);
                r *= p + (this.Strike ? 0 : 1);
                r *= p + this.ColorTheme;
                r *= p + this.ColorValue.GetHashCode();
                r *= p + this.Family.GetHashCode();
                r *= p + this.Name.GetHashCode();
                r *= p + this.Scheme.GetHashCode();
                r *= p + (int)this.VerticalAlign;
                r *= p + this.Charset.GetHashCode();
                r *= p + this.size;
                return r;
            }

            /// <summary>
            /// Method to copy the current object to a new one with casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public Font CopyFont()
            {
                return (Font)Copy();
            }

            #endregion
        }
        #endregion

        #region numberFormat
        /// <summary>
        /// Class representing a NumberFormat entry. The NumberFormat entry is used to define cell formats like currency or date
        /// </summary>
        public class NumberFormat : AbstractStyle
        {
            #region constants
            /// <summary>
            /// Start ID for custom number formats as constant
            /// </summary>
            public const int CUSTOMFORMAT_START_NUMBER = 124;
            #endregion

            #region enums
            /// <summary>
            /// Enum for predefined number formats
            /// </summary>
            public enum FormatNumber
            {
                /// <summary>No format / Default</summary>
                none = 0,
                /// <summary>Format: 0</summary>
                format_1 = 1,
                /// <summary>Format: 0.00</summary>
                format_2 = 2,
                /// <summary>Format: #,##0</summary>
                format_3 = 3,
                /// <summary>Format: #,##0.00</summary>
                format_4 = 4,
                /// <summary>Format: $#,##0_);($#,##0)</summary>
                format_5 = 5,
                /// <summary>Format: $#,##0_);[Red]($#,##0)</summary>
                format_6 = 6,
                /// <summary>Format: $#,##0.00_);($#,##0.00)</summary>
                format_7 = 7,
                /// <summary>Format: $#,##0.00_);[Red]($#,##0.00)</summary>
                format_8 = 8,
                /// <summary>Format: 0%</summary>
                format_9 = 9,
                /// <summary>Format: 0.00%</summary>
                format_10 = 10,
                /// <summary>Format: 0.00E+00</summary>
                format_11 = 11,
                /// <summary>Format: # ?/?</summary>
                format_12 = 12,
                /// <summary>Format: # ??/??</summary>
                format_13 = 13,
                /// <summary>Format: m/d/yyyy</summary>
                format_14 = 14,
                /// <summary>Format: d-mmm-yy</summary>
                format_15 = 15,
                /// <summary>Format: d-mmm</summary>
                format_16 = 16,
                /// <summary>Format: mmm-yy</summary>
                format_17 = 17,
                /// <summary>Format: mm AM/PM</summary>
                format_18 = 18,
                /// <summary>Format: h:mm:ss AM/PM</summary>
                format_19 = 19,
                /// <summary>Format: h:mm</summary>
                format_20 = 20,
                /// <summary>Format: h:mm:ss</summary>
                format_21 = 21,
                /// <summary>Format: m/d/yyyy h:mm</summary>
                format_22 = 22,
                /// <summary>Format: #,##0_);(#,##0)</summary>
                format_37 = 37,
                /// <summary>Format: #,##0_);[Red](#,##0)</summary>
                format_38 = 38,
                /// <summary>Format: #,##0.00_);(#,##0.00)</summary>
                format_39 = 39,
                /// <summary>Format: #,##0.00_);[Red](#,##0.00)</summary>
                format_40 = 40,
                /// <summary>Format: mm:ss</summary>
                format_45 = 45,
                /// <summary>Format: [h]:mm:ss</summary>
                format_46 = 46,
                /// <summary>Format: mm:ss.0</summary>
                format_47 = 47,
                /// <summary>Format: ##0.0E+0</summary>
                format_48 = 48,
                /// <summary>Format: #</summary>
                format_49 = 49,
                /// <summary>Custom Format (ID 164 and higher)</summary>
                custom = 164,
            }
            #endregion

            #region properties
            /// <summary>
            /// Gets or sets the custom format code in the notation of Excel
            /// </summary>
            public string CustomFormatCode { get; set; }
            /// <summary>
            /// Gets or sets the format number of the custom format. Must be higher or equal then predefined custom number (164) 
            /// </summary>
            public int CustomFormatID { get; set; }
            /// <summary>
            /// Gets whether the number format is a custom format (higher or equals 164). If true, the format is custom
            /// </summary>
            [AppendAttribute(Ignore = true)]
            public bool IsCustomFormat
            {
                get
                {
                    if (Number == FormatNumber.custom) { return true; }
                    else { return false; }
                }
            }
            /// <summary>
            /// Gets or sets the format number. Set this to custom (164) in case of custom number formats
            /// </summary>
            public FormatNumber Number { get; set; }
            #endregion

            #region constructors
            /// <summary>
            /// Default constructor
            /// </summary>
            public NumberFormat()
            {
                Number = FormatNumber.none;
                CustomFormatCode = string.Empty;
                CustomFormatID = CUSTOMFORMAT_START_NUMBER;
            }
            #endregion

            #region methods

            /// <summary>
            /// Override toString method
            /// </summary>
            /// <returns>String of a class</returns>
            public override string ToString()
            {
                return "NumberFormat:" + this.GetHashCode();
            }

            /// <summary>
            /// Method to copy the current object to a new one without casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public override AbstractStyle Copy()
            {
                NumberFormat copy = new NumberFormat();
                copy.CustomFormatCode = CustomFormatCode;
                copy.CustomFormatID = CustomFormatID;
                copy.Number = Number;
                return copy;
            }

            /// <summary>
            /// Method to copy the current object to a new one with casting
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public NumberFormat CopyNumberFormat()
            {
                return (NumberFormat)Copy();
            }

            public override int GetHashCode()
            {
                int p = 251;
                int r = 1;
                r *= p + this.CustomFormatCode.GetHashCode();
                r *= p + this.CustomFormatID;
                r *= p + (int)this.Number;
                return r;
            }

            #endregion
        }
        #endregion

        #region subClass_BasicStyles
        /// <summary>
        /// Factory class with the most important predefined styles
        /// </summary>
        public static class BasicStyles
        {
            #region enums
            /// <summary>
            /// Enum with style selection
            /// </summary>
            private enum StyleEnum
            {
                /// <summary>Format text bold</summary>
                bold,
                /// <summary>Format text italic</summary>
                italic,
                /// <summary>Format text bold and italic</summary>
                boldItalic,
                /// <summary>Format text with an underline</summary>
                underline,
                /// <summary>Format text with a double underline</summary>
                doubleUnderline,
                /// <summary>Format text with a strike-through</summary>
                strike,
                /// <summary>Format number as date</summary>
                dateFormat,
                /// <summary>Rounds number as an integer</summary>
                roundFormat,
                /// <summary>Format cell with a thin border</summary>
                borderFrame,
                /// <summary>Format cell with a thin border and a thick bottom line as header cell</summary>
                borderFrameHeader,
                /// <summary>Special pattern fill style for compatibility purpose </summary>
                dottedFill_0_125,
                /// <summary>Style to apply on merged cells </summary>
                mergeCellStyle,
            }
            #endregion

            #region staticFields
            private static Style bold, italic, boldItalic, underline, doubleUnderline, strike, dateFormat, roundFormat, borderFrame, borderFrameHeader, dottedFill_0_125, mergeCellStyle;
            #endregion

            #region staticProperties
            /// <summary>Gets the bold style</summary>
            public static Style Bold
            { get { return GetStyle(StyleEnum.bold); } }
            /// <summary>Gets the bold and italic style</summary>
            public static Style BoldItalic
            { get { return GetStyle(StyleEnum.boldItalic); } }
            /// <summary>Gets the border frame style</summary>
            public static Style BorderFrame
            { get { return GetStyle(StyleEnum.borderFrame); } }
            /// <summary>Gets the border style for header cells</summary>
            public static Style BorderFrameHeader
            { get { return GetStyle(StyleEnum.borderFrameHeader); } }
            /// <summary>Gets the date format style</summary>
            public static Style DateFormat
            { get { return GetStyle(StyleEnum.dateFormat); } }
            /// <summary>Gets the double underline style</summary>
            public static Style DoubleUnderline
            { get { return GetStyle(StyleEnum.doubleUnderline); } }
            /// <summary>Gets the special pattern fill style (for compatibility)</summary>
            public static Style DottedFill_0_125
            { get { return GetStyle(StyleEnum.dottedFill_0_125); } }
            /// <summary>Gets the italic style</summary>
            public static Style Italic
            { get { return GetStyle(StyleEnum.italic); } }
            /// <summary>Gets the style used when merging cells</summary>
            public static Style MergeCellStyle
            { get { return GetStyle(StyleEnum.mergeCellStyle); } }
            /// <summary>Gets the round format style</summary>
            public static Style RoundFormat
            { get { return GetStyle(StyleEnum.roundFormat); } }
            /// <summary>Gets the strike style</summary>
            public static Style Strike
            { get { return GetStyle(StyleEnum.strike); } }
            /// <summary>Gets the underline style</summary>
            public static Style Underline
            { get { return GetStyle(StyleEnum.underline); } }
            #endregion

            #region staticMethods
            /// <summary>
            /// Method to maintain the styles and to create singleton instances
            /// </summary>
            /// <param name="value">Enum value to maintain</param>
            /// <returns>The style according to the passed enum value</returns>
            private static Style GetStyle(StyleEnum value)
            {
                Style s = null;
                switch (value)
                {
                    case StyleEnum.bold:
                        if (bold == null)
                        {
                            bold = new Style();
                            bold.CurrentFont.Bold = true;
                        }
                        s = bold;
                        break;
                    case StyleEnum.italic:
                        if (italic == null)
                        {
                            italic = new Style();
                            italic.CurrentFont.Italic = true;
                        }
                        s = italic;
                        break;
                    case StyleEnum.boldItalic:
                        if (boldItalic == null)
                        {
                            boldItalic = new Style();
                            boldItalic.CurrentFont.Italic = true;
                            boldItalic.CurrentFont.Bold = true;
                        }
                        s = boldItalic;
                        break;
                    case StyleEnum.underline:
                        if (underline == null)
                        {
                            underline = new Style();
                            underline.CurrentFont.Underline = true;
                        }
                        s = underline;
                        break;
                    case StyleEnum.doubleUnderline:
                        if (doubleUnderline == null)
                        {
                            doubleUnderline = new Style();
                            doubleUnderline.CurrentFont.DoubleUnderline = true;
                        }
                        s = doubleUnderline;
                        break;
                    case StyleEnum.strike:
                        if (strike == null)
                        {
                            strike = new Style();
                            strike.CurrentFont.Strike = true;
                        }
                        s = strike;
                        break;
                    case StyleEnum.dateFormat:
                        if (dateFormat == null)
                        {
                            dateFormat = new Style();
                            dateFormat.CurrentNumberFormat.Number = NumberFormat.FormatNumber.format_14;
                        }
                        s = dateFormat;
                        break;
                    case StyleEnum.roundFormat:
                        if (roundFormat == null)
                        {
                            roundFormat = new Style();
                            roundFormat.CurrentNumberFormat.Number = NumberFormat.FormatNumber.format_1;
                        }
                        s = roundFormat;
                        break;
                    case StyleEnum.borderFrame:
                        if (borderFrame == null)
                        {
                            borderFrame = new Style();
                            borderFrame.CurrentBorder.TopStyle = Border.StyleValue.thin;
                            borderFrame.CurrentBorder.BottomStyle = Border.StyleValue.thin;
                            borderFrame.CurrentBorder.LeftStyle = Border.StyleValue.thin;
                            borderFrame.CurrentBorder.RightStyle = Border.StyleValue.thin;
                        }
                        s = borderFrame;
                        break;
                    case StyleEnum.borderFrameHeader:
                        if (borderFrameHeader == null)
                        {
                            borderFrameHeader = new Style();
                            borderFrameHeader.CurrentBorder.TopStyle = Border.StyleValue.thin;
                            borderFrameHeader.CurrentBorder.BottomStyle = Border.StyleValue.medium;
                            borderFrameHeader.CurrentBorder.LeftStyle = Border.StyleValue.thin;
                            borderFrameHeader.CurrentBorder.RightStyle = Border.StyleValue.thin;
                            borderFrameHeader.CurrentFont.Bold = true;
                        }
                        s = borderFrameHeader;
                        break;
                    case StyleEnum.dottedFill_0_125:
                        if (dottedFill_0_125 == null)
                        {
                            dottedFill_0_125 = new Style();
                            dottedFill_0_125.CurrentFill.PatternFill = Fill.PatternValue.gray125;
                        }
                        s = dottedFill_0_125;
                        break;
                    case StyleEnum.mergeCellStyle:
                        if (mergeCellStyle == null)
                        {
                            mergeCellStyle = new Style();
                            mergeCellStyle.CurrentCellXf.ForceApplyAlignment = true;
                        }
                        s = mergeCellStyle;
                        break;
                    default:
                        break;
                }
                return s.CopyStyle(); // Copy makes basic styles immutable
            }

            /// <summary>
            /// Gets a style to colorize the text of a cell
            /// </summary>
            /// <param name="rgb">RGB code in hex format (e.g. FF00AC). Alpha will be set to full opacity (FF)</param>
            /// <returns>Style with font color definition</returns>
            public static Style ColorizedText(string rgb)
            {
                Style s = new Style();
                s.CurrentFont.ColorValue = "FF" + rgb.ToUpper();
                return s;
            }

            /// <summary>
            /// Gets a style to colorize the background of a cell
            /// </summary>
            /// <param name="rgb">RGB code in hex format (e.g. FF00AC). Alpha will be set to full opacity (FF)</param>
            /// <returns>Style with background color definition</returns>
            public static Style ColorizedBackground(string rgb)
            {
                Style s = new Style();
                s.CurrentFill.SetColor("FF" + rgb.ToUpper(), Fill.FillType.fillColor);
                return s;
            }

            /// <summary>
            /// Gets a style with a user defined font
            /// </summary>
            /// <param name="fontName">Name of the font</param>
            /// <param name="fontSize">Size of the font in points (optional; default 11)</param>
            /// <param name="isBold">If true, the font will be bold (optional; default false)</param>
            /// <param name="isItalic">If true, the font will be italic (optional; default false)</param>
            /// <returns>Style with font definition</returns>
            /// <remarks>The font name as well as the availability of bold and italic on the font cannot be validated by PicoXLSX. The generated file may be corrupt or rendered with a fall-back font in case of an error</remarks>
            public static Style Font(string fontName, int fontSize = 11, bool isBold = false, bool isItalic = false)
            {
                Style s = new Style();
                s.CurrentFont.Name = fontName;
                s.CurrentFont.Size = fontSize;
                s.CurrentFont.Bold = isBold;
                s.CurrentFont.Italic = isItalic;
                return s;
            }
            #endregion
        }
        #endregion

    }

    /// <summary>
    /// Class represents an abstract style component
    /// </summary>
    public abstract class AbstractStyle : IComparable<AbstractStyle>
    {
        /// <summary>
        /// Gets or sets the internal ID for sorting purpose in the Excel style document (nullable)
        /// </summary>
        [AppendAttribute(Ignore = true)]
        public int? InternalID { get; set; }


        /// <summary>
        /// Abstract method to copy a component (dereferencing)
        /// </summary>
        /// <returns>Returns a copied component</returns>
        public abstract AbstractStyle Copy();

        /// <summary>
        /// Internal method to copy altered properties from a source object. The decision whether a property is copied is dependent on a untouched reference object
        /// </summary>
        /// <typeparam name="T">Style or sub-class of Style that extends AbstractStyle</typeparam>
        /// <param name="source">Source object with properties to copy</param>
        /// <param name="reference">Reference object to decide whether the properties from the source objects are altered or not</param>
        internal void CopyProperties<T>(T source, T reference) where T : AbstractStyle
        {
            if (GetType() != source.GetType() && GetType() != reference.GetType())
            {
                throw new StyleException("CopyPropertyException", "The objects of the source, target and reference for style appending are not of the same type");
            }
            bool ignore;
            PropertyInfo[] infos = GetType().GetProperties();
            PropertyInfo sourceInfo, referenceInfo;
            IEnumerable<AppendAttribute> attributes;
            foreach (PropertyInfo info in infos)
            {
                attributes = (IEnumerable< AppendAttribute>)info.GetCustomAttributes(typeof(AppendAttribute));
                if (attributes.Count() > 0)
                {
                    ignore = false;
                    foreach (AppendAttribute attribute in attributes)
                    {
                        if (attribute.Ignore == true || attribute.NestedProperty == true)
                        {
                            ignore = true;
                            break;
                        }
                    }
                    if (ignore == true) { continue; } // skip property
                }

                sourceInfo = source.GetType().GetProperty(info.Name);
                referenceInfo = reference.GetType().GetProperty(info.Name);
                if (sourceInfo.GetValue(source).Equals(referenceInfo.GetValue(reference)) == false)
                {
                    info.SetValue(this, sourceInfo.GetValue(source));
                }
            }
        }

        /// <summary>
        /// Method to compare two objects for sorting purpose
        /// </summary>
        /// <param name="other">Other object to compare with this object</param>
        /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
        public int CompareTo(AbstractStyle other)
        {
            if (InternalID.HasValue == false) { return -1; }
            else if (other.InternalID.HasValue == false) { return 1; }
            else { return InternalID.Value.CompareTo(other.InternalID.Value); }
        }

        /// <summary>
        /// Method to compare two objects for sorting purpose
        /// </summary>
        /// <param name="other">Other object to compare with this object</param>
        /// <returns>True if both objects are equal, otherwise false</returns>
        public bool Equals(AbstractStyle other)
        {
            return this.GetHashCode() == other.GetHashCode();
        }

        /// <summary>
        /// Method to cast values of the components to string values for the hash calculation (protected/internal static method)
        /// </summary>
        /// <param name="o">Value to cast</param>
        /// <param name="sb">StringBuilder reference to put the casted object in</param>
        /// <param name="delimiter">Delimiter character to append after the casted value</param>
        protected static void CastValue(object o, ref StringBuilder sb, char? delimiter)
        {
            if (o == null)
            {
                sb.Append('#');
            }
            else if (o.GetType() == typeof(bool))
            {
                if ((bool)o == true) { sb.Append(1); }
                else { sb.Append(0); }
            }
            else if (o.GetType() == typeof(int))
            {
                sb.Append((int)o);
            }
            else if (o.GetType() == typeof(double))
            {
                sb.Append((double)o);
            }
            else if (o.GetType() == typeof(float))
            {
                sb.Append((float)o);
            }
            else if (o.GetType() == typeof(string))
            {
                if (o.ToString() == "#")
                {
                    sb.Append("_#_");
                }
                else
                {
                    sb.Append((string)o);
                }
            }
            else if (o.GetType() == typeof(long))
            {
                sb.Append((long)o);
            }
            else if (o.GetType() == typeof(char))
            {
                sb.Append((char)o);
            }
            else
            {
                sb.Append(o);
            }
            if (delimiter.HasValue == true)
            {
                sb.Append(delimiter.Value);
            }
        }

        /// <summary>
        /// Attribute designated to control the copying of style properties
        /// </summary>
        /// <seealso cref="System.Attribute" />
        public class AppendAttribute : Attribute
        {
            /// <summary>
            /// Indicates whether the property annotated with the attribute is ignored during the copying of properties
            /// </summary>
            /// <value>
            ///   <c>true</c> if ignored, otherwise <c>false</c>.
            /// </value>
            public bool Ignore { get; set; }

            /// <summary>
            /// Indicates whether the property annotated with the attribute is a nested property. Nested properties are ignored but during the copying of properties but can be broken down to its sub-properties
            /// </summary>
            /// <value>
            ///   <c>true</c> if a nested property, otherwise <c>false</c>.
            /// </value>
            public bool NestedProperty { get; set; }

            /// <summary>
            /// Default constructor
            /// </summary>
            public AppendAttribute()
            {
                Ignore = false;
                NestedProperty = false;
            }
        }
    }

    /*  ************************************************************************************  */




}

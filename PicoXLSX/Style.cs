/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace PicoXLSX
{
    /// <summary>
    /// Class representing a Style with sub classes within a style sheet. An instance of this class is only a container for the different sub-classes. These sub-classes contains the actual styling information.
    /// </summary>
    public class Style : IComparable<Style>, IEquatable<Style>
    {
        /// <summary>
        /// Current Font object of the style
        /// </summary>
        public Font CurrentFont { get; set; }
        /// <summary>
        /// Current Fill object of the style
        /// </summary>
        public Fill CurrentFill { get; set; }
        /// <summary>
        /// Current Border object of the style
        /// </summary>
        public Border CurrentBorder { get; set; }
        /// <summary>
        /// Current NumberFormat object of the style
        /// </summary>
        public NumberFormat CurrentNumberFormat { get; set; }
        /// <summary>
        /// Current CellXf object of the style
        /// </summary>
        public CellXf CurrentCellXf { get; set; }
        /// <summary>
        /// Internal ID for sorting purpose
        /// </summary>
        public int InternalID { get; set; }

        private string name;

        /// <summary>
        /// Name of the style. If not defined, a random name will be generated when the style is created
        /// </summary>
        public string Name
        {
            get { return name; }
        }
        

        /// <summary>
        /// Default constructor
        /// </summary>
        public Style()
        {
            this.CurrentFont = new Font();
            this.CurrentFill = new Fill();
            this.CurrentBorder = new Border();
            this.CurrentNumberFormat = new NumberFormat();
            this.CurrentCellXf = new CellXf();
            this.name = "Style_" + LowLevel.CreateUniqueName(16);
        }

        /// <summary>
        /// Constructor with definition of the style name
        /// </summary>
        /// <param name="name">Name of the style</param>
        public Style(string name) : this()
        {
            this.name = name;
        }

        /// <summary>
        /// Class representing an XF entry. The XF entry is used to make reference to other style instances like Border oder Fill and for the positioning of the cell content
        /// </summary>
        public class CellXf : IComparable<CellXf>, IEquatable<CellXf>
        {
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
            /// Enum for the vertical alignment of a cell 
            /// </summary>
            public enum VerticallAlignValue
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
            /// If true, the style is used for locking / protection of cells or worksheets
            /// </summary>
            public bool Locked { get; set; }
            /// <summary>
            /// If true, the style is used for hiding cell values / protection of cells
            /// </summary>
            public bool Hidden { get; set; }
            /// <summary>
            /// Horizontal alignment of the style
            /// </summary>
            public HorizontalAlignValue HorizontalAlign { get; set; }
            /// <summary>
            /// Vertical alignment of the style
            /// </summary>
            public VerticallAlignValue VerticalAlign { get; set; }
            /// <summary>
            /// Text break options of the style
            /// </summary>
            public TextBreakValue Alignment { get; set; }
            /// <summary>
            /// Internal ID for sorting purpose
            /// </summary>
            public int InternalID { get; set; }
  
            private int textRotation;
            private TextDirectionValue textDirection; 

            /// <summary>
            /// Text rotation in degrees (from +90 to -90)
            /// </summary>
            public int TextRotation
            {
                get { return textRotation; }
                set 
                {
                    textRotation = value;
                    this.TextDirection = TextDirectionValue.horizontal;
                    CalculateInternalRotation();
                }
            }

            /// <summary>
            /// Direction of the text within the cell
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
            /// If true, the applyAlignment value of the style will be set to true (used to merge cells)
            /// </summary>
            public bool ForceApplyAlignment { get; set; }
            
            /// <summary>
            /// Default constructor
            /// </summary>
            public CellXf()
            {
                this.HorizontalAlign = HorizontalAlignValue.none;
                this.Alignment = TextBreakValue.none;
                this.textDirection = TextDirectionValue.horizontal;
                this.VerticalAlign = VerticallAlignValue.none;
                this.textRotation = 0;
            }

            /// <summary>
            /// Method to calculate the internal text rotation. The text direction and rotation are handled internally by the text rotation value
            /// </summary>
            /// <returns>Returns the valid rotation in degrees for internal uses (LowLevel)</returns>
            /// <exception cref="FormatException">Throws a FormatException if the rotation angle (-90 to 90) is out of range</exception>
            public int CalculateInternalRotation()
            {
                if (this.textRotation < -90 || this.textRotation > 90)
                {
                    throw new FormatException("The rotation value (" + this.textRotation.ToString() + "°) is out of range. Range is form -90° to +90°");
                }
                if (this.textDirection == TextDirectionValue.vertical)
                {
                    return 255;
                }
                else
                {
                    if (this.textRotation >= 0)
                    {
                        return this.textRotation;
                    }
                    else
                    {
                        return (90 - this.textRotation);
                    }
                }
            }

            /// <summary>
            /// Method to determine the equality of two objects
            /// </summary>
            /// <param name="other">Object to compare against this object</param>
            /// <returns>True if both objects are equal, otherwise false</returns>
            public bool Equals(CellXf other)
            {
                if (other == null) { return false; }
                if (this.HorizontalAlign != other.HorizontalAlign) { return false; }
                if (this.Alignment != other.Alignment) { return false; }
                if (this.TextDirection != other.TextDirection) { return false; }
                if (this.TextRotation != other.TextRotation) { return false; }
                if (this.VerticalAlign != other.VerticalAlign) { return false; }
                if (this.ForceApplyAlignment != other.ForceApplyAlignment) { return false; }
                if (this.Locked != other.Locked) { return false; }
                if (this.Hidden != other.Hidden) { return false; }
                return true;
            }

            /// <summary>
            /// method to compare two objects for sorting purpose
            /// </summary>
            /// <param name="other">Other object to compare with this object</param>
            /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
            public int CompareTo(CellXf other)
            {
                return this.InternalID.CompareTo(other.InternalID);
            }

            /// <summary>
            /// Method to copy the current object to a new one
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public CellXf Copy()
            {
                CellXf copy = new CellXf();
                copy.HorizontalAlign = this.HorizontalAlign;
                copy.Alignment = this.Alignment;
                copy.TextDirection = this.TextDirection;
                copy.TextRotation = this.TextRotation;
                copy.VerticalAlign = this.VerticalAlign;
                copy.ForceApplyAlignment = this.ForceApplyAlignment;
                copy.Locked = this.Locked;
                copy.Hidden = this.Hidden;
                return copy;
            }

        }

        /// <summary>
        /// Class representing a Font entry. The Font entry is used to define text formatting
        /// </summary>
        public class Font : IComparable<Font>, IEquatable<Font>
        {
            /// <summary>
            /// Default font family as constant
            /// </summary>
            public const string DEFAULTFONT = "Calibri";

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

            private int size;

            /// <summary>
            /// Font size. Valid range is from 8 to 75
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
            /// Font name (Default is Calibri)
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            ///  Font family (Default is 2)
            /// </summary>
            public string Family { get; set; }
            /// <summary>
            /// Font color theme (Default is 1)
            /// </summary>
            public int ColorTheme { get; set; }
            /// <summary>
            /// Font color (default is empty)
            /// </summary>
            public string ColorValue { get; set; }
            /// <summary>
            /// Font scheme (Default is minor)
            /// </summary>
            public SchemeValue Scheme { get; set; }
            /// <summary>
            /// Alignment of the font (Default is none)
            /// </summary>
            public VerticalAlignValue VerticalAlign { get; set; }
            /// <summary>
            /// If true, the font is bold
            /// </summary>
            public bool Bold { get; set; }
            /// <summary>
            /// If true, the font is italic
            /// </summary>
            public bool Italic { get; set; }
            /// <summary>
            /// If true, the font has one underline
            /// </summary>
            public bool Underline { get; set; }
            /// <summary>
            /// If true, the font has a double underline
            /// </summary>
            public bool DoubleUnderline { get; set; }
            /// <summary>
            /// If true, the font is stroked through
            /// </summary>
            public bool Strike { get; set; }
            /// <summary>
            /// Charset of the Font (Default is empty)
            /// </summary>
            public string Charset { get; set; }
            /// <summary>
            /// Internal ID for sorting purpose
            /// </summary>
            public int InternalID { get; set; }
            /// <summary>
            /// In true the font is equals the default font
            /// </summary>
            public bool IsDefaultFont
            {
                get 
                {
                    Font temp = new Font();
                    return this.Equals(temp); 
                }
            }

            /// <summary>
            /// Default constructor
            /// </summary>
            public Font()
            {
                this.size = 11;
                this.Name = DEFAULTFONT;
                this.Family = "2";
                this.ColorTheme = 1;
                this.ColorValue = string.Empty;
                this.Charset = string.Empty;
                this.Scheme = SchemeValue.minor;
                this.VerticalAlign = VerticalAlignValue.none;
            }

            /// <summary>
            /// Method to determine the equality of two objects
            /// </summary>
            /// <param name="other">Object to compare against this object</param>
            /// <returns>True if both objects are equal, otherwise false</returns>
            public bool Equals(Font other)
            {
                if (other == null) { return false; }
                if (this.Bold != other.Bold) { return false; }
                if (this.ColorTheme != other.ColorTheme) { return false; }
                if (this.DoubleUnderline != other.DoubleUnderline) { return false; }
                if (this.Family != other.Family) { return false; }
                if (this.Italic != other.Italic) { return false; }
                if (this.Name != other.Name) { return false; }
                if (this.Scheme != other.Scheme) { return false; }
                if (this.VerticalAlign != other.VerticalAlign) { return false; }
                if (this.Charset != other.Charset) { return false; }
                if (this.Size != other.Size) { return false; }
                if (this.Strike != other.Strike) { return false; }
                if (this.Underline != other.Underline) { return false; }
                else { return true; }
            }

            /// <summary>
            /// method to compare two objects for sorting purpose
            /// </summary>
            /// <param name="other">Other object to compare with this object</param>
            /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
            public int CompareTo(Font other)
            {
                return this.InternalID.CompareTo(other.InternalID);
            }

            /// <summary>
            /// Method to copy the current object to a new one
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public Font Copy()
            {
                Font copy = new Font();
                copy.Bold = this.Bold;
                copy.Charset = this.Charset;
                copy.ColorTheme = this.ColorTheme;
                copy.VerticalAlign = this.VerticalAlign;
                copy.DoubleUnderline = this.DoubleUnderline;
                copy.Family = this.Family;
                copy.Italic = this.Italic;
                copy.Name = this.Name;
                copy.Scheme = this.Scheme;
                copy.Size = this.Size;
                copy.Strike = this.Strike;
                copy.Underline = this.Underline;
                return copy;
            }
        }

        /// <summary>
        /// Class representing a Fill (background) entry. The Fill entry is used to define background colors and fill patterns
        /// </summary>
        public class Fill : IComparable<Fill>, IEquatable<Fill>
        {
            /// <summary>
            /// Default Color (foreground or background) as constant
            /// </summary>
            public const string DEFAULTCOLOR = "FF000000";

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
            /// Gets the pattern name from the enum
            /// </summary>
            /// <param name="pattern">Enum to process</param>
            /// <returns>The valid value of the pattern as String</returns>
            public static string GetPatternName(PatternValue pattern)
            {
                string output = "";
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

            /// <summary>
            /// Indexed color (Default is 64)
            /// </summary>
            public int IndexedColor { get; set; }
            /// <summary>
            /// Pattern type of the fill (Default is none)
            /// </summary>
            public PatternValue PatternFill { get; set; }
            /// <summary>
            /// Foreground color of the fill
            /// </summary>
            public string ForegroundColor { get; set; }
            /// <summary>
            /// Background color of the fill
            /// </summary>
            public string BackgroundColor { get; set; }
            /// <summary>
            /// Internal ID for sorting purpose
            /// </summary>
            public int InternalID { get; set; }
            /// <summary>
            /// Default constructor
            /// </summary>
            public Fill()
            {
                this.IndexedColor = 64;
                this.PatternFill = PatternValue.none;
                this.ForegroundColor = DEFAULTCOLOR;
                this.BackgroundColor = DEFAULTCOLOR;
            }
            /// <summary>
            /// Constructor with foreground and background color
            /// </summary>
            /// <param name="forground">Foreground color of the fill</param>
            /// <param name="background">Background color of the fill</param>
            public Fill(string forground, string background)
            {
                this.BackgroundColor = background;
                this.ForegroundColor = forground;
                this.IndexedColor = 64;
                this.PatternFill = PatternValue.solid;
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
                    this.BackgroundColor = value;
                    this.ForegroundColor = DEFAULTCOLOR;
                }
                else
                {
                    this.BackgroundColor = DEFAULTCOLOR;
                    this.ForegroundColor = value;
                }
                this.IndexedColor = 64;
                this.PatternFill = PatternValue.solid;
            }

            /// <summary>
            /// Seth the color an the depending fill type
            /// </summary>
            /// <param name="value">color value</param>
            /// <param name="filltype">fill type (fill or pattern)</param>
            public void SetColor(string value, FillType filltype)
            {
                if (filltype == FillType.fillColor)
                {
                    this.ForegroundColor = value;
                    this.BackgroundColor = DEFAULTCOLOR;
                }
                else
                {
                    this.ForegroundColor = DEFAULTCOLOR;
                    this.BackgroundColor = value;
                }
                this.PatternFill = PatternValue.solid;
            }

            /// <summary>
            /// Method to determine the equality of two objects
            /// </summary>
            /// <param name="other">Object to compare against this object</param>
            /// <returns>True if both objects are equal, otherwise false</returns>
            public bool Equals(Fill other)
            {
                if (other == null) { return false; }
                if (this.IndexedColor != other.IndexedColor) { return false; }
                if (this.PatternFill != other.PatternFill) { return false; }
                if (this.ForegroundColor != other.ForegroundColor) { return false; }
                if (this.BackgroundColor != other.BackgroundColor) { return false; }
                else { return true; }
            }

            /// <summary>
            /// method to compare two objects for sorting purpose
            /// </summary>
            /// <param name="other">Other object to compare with this object</param>
            /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
            public int CompareTo(Fill other)
            {
                return this.InternalID.CompareTo(other.InternalID);
            }

            /// <summary>
            /// Method to copy the current object to a new one
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public Fill Copy()
            {
                Fill copy = new Fill();
                copy.BackgroundColor = this.BackgroundColor;
                copy.ForegroundColor = this.ForegroundColor;
                copy.IndexedColor = this.IndexedColor;
                copy.PatternFill = this.PatternFill;
                return copy;
            }
        }

        /// <summary>
        /// Class representing a Border entry. The Border entry is used to define frames and cell borders
        /// </summary>
        public class Border : IComparable<Border>, IEquatable<Border>
        {
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

            /// <summary>
            /// Style of left cell border
            /// </summary>
            public StyleValue LeftStyle { get; set; }
            /// <summary>
            /// Style of right cell border
            /// </summary>
            public StyleValue RightStyle { get; set; }
            /// <summary>
            /// Style of top cell border
            /// </summary>
            public StyleValue TopStyle { get; set; }
            /// <summary>
            /// Style of bottom cell border
            /// </summary>
            public StyleValue BottomStyle { get; set; }
            /// <summary>
            /// Style of the diagonal lines
            /// </summary>
            public StyleValue DiagonalStyle { get; set; }
            /// <summary>
            /// If true, the downwards diagonal line is used
            /// </summary>
            public bool DiagonalDown { get; set; }
            /// <summary>
            /// If true, the upwards diagonal line is used
            /// </summary>
            public bool DiagonalUp { get; set; }
            /// <summary>
            /// Color code (ARGB) of the left border
            /// </summary>
            public string LeftColor { get; set; }
            /// <summary>
            /// Color code (ARGB) of the right border
            /// </summary>
            public string RightColor { get; set; }
            /// <summary>
            /// Color code (ARGB) of the top border
            /// </summary>
            public string TopColor { get; set; }
            /// <summary>
            /// Color code (ARGB) of the bottom border
            /// </summary>
            public string BottomColor { get; set; }
            /// <summary>
            /// Color code (ARGB) of the diagonal lines
            /// </summary>
            public string DiagonalColor { get; set; }
            /// <summary>
            /// Internal ID for sorting purpose
            /// </summary>
            public int InternalID { get; set; }
            /// <summary>
            /// Default constructor
            /// </summary>
            public Border()
            {
                this.BottomColor = string.Empty;
                this.TopColor = string.Empty;
                this.LeftColor = string.Empty;
                this.RightColor = string.Empty;
                this.DiagonalColor = string.Empty;
                this.LeftStyle = StyleValue.none;
                this.RightStyle = StyleValue.none;
                this.TopStyle = StyleValue.none;
                this.BottomStyle = StyleValue.none;
                this.DiagonalStyle = StyleValue.none;
                this.DiagonalDown = false;
                this.DiagonalUp = false;
            }
            /// <summary>
            /// Method to compare two objects for sorting purpose
            /// </summary>
            /// <param name="other">Other object to compare with this object</param>
            /// <returns>True if both objects are equal, otherwise false</returns>
            public bool Equals(Border other)
            {
                if (other == null) { return false; }
                if (this.BottomColor != other.BottomColor) { return false; }
                if (this.BottomStyle != other.BottomStyle) { return false; }
                if (this.DiagonalColor != other.DiagonalColor) { return false; }
                if (this.DiagonalDown != other.DiagonalDown) { return false; }
                if (this.DiagonalStyle != other.DiagonalStyle) { return false; }
                if (this.DiagonalUp != other.DiagonalUp) { return false; }
                if (this.LeftColor != other.LeftColor) { return false; }
                if (this.LeftStyle != other.LeftStyle) { return false; }
                if (this.RightColor != other.RightColor) { return false; }
                if (this.RightStyle != other.RightStyle) { return false; }
                if (this.TopColor != other.TopColor) { return false; }
                if (this.TopStyle != other.TopStyle) { return false; }
                else { return true; }
            }

            /// <summary>
            /// Method to determine whether the object has no values but the default values (means: is empty and must not be processed)
            /// </summary>
            /// <returns>True if empty, otherwise false</returns>
            public bool IsEmpty()
            {
                bool state = true;
                if (this.BottomColor != string.Empty) {state = false;}
                if (this.TopColor != string.Empty) {state = false;}
                if (this.LeftColor != string.Empty) {state = false;}
                if (this.RightColor != string.Empty) {state = false;}
                if (this.DiagonalColor != string.Empty) {state = false;}
                if (this.LeftStyle != StyleValue.none) {state = false;}
                if (this.RightStyle != StyleValue.none) {state = false;}
                if (this.TopStyle != StyleValue.none) {state = false;}
                if (this.BottomStyle != StyleValue.none) {state = false;}
                if (this.DiagonalStyle != StyleValue.none) {state = false;}
                if (this.DiagonalDown != false) {state = false;}
                if (this.DiagonalUp != false) { state = false; }
                return state;
            }

            /// <summary>
            /// Method to compare two objects for sorting purpose
            /// </summary>
            /// <param name="other">Other object to compare with this object</param>
            /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
            public int CompareTo(Border other)
            {
                return this.InternalID.CompareTo(other.InternalID);
            }

            /// <summary>
            /// Method to copy the current object to a new one
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public Border Copy()
            {
                Border copy = new Border();
                copy.BottomColor = this.BottomColor;
                copy.BottomStyle = this.BottomStyle;
                copy.DiagonalColor = this.DiagonalColor;
                copy.DiagonalDown = this.DiagonalDown;
                copy.DiagonalStyle = this.DiagonalStyle;
                copy.DiagonalUp = this.DiagonalUp;
                copy.LeftColor = this.LeftColor;
                copy.LeftStyle = this.LeftStyle;
                copy.RightColor = this.RightColor;
                copy.RightStyle = this.RightStyle;
                copy.TopColor = this.TopColor;
                copy.TopStyle = this.TopStyle;
                return copy;
            }
        }

        /// <summary>
        /// Class representing a NumberFormat entry. The NumberFormat entry is used to define cell formats like currency or date
        /// </summary>
        public class NumberFormat : IComparable<NumberFormat>, IEquatable<NumberFormat>
        {
            /// <summary>
            /// Start ID for custom number formats as constant
            /// </summary>
            public const int CUSTOMFORMAT_START_NUMBER = 124;
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
            /// <summary>
            /// Format number. Set this to custom (164) in case of custom number formats
            /// </summary>
            public FormatNumber Number { get; set; }
            /// <summary>
            /// Format number of the custom format. Must be higher or equal then predefined custom number (164) 
            /// </summary>
            public int CustomFormatID { get; set; }
            /// <summary>
            /// Custom format code in the notation of Excel
            /// </summary>
            public string CustomFormatCode { get; set; }
            /// <summary>
            /// Returns true in case of a custom format (higher or equals 164)
            /// </summary>
            public bool IsCustomFormat
            {
                get
                {
                    if (Number ==  FormatNumber.custom) { return true; }
                    else { return false; }
                }
            }
            /// <summary>
            /// Internal ID for sorting purpose
            /// </summary>
            public int InternalID { get; set; }
            /// <summary>
            /// Default constructor
            /// </summary>
            public NumberFormat()
            {
                this.Number = FormatNumber.none;
                this.CustomFormatCode = string.Empty;
                this.CustomFormatID = CUSTOMFORMAT_START_NUMBER;
            }

            /// <summary>
            /// Method to determine the equality of two objects
            /// </summary>
            /// <param name="other">Object to compare against this object</param>
            /// <returns>True if both objects are equal, otherwise false</returns>
            public bool Equals(NumberFormat other)
            {
                if (other == null) { return false; }
                if (this.CustomFormatCode != other.CustomFormatCode) { return false; }
                if (this.CustomFormatID != other.CustomFormatID) { return false; }
                if (this.Number != other.Number) { return false; }
                else { return true; }
            }

            /// <summary>
            /// method to compare two objects for sorting purpose
            /// </summary>
            /// <param name="other">Other object to compare with this object</param>
            /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
            public int CompareTo(NumberFormat other)
            {
                return this.InternalID.CompareTo(other.InternalID);
            }

            /// <summary>
            /// Method to copy the current object to a new one
            /// </summary>
            /// <returns>Copy of the current object without the internal ID</returns>
            public NumberFormat Copy()
            {
                NumberFormat copy = new NumberFormat();
                copy.CustomFormatCode = this.CustomFormatCode;
                copy.CustomFormatID = this.CustomFormatID;
                copy.Number = this.Number;
                return copy;
            }

        }

        /// <summary>
        /// Factory class with the most important predefined styles
        /// </summary>
        public static class BasicStyles
        {
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
                doubleUnderlien,
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

            private static Style bold, italic, boldItalic, underline, doubleUnderline, strike, dateFormat, roundFormat, borderFrame, borderFrameHeader, dottedFill_0_125, mergeCellStyle;

            /// <summary>Gets the bold style</summary>
            public static Style Bold
            { get { return GetStyle(StyleEnum.bold); } }
            /// <summary>Gets the italic style</summary>
            public static Style Italic
            { get { return GetStyle(StyleEnum.italic); } }
            /// <summary>Gets the bold and italic style</summary>
            public static Style BoldItalic
            { get { return GetStyle(StyleEnum.boldItalic); } }
            /// <summary>Gets the underline style</summary>
            public static Style Underline
            { get { return GetStyle(StyleEnum.underline); } }
            /// <summary>Gets the double underline style</summary>
            public static Style DoubleUnderline
            { get { return GetStyle(StyleEnum.doubleUnderlien); } }
            /// <summary>Gets the strike style</summary>
            public static Style Strike
            { get { return GetStyle(StyleEnum.strike); } }
            /// <summary>Gets the date format style</summary>
            public static Style DateFormat
            { get { return GetStyle(StyleEnum.dateFormat); } }
            /// <summary>Gets the round format style</summary>
            public static Style RoundFormat
            { get { return GetStyle(StyleEnum.roundFormat); } }
            /// <summary>Gets the border frame style</summary>
            public static Style BorderFrame
            { get { return GetStyle(StyleEnum.borderFrame); } }
            /// <summary>Gets the border style for header cells</summary>
            public static Style BorderFrameHeader
            { get { return GetStyle(StyleEnum.borderFrameHeader); } }
            /// <summary>Gets the special pattern fill style (for compatibility)</summary>
            public static Style DottedFill_0_125
            { get { return GetStyle(StyleEnum.dottedFill_0_125); } }
            /// <summary>Gets the style used when merging cells</summary>
            public static Style MergeCellStyle
            { get { return GetStyle(StyleEnum.mergeCellStyle); } }

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
                    case StyleEnum.doubleUnderlien:
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
                return s;
            }


        }

        /// <summary>
        /// Method to determine the equality of two objects
        /// </summary>
        /// <param name="other">Object to compare against this object</param>
        /// <returns>True if both objects are equal, otherwise false</returns>
        public bool Equals(Style other)
        {
            if (this.CurrentBorder != null && other.CurrentBorder != null)
            {
                if (this.CurrentBorder.Equals(other.CurrentBorder) == false) { return false; }
            }
            if ((this.CurrentBorder == null || other.CurrentBorder == null) && !(this.CurrentBorder == null && other.CurrentBorder == null)) { return false; }

            if (this.CurrentFill != null && other.CurrentFill != null)
            {
                if (this.CurrentFill.Equals(other.CurrentFill) == false) { return false; }
            }
            if ((this.CurrentFill == null || other.CurrentFill == null) && !(this.CurrentFill == null && other.CurrentFill == null)) { return false; }

            if (this.CurrentFont != null && other.CurrentFont != null)
            {
                if (this.CurrentFont.Equals(other.CurrentFont) == false) { return false; }
            }
            if ((this.CurrentFont == null || other.CurrentFont == null) && !(this.CurrentFont == null && other.CurrentFont == null)) { return false; }

            if (this.CurrentNumberFormat != null && other.CurrentNumberFormat != null)
            {
                if (this.CurrentNumberFormat.Equals(other.CurrentNumberFormat) == false) { return false; }
            }
            if ((this.CurrentNumberFormat == null || other.CurrentNumberFormat == null) && !(this.CurrentNumberFormat == null && other.CurrentNumberFormat == null)) { return false; }

            if (this.CurrentCellXf != null && other.CurrentCellXf != null)
            {
                if (this.CurrentCellXf.Equals(other.CurrentCellXf) == false) { return false; }
            }
            if ((this.CurrentCellXf == null || other.CurrentCellXf == null) && !(this.CurrentCellXf == null && other.CurrentCellXf == null)) { return false; }
            return true;
        }

        /// <summary>
        /// method to compare two objects for sorting purpose
        /// </summary>
        /// <param name="other">Other object to compare with this object</param>
        /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
        public int CompareTo(Style other)
        {
            return this.InternalID.CompareTo(other.InternalID);
        }

        /// <summary>
        /// Method to copy the current object to a new one
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public Style Copy()
        {
            Style copy = new Style(this.Name + "_copy");
            copy.CurrentBorder = this.CurrentBorder.Copy();
            copy.CurrentFill = this.CurrentFill.Copy();
            copy.CurrentFont = this.CurrentFont.Copy();
            copy.CurrentNumberFormat = this.CurrentNumberFormat.Copy();
            copy.CurrentCellXf = this.CurrentCellXf.Copy();
            return copy;
        }

        /// <summary>
        /// Method to copy the current object to a new one
        /// </summary>
        /// <param name="overwriteFormat">NumberFormat object to replace the original object of the current object in the copy</param>
        /// <returns>Copy of the current object without the internal ID</returns>
        public Style Copy(NumberFormat overwriteFormat)
        {
            Style copy = new Style(this.Name + "_copy");
            copy.CurrentBorder = this.CurrentBorder.Copy();
            copy.CurrentFill = this.CurrentFill.Copy();
            copy.CurrentFont = this.CurrentFont.Copy();
            copy.CurrentCellXf = this.CurrentCellXf.Copy();
            copy.CurrentNumberFormat = overwriteFormat;
            return copy;
        }

    }
}

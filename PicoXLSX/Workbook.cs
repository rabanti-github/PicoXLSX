/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;

namespace PicoXLSX
{

    /// <summary>
    /// PicoXLSX is a library to generate XLSX files in a easy and native way
    /// </summary>
    [System.Runtime.CompilerServices.CompilerGenerated]
    class NamespaceDoc // This class is only for documentation purpose (Sandcastle)
    { }

    /// <summary>
    /// Class representing a workbook
    /// </summary>
    /// 
    public class Workbook
    {
        private string filename;
        private List<Worksheet> worksheets;
        private Worksheet currentWorksheet;
        private List<Style> styles;
        private Metadata workbookMetadata;
        private string workbookProtectionPassword;
        private bool lockWindowsIfProtected;
        private bool lockStructureIfProtected;
        private int selectedWorksheet;

        /// <summary>
        /// Gets the selected worksheet. The selected worksheet is not the current worksheet while design time but the selected sheet in the output file
        /// </summary>
        public int SelectedWorksheet
        {
            get { return selectedWorksheet; }
        }

        /// <summary>
        /// Gets the current worksheet
        /// </summary>
        public Worksheet CurrentWorksheet
        {
            get { return currentWorksheet; }
        }

        /// <summary>
        /// Gets the list of worksheets in the workbook
        /// </summary>
        public List<Worksheet> Worksheets
        {
            get { return worksheets; }
        }

        /// <summary>
        /// Gets or sets the filename of the workbook
        /// </summary>
        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }

        /// <summary>
        /// Gets the list of Styles of the workbook
        /// </summary>
        public List<Style> Styles
        {
            get { return styles; }
        }

        /// <summary>
        /// Meta data object of the workbook
        /// </summary>
        public Metadata WorkbookMetadata
        {
            get { return workbookMetadata; }
            set { workbookMetadata = value; }
        }


        /// <summary>
        /// Gets or sets whether the workbook is protected
        /// </summary>
        public bool UseWorkbookProtection { get; set; }

        /// <summary>
        /// Gets the password used for workbook protection
        /// </summary>
        /// <see cref="SetWorkbookProtection"/>
        public string WorkbookProtectionPassword
        {
            get { return workbookProtectionPassword; }
        }

        /// <summary>
        /// Gets whether the windows are locked if workbook is protected
        /// </summary>
        /// <see cref="SetWorkbookProtection"/> 
        public bool LockWindowsIfProtected
        {
            get { return lockWindowsIfProtected; }
        }

        /// <summary>
        /// Gets whether the structure are locked if workbook is protected
        /// </summary>
        /// <see cref="SetWorkbookProtection"/>
        public bool LockStructureIfProtected
        {
            get { return lockStructureIfProtected; }
        }
        
        /// <summary>
        /// Default Constructor with additional parameter to create a default worksheet
        /// </summary>
        /// <param name="createWorkSheet">If true, a default worksheet will be crated and set as default worksheet</param>
        public Workbook(bool createWorkSheet)
        { 
            this.worksheets = new List<Worksheet>();
            this.styles = new List<Style>();
            this.styles.Add(new Style("default")); // Do not remove this (Default style)
            this.styles.Add(Style.BasicStyles.DottedFill_0_125);
            this.workbookMetadata = new Metadata();
            if (createWorkSheet == true)
            {
                AddWorksheet("Sheet1");
            }
        }

        /// <summary>
        /// Constructor with filename ant the name of the first worksheet
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="sheetName">Name of the first worksheet</param>
        public Workbook(string filename, string sheetName) : this(false)
        {
            this.filename = filename;
            AddWorksheet(sheetName);
        }

        /// <summary>
        /// Adding a new Worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <exception cref="WorksheetNameAlreadxExistsException">Throws a WorksheetNameAlreadxExistsException if the name of the worksheet already exists</exception>
        /// <exception cref="FormatException">Throws a FormatException if the name contains illegal characters or is out of range (length between 1 an 31 characters)</exception>
        public void AddWorksheet(string name)
        {
            foreach(Worksheet item in this.worksheets)
            {
                if (item.SheetName == name)
                {
                    throw new WorksheetNameAlreadxExistsException("The worksheet with the name '" + name + "' already exists.");
                }
            }
            int number = this.worksheets.Count + 1;
            this.currentWorksheet = new Worksheet(name, number);
            this.worksheets.Add(this.currentWorksheet);
        }

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>Returns the current worksheet</returns>
        /// <exception cref="UnknownWorksheetException">Throws a UnknownWorksheetException if the name of the worksheet is unknown</exception>
        public Worksheet SetCurrentWorksheet(string name)
        {
            bool exists = false;
            foreach (Worksheet item in this.worksheets)
            {
                if (item.SheetName == name)
                {
                    this.currentWorksheet = item;
                    exists = true;
                    break;
                }
            }
            if (exists == false)
            {
                throw new UnknownWorksheetException("The worksheet with the name '" + name + "' does not exist.");
            }
            return this.currentWorksheet;
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheetIndex">Zero-based worksheet index</param>
        /// <exception cref="OutOfRangeException">Throws a OutOfRangeException if the index of the worksheet is out of range</exception>
        public void SetSelectedWorksheet(int worksheetIndex)
        {
            if (worksheetIndex < 0 || worksheetIndex > this.worksheets.Count - 1)
            {
                throw new OutOfRangeException("The worksheet index " + worksheetIndex.ToString() + " is out of range");
            }
            this.selectedWorksheet = worksheetIndex;
        }
        
		/// <summary>
		/// Sets the selected worksheet in the output workbook
		/// </summary>
		/// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
		/// <param name="worksheet">Worksheet object (must be in the collection of worksheets)</param>
		/// <exception cref="UnknownWorksheetException">Throws a UnknownWorksheetException if the worksheet was not found in the worksheet collection</exception>
        public void SetSelectedWorksheet(Worksheet worksheet)
        {
        	bool check = false;
            for(int i = 0; i < this.worksheets.Count; i++)
            {
                if (this.worksheets[i].Equals(worksheet))
                {
                    this.selectedWorksheet = i;
                    check = true;
                    break;
                }
            }
            if (check == false)
            {
            	throw new UnknownWorksheetException("The passed worksheet object is not in the worksheet collection.");
            }
        }

        /// <summary>
        /// Removes the defined worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <exception cref="UnknownWorksheetException">Throws a UnknownWorksheetException if the name of the worksheet is unknown</exception>
        public void RemoveWorksheet(string name)
        {
            bool exists = false;
            bool resetCurrent = false;
            int index = 0;
            for (int i = 0; i < this.worksheets.Count; i++)
            {
                if (this.worksheets[i].SheetName == name)
                {
                    index = i;
                    exists = true;
                    break;
                }
            }
            if (exists == false)
            {
                throw new UnknownWorksheetException("The worksheet with the name '" + name + "' does not exist.");
            }
            if (this.worksheets[index].SheetName == this.currentWorksheet.SheetName )
            {
                resetCurrent = true;
            }
            this.worksheets.RemoveAt(index);
            if (this.worksheets.Count > 0)
            {
                for (int i = 0; i < this.worksheets.Count; i++)
                {
                    this.worksheets[i].SheetID = i + 1;
                    if (resetCurrent == true && i == 0)
                    {
                        this.currentWorksheet = this.worksheets[i];
                    }
                }
            }
            else
            {
                this.currentWorksheet = null;
            }
            if (this.selectedWorksheet > this.worksheets.Count - 1)
            {
                this.selectedWorksheet = this.worksheets.Count - 1;
            }
        }

        /// <summary>
        /// Sets or removes the workbook protection. If protectWindows and protectStructure are both false, the workbook will not be protected
        /// </summary>
        /// <param name="state">If true, the workbook will be protected, otherwise not</param>
        /// <param name="protectWindows">If true, the windows will be locked if the workbook is protected</param>
        /// <param name="protectStructure">If true, the structure will be locked if the workbook is protected</param>
        /// <param name="password">Optional password. If null or empty, no password will be set in case of protection</param>
        public void SetWorkbookProtection(bool state, bool protectWindows, bool protectStructure, string password)
        {
            this.lockWindowsIfProtected = protectWindows;
            this.lockStructureIfProtected = protectStructure;
            this.workbookProtectionPassword = password;
            if (protectWindows == false && protectStructure == false)
            {
                this.UseWorkbookProtection = false;
            }
            else
            {
                this.UseWorkbookProtection = state;
            }
        }

        /// <summary>
        /// Adds a style to the style sheet of the workbook
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <param name="distinct">If true, the passed style will be replaced by an identical style if existing. Otherwise an exception will be thrown in case of a duplicate</param>
        /// <returns>Returns the added style. In case of an existing style, the distinct style will be returned</returns>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if the style already exists and parameter 'distinct' is set to false</exception>
        public Style AddStyle(Style style, bool distinct)
        {
            bool styleExits = false;
            bool identicalStyle = false;
            Style s;
            for (int i = 0; i < this.styles.Count; i++)
            {
                if (this.styles[i].Name == style.Name)
                {
                    if (this.styles[i].Equals(style) && distinct == true)
                    {
                        identicalStyle = true;
                        s = this.styles[i];
                    }
                    styleExits = true;
                    break;
                }
            }
            if (styleExits == true)
            {
                if (distinct == false && identicalStyle == false)
                {
                    throw new UndefinedStyleException("The style with the name '" + style.Name + "' already exits");
                }
                else
                {
                    s = style;
                }
            }
            else
            {
                s = style;
                this.styles.Add(style);
            }
            return s;
        }

        /// <summary>
        /// Removes the passed style from the style sheet
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(Style style)
        {
            RemoveStyle(style, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(string styleName)
        {
            RemoveStyle(styleName, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(Style style, bool onlyIfUnused)
        {
            if (style == null)
            {
                throw new UndefinedStyleException("The style to remove is not defined");
            }
            RemoveStyle(style.Name, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(string styleName, bool onlyIfUnused)
        {
            if (string.IsNullOrEmpty(styleName))
            {
                throw new UndefinedStyleException("The style to remove is not defined (no name specified)");
            }
            int index = -1;
            for(int i = 0; i < this.styles.Count; i++)
            {
                if (this.styles[i].Name == styleName)
                {
                    index = i;
                    break;
                }
            }
            if (index < 0)
            {
                throw new UndefinedStyleException("The style with the name '" + styleName + "' to remove was not found in the list of styles");
            }
            else if (this.styles[index].Name == "default" || index == 0)
            {
                throw new UndefinedStyleException("The default style can not be removed");
            }
            else
            {
                if (onlyIfUnused == true)
                {
                    bool styleInUse = false;
                    foreach(Worksheet sheet in this.worksheets)
                    {
                        foreach(KeyValuePair<string,Cell> cell in sheet.Cells)
                        {
                            if (cell.Value.CellStyle.Name == styleName)
                            {
                                styleInUse = true;
                                break;
                            }
                        }
                        if (styleInUse == true)
                        {
                            break;
                        }
                    }
                    if (styleInUse == false)
                    {
                        this.styles.RemoveAt(index);
                    }
                }
                else
                {
                    this.styles.RemoveAt(index);
                }
            }
        }

        /// <summary>
        /// Method to prepare the styles before saving the workbook. Don't use the method otherwise. Styles will be reordered and probably removed from the style sheet
        /// </summary>
        /// <param name="borders">Out parameter for a sorted list of Style.Border objects</param>
        /// <param name="fills">Out parameter for a sorted list of Style.Fill objects</param>
        /// <param name="fonts">Out parameter for a sorted list of Style.Font objects</param>
        /// <param name="numberFormats">Out parameter for a sorted list of Style.NumberFormat objects</param>
        /// <param name="cellXfs">Out parameter for a sorted list of Style.CellXf objects</param>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if one of the passed style components cannot be referenced or is null</exception>
        /// <remarks>This method is for internal use but must be public. Otherwise it's not possible to access it from low level methods. Don't use it</remarks>
        public void ReorganizeStyles(out  List<Style.Border> borders, out List<Style.Fill> fills, out List<Style.Font> fonts, out List<Style.NumberFormat> numberFormats, out List<Style.CellXf> cellXfs)
        {
            List<Style.Border> tempBorders = new List<Style.Border>();
            List<Style.Fill> tempFills = new List<Style.Fill>();
            List<Style.Font> tempFonts = new List<Style.Font>();
            List<Style.NumberFormat> tempNumberFormats = new List<Style.NumberFormat>();
            List<Style.CellXf> tempCellXfs = new List<Style.CellXf>();
            Style dateStyle = AddStyle(Style.BasicStyles.DateFormat, true);
            int existingIndex = 0;
            bool existing;
            int customNumberFormat = Style.NumberFormat.CUSTOMFORMAT_START_NUMBER;
            for(int i = 0; i < this.styles.Count; i++)
            {
                this.styles[i].InternalID = i;
                existing = false;
                foreach(Style.Border item in tempBorders)
                {
                    if (item.Equals(this.styles[i].CurrentBorder) == true)
                    { 
                        existing = true;
                        existingIndex = item.InternalID;
                        break;
                    }
                }
                if (existing == false)
                {
                    this.styles[i].CurrentBorder.InternalID = tempBorders.Count;
                    tempBorders.Add(this.styles[i].CurrentBorder);
                }
                else
                {
                    this.styles[i].CurrentBorder.InternalID = existingIndex;
                }
                existing = false;
                foreach (Style.Fill item in tempFills)
                {
                    if (item.Equals(this.styles[i].CurrentFill) == true)
                    {
                        existing = true;
                        existingIndex = item.InternalID;
                        break;
                    }
                }
                if (existing == false)
                {
                    this.styles[i].CurrentFill.InternalID = tempFills.Count;
                    tempFills.Add(this.styles[i].CurrentFill);
                }
                else
                {
                    this.styles[i].CurrentFill.InternalID = existingIndex;
                }
                existing = false;
                foreach (Style.Font item in tempFonts)
                {
                    if (item.Equals(this.styles[i].CurrentFont) == true)
                    {
                        existing = true;
                        existingIndex = item.InternalID;
                        break;
                    }
                }
                if (existing == false)
                {
                    this.styles[i].CurrentFont.InternalID = tempFonts.Count;
                    tempFonts.Add(this.styles[i].CurrentFont);
                }
                else
                {
                    this.styles[i].CurrentFont.InternalID = existingIndex;
                }
                existing = false;
                foreach (Style.NumberFormat item in tempNumberFormats)
                {
                    if (item.Equals(this.styles[i].CurrentNumberFormat) == true)
                    {
                        existing = true;
                        existingIndex = item.InternalID;
                        break;
                    }
                }
                if (existing == false)
                {
                    this.styles[i].CurrentNumberFormat.InternalID = tempNumberFormats.Count;
                    tempNumberFormats.Add(this.styles[i].CurrentNumberFormat);
                }
                else
                {
                    this.styles[i].CurrentNumberFormat.InternalID = existingIndex;
                }
                if (this.styles[i].CurrentNumberFormat.IsCustomFormat == true)
                {
                    this.styles[i].CurrentNumberFormat.CustomFormatID = customNumberFormat;
                    customNumberFormat++;
                }
                existing = false;
                foreach (Style.CellXf item in tempCellXfs)
                {
                    if (item.Equals(this.styles[i].CurrentCellXf) == true)
                    {
                        existing = true;
                        existingIndex = item.InternalID;
                        break;
                    }
                }
                if (existing == false)
                {
                    this.styles[i].CurrentCellXf.InternalID = tempCellXfs.Count;
                    tempCellXfs.Add(this.styles[i].CurrentCellXf);
                }
                else
                {
                    this.styles[i].CurrentCellXf.InternalID = existingIndex;
                }
            }
            Style combiation;
            foreach(Worksheet sheet in this.worksheets)
            {
                foreach(KeyValuePair<string, Cell> cell in sheet.Cells)
                {
                    if (cell.Value.Fieldtype == PicoXLSX.Cell.CellType.DATE)
                    {
                        if (cell.Value.CellStyle == null)
                        {
                            combiation = dateStyle;
                        }
                        else
                        {
                            combiation = cell.Value.CellStyle.Copy(dateStyle.CurrentNumberFormat);
                        }
                        sheet.Cells[cell.Key].SetStyle(combiation, this);
                    }
                }
            }

            this.styles.Sort();
            tempBorders.Sort();
            tempFills.Sort();
            tempFonts.Sort();
            tempNumberFormats.Sort();
            tempCellXfs.Sort();
            borders = tempBorders;
            fonts = tempFonts;
            fills = tempFills;
            cellXfs = tempCellXfs;
            numberFormats = tempNumberFormats;

        }

        /// <summary>
        /// Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.
        /// </summary>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if one of the styles of the merged cells cannot be referenced or is null</exception>
        public void ResolveMergedCells()
        {
            Style mergStyle = Style.BasicStyles.MergeCellStyle;
            int pos;
            List<Cell.Address> addresses;
            Cell cell;
            foreach(Worksheet sheet in this.worksheets)
            {
                foreach(KeyValuePair<string, Cell.Range> range in sheet.MergedCells)
                {
                    pos = 0;
                    addresses = Cell.GetCellRange(range.Value.StartAddress, range.Value.EndAddress);
                    foreach(Cell.Address address in addresses)
                    {
                        if (sheet.Cells.ContainsKey(address.ToString()) == false)
                        {
                            cell = new Cell();
                            cell.Fieldtype = Cell.CellType.EMPTY;
                            cell.RowAddress = address.Row;
                            cell.ColumnAddress = address.Column;
                            sheet.AddCell(cell);
                        }
                        else
                        {
                            cell = sheet.Cells[address.ToString()];
                        }
                        if (pos != 0)
                        {
                            cell.Fieldtype = Cell.CellType.EMPTY;
                        }
                        cell.SetStyle(mergStyle, this);
                        pos++;
                    }

                }
            }
        }

        
        /// <summary>
        /// Saves the workbook
        /// </summary>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="OutOfRangeException">Throws an OutOfRangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void Save()
        {
            LowLevel l = new LowLevel(this);
            l.Save();
        }

        /// <summary>
        /// Saves the workbook with the defined name
        /// </summary>
        /// <param name="filename">filename of the saved workbook</param>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="OutOfRangeException">Throws an OutOfRangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void SaveAs(string filename)
        {
            string backup = this.filename;
            this.filename = filename;
            LowLevel l = new LowLevel(this);
            l.Save();
            this.filename = backup;
        }

    }
}

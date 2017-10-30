/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;

namespace PicoXLSX
{

    /// <summary>
    /// PicoXLSX is a library to generate XLSX files in an easy and native way
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
#region privateFields
        private string filename;
        private List<Worksheet> worksheets;
        private Worksheet currentWorksheet;
        private StyleManager styleManager;
        private Metadata workbookMetadata;
        private string workbookProtectionPassword;
        private bool lockWindowsIfProtected;
        private bool lockStructureIfProtected;
        private int selectedWorksheet;
#endregion

#region properties
        /// <summary>
        /// Gets the current worksheet
        /// </summary>
        public Worksheet CurrentWorksheet
        {
            get { return currentWorksheet; }
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
        /// Gets whether the structure are locked if workbook is protected
        /// </summary>
        /// <see cref="SetWorkbookProtection"/>
        public bool LockStructureIfProtected
        {
            get { return lockStructureIfProtected; }
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
        /// Meta data object of the workbook
        /// </summary>
        public Metadata WorkbookMetadata
        {
            get { return workbookMetadata; }
            set { workbookMetadata = value; }
        }

        /// <summary>
        /// Gets the selected worksheet. The selected worksheet is not the current worksheet while design time but the selected sheet in the output file
        /// </summary>
        public int SelectedWorksheet
        {
            get { return selectedWorksheet; }
        }

        /// <summary>
        /// Gets the style manager of this workbook
        /// </summary>
        public StyleManager Styles
        {
            get { return styleManager; }
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
        /// Gets the list of worksheets in the workbook
        /// </summary>
        public List<Worksheet> Worksheets
        {
            get { return worksheets; }
        }
#endregion

#region constructors
        /// <summary>
        /// Default Constructor with additional parameter to create a default worksheet
        /// </summary>
        /// <param name="createWorkSheet">If true, a default worksheet will be crated and set as default worksheet</param>
        public Workbook(bool createWorkSheet)
        {
            Init();
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
        public Workbook(string filename, string sheetName)
        {
            Init();
            this.filename = filename;
            AddWorksheet(sheetName);
        }

        /// <summary>
        /// Constructor with filename ant the name of the first worksheet
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="sheetName">Name of the first worksheet</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>
        public Workbook(string filename, string sheetName, bool sanitizeSheetName)
        {
            Init();
            this.filename = filename;
            AddWorksheet(Worksheet.SanitizeWorksheetName(sheetName, this));
        }


#endregion

#region methods



        /// <summary>
        /// Adds a style to the style manager
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Returns the managed style of the style manager</returns>
        
        public Style AddStyle(Style style)
        {
            return this.styleManager.AddStyle(style);
        }

        /// <summary>
        /// Adds a style component to a style
        /// </summary>
        /// <param name="baseStyle">Style to append a component</param>
        /// <param name="newComponent">Component to add to the baseStyle</param>
        /// <returns>Returns the managed style of the style manager</returns>
        public Style addStyleComponent(Style baseStyle, AbstractStyle newComponent)
        {
        
            if (newComponent.GetType() == typeof(Style.Border))
            {
                baseStyle.BorderStyle = (Style.Border)newComponent;
            }
            else if (newComponent.GetType() == typeof(Style.CellXf))
            {
                baseStyle.CellXfStyle = (Style.CellXf)newComponent;
            }
            else if (newComponent.GetType() == typeof(Style.Fill))
            {
                baseStyle.FillStyle = (Style.Fill)newComponent;
            }
            else if (newComponent.GetType() == typeof(Style.Font))
            {
                baseStyle.FontStyle = (Style.Font)newComponent;
            }
            else if (newComponent.GetType() == typeof(Style.NumberFormat))
            {
                baseStyle.NumberFormatStyle = (Style.NumberFormat)newComponent;
            }
            return this.styleManager.AddStyle(baseStyle);
        }


        /// <summary>
        /// Adding a new Worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetNameAlreadxExistsException if the name of the worksheet already exists</exception>
        /// <exception cref="FormatException">Throws a FormatException if the name contains illegal characters or is out of range (length between 1 an 31 characters)</exception>
        public void AddWorksheet(string name)
        {
            foreach (Worksheet item in this.worksheets)
            {
                if (item.SheetName == name)
                {
                    throw new WorksheetException("WorksheetNameAlreadxExistsException", "The worksheet with the name '" + name + "' already exists.");
                }
            }
            int number = this.worksheets.Count + 1;
            Worksheet newWs = new Worksheet(name, number, this);
            this.currentWorksheet = newWs;
            this.worksheets.Add(newWs);
        }

        /// <summary>
        /// Adding a new Worksheet with a sanitizing option
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists and sanitizeSheetName is false</exception>
        /// <exception cref="FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitizeSheetName is false</exception>
        public void AddWorksheet(String name, bool sanitizeSheetName)
        {
            if (sanitizeSheetName == true)
            {
                String sanitized = Worksheet.SanitizeWorksheetName(name, this);
                AddWorksheet(sanitized);
            }
            else
            {
                AddWorksheet(name);
            }
        }

        /// <summary>
        /// Adding a new Worksheet
        /// </summary>
        /// <param name="worksheet">Prepared worksheet object</param>
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists</exception>
        /// <exception cref="FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31</exception>
        public void AddWorksheet(Worksheet worksheet)
        {
            for (int i = 0; i < this.worksheets.Count; i++)
            {
                if (this.worksheets[i].SheetName == worksheet.SheetName)
                {
                    throw new WorksheetException("WorksheetNameAlreadxExistsException", "The worksheet with the name '" + worksheet.SheetName + "' already exists.");
                }
            }
            int number = this.worksheets.Count+ 1;
            worksheet.SheetID = number;
            worksheet.WorkbookReference = this;
            this.currentWorksheet = worksheet;
            this.worksheets.Add(worksheet);
        }

        /// <summary>
        /// Init method called in the constructors
        /// </summary>
        private void Init()
        {
            this.worksheets = new List<Worksheet>();
            this.styleManager = new StyleManager();
            this.styleManager.AddStyle(new Style("default", 0, true));
            Style borderStyle = new Style("default_border_style", 1, true);
            borderStyle.BorderStyle = Style.BasicStyles.DottedFill_0_125.BorderStyle;
            borderStyle.FillStyle = Style.BasicStyles.DottedFill_0_125.FillStyle;
            this.styleManager.AddStyle(borderStyle);
            this.workbookMetadata = new Metadata();
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
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(Style style, bool onlyIfUnused)
        {
            if (style == null)
            {
                throw new StyleException("UndefinedStyleException", "The style to remove is not defined");
            }
            RemoveStyle(style.Name, onlyIfUnused);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(string styleName, bool onlyIfUnused)
        {
            if (string.IsNullOrEmpty(styleName))
            {
                throw new StyleException("MissingReferenceException", "The style to remove is not defined (no name specified)");
            }
            if (onlyIfUnused == true)
            {
                    bool styleInUse = false;
                    for(int i = 0; i < this.worksheets.Count; i++)
                    {
                        foreach(KeyValuePair<string,Cell> cell in this.worksheets[i].Cells)
                        {
                            if (cell.Value.CellStyle == null) { continue; }
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
                        this.styleManager.RemoveStyle(styleName);
                    }
            }
            else
            {
                this.styleManager.RemoveStyle(styleName);
            }
        }

        /// <summary>
        /// Removes the defined worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <exception cref="WorksheetException">Throws a UnknownWorksheetException if the name of the worksheet is unknown</exception>
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
                throw new WorksheetException("UnknownWorksheetException", "The worksheet with the name '" + name + "' does not exist.");
            }
            if (this.worksheets[index].SheetName == this.currentWorksheet.SheetName)
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
        /// Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.
        /// </summary>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if one of the styles of the merged cells cannot be referenced or is null</exception>
        public void ResolveMergedCells()
        {
            Style mergeStyle = Style.BasicStyles.MergeCellStyle;
            int pos;
            List<Cell.Address> addresses;
            Cell cell;
            foreach (Worksheet sheet in this.worksheets)
            {
                foreach (KeyValuePair<string, Cell.Range> range in sheet.MergedCells)
                {
                    pos = 0;
                    addresses = Cell.GetCellRange(range.Value.StartAddress, range.Value.EndAddress);
                    foreach (Cell.Address address in addresses)
                    {
                        if (sheet.Cells.ContainsKey(address.ToString()) == false)
                        {
                            cell = new Cell();
                            cell.DataType = Cell.CellType.EMPTY;
                            cell.RowAddress = address.Row;
                            cell.ColumnAddress = address.Column;
                            cell.WorksheetReference = sheet;
                            sheet.AddCell(cell);
                        }
                        else
                        {
                            cell = sheet.Cells[address.ToString()];
                        }
                        if (pos != 0)
                        {
                            cell.DataType = Cell.CellType.EMPTY;
                        }
                        cell.SetStyle(mergeStyle);
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

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>Returns the current worksheet</returns>
        /// <exception cref="WorksheetException">Throws a MissingReferenceException if the name of the worksheet is unknown</exception>
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
                throw new WorksheetException("MissingReferenceException", "The worksheet with the name '" + name + "' does not exist.");
            }
            return this.currentWorksheet;
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheetIndex">Zero-based worksheet index</param>
        /// <exception cref="RangeException">Throws a OutOfRangeException if the index of the worksheet is out of range</exception>
        public void SetSelectedWorksheet(int worksheetIndex)
        {
            if (worksheetIndex < 0 || worksheetIndex > this.worksheets.Count - 1)
            {
                throw new RangeException("OutOfRangeException","The worksheet index " + worksheetIndex.ToString() + " is out of range");
            }
            this.selectedWorksheet = worksheetIndex;
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
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheet">Worksheet object (must be in the collection of worksheets)</param>
        /// <exception cref="WorksheetException">Throws a UnknownWorksheetException if the worksheet was not found in the worksheet collection</exception>
        public void SetSelectedWorksheet(Worksheet worksheet)
        {
            bool check = false;
            for (int i = 0; i < this.worksheets.Count; i++)
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
                throw new WorksheetException("UnknownWorksheetException", "The passed worksheet object is not in the worksheet collection.");
            }
        }

#endregion
    }
}

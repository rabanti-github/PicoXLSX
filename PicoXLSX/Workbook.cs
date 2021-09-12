/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static PicoXLSX.Style;

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
        private Metadata workbookMetadata;
        private string workbookProtectionPassword;
        private bool lockWindowsIfProtected;
        private bool lockStructureIfProtected;
        private int selectedWorksheet;
        private Shortener shortener;
        #endregion

        #region properties


        /// <summary>
        /// Gets the shortener object for the current worksheet
        /// </summary>
        public Shortener WS
        {
            get { return shortener; }
        }


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
        /// <remarks>
        /// Note that the file name is not sanitized. If a filename is set that is not compliant to the file system, saving of the workbook may fail
        /// </remarks>
        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }

        /// <summary>
        /// Gets whether the structure are locked if workbook is protected. See also <see cref="SetWorkbookProtection"/>
        /// </summary>
        public bool LockStructureIfProtected
        {
            get { return lockStructureIfProtected; }
        }

        /// <summary>
        /// Gets whether the windows are locked if workbook is protected. See also <see cref="SetWorkbookProtection"/>
        /// </summary> 
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
        /// Gets or sets whether the workbook is protected
        /// </summary>
        public bool UseWorkbookProtection { get; set; }

        /// <summary>
        /// Gets the password used for workbook protection. See also <see cref="SetWorkbookProtection"/>
        /// </summary>
        /// <remarks>The password of this property is stored in plan text. Encryption is performed when the workbook is saved</remarks>
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

        /// <summary>
        /// Gets or sets whether the whole workbook is hidden
        /// </summary>
        /// <remarks>A hidden workbook can only be made visible, using another, already visible Excel window</remarks>
        public bool Hidden { get; set; }

        #endregion

        #region constructors
        /// <summary>
        /// Default constructor. No initial worksheet is created. Use <see cref="AddWorksheet(string)"/> (or overloads) to add one
        /// </summary>
        public Workbook()
        {
            Init();
        }

        /// <summary>
        /// Constructor with additional parameter to create a default worksheet. This constructor can be used to define a workbook that is saved as stream
        /// </summary>
        /// <param name="createWorkSheet">If true, a default worksheet with the name 'Sheet1' will be crated and set as current worksheet</param>
        public Workbook(bool createWorkSheet)
        {
            Init();
            if (createWorkSheet)
            {
                AddWorksheet("Sheet1");
            }
        }

        /// <summary>
        /// Constructor with additional parameter to create a default worksheet with the specified name. This constructor can be used to define a workbook that is saved as stream
        /// </summary>
        /// <param name="sheetName">Filename of the workbook.  The name will be sanitized automatically according to the specifications of Excel</param>
        public Workbook(string sheetName)
        {
            Init();
            AddWorksheet(sheetName, true);
        }

        /// <summary>
        /// Constructor with filename ant the name of the first worksheet
        /// </summary>
        /// <param name="filename">Filename of the workbook.  The name will be sanitized automatically according to the specifications of Excel</param>
        /// <param name="sheetName">Name of the first worksheet. The name will be sanitized automatically according to the specifications of Excel</param>
        public Workbook(string filename, string sheetName)
        {
            Init();
            this.filename = filename;
            AddWorksheet(sheetName, true);
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
            if (sanitizeSheetName)
            {
                AddWorksheet(Worksheet.SanitizeWorksheetName(sheetName, this));
            }
            else
            {
                AddWorksheet(sheetName);
            }
        }

        #endregion

        #region methods

        /// <summary>
        /// Adds a style to the style repository. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Returns the managed style of the style repository</returns>
        /// 
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public Style AddStyle(Style style)
        {
            return StyleRepository.Instance.AddStyle(style);
        }

        /// <summary>
        /// Adds a style component to a style. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="baseStyle">Style to append a component</param>
        /// <param name="newComponent">Component to add to the baseStyle</param>
        /// <returns>Returns the modified style of the style repository</returns>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public Style AddStyleComponent(Style baseStyle, AbstractStyle newComponent)
        {

            if (newComponent.GetType() == typeof(Border))
            {
                baseStyle.CurrentBorder = (Border)newComponent;
            }
            else if (newComponent.GetType() == typeof(CellXf))
            {
                baseStyle.CurrentCellXf = (CellXf)newComponent;
            }
            else if (newComponent.GetType() == typeof(Fill))
            {
                baseStyle.CurrentFill = (Fill)newComponent;
            }
            else if (newComponent.GetType() == typeof(Font))
            {
                baseStyle.CurrentFont = (Font)newComponent;
            }
            else if (newComponent.GetType() == typeof(NumberFormat))
            {
                baseStyle.CurrentNumberFormat = (NumberFormat)newComponent;
            }
            return StyleRepository.Instance.AddStyle(baseStyle);
        }


        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet already exists</exception>
        /// <exception cref="FormatException">Throws a FormatException if the name contains illegal characters or is out of range (length between 1 an 31 characters)</exception>
        public void AddWorksheet(string name)
        {
            foreach (Worksheet item in worksheets)
            {
                if (item.SheetName == name)
                {
                    throw new WorksheetException("The worksheet with the name '" + name + "' already exists.");
                }
            }
            int number = GetNextWorksheetId();
            Worksheet newWs = new Worksheet(name, number, this);
            currentWorksheet = newWs;
            worksheets.Add(newWs);
            shortener.SetCurrentWorksheetInternal(currentWorksheet);
        }

        /// <summary>
        /// Adding a new Worksheet with a sanitizing option. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists and sanitizeSheetName is false</exception>
        /// <exception cref="FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitizeSheetName is false</exception>
        public void AddWorksheet(String name, bool sanitizeSheetName)
        {
            if (sanitizeSheetName)
            {
                string sanitized = Worksheet.SanitizeWorksheetName(name, this);
                AddWorksheet(sanitized);
            }
            else
            {
                AddWorksheet(name);
            }
        }

        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="worksheet">Prepared worksheet object</param>
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists</exception>
        /// <exception cref="FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31)</exception>
        public void AddWorksheet(Worksheet worksheet)
        {
            AddWorksheet(worksheet, false);
        }

        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="worksheet">Prepared worksheet object</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>    
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists, when sanitation is false</exception>
        /// <exception cref="FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitation is false</exception>
        public void AddWorksheet(Worksheet worksheet, bool sanitizeSheetName)
        {
            if (sanitizeSheetName)
            {
                string name = Worksheet.SanitizeWorksheetName(worksheet.SheetName, this);
                worksheet.SheetName = name;
            }
            else
            {
                if (string.IsNullOrEmpty(worksheet.SheetName))
                {
                    throw new WorksheetException("The name of the passed worksheet is null or empty.");
                }
                for (int i = 0; i < worksheets.Count; i++)
                {
                    if (worksheets[i].SheetName == worksheet.SheetName)
                    {
                        throw new WorksheetException("The worksheet with the name '" + worksheet.SheetName + "' already exists.");
                    }
                }
            }
            worksheet.SheetID = GetNextWorksheetId();
            currentWorksheet = worksheet;
            worksheets.Add(worksheet);
            worksheet.WorkbookReference = this;
        }

        /// <summary>
        /// Removes the passed style from the style sheet. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(Style style)
        {
            RemoveStyle(style, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(string styleName)
        {
            RemoveStyle(styleName, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(Style style, bool onlyIfUnused)
        {
            if (style == null)
            {
                throw new StyleException("MissingReferenceException", "The style to remove is not defined");
            }
            RemoveStyle(style.Name, onlyIfUnused);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(string styleName, bool onlyIfUnused)
        {
            if (string.IsNullOrEmpty(styleName))
            {
                throw new StyleException("MissingReferenceException", "The style to remove is not defined (no name specified)");
            }
            // noOp / deprecated
        }

        /// <summary>
        /// Removes the defined worksheet based on its name. If the worksheet is the current or selected worksheet, the current and / or the selected worksheet will be set to the last worksheet of the workbook.
        /// If the last worksheet is removed, the selected worksheet will be set to 0 and the current worksheet to null.
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        public void RemoveWorksheet(string name)
        {
            Worksheet worksheetToRemove = worksheets.FindLast(w => w.SheetName == name);
            if (worksheetToRemove == null)
            {
                throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
            }
            int index = worksheets.IndexOf(worksheetToRemove);
            bool resetCurrentWorksheet = worksheetToRemove == currentWorksheet;
            RemoveWorksheet(index, resetCurrentWorksheet);
        }

        /// <summary>
        /// Removes the defined worksheet based on its index. If the worksheet is the current or selected worksheet, the current and / or the selected worksheet will be set to the last worksheet of the workbook.
        /// If the last worksheet is removed, the selected worksheet will be set to 0 and the current worksheet to null.
        /// </summary>
        /// <param name="index">Index within the worksheets list</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the index is out of range</exception>

        public void RemoveWorksheet(int index)
        {
            if (index < 0 || index >= worksheets.Count)
            {
                throw new WorksheetException("The worksheet index " + index + " is out of range");
            }
            bool resetCurrentWorksheet = worksheets[index] == currentWorksheet;
            RemoveWorksheet(index, resetCurrentWorksheet);
        }

        /// <summary>
        /// Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.<br/>
        /// This is an internal method. There is no need to use it
        /// </summary>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the merged cells cannot be referenced or is null</exception>
        internal void ResolveMergedCells()
        {
            foreach (Worksheet worksheet in worksheets)
            {
                worksheet.ResolveMergedCells();
            }
        }

        /// <summary>
        /// Saves the workbook
        /// </summary>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void Save()
        {
            LowLevel l = new LowLevel(this);
            l.Save();
        }

        /// <summary>
        /// Saves the workbook asynchronous.
        /// </summary>
        /// <returns>Task object (void)</returns>
        /// <exception cref="IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="StyleException">May throw a StyleException if one of the styles of the workbook cannot be referenced or is null. The asynchronous operation may hide the exception.</exception>
        public async Task SaveAsync()
        {
            LowLevel l = new LowLevel(this);
            await l.SaveAsync();
        }

        /// <summary>
        /// Saves the workbook with the defined name
        /// </summary>
        /// <param name="fileName">filename of the saved workbook</param>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void SaveAs(string fileName)
        {
            string backup = fileName;
            filename = fileName;
            LowLevel l = new LowLevel(this);
            l.Save();
            filename = backup;
        }

        /// <summary>
        /// Saves the workbook with the defined name asynchronous.
        /// </summary>
        /// <param name="fileName">filename of the saved workbook</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="StyleException">May throw a StyleException if one of the styles of the workbook cannot be referenced or is null. The asynchronous operation may hide the exception.</exception>
        public async Task SaveAsAsync(string fileName)
        {
            string backup = fileName;
            filename = fileName;
            LowLevel l = new LowLevel(this);
            await l.SaveAsync();
            filename = backup;
        }

        /// <summary>
        /// Save the workbook to a writable stream
        /// </summary>
        /// <param name="stream">Writable stream</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void SaveAsStream(Stream stream, bool leaveOpen = false)
        {
            LowLevel l = new LowLevel(this);
            l.SaveAsStream(stream, leaveOpen);
        }

        /// <summary>
        /// Save the workbook to a writable stream asynchronous.
        /// </summary>
        /// <param name="stream">>Writable stream</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="IOException">Throws IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="StyleException">May throw a StyleException if one of the styles of the workbook cannot be referenced or is null. The asynchronous operation may hide the exception.</exception>
        public async Task SaveAsStreamAsync(Stream stream, bool leaveOpen = false)
        {
            LowLevel l = new LowLevel(this);
            await l.SaveAsStreamAsync(stream, leaveOpen);
        }

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>Returns the current worksheet</returns>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        public Worksheet SetCurrentWorksheet(string name)
        {
            currentWorksheet = worksheets.FirstOrDefault(w => w.SheetName == name);
            if (currentWorksheet == null)
            {
                throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
            }
            shortener.SetCurrentWorksheetInternal(currentWorksheet);
            return currentWorksheet;
        }

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="worksheetIndex">Zero-based worksheet index</param>
        /// <returns>Returns the current worksheet</returns>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        public Worksheet SetCurrentWorksheet(int worksheetIndex)
        {
            if (worksheetIndex < 0 || worksheetIndex > worksheets.Count - 1)
            {
                throw new RangeException("OutOfRangeException", "The worksheet index " + worksheetIndex + " is out of range");
            }
            currentWorksheet = worksheets[worksheetIndex];
            shortener.SetCurrentWorksheetInternal(currentWorksheet);
            return currentWorksheet;
        }

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="worksheet">Worksheet object (must be in the collection of worksheets)</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet was not found in the worksheet collection</exception>
        public void SetCurrentWorksheet(Worksheet worksheet)
        {
            int index = worksheets.IndexOf(worksheet);
            if (index < 0)
            {
                throw new WorksheetException("The passed worksheet object is not in the worksheet collection.");
            }
            currentWorksheet = worksheets[index];
            shortener.SetCurrentWorksheetInternal(worksheet);
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet to be set selected is hidden</exception>
        public void SetSelectedWorksheet(string name)
        {
            selectedWorksheet = worksheets.FindIndex(w => w.SheetName == name);
            if (selectedWorksheet < 0)
            {
                throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
            }
            ValidateWorksheets();
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheetIndex">Zero-based worksheet index</param>
        /// <exception cref="RangeException">Throws a RangeException if the index of the worksheet is out of range</exception>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet to be set selected is hidden</exception>
        public void SetSelectedWorksheet(int worksheetIndex)
        {
            if (worksheetIndex < 0 || worksheetIndex > worksheets.Count - 1)
            {
                throw new RangeException("OutOfRangeException", "The worksheet index " + worksheetIndex + " is out of range");
            }
            selectedWorksheet = worksheetIndex;
            ValidateWorksheets();
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheet">Worksheet object (must be in the collection of worksheets)</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet was not found in the worksheet collection</exception>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet to be set selected is hidden</exception>
        public void SetSelectedWorksheet(Worksheet worksheet)
        {
            selectedWorksheet = worksheets.IndexOf(worksheet);
            if (selectedWorksheet < 0)
            {
                throw new WorksheetException("The passed worksheet object is not in the worksheet collection.");
            }
            ValidateWorksheets();
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
            lockWindowsIfProtected = protectWindows;
            lockStructureIfProtected = protectStructure;
            workbookProtectionPassword = password;
            if (protectWindows == false && protectStructure == false)
            {
                UseWorkbookProtection = false;
            }
            else
            {
                UseWorkbookProtection = state;
            }
        }

        /// <summary>
        /// Removes the worksheet at the defined index and relocates current and selected worksheet references
        /// </summary>
        /// <param name="index">Index within the worksheets list</param>
        /// <param name="resetCurrentWorksheet">If true, the current worksheet will be relocated to the last worksheet in the list</param>
        private void RemoveWorksheet(int index, bool resetCurrentWorksheet)
        {
            worksheets.RemoveAt(index);
            if (worksheets.Count > 0)
            {
                for (int i = 0; i < worksheets.Count; i++)
                {
                    worksheets[i].SheetID = i + 1;
                }
                if (resetCurrentWorksheet)
                {
                    currentWorksheet = worksheets[worksheets.Count - 1];
                }
                if (selectedWorksheet == index || selectedWorksheet > worksheets.Count - 1)
                {
                    selectedWorksheet = worksheets.Count - 1;
                }
            }
            else
            {
                currentWorksheet = null;
                selectedWorksheet = 0;
            }
            ValidateWorksheets();
        }

        /// <summary>
        /// Validates the worksheets regarding several conditions that must be met:<br/>
        /// - At least one worksheet must be defined<br/>
        /// - A hidden worksheet cannot be the selected one<br/>
        /// - At least one worksheet must be visible<br/>
        /// If one of the conditions is not met, an exception is thrown
        /// </summary>
        internal void ValidateWorksheets()
        {
            int woksheetCount = worksheets.Count;
            if (woksheetCount == 0)
            {
                throw new WorksheetException("The workbook must contain at least one worksheet");
            }
            for (int i = 0; i < woksheetCount; i++)
            {
                if (worksheets[i].Hidden)
                {
                    if (i == selectedWorksheet)
                    {
                        throw new WorksheetException("The worksheet with the index " + selectedWorksheet + " cannot be set as selected, since it is set hidden");
                    }
                }
            }
        }

        /// <summary>
        /// Gets the next free worksheet ID
        /// </summary>
        /// <returns>Worksheet ID</returns>
        private int GetNextWorksheetId()
        {
            if (worksheets.Count == 0)
            {
                return 1;
            }
            return worksheets.Max(w => w.SheetID) + 1;
        }

        /// <summary>
        /// Init method called in the constructors
        /// </summary>
        private void Init()
        {
            worksheets = new List<Worksheet>();
            workbookMetadata = new Metadata();
            shortener = new Shortener(this);
        }

        #endregion

        #region sub-classes

        /// <summary>
        /// Class to provide access to the current worksheet with a shortened syntax. Note: The WS object can be null if the workbook was created without a worksheet. The object will be available as soon as the current worksheet is defined
        /// </summary>
        public class Shortener
        {
            private Worksheet currentWorksheet;
            private readonly Workbook workbookReference;

            /// <summary>
            /// Constructor with workbook reference
            /// </summary>
            /// <param name="reference">Workbook reference</param>
            public Shortener(Workbook reference)
            {
                this.workbookReference = reference;
                this.currentWorksheet = reference.CurrentWorksheet;
            }

            /// <summary>
            /// Sets the worksheet accessed by the shortener
            /// </summary>
            /// <param name="worksheet">Current worksheet</param>
            public void SetCurrentWorksheet(Worksheet worksheet)
            {
                workbookReference.SetCurrentWorksheet(worksheet);
                currentWorksheet = worksheet;
            }

            /// <summary>
            /// Sets the worksheet accessed by the shortener, invoked by the workbook
            /// </summary>
            /// <param name="worksheet">Current worksheet</param>
            internal void SetCurrentWorksheetInternal(Worksheet worksheet)
            {
                currentWorksheet = worksheet;
            }

            /// <summary>
            /// Sets a value into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
            /// </summary>
            /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
            /// <param name="value">Value to set</param>
            public void Value(object value)
            {
                NullCheck();
                currentWorksheet.AddNextCell(value);
            }

            /// <summary>
            /// Sets a value with style into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
            /// </summary>
            /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
            /// <param name="value">Value to set</param>
            /// <param name="style">Style to apply</param>
            public void Value(object value, Style style)
            {
                NullCheck();
                currentWorksheet.AddNextCell(value, style);
            }

            /// <summary>
            /// Sets a formula into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
            /// </summary>
            /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
            /// <param name="formula">Formula to set</param>
            public void Formula(string formula)
            {
                NullCheck();
                currentWorksheet.AddNextCellFormula(formula);
            }

            /// <summary>
            /// Sets a formula with style into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
            /// </summary>
            /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
            /// <param name="formula">Formula to set</param>
            /// <param name="style">Style to apply</param>
            public void Formula(string formula, Style style)
            {
                NullCheck();
                currentWorksheet.AddNextCellFormula(formula, style);
            }

            /// <summary>
            /// Moves the cursor one row down
            /// </summary>
            public void Down()
            {
                NullCheck();
                currentWorksheet.GoToNextRow();
            }

            /// <summary>
            /// Moves the cursor the number of defined rows down
            /// </summary>
            /// <param name="numberOfRows">Number of rows to move</param>
            /// <param name="keepColumnPosition">If true, the column position is preserved, otherwise set to 0</param>
            public void Down(int numberOfRows, bool keepColumnPosition = false)
            {
                NullCheck();
                currentWorksheet.GoToNextRow(numberOfRows, keepColumnPosition);
            }

            /// <summary>
            /// Moves the cursor one row up
            /// </summary>
            /// <remarks>An exception will be thrown if the row number is below 0/></remarks>
            public void Up()
            {
                NullCheck();
                currentWorksheet.GoToNextRow(-1);
            }

            /// <summary>
            /// Moves the cursor the number of defined rows up
            /// </summary>
            /// <param name="numberOfRows">Number of rows to move</param>
            /// <param name="keepColumnosition">If true, the column position is preserved, otherwise set to 0</param>
            /// <remarks>An exception will be thrown if the row number is below 0. Values can be also negative. However, this is the equivalent of the function <see cref="Down(int, bool)"/></remarks>
            public void Up(int numberOfRows, bool keepColumnosition = false)
            {
                NullCheck();
                currentWorksheet.GoToNextRow(-1 * numberOfRows, keepColumnosition);
            }

            /// <summary>
            /// Moves the cursor one column to the right
            /// </summary>
            public void Right()
            {
                NullCheck();
                currentWorksheet.GoToNextColumn();
            }

            /// <summary>
            /// Moves the cursor the number of defined columns to the right
            /// </summary>
            /// <param name="numberOfColumns">Number of columns to move</param>
            /// <param name="keepRowPosition">If true, the row position is preserved, otherwise set to 0</param>
            public void Right(int numberOfColumns, bool keepRowPosition = false)
            {
                NullCheck();
                currentWorksheet.GoToNextColumn(numberOfColumns, keepRowPosition);
            }

            /// <summary>
            /// Moves the cursor one column to the left
            /// </summary>
            /// <remarks>An exception will be thrown if the column number is below 0</remarks>
            public void Left()
            {
                NullCheck();
                currentWorksheet.GoToNextColumn(-1);
            }

            /// <summary>
            /// Moves the cursor the number of defined columns to the left
            /// </summary>
            /// <param name="numberOfColumns">Number of columns to move</param>
            /// <param name="keepRowRowPosition">If true, the row position is preserved, otherwise set to 0</param>
            /// <remarks>An exception will be thrown if the column number is below 0. Values can be also negative. However, this is the equivalent of the function <see cref="Right(int, bool)"/></remarks>
            public void Left(int numberOfColumns, bool keepRowRowPosition = false)
            {
                NullCheck();
                currentWorksheet.GoToNextColumn(-1 * numberOfColumns, keepRowRowPosition);
            }

            /// <summary>
            /// Internal method to check whether the worksheet is null
            /// </summary>
            private void NullCheck()
            {
                if (currentWorksheet == null)
                {
                    throw new WorksheetException("No worksheet was defined");
                }
            }


        }

        #endregion

    }
}

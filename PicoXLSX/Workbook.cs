/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
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
        /// Default constructor
        /// </summary>
        public Workbook()
        { 
            this.worksheets = new List<Worksheet>();
        }

        /// <summary>
        /// Constructor with filename ant the name of the first worksheet
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="sheetName">Name of the first worksheet</param>
        public Workbook(string filename, string sheetName) : this()
        {
            this.filename = filename;
            AddWorksheet(sheetName);
        }

        /// <summary>
        /// Adding a new Worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <exception cref="WorksheetNameAlreadxExistsException">Throws a WorksheetNameAlreadxExistsException if the name of the worksheet already exists</exception>
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
        /// <param name="name">Name ot the worksheet</param>
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
        }

        
        /// <summary>
        /// Saets the workbook
        /// </summary>
        /// <returns>Returns true if the save action was successful, otherweise false</returns>
        public bool Save()
        {
            LowLevel l = new LowLevel(this);
            return l.Save();
        }

        /// <summary>
        /// Saves the worksheet with the defined name
        /// </summary>
        /// <param name="filename">filename of the saved workbook</param>
        /// <returns>Returns true if the save action was successful, otherweise false</returns>
        public bool SaveAs(string filename)
        {
            string backup = this.filename;
            this.filename = filename;
            LowLevel l = new LowLevel(this);
            bool state = l.Save();
            this.filename = backup;
            return state;
        }

    }
}

﻿/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PicoXLSX;
using Testing;

namespace Demo.Testing
{
    /// <summary>
    /// Class for performance tests
    /// </summary>
    public static class Performance
    {
        /// <summary>
        /// Method to perform a stress test on PicoXLSX with a high amount of random data
        /// </summary>
        /// <param name="filename">filename of the output</param>
        /// <param name="sheetName">name of the worksheet</param>
        /// <param name="rows">Number of rows</param>
        /// <param name="cols">Number of columns</param>
        /// <remarks>The data type is determined per column randomly. In case of strings, random ASCII characters from 1 to 256 characters are written into the cells</remarks> 
        public static void StressTest(string filename, string sheetName, int rows, int cols)
        {
            System.Console.WriteLine("Starting performance test - Generating Array...");
            List<List<object>> field = new List<List<object>>();
            List<object> row;
            List<int> colTypes = new List<int>();
            DateTime min = new DateTime(1901, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
            DateTime max = new DateTime(2100, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
            int j;
            for (int i = 0; i < cols; i++)
            {
                colTypes.Add(Utils.PseudoRandomInteger(1, 6));
            }
            for (int i = 0; i < rows; i++)
            {
                row = new List<object>();
                for (j = 0; j < cols; j++)
                {
                    if (colTypes[j] == 1) { row.Add(Utils.PseduoRandomDate(min, max)); }
                    else if (colTypes[j] == 2) { row.Add(Utils.PseudoRandomBool()); }
                    else if (colTypes[j] == 3) { row.Add(Utils.PseudoRandomDouble(double.MinValue, double.MaxValue)); }
                    else if (colTypes[j] == 4) { row.Add(Utils.PseudoRandomInteger(int.MinValue, int.MaxValue)); }
                    else if (colTypes[j] == 5) { row.Add(Utils.PseudoRandomLong(long.MinValue, long.MaxValue)); }
                    else if (colTypes[j] == 6) { row.Add(Utils.PseudoRandomString(1, 256)); }
                }
                field.Add(row);
            }
            System.Console.WriteLine("Writing cells...");
            PicoXLSX.Workbook b = new PicoXLSX.Workbook(filename, sheetName);
            PicoXLSX.Worksheet s = b.CurrentWorksheet;
            s.CurrentCellDirection = PicoXLSX.Worksheet.CellDirection.ColumnToColumn;
            for (int i = 0; i < rows; i++)
            {
                for (j = 0; j < cols; j++)
                {
                    s.AddNextCell(field[i][j], Style.BasicStyles.Bold);
                }
                s.GoToNextRow();
            }
            System.Console.WriteLine("Saving workbook...");
            b.Save();
            System.Console.WriteLine("Workbook saved");
        }


    }
}

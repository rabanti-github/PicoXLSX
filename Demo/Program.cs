using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using PicoXLSX;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            Demo1();
            //Demo2();
            //Demo3();
        }

        private static void Demo1()
        {
            Workbook workbook = new Workbook("test1.xlsx", "Sheet1");   // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Test");              // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(123);                 // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(true);                // Add cell C1
            workbook.CurrentWorksheet.GoToNextRow();                    // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(123.456d);            // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(123.789f);            // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);        // Add cell C2
            workbook.CurrentWorksheet.GoToNextRow();                    // Go to Row 3
            workbook.CurrentWorksheet.AddNextCellFormula("B1*22");      // Add cell A3 as formula (B1 times 22)
            workbook.CurrentWorksheet.AddNextCellFormula("ROUNDDOWN(A2,1)"); // Add cell B3 as formula (Floor A2 with one decimal place)
            workbook.CurrentWorksheet.AddNextCellFormula("PI()");       // Add cell C3 as formula (Pi = 3.14.... )
            workbook.Save();                                            // Save the workbook
        }

        private static void Demo2()
        {
            Workbook workbook = new Workbook();                         // Create new workbook
            workbook.AddWorksheet("Sheet1");                            // Add a new Worksheet and set it as current sheet
            workbook.CurrentWorksheet.AddNextCell("月曜日");            // Add cell A1 (Unicode)
            workbook.CurrentWorksheet.AddNextCell(-987);                // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(false);               // Add cell C1
            workbook.CurrentWorksheet.GoToNextRow();                    // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(-123.456d);           // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(-123.789f);           // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);        // Add cell C3
            workbook.AddWorksheet("Sheet2");                            // Add a new Worksheet and set it as current sheet
            workbook.CurrentWorksheet.AddCell("ABC", "A1");             // Add cell A1
            workbook.CurrentWorksheet.AddCell(779, 2, 1);               // Add cell C2 (zero based addresses: column 2=C, row 1=2)
            workbook.CurrentWorksheet.AddCell(false, 3, 2);             // Add cell D3 (zero based addresses: column 3=D, row 2=3)
            workbook.CurrentWorksheet.AddNextCell(0);                   // Add cell E3 (direction: column to column)
            List<string> values = new List<string>() { "V1", "V2", "V3" }; // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, "A4:C4");    // Add a cell range to A4 - C4
            workbook.SaveAs("test2.xlsx");                              // Save the workbook
        }

        private static void Demo3()
        {
            Workbook workbook = new Workbook("test3.xlsx", "Sheet1");   // Create new workbook
            workbook.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow;  // Change the cell direction
            workbook.CurrentWorksheet.AddNextCell(1);                   // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(2);                   // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(3);                   // Add cell A3
            workbook.CurrentWorksheet.AddNextCell(4);                   // Add cell A4
            workbook.CurrentWorksheet.GoToNextColumn();                 // Go to Column B
            workbook.CurrentWorksheet.AddNextCell("A");                 // Add cell B1
            workbook.CurrentWorksheet.AddNextCell("B");                 // Add cell B2
            workbook.CurrentWorksheet.AddNextCell("C");                 // Add cell B3
            workbook.CurrentWorksheet.AddNextCell("D");                 // Add cell B4
            workbook.CurrentWorksheet.RemoveCell("A2");                 // Delete cell A2
            workbook.CurrentWorksheet.RemoveCell(1,1);                  // Delete cell B2
            workbook.Save();                                            // Save the workbook
        }


    }
}

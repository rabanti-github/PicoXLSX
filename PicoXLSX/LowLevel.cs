/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO.Packaging;
using System.Text.RegularExpressions;

namespace PicoXLSX
{
    /// <summary>
    /// Class for low level handling (XML, formatting, packing)
    /// </summary>
    class LowLevel
    {
        private static DocumentPath WORKBOOK = new DocumentPath("workbook.xml", "xl/");
        private static DocumentPath STYLES = new DocumentPath("styles.xml", "xl/");

        private CultureInfo culture;
        private Workbook workbook;

        /// <summary>
        /// Constructor with defined workbook object
        /// </summary>
        /// <param name="workbook">workbook to process</param>
        public LowLevel(Workbook workbook)
        {
            this.culture = CultureInfo.CreateSpecificCulture("en-US");
            this.workbook = workbook;
        }

        /// <summary>
        /// Method to save the workbook
        /// </summary>
        /// <returns>True, if the workbook could be saved, otherwise false</returns>
        public bool Save()
        {
            XmlDocument workbookDocument = CreateWorkbookDocument();
            XmlDocument worksheetDocument;
            DocumentPath sheetPath;
            List<Uri> sheetURIs = new List<Uri>();
            try
            {
                using (System.IO.Packaging.Package p = Package.Open(this.workbook.Filename, FileMode.Create))
                {
                    Uri workbookUri = new Uri(WORKBOOK.GetFullPath(), UriKind.Relative);
                    Uri stylesheetUri = new Uri(STYLES.GetFullPath(), UriKind.Relative);

                    PackagePart pp = p.CreatePart(workbookUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", CompressionOption.Normal);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        WriteXMLStream(ms, workbookDocument);
                        LowLevel.CopyStream(ms, pp.GetStream());
                    }
                    int styleId = this.workbook.Worksheets.Count + 1;
                    p.CreateRelationship(pp.Uri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                    pp.CreateRelationship(stylesheetUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "rId" + styleId.ToString());

                    foreach (Worksheet item in this.workbook.Worksheets)
                    {
                        sheetPath = new DocumentPath("sheet" + item.SheetID.ToString() + ".xml", "xl/worksheets");
                        sheetURIs.Add(new Uri(sheetPath.GetFullPath(), UriKind.Relative));
                        pp.CreateRelationship(sheetURIs[sheetURIs.Count - 1], TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId" + item.SheetID.ToString());
                    }
                    int i = 0;
                    foreach (Worksheet item in this.workbook.Worksheets)
                    {
                        pp = p.CreatePart(sheetURIs[i], @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
                        i++;
                        worksheetDocument = CreateWorksheetPart(item);
                        using (MemoryStream ms = new MemoryStream())
                        {
                            WriteXMLStream(ms, worksheetDocument);
                            LowLevel.CopyStream(ms, pp.GetStream());
                        }
                    }


                    pp = p.CreatePart(stylesheetUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", CompressionOption.Normal);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        WriteXMLStream(ms, CreateStyleSheetDocument());
                        LowLevel.CopyStream(ms, pp.GetStream());
                    }
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Method to create a worksheet part as XML document
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <returns>Formated XML document</returns>
        private XmlDocument CreateWorksheetPart(Worksheet worksheet)
        {
            XmlDocument worksheetDocument = new XmlDocument();
            List<List<Cell>> celldata = GetSortedSheetData(worksheet);
            StringBuilder sb = new StringBuilder();
            string line;
            sb.Append("<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n");
            sb.Append("<x:sheetData>\r\n");
            foreach(List<Cell> item in celldata)
            {
                line = CreateRowString(item);
                sb.Append(line + "\r\n");
            }
            sb.Append("</x:sheetData>\r\n");
            sb.Append("</x:worksheet>");
            worksheetDocument.LoadXml(sb.ToString());
            XmlDeclaration dec = worksheetDocument.CreateXmlDeclaration("1.0", null, null);
            dec.Encoding = "UTF-8";
            XmlElement root = worksheetDocument.DocumentElement;
            worksheetDocument.InsertBefore(dec, root);
            return worksheetDocument;
        }

        /// <summary>
        /// Method to sort the cells of a worksheet as preparation for the XML document
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Two dimensional array of cell objects</returns>
        private List<List<Cell>> GetSortedSheetData(Worksheet sheet)
        {
            List<Cell> temp = new List<Cell>();
            foreach(KeyValuePair<string, Cell> item in sheet.Cells)
            {
                temp.Add(item.Value);
            }
            temp.Sort();
            List<Cell> line = new List<Cell>();
            List<List<Cell>> output = new List<List<Cell>>();
            if (temp.Count > 0)
            {
                int rowNumber = temp[0].RowAddress;
                foreach (Cell item in temp)
                {
                    if (item.RowAddress != rowNumber)
                    {
                        output.Add(line);
                        line = new List<Cell>();
                        rowNumber = item.RowAddress;
                    }
                    line.Add(item);
                }
                if (line.Count > 0)
                {
                    output.Add(line);
                }
            }
            return output;
        }

        /// <summary>
        /// Method to create a row string
        /// </summary>
        /// <param name="columnFields">List of cells</param>
        /// <returns>Formated row string</returns>
        private string CreateRowString(List<Cell> columnFields)
        {
            StringBuilder sb = new StringBuilder();
            if (columnFields.Count > 0)
            {
                sb.Append("<x:row r=\"" + (columnFields[0].RowAddress + 1).ToString() + "\">\r\n");
            }
            else
            {
                sb.Append("<x:row>\r\n");
            }
            string typeAttribute;
            string sValue = "";
            string tValue = "";
            string value = "";
            bool bVal;

            DateTime dVal;
            int col = 0;
            foreach (Cell item in columnFields)
            {
                tValue = " ";
                sValue = "";
                if (item.Fieldtype == Cell.CellType.BOOL)
                {
                    typeAttribute = "b";
                    tValue = " t=\"" + typeAttribute + "\" ";
                    bVal = (bool)item.Value;
                    if (bVal == true) { value = "1"; }
                    else { value = "0"; }
                    
                }
                // Number casting
                else if (item.Fieldtype == Cell.CellType.NUMBER)
                {
                    typeAttribute = "n";
                    tValue = " t=\"" + typeAttribute + "\" ";
                    Type t = item.Value.GetType();


                    if (t == typeof(int))
                    {
                        value = ((int)item.Value).ToString("G", culture);
                    }
                    else if(t == typeof(double))
                    {
                        value = ((double)item.Value).ToString("G", culture);

                    }
                    else if(t == typeof(float))
                    {
                        value = ((float)item.Value).ToString("G", culture);
                    }

                }
                // Date parsing
                else if (item.Fieldtype == Cell.CellType.DATE)
                {
                    typeAttribute = "d";
                    dVal = (DateTime)item.Value;
                    value = LowLevel.GetOADateTimeString(dVal);
                    sValue = " s=\"1\" ";
                }
                // String parsing
                else
                {
                    typeAttribute = "str";
                    tValue = " t=\"" + typeAttribute + "\" ";
                    value = item.Value.ToString();
                }
                sb.Append("<x:c" + tValue + "r=\"" + item.GetCellAddress() + "\"" + sValue + ">\r\n");
                if (item.Fieldtype == Cell.CellType.FORMULA)
                {
                    sb.Append("<x:f>" + LowLevel.EscapeXMLChars(item.Value.ToString()) + "</x:f>\r\n");
                }
                else
                {
                    sb.Append("<x:v>" + LowLevel.EscapeXMLChars(value) + "</x:v>\r\n");
                }

                sb.Append("</x:c>\r\n");
                col++;
            }
            sb.Append("</x:row>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create a workbook as XML document
        /// </summary>
        /// <returns>Formated XML document</returns>
        private XmlDocument CreateWorkbookDocument()
        {
            if (this.workbook.Worksheets.Count == 0)
            {
                throw new OutOfRangeException("The workbook can not be created because no worksheet was defined.");
            }
            XmlDocument doc = new XmlDocument();
            StringBuilder sb = new StringBuilder();
            sb.Append("<x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n");
            sb.Append("<x:sheets>\r\n");
            foreach(Worksheet item in this.workbook.Worksheets)
            {
                sb.Append("<x:sheet r:id=\"rId" + item.SheetID.ToString() + "\" sheetId=\"" + item.SheetID.ToString() + "\" name=\"" + LowLevel.EscapeXMLAttributeChars(item.SheetName) + "\"/>\r\n");
            }
            sb.Append("</x:sheets>\r\n");
            sb.Append("</x:workbook>");
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", null, null);
            dec.Encoding = "UTF-8";
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            //this.workbookDocument = doc;
            return doc;
        }

        /// <summary>
        /// Method to create a stylesheet as XML document
        /// </summary>
        /// <returns>Formated XML document</returns>
        private XmlDocument CreateStyleSheetDocument()
        {
            XmlDocument doc = new XmlDocument();
            StringBuilder sb = new StringBuilder();
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">\r\n");
            sb.Append("<fonts x14ac:knownFonts=\"1\" count=\"1\">\r\n");
            sb.Append("<font>\r\n<sz val=\"11\" />\r\n<name val=\"Calibri\" />\r\n<family val=\"2\" />\r\n<scheme val=\"minor\" />\r\n</font>\r\n");
            sb.Append("</fonts>\r\n");

            sb.Append("<fills count=\"1\">\r\n");
            sb.Append("<fill>\r\n<patternFill patternType=\"none\" />\r\n</fill>\r\n");
            sb.Append("</fills>\r\n");

            sb.Append("<borders count=\"1\">\r\n");
            sb.Append("<border>\r\n<left />\r\n<right />\r\n<top />\r\n<bottom />\r\n<diagonal />\r\n</border>\r\n");
            sb.Append("</borders>\r\n");

            sb.Append("<cellXfs count=\"2\">\r\n");
            sb.Append("<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>\r\n");
            sb.Append("<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"14\" applyNumberFormat=\"1\" xfId=\"0\"/>\r\n");
            sb.Append("</cellXfs>\r\n");

            sb.Append("</styleSheet>");
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", null, null);
            dec.Encoding = "UTF-8";
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            return doc;
        }

        /// <summary>
        /// Method to write an XML document to a MemoryStream
        /// </summary>
        /// <param name="stream">Stream to write the XML document</param>
        /// <param name="document">XML document to process</param>
        private void WriteXMLStream(MemoryStream stream, XmlDocument document)
        {
            if (stream == null) { return; }
            if (stream.CanWrite == false) { return; }
            document.Save(stream);
        }

        /// <summary>
        /// Method to escape XML charactes between two XML tags
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        public static string EscapeXMLChars(string input)
        {
            input = input.Replace("<", "&lt;");
            input = input.Replace(">", "&gt;");
            return input;
        }

        /// <summary>
        /// Method to esacpe XML charactes in an XML attribute
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        public static string EscapeXMLAttributeChars(string input)
        {
            input = input.Replace("\"", "&quot;");
            return input;
        }

        /// <summary>
        /// Method to copy a memory stream into another memory stream
        /// </summary>
        /// <param name="sourceStream">Source stream</param>
        /// <param name="targetStream">Target stream</param>
        /// <exception cref="IOException">Throws a IOException if the memory stream could not be copied</exception>
        public static void CopyStream(System.IO.MemoryStream sourceStream, System.IO.Stream targetStream)
        {
            if (sourceStream == null || targetStream == null)
            {
                throw new IOException("The source or target memory stream to create a workbook part was not defined.");
            }
            try
            {
                byte[] buffer = sourceStream.GetBuffer();
                targetStream.Write(buffer, 0, (int)sourceStream.Length);
            }
            catch (Exception e)
            {
                throw new IOException("The memory stream to create a workbook part could not be copied.", e);
            }
        }

        /// <summary>
        /// Method to convert a date or date and time into the Excel time format
        /// </summary>
        /// <param name="date">Date to process</param>
        /// <exception cref="FormatException">Throws a FormatException if the date could not be converted to the OA format</exception>
        /// <returns>Date or date and time as Number</returns>
        public static string GetOADateTimeString(DateTime date)
        {
            try
            {
                return date.ToOADate().ToString();
            }
            catch (Exception e)
            {
                throw new FormatException("The date could not be transformed into Excel format (OADate).", e);
            }
        }

        /// <summary>
        /// Class to manage XML document paths
        /// </summary>
        public class DocumentPath
        {
            /// <summary>
            /// File name of the document
            /// </summary>
            public string Filename { get; set; }
            /// <summary>
            /// Path of the document
            /// </summary>
            public string Path { get; set; }

            /// <summary>
            /// Default constructor
            /// </summary>
            public DocumentPath()
            {
            }

            /// <summary>
            /// Constructor with defined file name and path
            /// </summary>
            /// <param name="fiename">File name of the document</param>
            /// <param name="path">Path of the document</param>
            public DocumentPath(string fiename, string path)
            {
                this.Filename = fiename;
                this.Path = path;
            }

            /// <summary>
            /// Method to return the full path of the document
            /// </summary>
            /// <returns>Full path</returns>
            public string GetFullPath()
            {
                if (this.Path == null) { return this.Filename; }
                if (this.Path == "") { return this.Filename; }
                if (this.Path[this.Path.Length - 1] == System.IO.Path.AltDirectorySeparatorChar || this.Path[this.Path.Length - 1] == System.IO.Path.DirectorySeparatorChar)
                {
                    return System.IO.Path.AltDirectorySeparatorChar.ToString() + this.Path + this.Filename;
                }
                else
                {
                    return System.IO.Path.AltDirectorySeparatorChar.ToString() + this.Path + System.IO.Path.AltDirectorySeparatorChar.ToString() + this.Filename;
                }
            }

        }

    }
}

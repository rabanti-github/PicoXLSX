/*
 * PicoXLSX is a small .NET library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace PicoXLSX
{
    /// <summary>
    /// Class for low level handling (XML, formatting, packing)
    /// </summary>
    /// <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    class LowLevel
    {
        private static DocumentPath WORKBOOK = new DocumentPath("workbook.xml", "xl/");
        private static DocumentPath STYLES = new DocumentPath("styles.xml", "xl/");
        private static DocumentPath APP_PROPERTIES = new DocumentPath("app.xml", "docProps/");
        private static DocumentPath CORE_PROPERTIES = new DocumentPath("core.xml", "docProps/");
        private static DocumentPath SHARED_STRINGS = new DocumentPath("sharedStrings.xml", "xl/");
        private static RNGCryptoServiceProvider RNGcsp = new RNGCryptoServiceProvider();

        private CultureInfo culture;
        private Workbook workbook;
        //private Dictionary<string, string> sharedStrings;
        private SortedMap sharedStrings;
        private int sharedStringsTotalCount;

        /// <summary>
        /// Constructor with defined workbook object
        /// </summary>
        /// <param name="workbook">Workbook to process</param>
        public LowLevel(Workbook workbook)
        {
            this.culture = CultureInfo.CreateSpecificCulture("en-US");
            this.workbook = workbook;
            //this.sharedStrings = new Dictionary<string, string>();
            this.sharedStrings = new SortedMap();
            this.sharedStringsTotalCount = 0;
        }

        /// <summary>
        /// Method to save the workbook
        /// </summary>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="OutOfRangeException">Throws an OutOfRangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        /// <remarks>The UndefinedStyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        public void Save()
        {
            this.workbook.ResolveMergedCells();
            XmlDocument workbookDocument = CreateWorkbookDocument();
            XmlDocument stylesheetDocument = CreateStyleSheetDocument();
            XmlDocument worksheetDocument;
            DocumentPath sheetPath;
            List<Uri> sheetURIs = new List<Uri>();

            try
            {
                using (System.IO.Packaging.Package p = Package.Open(this.workbook.Filename, FileMode.Create))
                {
                    Uri workbookUri = new Uri(WORKBOOK.GetFullPath(), UriKind.Relative);
                    Uri stylesheetUri = new Uri(STYLES.GetFullPath(), UriKind.Relative);
                    Uri appPropertiesUri = new Uri(APP_PROPERTIES.GetFullPath(), UriKind.Relative);
                    Uri corePropertiesUri = new Uri(CORE_PROPERTIES.GetFullPath(), UriKind.Relative);
                    Uri sharedStringsUri = new Uri(SHARED_STRINGS.GetFullPath(), UriKind.Relative);

                    PackagePart pp = p.CreatePart(workbookUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", CompressionOption.Normal);
                    p.CreateRelationship(pp.Uri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1");
                    p.CreateRelationship(corePropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "rId2"); //!
                    p.CreateRelationship(appPropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "rId3"); //!

                    using (MemoryStream ms = new MemoryStream()) // Write workbook.xml
                    {
                        WriteXMLStream(ms, workbookDocument);
                        LowLevel.CopyStream(ms, pp.GetStream());
                    }
                    int idCounter = this.workbook.Worksheets.Count + 1;
                    
                    pp.CreateRelationship(stylesheetUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "rId" + idCounter.ToString());
                    pp.CreateRelationship(sharedStringsUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "rId" + (idCounter + 1).ToString());

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

                    pp = p.CreatePart(sharedStringsUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", CompressionOption.Normal);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        WriteXMLStream(ms, CreateSharedStringsDocument());
                        LowLevel.CopyStream(ms, pp.GetStream());
                    }
                    

                    pp = p.CreatePart(stylesheetUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", CompressionOption.Normal);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        WriteXMLStream(ms, stylesheetDocument);
                        LowLevel.CopyStream(ms, pp.GetStream());
                    }

                    if (workbook.WorkbookMetadata != null)
                    {
                        pp = p.CreatePart(appPropertiesUri, @"application/vnd.openxmlformats-officedocument.extended-properties+xml", CompressionOption.Normal);                       
                        using (MemoryStream ms = new MemoryStream())
                        {
                            WriteXMLStream(ms, CreateAppPropertiesDocument());
                            LowLevel.CopyStream(ms, pp.GetStream());
                        }
                        pp = p.CreatePart(corePropertiesUri, @"application/vnd.openxmlformats-package.core-properties+xml", CompressionOption.Normal);                       
                        using (MemoryStream ms = new MemoryStream())
                        {
                            WriteXMLStream(ms, CreateCorePropertiesDocument());
                            LowLevel.CopyStream(ms, pp.GetStream());
                        }
                    }
     

                }
            }
            catch (Exception e)
            {
                throw new IOException("An error occurred while saving. See inner exception for details.", e);
            }
        }

        /// <summary>
        /// Method to create a worksheet part as XML document
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <returns>Formated XML document</returns>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        private XmlDocument CreateWorksheetPart(Worksheet worksheet)
        {
            worksheet.RecalculateAutoFilter();
            worksheet.RecalculateColumns();
            XmlDocument worksheetDocument = new XmlDocument();
            List<List<Cell>> celldata = GetSortedSheetData(worksheet);
            StringBuilder sb = new StringBuilder();
            string line;
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            
            if (worksheet.SelectedCells != null)
            {
            	sb.Append("<sheetViews><sheetView workbookViewId=\"0\"");
            	if (this.workbook.SelectedWorksheet == worksheet.SheetID - 1)
            	{
            		sb.Append(" tabSelected=\"1\"");
            	}
            	sb.Append("><selection sqref=\"");
            	sb.Append(worksheet.SelectedCells.ToString());
            	sb.Append("\" activeCell=\"");
            	sb.Append(worksheet.SelectedCells.Value.StartAddress.ToString());
            	sb.Append("\"/></sheetView></sheetViews>");
            }
            
            sb.Append("<sheetFormatPr x14ac:dyDescent=\"0.25\" defaultRowHeight=\"" + worksheet.DefaultRowHeight.ToString("G", culture) + "\" baseColWidth=\"" + worksheet.DefaultColumnWidth.ToString("G", culture) + "\"/>");
            
            string colWidths = CreateColsString(worksheet);
            if (string.IsNullOrEmpty(colWidths) == false)
            {
                sb.Append("<cols>");
                sb.Append(colWidths);
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            foreach(List<Cell> item in celldata)
            {
                line = CreateRowString(item, worksheet);
                sb.Append(line);
            }
            sb.Append("</sheetData>");

            sb.Append(CreateMergedCellsString(worksheet));
            sb.Append(CreateSheetProtectionString(worksheet));

            if (worksheet.AutoFilterRange != null)
            {
                sb.Append("<autoFilter ref=\"" + worksheet.AutoFilterRange.Value.ToString() + "\"/>");
            }

            sb.Append("</worksheet>");
            worksheetDocument.LoadXml(sb.ToString());
            XmlDeclaration dec = worksheetDocument.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            XmlElement root = worksheetDocument.DocumentElement;
            worksheetDocument.InsertBefore(dec, root);
            return worksheetDocument;
        }

        /// <summary>
        /// Method to create a style sheet as XML document
        /// </summary>
        /// <returns>Formated XML document</returns>
        /// <exception cref="UndefinedStyleException">Throws an UndefinedStyleException if one of the styles cannot be referenced or is null</exception>
        /// <remarks>The UndefinedStyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        
        private XmlDocument CreateStyleSheetDocument()
        {
            List<Style.Border> borders;
            List<Style.Fill> fills;
            List<Style.Font> fonts;
            List<Style.NumberFormat> numberFormats;
            List<Style.CellXf> cellXfs; // Not used at the moment
            int numFormatCount = 0;
            this.workbook.ReorganizeStyles(out borders, out fills, out fonts, out numberFormats, out cellXfs);
            string bordersString = CreateStyleBorderString(borders);
            string fillsString = CreateStyleFillString(fills);
            string fontsString = CreateStyleFontString(fonts);
            string numberFormatsString = CreateStyleNumberFormatString(numberFormats, out numFormatCount);
            string xfsStings = CreateStyleXfsString(this.workbook.Styles);
            string mruColorString = CreateMruColorsString(fonts, fills);
            StringBuilder sb = new StringBuilder();
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            if (numFormatCount > 0)
            {
                sb.Append("<numFmts count=\"" + numFormatCount.ToString("G", culture) + "\">");
                sb.Append(numberFormatsString + "</numFmts>");
            }
            sb.Append("<fonts x14ac:knownFonts=\"1\" count=\"" + fonts.Count.ToString("G", culture) + "\">");
            sb.Append(fontsString + "</fonts>");
            sb.Append("<fills count=\"" + fills.Count.ToString("G", culture) + "\">");
            sb.Append(fillsString + "</fills>");
            sb.Append("<borders count=\"" + borders.Count.ToString("G", culture) + "\">");
            sb.Append(bordersString + "</borders>");
            sb.Append("<cellXfs count=\"" + this.workbook.Styles.Count.ToString("G", culture) + "\">");
            sb.Append(xfsStings + "</cellXfs>");
            if (this.workbook.WorkbookMetadata != null)
            {
                if (string.IsNullOrEmpty(mruColorString) == false && this.workbook.WorkbookMetadata.UseColorMRU == true)
                {
                    sb.Append("<colors>");
                    sb.Append(mruColorString);
                    sb.Append("</colors>");
                }
            }
            sb.Append("</styleSheet>");
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            return doc;
        }

        /// <summary>
        /// Method to create the app-properties (part of meta data) as XML document
        /// </summary>
        /// <returns>Formated XML document</returns>
        private XmlDocument CreateAppPropertiesDocument()
        {
            XmlDocument doc = new XmlDocument();
            StringBuilder sb = new StringBuilder();
            sb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
            sb.Append(CreateAppString());
            sb.Append("</Properties>");
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            return doc;
        }

        /// <summary>
        /// Method to create the core-properties (part of meta data) as XML document
        /// </summary>
        /// <returns>Formated XML document</returns>
        private XmlDocument CreateCorePropertiesDocument()
        {
            XmlDocument doc = new XmlDocument();
            StringBuilder sb = new StringBuilder();
            sb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            sb.Append(CreateCorePropertiesString());
            sb.Append("</cp:coreProperties>");
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            return doc;
        }

        /// <summary>
        /// Method to create a workbook as XML document
        /// </summary>
        /// <returns>Formated XML document</returns>
        /// <exception cref="OutOfRangeException">Throws an OutOfRangeException if an address was out of range</exception>
        private XmlDocument CreateWorkbookDocument()
        {
            if (this.workbook.Worksheets.Count == 0)
            {
                throw new OutOfRangeException("The workbook can not be created because no worksheet was defined.");
            }
            XmlDocument doc = new XmlDocument();
            StringBuilder sb = new StringBuilder();
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            if (this.workbook.SelectedWorksheet > 0)
            {
                sb.Append("<bookViews><workbookView activeTab=\"");
                sb.Append(this.workbook.SelectedWorksheet.ToString("G", culture));
                sb.Append("\"/></bookViews>");
            }
            if (this.workbook.UseWorkbookProtection == true)
            {
                sb.Append("<workbookProtection");
                if (this.workbook.LockWindowsIfProtected == true)
                {
                    sb.Append(" lockWindows=\"1\"");
                }
                if (this.workbook.LockStructureIfProtected == true)
                {
                    sb.Append(" lockStructure=\"1\"");
                }
                if (string.IsNullOrEmpty(this.workbook.WorkbookProtectionPassword) == false)
                {
                    sb.Append("workbookPassword=\"");
                    sb.Append(GeneratePasswordHash(this.workbook.WorkbookProtectionPassword));
                    sb.Append("\"");
                }
                sb.Append("/>");
            }
            sb.Append("<sheets>");
            foreach (Worksheet item in this.workbook.Worksheets)
            {
                sb.Append("<sheet r:id=\"rId" + item.SheetID.ToString() + "\" sheetId=\"" + item.SheetID.ToString() + "\" name=\"" + LowLevel.EscapeXMLAttributeChars(item.SheetName) + "\"/>");
            }
            sb.Append("</sheets>");
            sb.Append("</workbook>");
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            //this.workbookDocument = doc;
            return doc;
        }

        /// <summary>
        /// Method to create a style sheet as XML document (OBSOLETE / fall-back method)
        /// </summary>
        /// <returns>Formated XML document</returns>
        [Obsolete("This method was superseded by the method CreateStyleSheetDocument. Only use this as fall-back if the Style class became broken")]
        private XmlDocument CreateStyleSheetDocumentFallback()
        {
            XmlDocument doc = new XmlDocument();
            StringBuilder sb = new StringBuilder();
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            sb.Append("<fonts x14ac:knownFonts=\"1\" count=\"4\">");
            sb.Append("<font><sz val=\"11\" /><name val=\"Calibri\" /><family val=\"2\" /><scheme val=\"minor\" /></font>");            // Date
            sb.Append("<font><b/><sz val=\"11\" /><name val=\"Calibri\" /><family val=\"2\" /><scheme val=\"minor\" /></font>");    // Bold
            sb.Append("<font><i/><sz val=\"11\" /><name val=\"Calibri\" /><family val=\"2\" /><scheme val=\"minor\" /></font>");    // Italic
            sb.Append("<font><u/><sz val=\"11\" /><name val=\"Calibri\" /><family val=\"2\" /><scheme val=\"minor\" /></font>");    // Underline
            sb.Append("</fonts>");

            sb.Append("<fills count=\"1\">");
            sb.Append("<fill><patternFill patternType=\"none\" /></fill>");
            sb.Append("</fills>");

            sb.Append("<borders count=\"1\">");
            sb.Append("<border><left /><right /><top /><bottom /><diagonal /></border>");
            sb.Append("</borders>");

            sb.Append("<cellXfs count=\"3\">");
            sb.Append("<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>");
            sb.Append("<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"14\" applyNumberFormat=\"1\" xfId=\"0\"/>"); // DateFormat (s="1")
            sb.Append("<xf borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" applyNumberFormat=\"1\" xfId=\"0\"/>"); // Bold  (s="2")
            sb.Append("<xf borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\" applyNumberFormat=\"1\" xfId=\"0\"/>"); // Italic  (s="3")
            sb.Append("<xf borderId=\"0\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" applyNumberFormat=\"1\" xfId=\"0\"/>"); // Underline  (s="4")
            sb.Append("</cellXfs>");

            sb.Append("</styleSheet>");
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            return doc;
        }

        /// <summary>
        /// Method to create shared strings as XML document
        /// </summary>
        /// <returns>Formated XML document</returns>
        private XmlDocument CreateSharedStringsDocument()
        {
            XmlDocument doc = new XmlDocument();
            StringBuilder sb = new StringBuilder();
            sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
            sb.Append(this.sharedStringsTotalCount.ToString("G", culture));
            sb.Append("\" uniqueCount=\"");
            sb.Append(this.sharedStrings.Count.ToString("G", culture));
            sb.Append("\">");
           // foreach(KeyValuePair<string,string>str in this.sharedStrings)
           foreach(SortedMap.Tuple str in this.sharedStrings)
            {
                sb.Append("<si><t>");
                sb.Append(EscapeXMLChars(str.Key));
                sb.Append("</t></si>");
            }
            sb.Append("</sst>");
            doc.LoadXml(sb.ToString());
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(dec, root);
            return doc;
        }

        /// <summary>
        /// Method to sort the cells of a worksheet as preparation for the XML document
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Two dimensional array of Cell objects</returns>
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
        /// <param name="worksheet">Worksheet to process</param>
        /// <returns>Formated row string</returns>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        private string CreateRowString(List<Cell> columnFields, Worksheet worksheet)
        {
            int rowNumber = columnFields[0].RowAddress;
            string heigth = "";
            string hidden = "";
            if (worksheet.RowHeights.ContainsKey(rowNumber))
            {
                if (worksheet.RowHeights[rowNumber] != worksheet.DefaultRowHeight)
                {
                    heigth = " x14ac:dyDescent=\"0.25\" customHeight=\"1\" ht=\"" + worksheet.RowHeights[rowNumber].ToString("G", culture) + "\"";
                }
            }
            if (worksheet.HiddenRows.ContainsKey(rowNumber))
            {
                if (worksheet.HiddenRows[rowNumber] == true)
                {
                    hidden = " hidden=\"1\"";
                }
            }
            StringBuilder sb = new StringBuilder();
            if (columnFields.Count > 0)
            {
                sb.Append("<row r=\"" + (rowNumber + 1).ToString() + "\"" + heigth + hidden + ">");
            }
            else
            {
                sb.Append("<row" + heigth + ">");
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
                if (item.CellStyle != null)
                {
                    sValue = " s=\"" + item.CellStyle.InternalID.ToString("G", culture) + "\" ";
                }
                else
                {
                    sValue = "";
                }
                item.ResolveCellType(); // Recalculate the type (for handling DEFAULT)
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
                    else if (t == typeof(long))
                    {
                        value = ((long)item.Value).ToString("G", culture);
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
                }
                // String parsing
                else
                {
                    if (item.Value == null)
                    {
                        typeAttribute = "str";
                        value = string.Empty;
                    }
                    else // Handle sharedStrings
                    {
                        if (item.Fieldtype == Cell.CellType.FORMULA)
                        {
                            typeAttribute = "str";
                            value = item.Value.ToString();
                        }
                        else
                        {
                            typeAttribute = "s";
                            value = item.Value.ToString();
                            if (this.sharedStrings.ContainsKey(value) == false)
                            {
                                this.sharedStrings.Add(value, sharedStrings.Count.ToString("G", culture));
                            }
                            value = this.sharedStrings[value];
                            this.sharedStringsTotalCount++;
                        }
                    }
                    tValue = " t=\"" + typeAttribute + "\" ";
                }
                if (item.Fieldtype != Cell.CellType.EMPTY)
                {
                    sb.Append("<c" + tValue + "r=\"" + item.GetCellAddress() + "\"" + sValue + ">");
                    if (item.Fieldtype == Cell.CellType.FORMULA)
                    {
                        sb.Append("<f>" + LowLevel.EscapeXMLChars(item.Value.ToString()) + "</f>");
                    }
                    else
                    {
                        sb.Append("<v>" + LowLevel.EscapeXMLChars(value) + "</v>");
                    }
                    sb.Append("</c>");
                }
                else // Empty cell
                {
                    sb.Append("<c" + tValue + "r=\"" + item.GetCellAddress() + "\"" + sValue + "/>");
                }
                col++;
            }
            sb.Append("</row>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the merged cells string of the passed worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Formated string with merged cell ranges</returns>
        private string CreateMergedCellsString(Worksheet sheet)
        {
            if (sheet.MergedCells.Count < 1)
            {
                return string.Empty;
            }
                StringBuilder sb = new StringBuilder();
                sb.Append("<mergeCells count=\"" + sheet.MergedCells.Count.ToString("G", culture) + "\">");
                foreach (KeyValuePair<string, Cell.Range> item in sheet.MergedCells)
                {
                    sb.Append("<mergeCell ref=\"" + item.Value.ToString() + "\"/>");
                }
            sb.Append("</mergeCells>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the protection string of the passed worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Formated string with protection statement of the worksheet</returns>
        private string CreateSheetProtectionString(Worksheet sheet)
        {
            if (sheet.UseSheetProtection == false)
            {
                return string.Empty;
            }
            Dictionary<Worksheet.SheetProtectionValue, int> actualLockingValues = new Dictionary<Worksheet.SheetProtectionValue,int>();
            if (sheet.SheetProtectionValues.Count == 0)
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.selectLockedCells, 1);
                actualLockingValues.Add(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
            }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.objects) == false)
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.objects, 1);
            }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.scenarios) == false)
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.scenarios, 1);
            }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectLockedCells) == false )
            {
                if (actualLockingValues.ContainsKey(Worksheet.SheetProtectionValue.selectLockedCells) == false)
                {
                    actualLockingValues.Add(Worksheet.SheetProtectionValue.selectLockedCells, 1);
                }
            }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectUnlockedCells) == false || sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectLockedCells) == false)
            {
                if (actualLockingValues.ContainsKey(Worksheet.SheetProtectionValue.selectUnlockedCells) == false)
                {
                    actualLockingValues.Add(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
                }
            }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatCells)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatCells, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertHyperlinks)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertHyperlinks, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.deleteColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.deleteColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.deleteRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.deleteRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.sort)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.sort, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.autoFilter)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.autoFilter, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.pivotTables)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.pivotTables, 0); }
            StringBuilder sb = new StringBuilder();
            sb.Append("<sheetProtection");
            string temp;
            foreach(KeyValuePair<Worksheet.SheetProtectionValue, int>item in actualLockingValues)
            {
               try
               {
                   temp = Enum.GetName(typeof(Worksheet.SheetProtectionValue), item.Key); // Note! If the enum names differs from the OOXML definitions, this method will cause invalid OOXML entries
                   sb.Append(" " + temp + "=\"" + item.Value.ToString("G", culture)  + "\"");
               }
               catch { }
            }
            if (string.IsNullOrEmpty(sheet.SheetProtectionPassword) == false)
            {
                string hash = GeneratePasswordHash(sheet.SheetProtectionPassword);
                sb.Append(" password=\"" + hash + "\"");
            }
            sb.Append(" sheet=\"1\"/>");
           return sb.ToString();
        }


        /// <summary>
        /// Method to create the XML string for the app-properties document
        /// </summary>
        /// <returns>String with formated XML data</returns>
        private string CreateAppString()
        {
            if (this.workbook.WorkbookMetadata == null) { return string.Empty; }
            Metadata md = this.workbook.WorkbookMetadata;
            StringBuilder sb = new StringBuilder();
            AppendXMLtag(sb, "0", "TotalTime", null);
            AppendXMLtag(sb, md.Application, "Application", null);
            AppendXMLtag(sb, "0", "DocSecurity", null);
            AppendXMLtag(sb, "false", "ScaleCrop", null);
            AppendXMLtag(sb, md.Manager, "Manager", null);
            AppendXMLtag(sb, md.Company, "Company", null);
            AppendXMLtag(sb, "false", "LinksUpToDate", null);
            AppendXMLtag(sb, "false", "SharedDoc", null);
            AppendXMLtag(sb, md.HyperlinkBase, "HyperlinkBase", null);
            AppendXMLtag(sb, "false", "HyperlinksChanged", null);
            AppendXMLtag(sb, md.ApplicationVersion, "AppVersion", null);
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the core-properties document
        /// </summary>
        /// <returns>String with formated XML data</returns>
        private string CreateCorePropertiesString()
        {
            if (this.workbook.WorkbookMetadata == null) { return string.Empty; }
            Metadata md = this.workbook.WorkbookMetadata;
            StringBuilder sb = new StringBuilder();
            AppendXMLtag(sb, md.Title, "title", "dc");
            AppendXMLtag(sb, md.Subject, "subject", "dc");
            AppendXMLtag(sb, md.Creator, "creator", "dc");
            AppendXMLtag(sb, md.Creator, "lastModifiedBy", "cp");
            AppendXMLtag(sb, md.Keywords, "keywords", "cp");
            AppendXMLtag(sb, md.Description, "description", "dc");

            string time = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ");
            sb.Append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + time + "</dcterms:created>");
            sb.Append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + time + "</dcterms:modified>");

            AppendXMLtag(sb, md.Category, "category", "cp");
            AppendXMLtag(sb, md.ContentStatus, "contentStatus", "cp");

            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the font part of the style sheet document
        /// </summary>
        /// <param name="fontStyles">List of Style.Font objects</param>
        /// <returns>String with formated XML data</returns>
        private string CreateStyleFontString(List<Style.Font> fontStyles)
        {
            StringBuilder sb = new StringBuilder();
            foreach(Style.Font item in fontStyles)
            {
                sb.Append("<font>");
                if (item.Bold == true) { sb.Append("<b/>"); }
                if (item.Italic == true) { sb.Append("<i/>"); }
                if (item.Underline == true) { sb.Append("<u/>"); }
                if (item.DoubleUnderline == true) { sb.Append("<u val=\"double\"/>"); }
                if (item.Strike == true) { sb.Append("<strike/>"); }
                if (item.VerticalAlign == Style.Font.VerticalAlignValue.subscript) { sb.Append("<vertAlign val=\"subscript\"/>"); }
                else if (item.VerticalAlign == Style.Font.VerticalAlignValue.superscript) { sb.Append("<vertAlign val=\"superscript\"/>"); }
                sb.Append("<sz val=\"" + item.Size.ToString("G", culture) + "\"/>");
                if (string.IsNullOrEmpty(item.ColorValue))
                {
                    sb.Append("<color theme=\"" + item.ColorTheme.ToString("G", culture) + "\"/>");
                }
                else
                {
                    sb.Append("<color rgb=\"" + item.ColorValue + "\"/>");
                }
                sb.Append("<name val=\"" + item.Name + "\"/>");
                sb.Append("<family val=\"" + item.Family + "\"/>");
                if (item.Scheme != Style.Font.SchemeValue.none)
                {
                    if (item.Scheme == Style.Font.SchemeValue.major)
                    { sb.Append("<scheme val=\"major\"/>"); }
                    else if (item.Scheme == Style.Font.SchemeValue.minor)
                    { sb.Append("<scheme val=\"minor\"/>"); }
                }
                if (string.IsNullOrEmpty(item.Charset) == false)
                {
                    sb.Append("<charset val=\"" + item.Charset + "\"/>");
                }
                sb.Append("</font>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the border part of the style sheet document
        /// </summary>
        /// <param name="borderStyles">List of Style.Border objects</param>
        /// <returns>String with formated XML data</returns>
        private string CreateStyleBorderString(List<Style.Border> borderStyles)
        {
            StringBuilder sb = new StringBuilder();
            foreach (Style.Border item in borderStyles)
            {
                if (item.DiagonalDown == true && item.DiagonalUp == false) { sb.Append("<border diagonalDown=\"1\">"); }
                else if (item.DiagonalDown == false && item.DiagonalUp == true) { sb.Append("<border diagonalUp=\"1\">"); }
                else if (item.DiagonalDown == true && item.DiagonalUp == true) { sb.Append("<border diagonalDown=\"1\" diagonalUp=\"1\">"); }
                else { sb.Append("<border>"); }
                
                if (item.LeftStyle != Style.Border.StyleValue.none)
                {
                    sb.Append("<left style=\"" + Style.Border.GetStyleName(item.LeftStyle) + "\">");
                    if (string.IsNullOrEmpty(item.LeftColor) == true) { sb.Append("<color rgb=\"" + item.LeftColor + "\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</left>");
                }
                else
                {
                    sb.Append("<left/>");
                }
                if (item.RightStyle != Style.Border.StyleValue.none)
                {
                    sb.Append("<right style=\"" + Style.Border.GetStyleName(item.RightStyle) + "\">");
                    if (string.IsNullOrEmpty(item.RightColor) == true) { sb.Append("<color rgb=\"" + item.RightColor + "\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</right>");
                }
                else
                {
                    sb.Append("<right/>");
                }
                if (item.TopStyle != Style.Border.StyleValue.none)
                {
                    sb.Append("<top style=\"" + Style.Border.GetStyleName(item.TopStyle) + "\">");
                    if (string.IsNullOrEmpty(item.TopColor) == true) { sb.Append("<color rgb=\"" + item.TopColor + "\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</top>");
                }
                else
                {
                    sb.Append("<top/>");
                }
                if (item.BottomStyle != Style.Border.StyleValue.none)
                {
                    sb.Append("<bottom style=\"" + Style.Border.GetStyleName(item.BottomStyle) + "\">");
                    if (string.IsNullOrEmpty(item.BottomColor) == true) { sb.Append("<color rgb=\"" + item.BottomColor + "\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</bottom>");
                }
                else
                {
                    sb.Append("<bottom/>");
                }
                if (item.DiagonalStyle != Style.Border.StyleValue.none)
                {
                    sb.Append("<diagonal style=\"" + Style.Border.GetStyleName(item.DiagonalStyle) + "\">");
                    if (string.IsNullOrEmpty(item.DiagonalColor) == true) { sb.Append("<color rgb=\"" + item.DiagonalColor + "\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</diagonal>");
                }
                else
                {
                    sb.Append("<diagonal/>");
                }

                sb.Append("</border>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the fill part of the style sheet document
        /// </summary>
        /// <param name="fillStyles">List of Style.Fill objects</param>
        /// <returns>String with formated XML data</returns>
        private string CreateStyleFillString(List<Style.Fill> fillStyles)
        {
            StringBuilder sb = new StringBuilder();
            foreach (Style.Fill item in fillStyles)
            {
                sb.Append("<fill>");
                sb.Append("<patternFill patternType=\"" + Style.Fill.GetPatternName(item.PatternFill) + "\"");
                if (item.PatternFill == Style.Fill.PatternValue.solid)
                {
                    sb.Append(">");
                    sb.Append("<fgColor rgb=\"" + item.ForegroundColor + "\"/>");
                    sb.Append("<bgColor indexed=\"" + item.IndexedColor.ToString("G", culture) + "\"/>");
                    sb.Append("</patternFill>");
                }
                else if (item.PatternFill == Style.Fill.PatternValue.mediumGray || item.PatternFill == Style.Fill.PatternValue.lightGray || item.PatternFill == Style.Fill.PatternValue.gray0625 || item.PatternFill == Style.Fill.PatternValue.darkGray)
                {
                    sb.Append(">");
                    sb.Append("<fgColor rgb=\"" + item.ForegroundColor + "\"/>");
                    if (string.IsNullOrEmpty(item.BackgroundColor) == false)
                    {
                        sb.Append("<bgColor rgb=\"" + item.BackgroundColor + "\"/>");
                    }
                    sb.Append("</patternFill>");
                }
                else
                {
                    sb.Append("/>");
                }
                sb.Append("</fill>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the color-MRU part of the style sheet document (recent colors)
        /// </summary>
        /// <param name="fills">List of Style.Fill objects</param>
        /// <param name="fonts">List of Style.Font objects</param>
        /// <returns>String with formated XML data</returns>
        private string CreateMruColorsString(List<Style.Font> fonts, List<Style.Fill> fills)
        {
            StringBuilder sb = new StringBuilder();
            List<string> tempColors = new List<string>();
            foreach (Style.Font item in fonts)
            {
                if (string.IsNullOrEmpty(item.ColorValue) == true) { continue; }
                if (item.ColorValue == Style.Fill.DEFAULTCOLOR) { continue; }
                if (tempColors.Contains(item.ColorValue) == false) { tempColors.Add(item.ColorValue); }
            }
            foreach (Style.Fill item in fills)
            {
                if (string.IsNullOrEmpty(item.BackgroundColor) == false)
                {
                    if (item.BackgroundColor != Style.Fill.DEFAULTCOLOR)
                    {
                        if (tempColors.Contains(item.BackgroundColor) == false) { tempColors.Add(item.BackgroundColor); }
                    }
                }
                if (string.IsNullOrEmpty(item.ForegroundColor) == false)
                {
                    if (item.ForegroundColor != Style.Fill.DEFAULTCOLOR)
                    {
                        if (tempColors.Contains(item.ForegroundColor) == false) { tempColors.Add(item.ForegroundColor); }
                    }
                }
            }
            if (tempColors.Count > 0)
            {
                sb.Append("<mruColors>");
                foreach(string item in tempColors)
                {
                    sb.Append("<color rgb=\"" + item + "\"/>");
                }
                sb.Append("</mruColors>");
                return sb.ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Method to create the XML string for the Xf part of the style sheet document
        /// </summary>
        /// <param name="styles">List of Style objects</param>
        /// <returns>String with formated XML data</returns>
        private string CreateStyleXfsString(List<Style> styles)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            string alignmentString, protectionString;
            int formatNumber, textRotation;
            foreach (Style item in styles)
            {
                textRotation = item.CurrentCellXf.CalculateInternalRotation();
                alignmentString = string.Empty;
                protectionString = string.Empty;
                if (item.CurrentCellXf.HorizontalAlign != Style.CellXf.HorizontalAlignValue.none || item.CurrentCellXf.VerticalAlign != Style.CellXf.VerticallAlignValue.none || item.CurrentCellXf.Alignment != Style.CellXf.TextBreakValue.none || textRotation != 0)
                {
                    sb2.Clear();
                    sb2.Append("<alignment");
                    if (item.CurrentCellXf.HorizontalAlign != Style.CellXf.HorizontalAlignValue.none)
                    {
                        sb2.Append(" horizontal=\"");
                        if (item.CurrentCellXf.HorizontalAlign == Style.CellXf.HorizontalAlignValue.center) { sb2.Append("center"); }
                        else if (item.CurrentCellXf.HorizontalAlign == Style.CellXf.HorizontalAlignValue.right) { sb2.Append("right"); }
                        else if (item.CurrentCellXf.HorizontalAlign == Style.CellXf.HorizontalAlignValue.centerContinuous) { sb2.Append("centerContinuous"); }
                        else if (item.CurrentCellXf.HorizontalAlign == Style.CellXf.HorizontalAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (item.CurrentCellXf.HorizontalAlign == Style.CellXf.HorizontalAlignValue.fill) { sb2.Append("fill"); }
                        else if (item.CurrentCellXf.HorizontalAlign == Style.CellXf.HorizontalAlignValue.general) { sb2.Append("general"); }
                        else if (item.CurrentCellXf.HorizontalAlign == Style.CellXf.HorizontalAlignValue.justify) { sb2.Append("justify"); }
                        else { sb2.Append("left"); }
                        sb2.Append("\"");
                    }
                    if (item.CurrentCellXf.VerticalAlign != Style.CellXf.VerticallAlignValue.none)
                    {
                        sb2.Append(" vertical=\"");
                        if (item.CurrentCellXf.VerticalAlign == Style.CellXf.VerticallAlignValue.center) { sb2.Append("center"); }
                        else if (item.CurrentCellXf.VerticalAlign == Style.CellXf.VerticallAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (item.CurrentCellXf.VerticalAlign == Style.CellXf.VerticallAlignValue.justify) { sb2.Append("justify"); }
                        else if (item.CurrentCellXf.VerticalAlign == Style.CellXf.VerticallAlignValue.top) { sb2.Append("top"); }
                        else { sb2.Append("bottom"); }
                        sb2.Append("\"");
                    }
                    
                    if (item.CurrentCellXf.Alignment != Style.CellXf.TextBreakValue.none)
                    {
                        if (item.CurrentCellXf.Alignment == Style.CellXf.TextBreakValue.shrinkToFit) { sb2.Append(" shrinkToFit=\"1"); }
                        else { sb2.Append(" wrapText=\"1"); }
                        sb2.Append("\"");
                    }
                    if (textRotation != 0)
                    {
                        sb2.Append(" textRotation=\"");
                        sb2.Append(textRotation.ToString("G", culture));
                        sb2.Append("\"");
                    }
                    sb2.Append("/>"); // </xf>
                    alignmentString = sb2.ToString();
                }

                if (item.CurrentCellXf.Hidden == true || item.CurrentCellXf.Locked == true)
                {
                    if (item.CurrentCellXf.Hidden == true && item.CurrentCellXf.Locked == true)
                    {
                        protectionString = "<protection locked=\"1\" hidden=\"1\"/>";
                    }
                    else if (item.CurrentCellXf.Hidden == true && item.CurrentCellXf.Locked == false)
                    {
                        protectionString = "<protection hidden=\"1\" locked=\"0\"/>";
                    }
                    else
                    {
                        protectionString = "<protection hidden=\"0\" locked=\"1\"/>";
                    }
                }

                sb.Append("<xf numFmtId=\"");
                if (item.CurrentNumberFormat.IsCustomFormat == true)
                {
                    sb.Append(item.CurrentNumberFormat.CustomFormatID.ToString("G", culture));
                }
                else
                {
                    formatNumber = (int)item.CurrentNumberFormat.Number;
                    sb.Append(formatNumber.ToString("G", culture));
                }
                sb.Append("\" borderId=\"" + item.CurrentBorder.InternalID.ToString("G", culture));
                sb.Append("\" fillId=\"" + item.CurrentFill.InternalID.ToString("G", culture));
                sb.Append("\" fontId=\"" + item.CurrentFont.InternalID.ToString("G", culture));
                if (item.CurrentFont.IsDefaultFont == false)
                {
                    sb.Append("\" applyFont=\"1");
                }
                if (item.CurrentFill.PatternFill != Style.Fill.PatternValue.none)
                {
                    sb.Append("\" applyFill=\"1");
                }
                if (item.CurrentBorder.IsEmpty() == false)
                {
                    sb.Append("\" applyBorder=\"1");
                }
                if (alignmentString != string.Empty || item.CurrentCellXf.ForceApplyAlignment == true)
                {
                    sb.Append("\" applyAlignment=\"1");
                }
                if (protectionString != string.Empty)
                {
                    sb.Append("\" applyProtection=\"1");
                }
                if (item.CurrentNumberFormat.Number != Style.NumberFormat.FormatNumber.none)
                {
                    sb.Append("\" applyNumberFormat=\"1\"");
                }
                else
                {
                    sb.Append("\""); 
                }
                if (alignmentString != string.Empty || protectionString != string.Empty)
                {
                    sb.Append(">");
                    sb.Append(alignmentString);
                    sb.Append(protectionString);
                    sb.Append("</xf>");
                }
                else
                {
                    sb.Append("/>");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the columns as XML string. This is used to define the width of columns
        /// </summary>
        /// <param name="worksheet">Worksheet to process</param>
        /// <returns>String with formated XML data</returns>
        private string CreateColsString(Worksheet worksheet)
        {
            if (worksheet.Columns.Count > 0)
            {
                string col;
                string hidden = "";
                StringBuilder sb = new StringBuilder();
                foreach(KeyValuePair<int, Worksheet.Column> column in worksheet.Columns)
                { 
                    if (column.Value.Width == worksheet.DefaultColumnWidth && column.Value.IsHidden == false) { continue; }
                    if (worksheet.Columns.ContainsKey(column.Key))
                    {
                        if (worksheet.Columns[column.Key].IsHidden == true)
                        {
                            hidden = " hidden=\"1\"";
                        }
                    }
                    col = (column.Key + 1).ToString("G", culture); // Add 1 for Address
                    sb.Append("<col customWidth=\"1\" width=\"" + column.Value.Width.ToString("G", culture) + "\" max=\"" + col + "\" min=\"" + col + "\"" + hidden + "/>");
                }
                string value = sb.ToString();
                if (value.Length > 0)
                {
                    return value;
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Method to create the XML string for the number format part of the style sheet document 
        /// </summary>
        /// <param name="numberFormatStyles">List of Style.NumberFormat objects</param>
        /// <param name="counter">Out-parameter for the number of custom number formats</param>
        /// <returns>String with formated XML data</returns>
        private string CreateStyleNumberFormatString(List<Style.NumberFormat> numberFormatStyles, out int counter)
        {
            counter = 0;
            StringBuilder sb = new StringBuilder();
            foreach (Style.NumberFormat item in numberFormatStyles)
            {
                if (item.IsCustomFormat == true)
                {
                    sb.Append("<numFmt formatCode=\"" + item.CustomFormatCode + "\" numFmtId=\"" + item.CustomFormatID.ToString("G", culture) + "\"/>");
                    counter++;
                }
            }
            return sb.ToString();
        }



#region utilMethods


		/// <summary>
		/// Creates an unique, random string with the stated length
		/// </summary>
		/// <param name="length">Lenth of the name in characters</param>
		/// <returns>Unique (random) name</returns>
        public static string CreateUniqueName(int length)
        {
            byte[] rndByte = new byte[4];
            RNGcsp.GetBytes(rndByte);
            int res = BitConverter.ToInt32(rndByte, 0);
            int nds;
            Random rnd = new Random(res);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < length; i++)
            {
                nds = rnd.Next(0, 2);
                if (nds == 0)
                {
                    sb.Append((char)rnd.Next(48, 57));
                }
                else
                {
                    sb.Append((char)rnd.Next(65, 90));
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to append a simple XML tag with an enclosed value to the passed StringBuilder
        /// </summary>
        /// <param name="sb">StringBuilder to append</param>
        /// <param name="value">Value of the XML element</param>
        /// <param name="tagName">Tag name of the XML element</param>
        /// <param name="nameSpace">Optional XML name space. Can be empty or null</param>
        /// <returns>Returns false if no tag was appended, because the value or tag name was null or empty</returns>
        private bool AppendXMLtag(StringBuilder sb, string value, string tagName, string nameSpace)
        {
            if (string.IsNullOrEmpty(value)) { return false; }
            if (sb == null || string.IsNullOrEmpty(tagName)) { return false; }
            bool hasNoNs = string.IsNullOrEmpty(nameSpace);
            sb.Append('<');
            if (hasNoNs == false)
            {
                sb.Append(nameSpace);
                sb.Append(':');
            }
            sb.Append(tagName + ">");
            sb.Append(EscapeXMLChars(value));
            sb.Append("</");
            if (hasNoNs == false)
            {
                sb.Append(nameSpace);
                sb.Append(':');
            }
            sb.Append(tagName);
            sb.Append(">");
            return true;
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
        /// Method to escape XML characters between two XML tags
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        /// <remarks>Note: The XML specs allow characters up to the character value of 0x10FFFF. However, the C# char range is only up to 0xFFFF. PicoXLSX will neglect all values above this level in the sanitizing check. Illegal characters like 0x1 will be replaced with a white space (0x20)</remarks>
        public static string EscapeXMLChars(string input)
        {
            int len = input.Length;
            List<int> illegalCharacters = new List<int>(len);
            List<byte> characterTypes = new List<byte>(len);
            int i;
            for (i = 0; i < len; i++)
            {
                if ((input[i] < 0x9) || (input[i] > 0xA && input[i] < 0xD) || (input[i] > 0xD && input[i] < 0x20) || (input[i] > 0xD7FF && input[i] < 0xE000) || (input[i] > 0xFFFD))
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(0);
                    continue;
                } // Note: XML specs allow characters up to 0x10FFFF. However, the C# char range is only up to 0xFFFF; Higher values are neglected here 
                if (input[i] == 0x3C) // <
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(1);
                }
                else if (input[i] == 0x3E) // >
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(2);
                }
                else if (input[i] == 0x26) // &
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(3);
                }
            }
            if (illegalCharacters.Count == 0)
            {
                return input;
            }

            StringBuilder sb = new StringBuilder(len);
            int lastIndex = 0;
            len = illegalCharacters.Count;
            for (i = 0; i < len; i++)
            {
                sb.Append(input.Substring(lastIndex, illegalCharacters[i] - lastIndex));
                if (characterTypes[i] == 0)
                {
                    sb.Append(' '); // Whitespace as fall back on illegal character
                }
                else if (characterTypes[i] == 1) // replace <
                {
                    sb.Append("&lt;");
                }
                else if (characterTypes[i] == 2) // replace >
                {
                    sb.Append("&gt;");
                }
                else if (characterTypes[i] == 3) // replace &
                {
                    sb.Append("&amp;");
                }
                lastIndex = illegalCharacters[i] + 1;
            }
            sb.Append(input.Substring(lastIndex));
            return sb.ToString();
        }

        /// <summary>
        /// Method to escape XML characters in an XML attribute
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        public static string EscapeXMLAttributeChars(string input)
        {
            input = EscapeXMLChars(input); // Sanitize string from illegal characters beside quotes
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
        /// <returns>Date or date and time as Number</returns>
        /// <exception cref="FormatException">Throws a FormatException if the passed date cannot be translated to OADate format</exception>
        /// <remarks>OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances</remarks>
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
        /// Method to generate an Excel internal password hash to protect workbooks or worksheets<br></br>This method is derived from the c++ implementation by Kohei Yoshida (<a href="http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)
        /// </summary>
        /// <remarks>WARNING! Do not use this method to encrypt 'real' passwords or data outside from PicoXLSX. This is only a minor security feature. Use a proper cryptography method instead.</remarks>
        /// <param name="password">Password string in UTF-8 to encrypt</param>
        /// <returns>16 bit hash as hex string</returns>
        public static string GeneratePasswordHash(string password)
        {
            if (string.IsNullOrEmpty(password)) { return string.Empty; }
            int PasswordLength = password.Length;
            int passwordHash = 0;
            char character;
            for (int i = PasswordLength; i > 0; i--)
            {
                character = password[i - 1];
                passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
                passwordHash ^= character;
            }
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= (0x8000 | ('N' << 8) | 'K');
            passwordHash ^= PasswordLength;
            return passwordHash.ToString("X");
        }

#endregion
	
	/// <summary>
	/// Class to manage key value pairs (string / string). The entries are in the order how they were added
	/// </summary>
	public class SortedMap : IEnumerable<SortedMap.Tuple>
		{
			private List<Tuple> entries;
			
			/// <summary>
			/// Gets the number of entries in the map
			/// </summary>
			public int Count {
				get { return this.entries.Count; }
			}
			
			/// <summary>
			/// Indexer to get the specific value by the key
			/// </summary>
			public string this[string key]
			{
				get
				{
					int s = this.entries.Count;
			        for(int i = 0; i < s; i++ )
			        {
			            if (this.entries[i].Key == key)
			            {
			            	return this.entries[i].Value;
			            }
			        }
			        return null;
				}
			}
			
			/// <summary>
			/// Default constructor
			/// </summary>
			public SortedMap()
			{
				this.entries = new List<Tuple>();
			}
			
			
			/// <summary>
			/// Method to add a key value pair
			/// </summary>
			/// <param name="key">Key as string</param>
			/// <param name="value">Value as string</param>
			/// <returns>Returns the index of the inserted or replaced entry in the map</returns>
			public int Add(string key, string value)
		    {
				int s = this.entries.Count;
		        for(int i = 0; i < s; i++ )
		        {
		            if (this.entries[i].Key == key)
		            {
		            	this.entries[i] = new Tuple(key, value);
		                return i;
		            }
		        }
		        this.entries.Add(new Tuple(key, value));
		        return s;
		    }
			
			/// <summary>
			/// Gets whether the specified key exists in the map
			/// </summary>
			/// <param name="key">Key to check</param>
			/// <returns>True if the entry exists, otherwise false</returns>
			public bool ContainsKey(string key)
			{
				int s = this.entries.Count;
		        for(int i = 0; i < s; i++ )
		        {
		            if (this.entries[i].Key == key)
		            {
		            	return true;
		            }
		        }
		        return false;			
			}
			
			/// <summary>
			/// Removes all entries
			/// </summary>
			public void Clear()
			{
				this.entries.Clear();
			}
			
			/// <summary>
			/// Gets a list of key
			/// </summary>
			/// <returns>Keys of the map</returns>
			public List<string> Keys()
			{
				List<string> output = new List<string>();
				int s = this.entries.Count;
		        for(int i = 0; i < s; i++ )
		        {
		        	output.Add(this.entries[i].Key);
		        }
		        return output;		
			}
			
			/// <summary>
			/// Gets a List of Values
			/// </summary>
			/// <returns>Values of the map</returns>
			public List<string> Values()
			{
				List<string> output = new List<string>();
				int s = this.entries.Count;
		        for(int i = 0; i < s; i++ )
		        {
		        	output.Add(this.entries[i].Value);
		        }
		        return output;		
			}

			/// <summary>
			/// Sub-Class to manage string/string tuples 
			/// </summary>
			public class Tuple
			{
				/// <summary>
				/// Key of the tuple
				/// </summary>
				public string Key { get; set; }
				/// <summary>
				/// Value of the tuple
				/// </summary>
				public string Value { get; set; }
				/// <summary>
				/// Default constructor with parameters
				/// </summary>
				/// <param name="key">Key of the tuple</param>
				/// <param name="value">Value of the tuple</param>
				public Tuple(string key, string value)
				{
					this.Key = key;
					this.Value = value;
				}
			}

			#region IEnumerable implementation
			public IEnumerator<SortedMap.Tuple> GetEnumerator()
			{
				return this.entries.GetEnumerator();
			}
			#endregion
			#region IEnumerable implementation
			IEnumerator IEnumerable.GetEnumerator()
			{
				return this.GetEnumerator();
			}
			#endregion
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
            /// <param name="filename">File name of the document</param>
            /// <param name="path">Path of the document</param>
            public DocumentPath(string filename, string path)
            {
                this.Filename = filename;
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

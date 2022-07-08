using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelExport.Export
{
    public class ExcelExportLib
    {
        const int DEFAULT_COLUMN_WIDTH = 30;        // Set each column to be 20 wide
        static string HeaderColorCode;

        public MemoryStream CreateExcelDocument<T>(List<T> list, string headerColorCode)
        {
            DataSet ds = new DataSet();
            HeaderColorCode = headerColorCode;
            ds.Tables.Add(ListToDataTable(list));

            return CreateExcelDocument(ds);
        }

        public static DataTable ListToDataTable<T>(List<T> list)
        {
            DataTable dt = new DataTable();

            foreach (PropertyInfo info in typeof(T).GetProperties())
            {
                dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));
            }
            foreach (T t in list)
            {
                DataRow row = dt.NewRow();
                foreach (PropertyInfo info in typeof(T).GetProperties())
                {
                    if (!IsNullableType(info.PropertyType))
                        row[info.Name] = info.GetValue(t, null);
                    else
                        row[info.Name] = (info.GetValue(t, null) ?? DBNull.Value);
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
        private static Type GetNullableType(Type t)
        {
            Type returnType = t;
            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                returnType = Nullable.GetUnderlyingType(t);
            }
            return returnType;
        }
        private static bool IsNullableType(Type type)
        {
            return (type == typeof(string) ||
                    type.IsArray ||
                    (type.IsGenericType &&
                     type.GetGenericTypeDefinition().Equals(typeof(Nullable<>))));
        }

        public static MemoryStream CreateExcelDocument(DataTable dt)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            MemoryStream result = CreateExcelDocument(ds);
            ds.Tables.Remove(dt);
            return result;
        }

        public static MemoryStream CreateExcelDocument(DataSet ds)
        {
            MemoryStream mem = new MemoryStream();
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook))
            {
                WriteExcelFile(ds, spreadsheet);
            }
            return mem;
        }

        private static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet)
        {
            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            DefinedNames definedNamesCol = new DefinedNames();
            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));
            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            workbookStylesPart.Stylesheet = GenerateStyleSheet();
            workbookStylesPart.Stylesheet.Save();

            uint worksheetNumber = 1;
            Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            foreach (DataTable dt in ds.Tables)
            {
                string worksheetName = "Articles Report";
                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart), SheetId = worksheetNumber, Name = worksheetName };

                sheets.Append(sheet);
                WriteDataTableToExcelWorksheet(dt, newWorksheetPart, definedNamesCol);

                worksheetNumber++;
            }
            spreadsheet.WorkbookPart.Workbook.Append(definedNamesCol);
            spreadsheet.WorkbookPart.Workbook.Save();
        }

        private static Stylesheet GenerateStyleSheet()
        {
            uint iExcelIndex = 164;

            return new Stylesheet(
                new NumberingFormats(
                    //  
                    new NumberingFormat()                                                  // Custom number format # 164: especially for date-times
                    {
                        NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                        FormatCode = StringValue.FromString("dd/MMM/yyyy hh:mm:ss")
                    },
                    new NumberingFormat()                                                   // Custom number format # 165: especially for date times (with a blank time)
                    {
                        NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                        FormatCode = StringValue.FromString("dd/MMM/yyyy")
                    }
               ),
                new Fonts(
                    new Font(                                                               // Index 0 - The default font.
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Arial" }),
                    new Font(                                                               // Index 1 - A 12px bold font, in white.
                        new Bold(),
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } },
                        new FontName() { Val = "Arial" }),
                    new Font(                                                               // Index 2 - An Italic font.
                        new Italic(),
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" }),
                     new Font(                                                               // Index 3 - Uri font.
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "0000FF" } },
                        new FontName() { Val = "Arial" },
                        new Underline() { Val = UnderlineValues.Single })
                ),
                new Fills(
                    new Fill(                                                           // Index 0 - The default fill.
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                        new PatternFill() { PatternType = PatternValues.Gray125 }),
                    new Fill(                                                           // Index 2 - The yellow fill.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
                        )
                        { PatternType = PatternValues.Solid }),
                    new Fill(                                                           // Index 3 - Dark-gray fill.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = HeaderColorCode } }
                        )
                        { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(                                                         // Index 0 - The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                         // Style # 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { NumberFormatId = 164 },                                         // Style # 1 - DateTimes
                    new CellFormat() { NumberFormatId = 165 },                                         // Style # 2 - Dates (with a blank time)
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center })
                    { FontId = 1, FillId = 3, BorderId = 0, ApplyFont = true, ApplyAlignment = true },       // Style # 3 - Header row 
                    new CellFormat() { NumberFormatId = 3 },                                           // Style # 4 - Number format: #,##0
                    new CellFormat() { NumberFormatId = 4 },                                           // Style # 5 - Number format: #,##0.00
                    new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Style # 6 - Bold 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Style # 7 - Italic
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Style # 8 - Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Style # 9 - Yellow Fill
                    new CellFormat(                                                                    // Style # 10 - Alignment
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // Style # 11 - Border
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 0 }
                )
            ); // return
        }

        private static void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart, DefinedNames definedNamesCol)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart, Encoding.ASCII);
            writer.WriteStartElement(new Worksheet());

            //  To demonstrate how to set column-widths in Excel, here's how to set the width of all columns to our default of "25":
            UInt32 inx = 1;
            writer.WriteStartElement(new Columns());
            foreach (DataColumn dc in dt.Columns)
            {
                writer.WriteElement(new Column { Min = inx, Max = inx, CustomWidth = true, Width = DEFAULT_COLUMN_WIDTH });
                inx++;
            }
            writer.WriteEndElement();


            writer.WriteStartElement(new SheetData());

            string cellValue = "";
            string cellReference = "";

            int numberOfColumns = dt.Columns.Count;
            bool[] IsIntegerColumn = new bool[numberOfColumns];
            bool[] IsFloatColumn = new bool[numberOfColumns];
            bool[] IsDateColumn = new bool[numberOfColumns];
            bool[] IsUriColumn = new bool[numberOfColumns];

            string[] excelColumnNames = new string[numberOfColumns];
            for (int n = 0; n < numberOfColumns; n++)
                excelColumnNames[n] = GetExcelColumnName(n);

            uint rowIndex = 1;

            writer.WriteStartElement(new Row { RowIndex = rowIndex, Height = 40, CustomHeight = true });
            for (int colInx = 0; colInx < numberOfColumns; colInx++)
            {
                DataColumn col = dt.Columns[colInx];
                AppendHeaderTextCell(excelColumnNames[colInx] + "1", col.ColumnName.TrimStart('_').Replace("_", " "), writer);

                IsIntegerColumn[colInx] = (col.DataType.FullName.StartsWith("System.Int"));
                IsFloatColumn[colInx] = (col.DataType.FullName == "System.Decimal") || (col.DataType.FullName == "System.Double") || (col.DataType.FullName == "System.Single");
                IsDateColumn[colInx] = (col.DataType.FullName == "System.DateTime");
                IsUriColumn[colInx] = (col.DataType.FullName == "System.Uri");
            }
            writer.WriteEndElement();   //  End of header "Row"

            double cellFloatValue = 0;
            CultureInfo ci = new CultureInfo("en-US");
            foreach (DataRow dr in dt.Rows)
            {
                // ...create a new row, and append a set of this row's data to it.
                ++rowIndex;

                writer.WriteStartElement(new Row { RowIndex = rowIndex, Height = 40, CustomHeight = true });

                for (int colInx = 0; colInx < numberOfColumns; colInx++)
                {
                    cellValue = dr.ItemArray[colInx].ToString();
                    cellValue = ReplaceHexadecimalSymbols(cellValue);
                    cellReference = excelColumnNames[colInx] + rowIndex.ToString();

                    // Create cell with data
                    if (IsIntegerColumn[colInx] || IsFloatColumn[colInx])
                    {
                        cellFloatValue = 0;
                        bool bIncludeDecimalPlaces = IsFloatColumn[colInx];
                        if (double.TryParse(cellValue, out cellFloatValue))
                        {
                            cellValue = cellFloatValue.ToString(ci);
                            AppendNumericCell(cellReference, cellValue, bIncludeDecimalPlaces, writer);
                        }
                    }
                    else if (IsDateColumn[colInx])
                    {
                        DateTime dateValue;
                        if (DateTime.TryParse(cellValue, out dateValue))
                        {
                            AppendDateCell(cellReference, dateValue, writer);
                        }
                        else
                        {
                            AppendTextCell(cellReference, cellValue, writer);
                        }
                    }
                    else if (IsUriColumn[colInx])
                    {
                        AppendUriCell(cellReference, cellValue, writer);
                    }
                    else
                    {
                        AppendTextCell(cellReference, cellValue, writer);
                    }
                }
                writer.WriteEndElement(); //  End of Row
            }
            writer.WriteEndElement(); //  End of SheetData
            writer.WriteEndElement(); //  End of worksheet

            writer.Close();
        }

        private static void AppendHeaderTextCell(string cellReference, string cellStringValue, OpenXmlWriter writer)
        {
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.String,
                StyleIndex = 3
            });
        }

        private static void AppendTextCell(string cellReference, string cellStringValue, OpenXmlWriter writer)
        {
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.String
            });
        }

        private static void AppendDateCell(string cellReference, DateTime dateTimeValue, OpenXmlWriter writer)
        {
            string cellStringValue = dateTimeValue.ToOADate().ToString(CultureInfo.InvariantCulture);
            bool bHasBlankTime = (dateTimeValue.Date == dateTimeValue);

            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                StyleIndex = UInt32Value.FromUInt32(bHasBlankTime ? (uint)2 : (uint)1),
                DataType = CellValues.Number        //  Use this, rather than CellValues.Date
            });
        }

        private static void AppendFormulaCell(string cellReference, string cellStringValue, OpenXmlWriter writer)
        {
            //  Add a new "formula" Excel Cell to our Row 
            writer.WriteElement(new Cell
            {
                CellFormula = new CellFormula(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.Number
            });
        }

        private static void AppendNumericCell(string cellReference, string cellStringValue, bool bIncludeDecimalPlaces, OpenXmlWriter writer)
        {
            //  Add a new numeric Excel Cell to our Row.
            UInt32 cellStyle = (UInt32)(bIncludeDecimalPlaces ? 5 : 4);
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                StyleIndex = cellStyle,                                 //  Style #4 formats with 0 decimal places, style #5 formats with 2 decimal places
                DataType = CellValues.Number
            });
        }

        private static void AppendUriCell(string cellReference, string cellStringValue, OpenXmlWriter writer)
        {
            string dataVal = $@"HYPERLINK(""{cellStringValue}"", ""Article Image"")";
            writer.WriteElement(new Cell
            {
                CellFormula = new CellFormula(dataVal),
                CellReference = cellReference,
                StyleIndex = 12,
                DataType = CellValues.InlineString
            });
        }

        private static string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }

        //  Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Z, AA, AB, AC... AY, AZ, BA, BB..)
        public static string GetExcelColumnName(int columnIndex)
        {
            int firstInt = columnIndex / 676;
            int secondInt = (columnIndex % 676) / 26;
            if (secondInt == 0)
            {
                secondInt = 26;
                firstInt = firstInt - 1;
            }
            int thirdInt = (columnIndex % 26);

            char firstChar = (char)('A' + firstInt - 1);
            char secondChar = (char)('A' + secondInt - 1);
            char thirdChar = (char)('A' + thirdInt);

            if (columnIndex < 26)
                return thirdChar.ToString();

            if (columnIndex < 702)
                return string.Format("{0}{1}", secondChar, thirdChar);

            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }
    }
}

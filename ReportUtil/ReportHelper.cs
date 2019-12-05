using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

namespace ReportUtil
{
    public interface IReportHelper
    {
        Stream GenerateReport<TMaster, TDetail>(IList<TMaster> masters, ColumnDefBase[] columnDefs, Func<TMaster, IList<TDetail>> getDetailFunc);
        Stream GenerateReportWithTemplate<TMaster, TDetail>(Stream targetStream, IList<TMaster> masters, ColumnDefBase[] columnDefs, Func<TMaster, IList<TDetail>> getDetailFunc, string sheetName = "Sheet1", int beginRowIndex = 2);
        Stream GenerateReportWithTemplate<T>(Stream targetStream, IList<T> records, ColumnDef<T>[] columnDefs, string sheetName = "Sheet1", int beginRowIndex = 2);
    }
    public class ColumnDefBase
    {
        public string Captain { get; set; }

        public EnumValue<CellValues> TargetDataType { get; set; }

    }

    public class ColumnDef<T> : ColumnDefBase
    {
        public Func<T, CellValue> GetValueFunc { get; set; }
    }

    internal enum Directions
    {
        Left = 1,
        Top = 2,
        Right = 4,
        Bottom = 8
    }

    public class ReportBuilder
    {
        SpreadsheetDocument spreadsheetDocument;
        WorkbookPart workbookpart;
        WorkbookStylesPart sp;
        Stream stream;
        int numberStyleIndex, dateStyleIndex, stringStyleIndex;
        public ReportBuilder():this(new MemoryStream())
        {
            
        }

        public ReportBuilder(Stream stream)
        {
            this.stream = stream;
            this.spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            this.Init();
        }

        private void Init()
        {
            this.workbookpart = spreadsheetDocument.AddWorkbookPart();
            this.workbookpart.Workbook = new Workbook();
            this.sp = this.workbookpart.AddNewPart<WorkbookStylesPart>("rId3");

            int numberStyleIndex, dateStyleIndex, stringStyleIndex;
            ReportHelper.CreateCellFormats(sp, out numberStyleIndex, out dateStyleIndex, out stringStyleIndex);
            this.numberStyleIndex = numberStyleIndex;
            this.dateStyleIndex = dateStyleIndex;
            this.stringStyleIndex = stringStyleIndex;

        }

        public ReportBuilder AddSheet<T>(IList<T> records, ColumnDef<T>[] columnDefs)
        {
            WorksheetPart worksheetPart = this.workbookpart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            uint[] styleIndexies = ReportHelper.GetStylesIndexes(columnDefs, this.numberStyleIndex, this.dateStyleIndex, this.stringStyleIndex);
            CellFormat cellFormat = this.sp.Stylesheet.CellFormats.ChildElements[(int)styleIndexies[0]] as CellFormat;

            #region captain 
            Row captainRow = new Row();
            captainRow.RowIndex = (UInt32)1;
            int captainSyleIndex = ReportHelper.AddCaptainFormat(sp.Stylesheet, cellFormat.BorderId);
            for (int i = 0; i < columnDefs.Length; i++)
            {

                Cell cellCaptain = new Cell();

                cellCaptain.CellReference = $"{(char)((int)'A' + i)}{captainRow.RowIndex}";
                cellCaptain.CellValue = new CellValue(columnDefs[i].Captain);
                cellCaptain.DataType = new EnumValue<CellValues>(CellValues.String);
                cellCaptain.StyleIndex = (uint)captainSyleIndex;
                captainRow.AppendChild(cellCaptain);
            }

            sheetData.AppendChild(captainRow);
            uint rowIndex = captainRow.RowIndex;
            #endregion
            foreach (var record in records)
            {
                var dataRow = new Row();
                dataRow.RowIndex = (uint)++rowIndex;
                sheetData.Append(dataRow);

                for (int k = 0; k < columnDefs.Length; k++)
                {
                    Cell cellData = new Cell();
                    cellData.CellReference = $"{(char)((int)'A' + k)}{dataRow.RowIndex}";
                    cellData.DataType = columnDefs[k].TargetDataType;
                    dataRow.Append(cellData);
                    cellData.CellValue = columnDefs[k].GetValueFunc(record);
                    cellData.StyleIndex = styleIndexies[k];
                }
            }

            ReportHelper.SetColumnWidth(columnDefs, worksheetPart, sheetData, styleIndexies);

            ReportHelper.AppendNewSheet(spreadsheetDocument, worksheetPart);

            return this;
        }

        public Stream Build()
        {
            this.spreadsheetDocument.Close();
            stream.Position = 0;
            return stream;
        }

    }

    public class ReportHelper : IReportHelper
    {
        public Stream GenerateReport<T>(IList<T> records, ColumnDef<T>[] columnDefs)
        {
            var stream = new MemoryStream();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                WorkbookStylesPart sp = workbookpart.AddNewPart<WorkbookStylesPart>("rId3");
                uint[] styleIndexies = AddCellStyles(columnDefs, sp);
                CellFormat cellFormat = sp.Stylesheet.CellFormats.ChildElements[(int)styleIndexies[0]] as CellFormat;

                #region captain 
                Row captainRow = new Row();
                captainRow.RowIndex = (UInt32)1;
                int captainSyleIndex = AddCaptainFormat(sp.Stylesheet, cellFormat.BorderId);
                for (int i = 0; i < columnDefs.Length; i++)
                {

                    Cell cellCaptain = new Cell();

                    cellCaptain.CellReference = $"{(char)((int)'A' + i)}{captainRow.RowIndex}";
                    cellCaptain.CellValue = new CellValue(columnDefs[i].Captain);
                    cellCaptain.DataType = new EnumValue<CellValues>(CellValues.String);
                    cellCaptain.StyleIndex = (uint)captainSyleIndex;
                    captainRow.AppendChild(cellCaptain);
                }

                sheetData.AppendChild(captainRow);
                uint rowIndex = captainRow.RowIndex;
                #endregion
                foreach (var record in records)
                {
                    var dataRow = new Row();
                    dataRow.RowIndex = (uint)++rowIndex;
                    sheetData.Append(dataRow);

                    for (int k = 0; k < columnDefs.Length; k++)
                    {
                        Cell cellData = new Cell();
                        cellData.CellReference = $"{(char)((int)'A' + k)}{dataRow.RowIndex}";
                        cellData.DataType = columnDefs[k].TargetDataType;
                        dataRow.Append(cellData);
                        cellData.CellValue = columnDefs[k].GetValueFunc(record);
                        cellData.StyleIndex = styleIndexies[k];
                    }
                }

                SetColumnWidth(columnDefs, worksheetPart, sheetData, styleIndexies);

                AppendNewSheet(spreadsheetDocument, worksheetPart);
            }

            stream.Position = 0;
            return stream;
        }
        public Stream GenerateReport<TMaster, TDetail>(IList<TMaster> masters, ColumnDefBase[] columnDefs, Func<TMaster, IList<TDetail>> getDetailFunc)
        {
            var stream = new MemoryStream();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                WorkbookStylesPart sp = workbookpart.AddNewPart<WorkbookStylesPart>("rId3");
                uint[] styleIndexies = AddCellStyles(columnDefs, sp);
                CellFormat cellFormat = sp.Stylesheet.CellFormats.ChildElements[(int)styleIndexies[0]] as CellFormat;

                #region captain 
                Row captainRow = new Row();
                captainRow.RowIndex = (UInt32)1;
                int captainSyleIndex = AddCaptainFormat(sp.Stylesheet, cellFormat.BorderId);
                for (int i = 0; i < columnDefs.Length; i++)
                {

                    Cell cellCaptain = new Cell();

                    cellCaptain.CellReference = $"{(char)((int)'A' + i)}{captainRow.RowIndex}";
                    cellCaptain.CellValue = new CellValue(columnDefs[i].Captain);
                    cellCaptain.DataType = new EnumValue<CellValues>(CellValues.String);
                    cellCaptain.StyleIndex = (uint)captainSyleIndex;
                    captainRow.AppendChild(cellCaptain);
                }

                sheetData.AppendChild(captainRow);
                #endregion

                AssignData(masters, columnDefs, getDetailFunc, worksheetPart, sheetData, styleIndexies);

                SetColumnWidth(columnDefs, worksheetPart, sheetData, styleIndexies);

                AppendNewSheet(spreadsheetDocument, worksheetPart);

            }
            stream.Position = 0;
            return stream;
        }

        public Stream GenerateReportWithTemplate<TMaster, TDetail>(Stream targetStream, IList<TMaster> masters, ColumnDefBase[] columnDefs, Func<TMaster, IList<TDetail>> getDetailFunc, string sheetName = "Sheet1", int beginRowIndex = 2)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(targetStream, true))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(doc, sheetName);
                Worksheet worksheet;
                if (worksheetPart == null)
                {
                    worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
                    AppendNewSheet(doc, worksheetPart, sheetName);
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    worksheet = worksheetPart.Worksheet;
                }
                else
                {
                    worksheet = worksheetPart.Worksheet;
                }

                var sheetData = worksheet.GetFirstChild<SheetData>();

                WorkbookStylesPart sp = doc.WorkbookPart.WorkbookStylesPart;
                uint[] styleIndexies = AddCellStyles(columnDefs, sp);
                AssignData(masters, columnDefs, getDetailFunc, worksheetPart, sheetData, styleIndexies);
            }

            targetStream.Position = 0;
            return targetStream;
        }

        internal static WorksheetPart AppendNewSheet(SpreadsheetDocument doc, WorksheetPart worksheetPart, string sheetName = null)
        {
            Sheets sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
            {
                sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
            }
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId =
                    sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = $"Sheet{sheetId}";
            }

            sheets.AppendChild(new Sheet()
            {
                Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = sheetName
            });
            return worksheetPart;
        }

        public Stream GenerateReportWithTemplate<T>(Stream targetStream, IList<T> records, ColumnDef<T>[] columnDefs, string sheetName = "Sheet1", int beginRowIndex = 2)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(targetStream, true))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(doc, sheetName);
                SheetData sheetData;
                if (worksheetPart == null)
                {
                    worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
                    AppendNewSheet(doc, worksheetPart, sheetName);
                    sheetData = new SheetData();
                    if (worksheetPart.Worksheet == null)
                        worksheetPart.Worksheet = new Worksheet();
                    worksheetPart.Worksheet.AppendChild(sheetData);
                }
                else
                {
                    sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                }

                WorkbookStylesPart sp = doc.WorkbookPart.WorkbookStylesPart;
                uint[] styleIndexies = AddCellStyles(columnDefs, sp);


                int rowIndex = beginRowIndex - 1;
                foreach (var record in records)
                {
                    var dataRow = new Row();
                    dataRow.RowIndex = (uint)++rowIndex;
                    sheetData.Append(dataRow);

                    for (int k = 0; k < columnDefs.Length; k++)
                    {
                        Cell cellData = new Cell();
                        cellData.CellReference = $"{(char)((int)'A' + k)}{dataRow.RowIndex}";
                        cellData.DataType = columnDefs[k].TargetDataType;
                        dataRow.Append(cellData);
                        cellData.CellValue = columnDefs[k].GetValueFunc(record);
                        cellData.StyleIndex = styleIndexies[k];
                    }
                }

            }

            targetStream.Position = 0;
            return targetStream;
        }

        internal static void SetColumnWidth(ColumnDefBase[] columnDefs, WorksheetPart worksheetPart, SheetData sheetData, uint[] styleIndexies)
        {
            var col = columnDefs.FirstOrDefault(c => c.TargetDataType.Value == CellValues.Number);
            if (col == null) return;
            int numberColumnIndex = Array.IndexOf<ColumnDefBase>(columnDefs, col);
            uint numberStyleIndex = styleIndexies[numberColumnIndex];
            Columns excelColumns = AutoSize(sheetData, numberStyleIndex);

            worksheetPart.Worksheet.InsertBefore(excelColumns, sheetData);
        }

        private static uint[] AddCellStyles(ColumnDefBase[] columnDefs, WorkbookStylesPart sp)
        {
            int numberStyleIndex, dateStyleIndex, stringStyleIndex;
            CreateCellFormats(sp, out numberStyleIndex, out dateStyleIndex, out stringStyleIndex);

            return GetStylesIndexes(columnDefs, numberStyleIndex, dateStyleIndex, stringStyleIndex);
        }

        internal static uint[] GetStylesIndexes(ColumnDefBase[] columnDefs, int numberStyleIndex, int dateStyleIndex, int stringStyleIndex)
        {
            uint[] styleIndexies = new uint[columnDefs.Length];

            for (int i = 0; i < columnDefs.Length; i++)
            {
                if (columnDefs[i].TargetDataType.Value == CellValues.Number)
                {
                    styleIndexies[i] = (UInt32)numberStyleIndex;
                }
                else if (columnDefs[i].TargetDataType.Value == CellValues.Date)
                {
                    styleIndexies[i] = (UInt32)dateStyleIndex;
                }
                else
                {
                    styleIndexies[i] = (UInt32)stringStyleIndex;
                }
            }

            return styleIndexies;
        }

        internal static void CreateCellFormats(WorkbookStylesPart sp, out int numberStyleIndex, out int dateStyleIndex, out int stringStyleIndex)
        {
            if (sp.Stylesheet == null)
                sp.Stylesheet = CreateStyleSheet();
            uint allBorderStyle = AddBorderStyle(sp.Stylesheet, Directions.Left | Directions.Top | Directions.Bottom | Directions.Right);

            numberStyleIndex = AddNumberFormat(sp.Stylesheet, allBorderStyle);
            dateStyleIndex = AddDateFormat(sp.Stylesheet, allBorderStyle);
            stringStyleIndex = AddStringFormat(sp.Stylesheet, allBorderStyle);
        }

        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                 document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        internal static int AddCaptainFormat(Stylesheet stylesheet, uint borderId)
        {
            Fill fill = new Fill();

            PatternFill patternFill = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FF305496" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = 64 };

            patternFill.Append(foregroundColor1);
            patternFill.Append(backgroundColor1);

            fill.Append(patternFill);
            if (stylesheet.Fills == null)
            {
                stylesheet.Fills = new Fills();
                stylesheet.Fills.Count = 0U;
            }

            stylesheet.Fills.AppendChild<Fill>(fill);
            uint fillId = (uint)stylesheet.Fills.Count++;

            Font font = new Font() { Color = new Color() { Rgb = "FFFFFFFF" } };
            stylesheet.Fonts.AppendChild(font);

            uint fontId = stylesheet.Fonts.Count++;

            var cellFormat = new CellFormat();
            cellFormat.FontId = fontId;
            cellFormat.FillId = fillId;
            cellFormat.BorderId = borderId;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;
            cellFormat.AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Center, Horizontal = HorizontalAlignmentValues.Center });

            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

            stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)stylesheet.CellFormats.ChildElements.Count);


            stylesheet.Save();
            return stylesheet.CellFormats.ChildElements.Count - 1;
        }
        internal static int AddNumberFormat(Stylesheet stylesheet, uint borderStyleId)
        {
            if (stylesheet.NumberingFormats == null)
                stylesheet.NumberingFormats = new NumberingFormats();
            NumberingFormat nf2decimal = new NumberingFormat();
            nf2decimal.NumberFormatId = UInt32Value.FromUInt32(3453);
            nf2decimal.FormatCode = StringValue.FromString("0.00");
            stylesheet.NumberingFormats.AppendChild<NumberingFormat>(nf2decimal);

            var cellFormat = new CellFormat();
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = borderStyleId;
            cellFormat.FormatId = 0;
            cellFormat.NumberFormatId = nf2decimal.NumberFormatId;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;
            cellFormat.AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Center });

            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats();
            }
            //append cell format for cells of header row
            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);


            //update font count 
            stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)stylesheet.CellFormats.ChildElements.Count);

            stylesheet.Save();
            return stylesheet.CellFormats.ChildElements.Count - 1;

            //save the changes to the style sheet part   

        }

        internal static int AddDateFormat(Stylesheet stylesheet, uint borderStyleId)
        {
            if (stylesheet.NumberingFormats == null)
                stylesheet.NumberingFormats = new NumberingFormats();

            NumberingFormat nf2decimal = new NumberingFormat();
            nf2decimal.NumberFormatId = UInt32Value.FromUInt32(3454);
            nf2decimal.FormatCode = StringValue.FromString(@"[$-404]e""-""m""-""d""-""h"":""mm"":""ss;@");
            stylesheet.NumberingFormats.AppendChild<NumberingFormat>(nf2decimal);

            var cellFormat = new CellFormat();
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = borderStyleId;
            cellFormat.FormatId = 0;
            cellFormat.NumberFormatId = nf2decimal.NumberFormatId;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;
            cellFormat.AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Center, Horizontal = HorizontalAlignmentValues.Left });

            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

            stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)stylesheet.CellFormats.ChildElements.Count);


            stylesheet.Save();
            return stylesheet.CellFormats.ChildElements.Count - 1;

        }

        internal static void AssignData<TMaster, TDetail>(IList<TMaster> masters, ColumnDefBase[] columns, Func<TMaster, IList<TDetail>> getDetailFunc, WorksheetPart worksheetPart, SheetData sheetData, uint[] styleIndexies, int startDataRowIndex = 2)
        {
            var mergeCells = new MergeCells();
            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());

            int rowIndex = startDataRowIndex - 1;
            for (int i = 0; i < masters.Count; i++)
            {
                var master = masters[i];
                var dataRow = new Row();
                dataRow.RowIndex = (uint)++rowIndex;
                sheetData.Append(dataRow);

                uint mainRowStartIndex = dataRow.RowIndex;

                var details = getDetailFunc(master);
                for (int j = 0; j < details.Count(); j++)
                {
                    if (j > 0)
                    {
                        dataRow = new Row();
                        dataRow.RowIndex = (uint)++rowIndex;
                        sheetData.Append(dataRow);
                    }
                    var detail = details[j];
                    for (int k = 0; k < columns.Length; k++)
                    {
                        //if (j > 0 && columns[k] is ColumnDef<TMaster>) continue;

                        Cell cellData = new Cell();
                        cellData.CellReference = $"{(char)((int)'A' + k)}{dataRow.RowIndex}";
                        cellData.DataType = columns[k].TargetDataType;
                        cellData.StyleIndex = styleIndexies[k];
                        dataRow.Append(cellData);

                        if (columns[k] is ColumnDef<TMaster>)
                        {
                            cellData.CellValue = ((ColumnDef<TMaster>)columns[k]).GetValueFunc(master);
                        }
                        else
                        {
                            cellData.CellValue = ((ColumnDef<TDetail>)columns[k]).GetValueFunc(detail);
                        }

                    }
                }

                //merge
                if (mainRowStartIndex < rowIndex)
                {
                    for (int k = 0; k < columns.Length; k++)
                    {
                        if (columns[k] is ColumnDef<TDetail>) continue;

                        string cloumnReference = $"{(char)((int)'A' + k)}";

                        MergeCell mergeCell = new MergeCell()
                        {
                            Reference =
                            new StringValue($"{cloumnReference}{mainRowStartIndex}" + ":" + $"{cloumnReference}{rowIndex}")
                        };
                        mergeCells.Append(mergeCell);

                    }
                }

            }
        }

        internal static int AddStringFormat(Stylesheet stylesheet, uint borderStyleId)
        {
            var cellFormat = new CellFormat();
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = borderStyleId;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;
            cellFormat.AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Center, Horizontal = HorizontalAlignmentValues.Left });

            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

            stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)stylesheet.CellFormats.ChildElements.Count);


            stylesheet.Save();
            return stylesheet.CellFormats.ChildElements.Count - 1;
        }

        internal static uint AddBorderStyle(Stylesheet stylesheet, Directions borders)
        {
            Border border = new Border();

            Color borderColor = new Color() { Indexed = (UInt32Value)64U };
            if ((borders & Directions.Left) == Directions.Left)
            {
                border.LeftBorder = new LeftBorder()
                {
                    Style = BorderStyleValues.Thin,
                    Color = new Color() { Indexed = (UInt32Value)64U }
                };
            }

            if ((borders & Directions.Right) == Directions.Right)
            {
                border.RightBorder = new RightBorder()
                {
                    Style = BorderStyleValues.Thin,
                    Color = new Color()
                    {
                        Indexed = (UInt32Value)64U
                    }
                };

            }

            if ((borders & Directions.Bottom) == Directions.Bottom)
            {
                border.BottomBorder = new BottomBorder()
                {
                    Style = BorderStyleValues.Thin,
                    Color = new Color()
                    {
                        Indexed = (UInt32Value)64U
                    }
                };

            }

            if ((borders & Directions.Top) == Directions.Top)
            {
                border.TopBorder = new TopBorder()
                {
                    Style = BorderStyleValues.Thin,
                    Color = new Color()
                    {
                        Indexed = (UInt32Value)64U
                    }
                };
            }

            border.Append(new DiagonalBorder());

            if (stylesheet.Borders == null)
            {
                stylesheet.Borders = new DocumentFormat.OpenXml.Spreadsheet.Borders();
                stylesheet.Borders.Count = (uint)0;
            }
            stylesheet.Borders.Append(border);

            return stylesheet.Borders.Count++;
        }
        internal static Stylesheet CreateStyleSheet()
        {
            Console.WriteLine("Creating styles");
            var stylesheet = new Stylesheet();
            // blank font list
            stylesheet.Fonts = new Fonts();
            stylesheet.Fonts.Count = 1;
            stylesheet.Fonts.AppendChild(new Font());

            // create fills
            stylesheet.Fills = new Fills();

            // create a solid red fill
            var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
            solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }; // red fill
            solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };

            stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
            stylesheet.Fills.Count = 3;

            // blank border list
            stylesheet.Borders = new Borders();
            stylesheet.Borders.Count = 1;
            stylesheet.Borders.AppendChild(new Border());

            // blank cell format list
            stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesheet.CellStyleFormats.Count = 1;
            stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            // cell format list
            stylesheet.CellFormats = new CellFormats();
            // empty one for index 0, seems to be required
            stylesheet.CellFormats.AppendChild(new CellFormat());
            // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true, })
                                  .AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Center });
            stylesheet.CellFormats.Count = 2;

            return stylesheet;
        }

        private static Columns AutoSize(SheetData sheetData, uint numberStyleIndex)
        {
            var maxColWidth = GetMaxCharacterWidth(sheetData, numberStyleIndex);

            Columns columns = new Columns();
            //this is the width of the font
            double maxWidth = 1;
            foreach (var item in maxColWidth)
            {
                double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;

                double pixels = Math.Truncate(((256 * width + Math.Truncate(128 / maxWidth)) / 256) * maxWidth);

                double charWidth = Math.Truncate((pixels - 5) / maxWidth * 100 + 0.5) / 100;

                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
                columns.Append(col);
            }

            return columns;
        }


        private static Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData, uint numberStyleIndex)
        {
            //iterate over all cells getting a max char value for each column
            Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
            var rows = sheetData.Elements<Row>();
            foreach (var r in rows)
            {
                var cells = r.Elements<Cell>().ToArray();

                //using cell index as my column
                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                    var cellTextLength = cellValue.Length;

                    if (cell.StyleIndex != null && cell.StyleIndex == numberStyleIndex)
                    {
                        int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                        //add 3 for '.00' 
                        cellTextLength += (3 + thousandCount);
                    }

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }

    }
}

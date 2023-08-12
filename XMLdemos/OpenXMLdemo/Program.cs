using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace OpenXMLdemo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            TestModelList tmList = new()
            {
                testData = new List<TestModel>()
            };
            TestModel tm = new()
            {
                TestId = 1,
                TestName = "Test1",
                TestDesc = "Tested 1 time",
                TestDate = DateTime.Now.Date
            };
            tmList.testData.Add(tm);

            TestModel tm1 = new()
            {
                TestId = 2,
                TestName = "Test2",
                TestDesc = "Tested 2 times",
                TestDate = DateTime.Now.AddDays(-1)
            };
            tmList.testData.Add(tm1);

            TestModel tm2 = new()
            {
                TestId = 3,
                TestName = "Test3",
                TestDesc = "Tested 3 times",
                TestDate = DateTime.Now.AddDays(-2)
            };
            tmList.testData.Add(tm2);

            TestModel tm3 = new()
            {
                TestId = 4,
                TestName = "Test4",
                TestDesc = "Tested 4 times",
                TestDate = DateTime.Now.AddDays(-3)
            };
            tmList.testData.Add(tm);

            Program p = new();
            p.CreateExcelFile(tmList, "C:\\Users\\goncalo.amador\\source\\repos");
        }

        public void CreateExcelFile(TestModelList data, string OutPutFileDirectory)
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            string fileFullname = Path.Combine(OutPutFileDirectory, "OpenXMLdemo_Output.xlsx");

            if (File.Exists(fileFullname))
            {
                fileFullname = Path.Combine(OutPutFileDirectory, "OpenXMLdemo_Output_" + datetime + ".xlsx");
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data);
            }
        }

        private void CreatePartsForExcel(SpreadsheetDocument document, TestModelList data)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }

        private void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new();
            Sheets sheets1 = new();
            Sheet sheet1 = new() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }

        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1)
        {
            Worksheet worksheet1 = new() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new() { Reference = "A1" };

            SheetViews sheetViews1 = new();

            SheetView sheetView1 = new() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet1;
        }

        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font1 = new();
            FontSize fontSize1 = new() { Val = 11D };
            Color color1 = new() { Theme = (UInt32Value)1U };
            FontName fontName1 = new() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new() { Val = 2 };
            FontScheme fontScheme1 = new() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new();
            Bold bold1 = new();
            FontSize fontSize2 = new() { Val = 11D };
            Color color2 = new() { Theme = (UInt32Value)1U };
            FontName fontName2 = new() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new() { Val = 2 };
            FontScheme fontScheme2 = new() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new() { Count = (UInt32Value)2U };

            Fill fill1 = new();
            PatternFill patternFill1 = new() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new();
            PatternFill patternFill2 = new() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new() { Count = (UInt32Value)2U };

            Border border1 = new();
            LeftBorder leftBorder1 = new();
            RightBorder rightBorder1 = new();
            TopBorder topBorder1 = new();
            BottomBorder bottomBorder1 = new();
            DiagonalBorder diagonalBorder1 = new();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new();

            LeftBorder leftBorder2 = new() { Style = BorderStyleValues.Thin };
            Color color3 = new() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new() { Style = BorderStyleValues.Thin };
            Color color4 = new() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new() { Style = BorderStyleValues.Thin };
            Color color5 = new() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new() { Style = BorderStyleValues.Thin };
            Color color6 = new() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat4 = new() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new();

            StylesheetExtension stylesheetExtension1 = new() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        private SheetData GenerateSheetdataForDetails(TestModelList data)
        {
            SheetData sheetData1 = new();
            sheetData1.Append(CreateHeaderRowForExcel());

            foreach (TestModel testmodel in data.testData)
            {
                Row partsRows = GenerateRowForChildPartDetail(testmodel);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }

        private Row CreateHeaderRowForExcel()
        {
            Row workRow = new();
            workRow.Append(CreateCell("Test Id", 2U));
            workRow.Append(CreateCell("Test Name", 2U));
            workRow.Append(CreateCell("Test Description", 2U));
            workRow.Append(CreateCell("Test Date", 2U));
            return workRow;
        }

        private Row GenerateRowForChildPartDetail(TestModel testmodel)
        {
            Row tRow = new();
            tRow.Append(CreateCell(testmodel.TestId.ToString()));
            tRow.Append(CreateCell(testmodel.TestName));
            tRow.Append(CreateCell(testmodel.TestDesc));
            tRow.Append(CreateCell(testmodel.TestDate.ToShortDateString()));

            return tRow;
        }

        private Cell CreateCell(string text)
        {
            Cell cell = new()
            {
                StyleIndex = 1U,
                DataType = ResolveCellDataTypeOnValue(text),
                CellValue = new CellValue(text)
            };
            return cell;
        }

        private Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new()
            {
                StyleIndex = styleIndex,
                DataType = ResolveCellDataTypeOnValue(text),
                CellValue = new CellValue(text)
            };
            return cell;
        }

        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            if (int.TryParse(text, out int intVal) || double.TryParse(text, out double doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
    }
}

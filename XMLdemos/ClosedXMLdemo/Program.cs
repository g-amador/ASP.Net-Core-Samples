using ClosedXML.Excel;
using System;

namespace ClosedXMLdemo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Program p = new();
            p.CreateExcelFile("C:\\Users\\goncalo.amador\\source\\repos\\ClosedXMLdemo_Output.xlsx");
        }

        public void CreateExcelFile(String filePath)
        {
            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sample Sheet");

            ws.Cell(2, 2).Value = "Hello World!";

            wb.SaveAs(filePath);
        }
    }
}

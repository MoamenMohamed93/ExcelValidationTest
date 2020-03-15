using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace ExcelValidationTest
{
    class Program
    {
        private const string Path = "./test.xlsx";

        static void Main()
        {
            GenerateExcel();
            ValidateExcel();
        }


        private static void GenerateExcel()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(Path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();

                workbookPart.Workbook = new Workbook();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                // create sheet data
                var sheetData = worksheetPart.Worksheet.
                  AppendChild(new SheetData());

                sheetData.AppendChild(new Row(new Cell()
                {
                    CellValue = new CellValue("StringValue"),
                    DataType = CellValues.Number
                }));

                worksheetPart.Worksheet.Save();

                document.WorkbookPart.Workbook.Sheets = new Sheets();

                document.WorkbookPart.Workbook.Sheets.AppendChild(new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)1,
                    Name = "MyFirstSheet"
                });
                document.WorkbookPart.Workbook.Save();
            }
        }
        private static void ValidateExcel()
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            int count = 0;
            foreach (ValidationErrorInfo error in validator.Validate(SpreadsheetDocument.Open(Path, true)))
            {
                count++;
                Console.WriteLine("Error " + count);
                Console.WriteLine("Description: " + error.Description);
                Console.WriteLine("Path: " + error.Path.XPath);
                Console.WriteLine("Part: " + error.Part.Uri);
                Console.WriteLine("-------------------------------------------");
            }
        }
    }
}

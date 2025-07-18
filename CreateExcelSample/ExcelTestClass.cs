using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CreateExcelSample;

/// <summary>
/// Just a test class to create an Excel spreadsheet using DocumentFormat.OpemXML
/// </summary>
public class ExcelTestClass
{
    /// <summary>
    /// A method to create a test Excel document
    /// </summary>
    /// <param name="fileName">String value of path and filename to create</param>
    /// <param name="dataList">List<FakeDataRecord> that contains the data to use</param>
    public void CreateExcelDocument(string fileName, List<FakeDataRecord> dataList)
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
        {
            // Add a WorkbookPart to the document
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook
            Sheet sheet = new()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "ExcelTab1"   // name of tab
            };
            sheets.Append(sheet);

            // // Access the SheetData to write your data
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // a row for the header values
            Row headerRow = new()
            {
                RowIndex = 1, // first row, apparently no need to increment below
            };

            // header row values
            headerRow.Append(new Cell()
            {
                CellValue = new CellValue("Value1"),
                DataType = CellValues.String,
            });
            headerRow.Append(new Cell()
            {
                CellValue = new CellValue("Value2"),
                DataType = CellValues.String
            });
            headerRow.Append(new Cell()
            {
                CellValue = new CellValue("Value3"),
                DataType = CellValues.String
            });
            headerRow.Append(new Cell()
            {
                CellValue = new CellValue("Value4"),
                DataType = CellValues.String
            });

            // add the header row
            sheetData.Append(headerRow);

            // Loop through some data and add rows and cells
            foreach (FakeDataRecord rec in dataList)
            {
                // a row for the DSEMS data
                Row dataRow = new();

                // CM#
                dataRow.Append(new Cell()
                {
                    CellValue = new CellValue(rec.Value1),
                    DataType = CellValues.String
                });

                // Release Name
                dataRow.Append(new Cell()
                {
                    CellValue = new CellValue(rec.Value2),
                    DataType = CellValues.String
                });

                // Prefix
                dataRow.Append(new Cell()
                {
                    CellValue = new CellValue(rec.Value3),
                    DataType = CellValues.String
                });

                // Program
                dataRow.Append(new Cell()
                {
                    CellValue = new CellValue(rec.Value4),
                    DataType = CellValues.String
                });

                // add the data row
                sheetData.Append(dataRow);
            }
            // save the document
            workbookPart.Workbook.Save();
        }
    }
}
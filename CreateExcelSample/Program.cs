using CreateExcelSample;

// A List of FakeDataRecord objects
List<FakeDataRecord> dataList = new List<FakeDataRecord>();

// A loop to build some test records
for (int i = 0; i < 299; i++)
{
    // NOTE: FakeDataReader is faking something like an OracleDataReader
    // Get a row of fake data & add to the dataList object
    dataList.Add(FakeDataReader.Read());
}

// An Excel class to build a spreadsheet, uses DocumentFormat.OpenXml SDK
ExcelTestClass myExcel = new();

// Pass the List<FakeDataRecord> to build a spreadsheet
myExcel.CreateExcelDocument("/home/ronmcgough/gitHubProjects/CreateExcelSample/Fake_Report.xlsx", dataList);

Console.WriteLine("Done");
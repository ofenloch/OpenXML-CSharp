using System;
using DocumentFormat.OpenXml; // dotnet add package DocumentFormat.OpenXml
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            string sSpreadsheetFileName = "hello.xlsx";


            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(sSpreadsheetFileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document (A SpreadsheetDocument must have at least a WorkbookPart and a WorkSheetPart.)
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart (A SpreadsheetDocument must have at least a WorkbookPart and a WorkSheetPart.)
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // The Worksheet will contain SheetData and Columns. It is the Sheet data which the actual values 
            // goes in rows and cells. By initializing the Worksheet we can append a SheetData as its child 
            // by passing it as argument.
            // Append a “Sheets” to the Workbook. The Sheets will contain one or many Sheet which each associated 
            // with a WorksheetPart.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Hello, World!"
            };
            sheets.Append(sheet);

            // At the end save the Workbook.
            workbookpart.Workbook.Save();

            // Append a SheetData class to the worksheet. The SheetData acts as the container where all the rows and columns will be going.
            SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

            // Constructing header row
            Row headerRow = new Row();
            // Add cells to the header row
            headerRow.Append(
                ConstructCell("Id", CellValues.String),
                ConstructCell("Name", CellValues.String),
                ConstructCell("Birth Date", CellValues.String));
            // Insert the header row to the Sheet Data
            sheetData.AppendChild(headerRow);

            // Constructing data row
            Row dataRow = new Row();
            // Add cells to the data row
            dataRow.Append(
                ConstructCell("1", CellValues.String),
                ConstructCell("Oliver Ofenloch", CellValues.String),
                ConstructCell("2020-08-24", CellValues.String));
            // Insert the header row to the Sheet Data
            sheetData.AppendChild(dataRow);

            // Close the document.
            spreadsheetDocument.Close();


        } // static void Main(string[] args)

        static public Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }


        // This is from the MS docs at https://docs.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name#sample-code
        public static void CreateSpreadsheetWorkbook(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        } // public static void CreateSpreadsheetWorkbook(string filepath)

    } // class Program
} // namespace OpenXML

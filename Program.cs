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

            // Create a SpreadsheetDocument
            SpreadsheetDocument spreadsheetDocument = createSpreadsheetWorkbook(sSpreadsheetFileName);

            // Close the SpreadsheetDocument.
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

        static public Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        // This is from http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-3-add-stylesheet-to-the-spreadsheet/
        static public Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true } // header
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }


        public static SpreadsheetDocument createSpreadsheetWorkbook(string filepath) {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

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

            // Add a WorkbookStylePart to the WorkbookPart and initialize its Stylesheet:
            WorkbookStylesPart stylePart = workbookpart.AddNewPart<WorkbookStylesPart>();
            stylePart.Stylesheet = GenerateStylesheet();
            stylePart.Stylesheet.Save();

            // Setting up columns
            // Adding custom width to specific columns is very easy. First we need to create a Columns 
            // object and then add one or more Column as its children which each will define the custom 
            // width of a range of columns in the spreadsheet. You can explore the properties of the 
            // column to specify more customization on columns. Here we are only interested in specifying 
            // the width of the columns.
            Columns columns = new Columns(
                    new Column // Id column
                        {
                        Min = 1,
                        Max = 1,
                        Width = 4,
                        CustomWidth = true
                    },
                    new Column // Name and Birthday columns
                        {
                        Min = 2,
                        Max = 3,
                        Width = 15,
                        CustomWidth = true
                    });
            worksheetPart.Worksheet.AppendChild(columns);

            // At the end save the Workbook.
            workbookpart.Workbook.Save();

            // Append a SheetData class to the worksheet. The SheetData acts as the container where all the rows and columns will be going.
            SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

            // Constructing header row
            Row headerRow = new Row();
            // Add cells to the header row
            headerRow.Append(
                ConstructCell("Id", CellValues.String, 2),
                ConstructCell("Name", CellValues.String, 2),
                ConstructCell("Birth Date", CellValues.String, 2));
            // Insert the header row to the Sheet Data
            sheetData.AppendChild(headerRow);

            // Constructing data row
            Row dataRow = new Row();
            // Add cells to the data row
            dataRow.Append(
                ConstructCell("1", CellValues.String, 1),
                ConstructCell("Oliver Ofenloch", CellValues.String, 1),
                ConstructCell("2020-08-24", CellValues.String, 1));
            // Insert the data row to the Sheet Data
            sheetData.AppendChild(dataRow);

            return spreadsheetDocument;
        } // public static SpreadsheetDocument createSpreadsheetWorkbook(string filepath)


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

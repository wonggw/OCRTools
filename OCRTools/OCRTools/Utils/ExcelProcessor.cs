using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OCRTools.Utils
{
    class ExcelProcessor
    {
        public static void CreateSpreadsheetWorkbook(Stream filepath , string text)
        {

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);
            WriteTextToCells(worksheetPart, text);
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        public static void WriteTextToCells(WorksheetPart worksheetPart, string text)
        {
            List<string> textPhase = text.Split("\n").ToList();
            List<List<string>> textWord = new List<List<string>>();
            foreach (var phase in textPhase)
            {
                textWord.Add(phase.Split(" ").ToList());
            }

            // Get the sheetData cell table.
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            char columnNameStart = 'A';
            char columnName = columnNameStart;
            uint rowNumber = 1;
            foreach (var phase in textWord)
            {
                foreach (var word in phase)
                {
                    string cellAddress = String.Concat(columnName, rowNumber);
                    // Add a row to the cell table.
                    Row row;
                    row = new Row() { RowIndex = rowNumber };
                    sheetData.Append(row);

                    // In the new row, find the column location to insert a cell.
                    Cell refCell = null;
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (string.Compare(cell.CellReference.Value, cellAddress, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }

                    // Add the cell to the cell table.
                    Cell newCell = new Cell() { CellReference = cellAddress };
                    row.InsertBefore(newCell, refCell);
                    // Set the cell value to be a numeric value.
                    newCell.CellValue = new CellValue(word);
                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);

                    int tempColumn = (int)columnName;
                    columnName = (char)++tempColumn;
                }
                columnName = columnNameStart;
                ++rowNumber;
            }

        }
    }
}

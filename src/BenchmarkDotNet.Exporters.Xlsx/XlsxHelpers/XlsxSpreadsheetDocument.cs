using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// A wrapper for <see cref="SpreadsheetDocument"/>. 
    /// </summary>
    public class XlsxSpreadsheetDocument
    {
        private readonly SpreadsheetDocument _spreadsheetDocument;

        public XlsxSpreadsheetDocument(SpreadsheetDocument spreadsheetDocument)
        {
            _spreadsheetDocument = spreadsheetDocument ?? throw new ArgumentNullException(nameof(spreadsheetDocument));
        }

        /// <summary>
        /// Initializes a workbook.
        /// </summary>
        public void InitializeWorkbook()
        {
            if (_spreadsheetDocument.WorkbookPart is null)
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = _spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook
                {
                    Sheets = new Sheets()
                };
            }
        }
        
        /// <summary>
        /// Adds a new sheet.
        /// </summary>
        /// <param name="name">The sheet name.</param>
        public XlsxSheet AddSheet(string name)
        {
            Guard.EnsureNotEmpty(name, nameof(name));

            var workbookPart = _spreadsheetDocument.WorkbookPart;
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var workSheet = new Worksheet();
            var sheetData = new SheetData();
            workSheet.AppendChild(sheetData);
            worksheetPart.Worksheet = workSheet;

            var sheetId = 1u;
            var allSheets = workbookPart.Workbook.Sheets.OfType<Sheet>();
            if (allSheets.Any())
            {
                sheetId = allSheets.Max(sh => sh.SheetId.Value) + 1;
            }

            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = name
            };
            workbookPart.Workbook.Sheets.Append(sheet);
            return new XlsxSheet(worksheetPart, sheetData, sheet);
        }

        /// <summary>
        /// Gets a sheet by name.
        /// </summary>
        /// <param name="name">The name.</param>
        public XlsxSheet GetSheet(string name)
        {
            Guard.EnsureNotEmpty(name, nameof(name));

             var workbookPart = _spreadsheetDocument.WorkbookPart;
            var sheets = workbookPart.Workbook.Sheets.OfType<Sheet>();
            var sheet = sheets.SingleOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (sheet is null)
            {
                return null;
            }
            else
            {
                var workbooksheets = workbookPart.GetPartsOfType<WorksheetPart>();
                var worksheetPart = workbooksheets.SingleOrDefault(wsp => workbookPart.GetIdOfPart(wsp) == sheet.Id);
                var workSheet = worksheetPart.Worksheet;
                var worksheet = workSheet.OfType<SheetData>().Single();
                return new XlsxSheet(worksheetPart, worksheet, sheet);
            }
        }

        public void Save()
        {
            _spreadsheetDocument.Save();
        }

    }
}

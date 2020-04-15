using BenchmarkDotNet.Exporters.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    /// <summary>
    /// Encapsulates common methods to test parts of the the XlsxSpreadsheetDocument.
    /// </summary>
    public class XlsxSpreadsheetDocumentFactoryTest
    {
        public void SetupXlsxSpreadSheetDocumentWithRow(Action<XlsxRow> func)
        {
            SetupXlsxSpreadSheetDocumentWithSheet(sheet =>
            {
                var row = sheet.AddNewRow();
                func(row);
            });
        }

        public void SetupXlsxSpreadSheetDocumentWithSheet(Action<XlsxSheet> func)
        {
            SetupXlsxSpreadSheetDocument(spreadsheet =>
            {
                var sheet = spreadsheet.AddSheet("Test");
                func(sheet);
            });
        }

        public void SetupXlsxSpreadSheetDocument(Action<XlsxSpreadsheetDocument> func)
        {
            if (func is null)
            {
                throw new ArgumentNullException(nameof(func));
            }
            using (var ms = new MemoryStream())
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    var spreadsheet = new XlsxSpreadsheetDocument(spreadsheetDocument);
                    spreadsheet.InitializeWorkbook();
                    spreadsheet.Save();
                    func(spreadsheet);
                }
            }
        }
    }
}


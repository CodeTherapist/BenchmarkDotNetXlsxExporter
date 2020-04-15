using BenchmarkDotNet.Exporters.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class SpreadsheetDocumentTests
    {
        [Fact]
        public void SimpleTest()
        {
            var ex = Record.Exception(() =>
            {
                using (var ms = new MemoryStream())
                {
                    using (var spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                    {
                        var spreadsheet = new XlsxSpreadsheetDocument(spreadsheetDocument);
                        spreadsheet.InitializeWorkbook();
                        spreadsheet.Save();
                    }
                    Assert.True(ms.Length > 0);
                }
            });
            Assert.Null(ex);
        }

        [Fact]
        public void SheetTests()
        {
            var ex = Record.Exception(() =>
            {
                using (var ms = new MemoryStream())
                {
                    using (var spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                    {
                        var spreadsheet = new XlsxSpreadsheetDocument(spreadsheetDocument);
                        spreadsheet.InitializeWorkbook();
                        spreadsheet.Save();

                        var addSheet = spreadsheet.AddSheet("test");
                        var getSheet = spreadsheet.GetSheet("test");

                        Assert.Same(getSheet.SheetName, addSheet.SheetName);
                        Assert.Equal(getSheet.SheetIndex, addSheet.SheetIndex);
                    }
                }

            });
            Assert.Null(ex);
        }
    }
}


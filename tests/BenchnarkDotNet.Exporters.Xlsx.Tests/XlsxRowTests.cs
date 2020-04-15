using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class XlsxRowTests : IClassFixture<XlsxSpreadsheetDocumentFactoryTest>
    {
        public XlsxRowTests(XlsxSpreadsheetDocumentFactoryTest xlsxSpreadsheetDocumentFactoryTest)
        {
            XlsxSpreadsheetDocumentFactoryTest = xlsxSpreadsheetDocumentFactoryTest;
        }

        private XlsxSpreadsheetDocumentFactoryTest XlsxSpreadsheetDocumentFactoryTest { get; }

        [Fact]
        public void SetCellValueDoesNotThrow()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithRow((row) =>
            {
                Assert.Null(Record.Exception(() => row.SetCellValue(2, "text")));
            });
        }

        [Fact]
        public void CreateTextCell()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithRow((row) =>
            {
                var xlsxCell = row.GetOrCreateCell(1);
                xlsxCell.SetValue("Test");
                Assert.Equal(CellValues.String, xlsxCell.DataType);
            });
        }
    }
}


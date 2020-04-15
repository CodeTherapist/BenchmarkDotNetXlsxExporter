using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class XlsxSheetTests : IClassFixture<XlsxSpreadsheetDocumentFactoryTest>
    {
        public XlsxSheetTests(XlsxSpreadsheetDocumentFactoryTest xlsxSpreadsheetDocumentFactoryTest)
        {
            XlsxSpreadsheetDocumentFactoryTest = xlsxSpreadsheetDocumentFactoryTest;
        }

        private XlsxSpreadsheetDocumentFactoryTest XlsxSpreadsheetDocumentFactoryTest { get; }

        [Fact]
        public void AddRowToSheetDoesNotThrow()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithSheet((sheet) =>
            {
                Assert.Null(Record.Exception(() => sheet.AddNewRow()));
            });
        }

        [Fact]
        public void GetNextRowIndexDoesNotThrow()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithSheet((sheet) =>
            {
                Assert.Null(Record.Exception(() => sheet.GetNextRowIndex()));
            });
        }

        [Fact]
        public void GetNextRowIndexMustIncreaseWhenRowIsAdded()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithSheet((sheet) =>
            {
                var nextRowIndex = sheet.GetNextRowIndex();
                Assert.Equal(1u, nextRowIndex);
                sheet.AddNewRow();
                var nextRowIndexAfterNewRow = sheet.GetNextRowIndex();
                Assert.Equal(2u, nextRowIndexAfterNewRow);
            });
        }


        [Fact]
        public void GetOrCreateRowDoesNotThrow()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithSheet((sheet) =>
            {
                Assert.Null(Record.Exception(() => sheet.GetOrCreateRow(1)));
            });
        }


        [Fact]
        public void GetOrCreateRowDoes()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithSheet((sheet) =>
            {
                Assert.Null(Record.Exception(() => sheet.GetOrCreateRow(1)));
            });
        }

    }
}


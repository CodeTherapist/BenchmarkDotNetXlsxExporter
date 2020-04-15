using DocumentFormat.OpenXml.Spreadsheet;
using System;
using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class XlsxCellTests : IClassFixture<XlsxSpreadsheetDocumentFactoryTest>
    {
        public XlsxCellTests(XlsxSpreadsheetDocumentFactoryTest xlsxSpreadsheetDocumentFactoryTest)
        {
            XlsxSpreadsheetDocumentFactoryTest = xlsxSpreadsheetDocumentFactoryTest;
        }

        private XlsxSpreadsheetDocumentFactoryTest XlsxSpreadsheetDocumentFactoryTest { get; }

        [Fact]
        public void CellAndCellReferenceIsNotNullWhenCreatingACell()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithRow((row) =>
            {
                var xlsxCell = row.GetOrCreateCell(1);
                Assert.NotNull(xlsxCell.Cell);
                Assert.NotNull(xlsxCell.CellReference);
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

        [Fact]
        public void CreateDateTimeCell()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithRow((row) =>
            {
                var xlsxCell = row.GetOrCreateCell(1);
                xlsxCell.SetValue(DateTime.Now);
                Assert.Equal(CellValues.Date, xlsxCell.DataType);
            });
        }

        [Fact]
        public void CreateNumberCell()
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithRow((row) =>
            {
                var xlsxCell = row.GetOrCreateCell(1);
                xlsxCell.SetValue(578);
                Assert.Equal(CellValues.Number, xlsxCell.DataType);
            });
        }

        [Theory()]
        [InlineData(true)]
        [InlineData(false)]
        public void CreateBooleanCell(bool value)
        {
            XlsxSpreadsheetDocumentFactoryTest.SetupXlsxSpreadSheetDocumentWithRow((row) =>
            {
                var xlsxCell = row.GetOrCreateCell(1);
                xlsxCell.SetValue(value);
                Assert.Equal(CellValues.Boolean, xlsxCell.DataType);
            });
        }
    }
}


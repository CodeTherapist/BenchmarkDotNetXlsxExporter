using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// A wrapper for <see cref="Sheet"/>.
    /// </summary>
    public class XlsxSheet
    {
        private readonly WorksheetPart _worksheetPart;
        private readonly SheetData _sheetData;
        private readonly Sheet _sheet;

        public XlsxSheet(WorksheetPart worksheetPart, SheetData sheetData, Sheet sheet)
        {
            _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
            _sheetData = sheetData ?? throw new ArgumentNullException(nameof(sheetData));
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        }

        /// <summary>
        /// Gets the underlying sheet data object.
        /// </summary>
        public SheetData SheetData => _sheetData;

        /// <summary>
        /// Gets the sheet name.
        /// </summary>
        public string SheetName => _sheet.Name;

        /// <summary>
        /// Gets the index of the sheet in the list of sheets.
        /// </summary>
        public uint SheetIndex => _sheet.SheetId.Value;

        /// <summary>
        /// Gets the next available row index (LastRow + 1).
        /// </summary>
        public uint GetNextRowIndex()
        {
            var rows = _sheetData.Elements<Row>();

            if (rows.Any())
            {
                return rows.Max(r => r.RowIndex.Value) + 1;
            }
            return 1;
        }

        /// <summary>
        /// Adds a new row.
        /// </summary>
        public XlsxRow AddNewRow()
        {
            var index = GetNextRowIndex();
            return GetOrCreateRow(index);
        }

        /// <summary>
        /// Gets or creates a new row.
        /// </summary>
        /// <param name="rowIndex">The row index.</param>
        public XlsxRow GetOrCreateRow(uint rowIndex)
        {
            var row = _sheetData.Elements<Row>().SingleOrDefault(r => r.RowIndex == rowIndex);
            if (row is null)
            {
                return CreateRow(rowIndex);
            }
            return new XlsxRow(row);
        }

        private XlsxRow CreateRow(uint rowIndex)
        {
            var row = new Row() { RowIndex = rowIndex };
            _sheetData.Append(row);
            return new XlsxRow(row);
        }
    }
}

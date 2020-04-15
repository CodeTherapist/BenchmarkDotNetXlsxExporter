using System.Collections.Generic;
using System.Linq;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Reports;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// The base class for a xlsx handler that exports a set of columns to a xlsx sheet.
    /// </summary>
    public abstract class ColumnsToSheetXlsxExporterHandlerBase : XlsxExporterHandlerBase
    {
        /// <summary>
        /// Gets the name of the sheet. Should be unique within the sheet.
        /// </summary>
        public abstract string SheetName { get; }

        protected abstract IReadOnlyCollection<IColumn> GetColumns(Summary summary);

        protected override void HandleCore(XlsxSpreadsheetDocument xlsxSpreadsheetDocument, Summary summary)
        {
            var sheet = xlsxSpreadsheetDocument.AddSheet(SheetName);
            var columns = GetColumns(summary);
            var relevantColumns = summary.Table.Columns
                .Where(c => columns.Any(sc => string.Equals(c.OriginalColumn.Id, sc.Id, System.StringComparison.OrdinalIgnoreCase)))
                .ToArray();

            var headerColIndex = 0;
            var headerRow = sheet.GetOrCreateRow(1);
            foreach (var col in relevantColumns)
            {
                headerColIndex++;
                headerRow.SetCellValue(headerColIndex, col.Header);
            }

            var rowIndex = 1u;
            var colIndex = 1;
            foreach (var col in relevantColumns)
            {
                foreach (var val in col.Content)
                {
                    var contentRow = sheet.GetOrCreateRow(rowIndex + 1);
                    contentRow.SetCellValue(colIndex, val);
                    rowIndex++;
                }
                colIndex++;
                rowIndex = 1u;
            }
        }
    }
}


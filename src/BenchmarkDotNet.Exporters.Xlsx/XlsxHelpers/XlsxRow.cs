using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// A wrapper for xlsx rows.
    /// </summary>
    public class XlsxRow 
    {
        private readonly Row _row;

        public XlsxRow(Row row)
        {
            _row = row ?? throw new ArgumentNullException(nameof(row));
        }

        /// <summary>
        /// Gets the underlying xlsx object.
        /// </summary>
        public Row Row => _row;

        /// <summary>
        /// Create or returns a cell.
        /// </summary>
        /// <param name="columnIndex">The column index.</param>
        public XlsxCell GetOrCreateCell(int columnIndex)
        {
            var cellReference = $"{ColumnNameHelper.GetExcelColumnName(columnIndex)}{_row.RowIndex}";
            var cell = _row.Elements<Cell>().SingleOrDefault(c => c.CellReference.Value == cellReference);
            if (cell is null)
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell c in _row.Elements<Cell>())
                {
                    if (c.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(c.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = c;
                            break;
                        }
                    }
                }

                cell = new Cell() { CellReference = cellReference };
                _row.InsertBefore(cell, refCell);
            }
            return new XlsxCell(cell, this);
        }

        /// <summary>
        /// Sets a value in a cell.
        /// </summary>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="textValue">The value.</param>
        public void SetCellValue(int columnIndex, string textValue)
        {
           var cell = GetOrCreateCell(columnIndex);
            cell.SetValue(textValue);
        }
    }
}

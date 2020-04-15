using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// A wrapper for <see cref="DocumentFormat.OpenXml.Spreadsheet.Cell"/>.
    /// </summary>
    public class XlsxCell
    {
        private readonly Cell _cell;
        private readonly XlsxRow _xlsxRow;

        public XlsxCell(Cell cell, XlsxRow xlsxRow)
        {
            _cell = cell ?? throw new ArgumentNullException(nameof(cell));
            _xlsxRow = xlsxRow ?? throw new ArgumentNullException(nameof(xlsxRow));
        }

        /// <summary>
        /// Gets the cell reference.
        /// </summary>
        public string CellReference => _cell.CellReference;

        /// <summary>
        /// Gets the underlying cell.
        /// </summary>
        public Cell Cell => _cell;

        /// <summary>
        /// Gets or sets the cell value type.
        /// </summary>
        public CellValues DataType => _cell.DataType;

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value. Cannot be null.</param>
        public void SetValue(string value)
        {
            _cell.CellValue = new CellValue(value);
            _cell.DataType = CellValues.String;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetValue(DateTime value)
        {
            _cell.CellValue = new CellValue(value);
            _cell.DataType = CellValues.Date;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetValue(int value)
        {
            _cell.CellValue = new CellValue(value.ToString());
            _cell.DataType = CellValues.Number;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetValue(double value)
        {
            _cell.CellValue = new CellValue(value.ToString());
            _cell.DataType = CellValues.Number;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetValue(float value)
        {
            _cell.CellValue = new CellValue(value.ToString());
            _cell.DataType = CellValues.Number;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetValue(byte value)
        {
            _cell.CellValue = new CellValue(value.ToString());
            _cell.DataType = CellValues.Number;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetValue(decimal value)
        {
            _cell.CellValue = new CellValue(value.ToString());
            _cell.DataType = CellValues.Number;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value.</param>
        public void SetValue(bool value)
        {
            _cell.CellValue = new CellValue(value.ToString());
            _cell.DataType = CellValues.Boolean;
        }

        /// <summary>
        /// Sets the value of the cell.
        /// </summary>
        /// <param name="value">The value. Cannot be null.</param>
        public void SetValue(object value)
        {
            if (value is null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            _cell.CellValue = new CellValue(value.ToString());
            _cell.DataType = CellValues.String;
        }
    }
}

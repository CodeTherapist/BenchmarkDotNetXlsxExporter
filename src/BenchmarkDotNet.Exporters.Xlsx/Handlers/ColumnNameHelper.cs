using System;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// A helper class for xlsx column names.
    /// </summary>
    public static class ColumnNameHelper
    {
        /// <summary>
        /// Gets an alpha character from an index.
        /// <para>1 => 'A'. 17 => 'Q'. '38' => 'AL'.</para>
        /// </summary>
        /// <param name="columnIndex">The column index.</param>
        public static string GetExcelColumnName(int columnIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentException("The index cannot be less than zero.", nameof(columnIndex));
            }

            int dividend = columnIndex;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
    }
}

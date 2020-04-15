using BenchmarkDotNet.Reports;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// A xlsx handler for the <see cref="XlsxExporter"/>. 
    /// </summary>
    public interface IXlsxExporterHandler
    {
        /// <summary>
        /// Extends the <paramref name="xlsxSpreadsheetDocument"/> with data from the <paramref name="summary"/>. 
        /// </summary>
        /// <param name="xlsxSpreadsheetDocument">The spreadsheet document. Cannot be null.</param>
        /// <param name="summary">The summary. Cannot be null.</param>
        void Handle(XlsxSpreadsheetDocument xlsxSpreadsheetDocument, Summary summary);
    }
}
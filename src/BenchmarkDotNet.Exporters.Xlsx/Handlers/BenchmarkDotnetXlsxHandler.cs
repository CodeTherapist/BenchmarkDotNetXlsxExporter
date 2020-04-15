using System;
using BenchmarkDotNet.Reports;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// The base class for all xlsx handler.
    /// </summary>
    public abstract class XlsxExporterHandlerBase : IXlsxExporterHandler
    {
        public void Handle(XlsxSpreadsheetDocument xlsxSpreadsheetDocument, Summary summary)
        {
            if (xlsxSpreadsheetDocument is null)
            {
                throw new ArgumentNullException(nameof(xlsxSpreadsheetDocument));
            }

            if (summary is null)
            {
                throw new ArgumentNullException(nameof(summary));
            }

            HandleCore(xlsxSpreadsheetDocument, summary);
        }

        protected abstract void HandleCore(XlsxSpreadsheetDocument xlsxSpreadsheetDocument, Summary summary);
    }
}


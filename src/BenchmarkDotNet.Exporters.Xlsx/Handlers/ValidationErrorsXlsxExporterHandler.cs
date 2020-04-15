using BenchmarkDotNet.Reports;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// This handler exports validation errors from BenchmarkDotnet.
    /// </summary>
    public class ValidationErrorsXlsxExporterHandler : XlsxExporterHandlerBase
    {
        protected override void HandleCore(XlsxSpreadsheetDocument xlsxSpreadsheetDocument, Summary summary)
        {
            var validationErrorsSheet = xlsxSpreadsheetDocument.AddSheet("ValidationErrors");
            foreach (var item in summary.ValidationErrors)
            {
                var row = validationErrorsSheet.AddNewRow();
                row.SetCellValue(1, item.Message);
                row.SetCellValue(2, item.IsCritical.ToString());
            }
        }
    }
}


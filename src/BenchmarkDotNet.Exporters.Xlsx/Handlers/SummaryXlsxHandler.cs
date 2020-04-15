using System.Collections.Generic;
using System.Linq;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Reports;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// This handler exports only visible summary columns to a sheet.
    /// </summary>
    public class SummaryXlsxHandler : ColumnsToSheetXlsxExporterHandlerBase
    {
        public override string SheetName => "Summary";

        protected override IReadOnlyCollection<IColumn> GetColumns(Summary summary)
        {
            return summary.GetColumns().Where(c => c.AlwaysShow).ToArray();
        }
    }
}


using System.Collections.Generic;
using System.Linq;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Reports;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// This handler exports all summary columns to a sheet.
    /// </summary>
    public class FullSummaryXlsxHandler : ColumnsToSheetXlsxExporterHandlerBase
    {
        public override string SheetName => "FullSummary";

        protected override IReadOnlyCollection<IColumn> GetColumns(Summary summary)
        {
            return summary.GetColumns().ToArray();
        }
    }


}


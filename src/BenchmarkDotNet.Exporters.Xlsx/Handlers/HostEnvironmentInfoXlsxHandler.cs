using System.Collections.Generic;
using BenchmarkDotNet.Environments;
using BenchmarkDotNet.Reports;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// A xlsx handler that exports host environment information.
    /// </summary>
    public class HostEnvironmentInfoXlsxHandler : XlsxExporterHandlerBase
    {
        protected override void HandleCore(XlsxSpreadsheetDocument xlsxSpreadsheetDocument, Summary summary)
        {
            var hostingEnvSheet = xlsxSpreadsheetDocument.AddSheet("HostingEnvironment");
            var infos = GetInfos(summary);
            foreach (var item in infos)
            {
                var row = hostingEnvSheet.AddNewRow();

                var key = row.GetOrCreateCell(1);
                key.SetValue(item.Key);
                var value = row.GetOrCreateCell(2);
                value.SetValue(item.Value);
            }
        }

        protected virtual Dictionary<string, string> GetInfos(Summary summary)
        {
            var env = summary.HostEnvironmentInfo;
            return new Dictionary<string, string>()
            {
                { nameof(HostEnvironmentInfo.Architecture), env.Architecture },
                { nameof(HostEnvironmentInfo.BenchmarkDotNetVersion), env.BenchmarkDotNetVersion },
                { nameof(HostEnvironmentInfo.ChronometerFrequency), env.ChronometerFrequency.ToString() },
                { nameof(HostEnvironmentInfo.ChronometerResolution), env.ChronometerResolution.ToString() },
                { nameof(HostEnvironmentInfo.Configuration), env.Configuration },
                { nameof(HostEnvironmentInfo.GCAllocationQuantum), env.GCAllocationQuantum.ToString() },
                { nameof(HostEnvironmentInfo.HardwareTimerKind), env.HardwareTimerKind.ToString() },
                { nameof(HostEnvironmentInfo.HasAttachedDebugger), env.HasAttachedDebugger.ToString() },
                { nameof(HostEnvironmentInfo.HasRyuJit), env.HasRyuJit.ToString() },
                { nameof(HostEnvironmentInfo.InDocker), env.InDocker.ToString() },
                { nameof(HostEnvironmentInfo.IsConcurrentGC), env.IsConcurrentGC.ToString() },
                { nameof(HostEnvironmentInfo.IsServerGC), env.IsServerGC.ToString() },
                { nameof(HostEnvironmentInfo.JitInfo), env.JitInfo },
                { nameof(HostEnvironmentInfo.RuntimeVersion), env.RuntimeVersion },
            };
        }
    }
}


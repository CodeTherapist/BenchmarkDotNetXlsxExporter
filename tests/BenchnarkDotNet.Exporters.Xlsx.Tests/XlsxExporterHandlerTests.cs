using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using BenchmarkDotNet.Exporters.Xlsx;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Validators;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class XlsxExporterHandlerTests
    {
        [Theory]
        [MemberData(nameof(GetHandler))]
        public void HandlerArgumentXlsxSpreadsheetDocumentCannotBeNull(IXlsxExporterHandler benchmarkDotnetXlsxHandler)
        {
            var ex = Record.Exception(() => benchmarkDotnetXlsxHandler.Handle(null,
                new Summary(string.Empty, new System.Collections.Immutable.ImmutableArray<BenchmarkReport>(),
                BenchmarkDotNet.Environments.HostEnvironmentInfo.GetCurrent(),
                string.Empty, string.Empty,
                TimeSpan.Zero, CultureInfo.CurrentCulture,
                new System.Collections.Immutable.ImmutableArray<ValidationError>())));
            Assert.NotNull(ex);
        }

        [Theory]
        [MemberData(nameof(GetHandler))]
        public void HandlerArgumentXlsxSummaryCannotBeNull(IXlsxExporterHandler benchmarkDotnetXlsxHandler)
        {
            using (var ms = new MemoryStream())
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    var ex = Record.Exception(() => benchmarkDotnetXlsxHandler.Handle(new XlsxSpreadsheetDocument(spreadsheetDocument), null));
                    Assert.NotNull(ex);
                }
            }
        }

        public static IEnumerable<object[]> GetHandler()
        {
            return new[] {
                new object[] { new SummaryXlsxHandler()},
                new object[] { new FullSummaryXlsxHandler()},
                new object[] { new HostEnvironmentInfoXlsxHandler()},
                new object[] { new ValidationErrorsXlsxExporterHandler()}
            };
        }
    }
}


using System;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Exporters.Xlsx;
using BenchmarkDotNet.Loggers;
using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class XlsxExporterTests : IClassFixture<TestBenchmarkRunner>
    {
        public XlsxExporterTests(TestBenchmarkRunner benchmarkForTests)
        {
            TestBenchmarkRunner = benchmarkForTests;
        }

        private TestBenchmarkRunner TestBenchmarkRunner { get; }

        [Fact]
        public void ExportToFilesTest()
        {
            string file = null;
            try
            {
                var summary = TestBenchmarkRunner.GetSummary();
                var xlsxReporter = new XlsxExporter();
                var dateTime = DateTime.Now;
                var files = xlsxReporter.ExportToFiles(summary, NullLogger.Instance);
                Assert.True(files.Any());
                file = files.First();
                Assert.True(File.Exists(file));
                Assert.True(File.GetLastWriteTime(file) > dateTime);
            }
            finally
            {
                if (!(file is null))
                    File.Delete(file);
            }
        }

        [Fact]
        public void CreateSpreadsheetWorkbookTest()
        {
            using (var filePath = TestHelper.GetTemporaryFilePath("CreateSpreadsheetWorkbookTest.xlsx"))
            {
                var summary = TestBenchmarkRunner.GetSummary();
                var xlsxReporter = new XlsxExporter();
                xlsxReporter.CreateSpreadsheetWorkbook(filePath, summary, NullLogger.Instance);
                Assert.True(File.Exists(filePath));
            }
        }

        [Fact]
        public void ExportToFilesThrowsWhenSummaryParameterIsNull()
        {
            var exporter = new XlsxExporter();
            Assert.NotNull(Record.Exception(() => exporter.ExportToFiles(null, NullLogger.Instance)));
        }

        [Fact]
        public void ExportToFilesThrowsWhenLoggerParameterIsNull()
        {
            var exporter = new XlsxExporter();
            Assert.NotNull(Record.Exception(() => exporter.ExportToFiles(TestBenchmarkRunner.EmptySummary, null)));
        }

        [Fact]
        public void XlsxExporterExportToLogDoesNotThrowOnEmptySummary()
        {
            var exporter = new XlsxExporter();
            Assert.Null(Record.Exception(() => exporter.ExportToLog(TestBenchmarkRunner.EmptySummary, NullLogger.Instance)));
        }

        [Fact]
        public void XlsxExporterExportToLogDoesNotThrowWhenParametersAreNull()
        {
            var exporter = new XlsxExporter();
            Assert.Null(Record.Exception(() => exporter.ExportToLog(null, null)));
        }

        [Fact]
        public void NewExporterTest()
        {
            Assert.Null(Record.Exception(() => new XlsxExporter()));
        }

        [Fact]
        public void ExporterConstructorCannotBeNullTest()
        {
            Assert.NotNull(Record.Exception(() => new XlsxExporter(null)));
        }

        [Fact]
        public void ExporterConstructorCanBeEmpty()
        {
            Assert.Null(Record.Exception(() => new XlsxExporter(Array.Empty<IXlsxExporterHandler>())));
        }

        [Fact]
        public void DefaultXlsxExporterDoesNotThrow()
        {
            Assert.Null(Record.Exception(() => XlsxExporter.Default));
        }

        [Fact]
        public void XlsxExporterDefaultXlsxHandlersDoesNotThrow()
        {
            Assert.Null(Record.Exception(() => XlsxExporter.DefaultXlsxHandlers));
        }

        [Fact]
        public void XlsxExporterMinimalXlsxHandlersDoesNotThrow()
        {
            Assert.Null(Record.Exception(() => XlsxExporter.MinimalXlsxHandlers));
        }
    }
}
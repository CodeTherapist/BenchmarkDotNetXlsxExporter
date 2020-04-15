using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Helpers;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Reports;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// The xlsx exporter. 
    /// </summary>
    public class XlsxExporter : IExporter
    {
        private readonly IXlsxExporterHandler[] _benchmarkXlsxHandlers;

        /// <summary>
        /// Represents a minimal set of xlsx handlers.
        /// <para>Consist of: <see cref="SummaryXlsxHandler"/>.</para>
        /// </summary>
        public static readonly IXlsxExporterHandler[] MinimalXlsxHandlers = new IXlsxExporterHandler[] { new SummaryXlsxHandler() };

        /// <summary>
        /// Represents a default set of xlsx handlers.
        /// <para>Consist of: <see cref="SummaryXlsxHandler"/>; <see cref="FullSummaryXlsxHandler"/>; <see cref="HostEnvironmentInfoXlsxHandler"/>; <see cref="ValidationErrorsXlsxExporterHandler"/>.</para>
        /// </summary>
        public static readonly IXlsxExporterHandler[] DefaultXlsxHandlers = new IXlsxExporterHandler[] {
            new SummaryXlsxHandler() ,
            new FullSummaryXlsxHandler(),
            new HostEnvironmentInfoXlsxHandler(),
            new ValidationErrorsXlsxExporterHandler(),
        };

        /// <summary>
        /// Gets the default Xlsx exporter.
        /// </summary>
        public static readonly IExporter Default = new XlsxExporter();

        /// <summary>
        /// Initializes a new instance of <see cref="XlsxExporter"/>.
        /// </summary>
        public XlsxExporter() : this(DefaultXlsxHandlers)
        {

        }

        /// <summary>
        /// Initializes a new instance of <see cref="XlsxExporter"/>.
        /// </summary>
        /// <param name="benchmarkXlsxHandler">The list of xlsx handlers. Cannot be null.</param>
        /// <exception cref="ArgumentNullException">Occures when <paramref name="benchmarkXlsxHandler"/> is null.</exception>
        public XlsxExporter(IXlsxExporterHandler[] benchmarkXlsxHandler)
        {
            _benchmarkXlsxHandlers = benchmarkXlsxHandler ?? throw new ArgumentNullException(nameof(benchmarkXlsxHandler));
        }

        /// <summary>
        /// Gets the name of the exporter.
        /// </summary>
        public string Name => nameof(XlsxExporter);

        /// <summary>
        /// Exports the summary to a file.
        /// </summary>
        /// <param name="summary">The summry.</param>
        /// <param name="consoleLogger">A logger.</param>
        /// <exception cref="ArgumentNullException">Occures when <paramref name="summary"/> or <paramref name="consoleLogger"/> is null.</exception>
        /// <returns>A list of files.</returns>
        public IEnumerable<string> ExportToFiles(Summary summary, ILogger consoleLogger)
        {
            if (summary is null)
            {
                throw new ArgumentNullException(nameof(summary));
            }

            if (consoleLogger is null)
            {
                throw new ArgumentNullException(nameof(consoleLogger), $"Use 'BenchmarkDotNet.Loggers.{nameof(NullLogger.Instance)}' instead.");
            }

            string fileName = GetFileName(summary);
            string filePath = GetArtifactFullName(summary);
            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (IOException)
                {
                    string uniqueString = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                    string alternativeFilePath = $"{Path.Combine(summary.ResultsDirectoryPath, fileName)}-report-{uniqueString}.xlsx";
                    consoleLogger.WriteLineError($"Could not overwrite file {filePath}. Exporting to {alternativeFilePath}");
                    filePath = alternativeFilePath;
                }
            }
            CreateSpreadsheetWorkbook(filePath, summary, consoleLogger);
            return new[] { filePath };
        }

        /// <summary>
        /// Exports the summary to a logger.
        /// <para>This is not supported by the <see cref="XlsxExporter"/> and does nothing.</para>
        /// </summary>
        public void ExportToLog(Summary summary, ILogger consoleLogger)
        {
            // This exporter can't write to a logger.
        }

        /// <summary>
        /// Creates a xlsx file of the summary.
        /// </summary>
        /// <param name="fullFileName">The full file name of the xlsx file.</param>
        /// <param name="summary">The summary.</param>
        /// <param name="consoleLogger">The logger.</param>
        /// <exception cref="ArgumentException">Occures when <paramref name="fullFileName"/> is null or empty.</exception>
        /// <exception cref="ArgumentNullException">Occures when <paramref name="summary"/> or <paramref name="consoleLogger"/> is null.</exception>
        /// <remarks>Throws any possible exceptions from <see cref="FileStream.FileStream(string, FileMode, FileAccess, FileShare)"/>.</remarks>
        public void CreateSpreadsheetWorkbook(string fullFileName, Summary summary, ILogger consoleLogger)
        {
            if (string.IsNullOrWhiteSpace(fullFileName))
            {
                throw new ArgumentException("Cannot be null or empty.", nameof(fullFileName));
            }

            using (var fs = new FileStream(fullFileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                CreateSpreadsheetWorkbook(fs, summary, consoleLogger);
            }
        }

        /// <summary>
        /// Creates a xlsx file of the summary.
        /// </summary>
        /// <param name="stream">The output stream.</param>
        /// <param name="summary">The summary.</param>
        /// <param name="consoleLogger">The logger.</param>
        /// <exception cref="ArgumentNullException">Occures when <paramref name="stream"/>, <paramref name="summary"/> or <paramref name="consoleLogger"/> is null.</exception>
        public void CreateSpreadsheetWorkbook(Stream stream, Summary summary, ILogger consoleLogger)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (summary is null)
            {
                throw new ArgumentNullException(nameof(summary));
            }

            if (consoleLogger is null)
            {
                throw new ArgumentNullException(nameof(consoleLogger), $"Use 'BenchmarkDotNet.Loggers.{nameof(NullLogger.Instance)}' instead.");
            }

            using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var spreadsheet = new XlsxSpreadsheetDocument(spreadsheetDocument);
                spreadsheet.InitializeWorkbook();
                var handlers = _benchmarkXlsxHandlers.Any() ? _benchmarkXlsxHandlers : MinimalXlsxHandlers;
                foreach (var handler in handlers)
                {
                    try
                    {
                        handler.Handle(spreadsheet, summary);
                    }
                    catch (Exception ex)
                    {
                        consoleLogger.WriteLineError($"Cannot execute {handler.GetType()}: {ex.ToString()}.");
                    }
                }
                spreadsheet.Save();
            }
        }

        public static string GetArtifactFullName(Summary summary)
        {
            if (summary is null)
            {
                throw new ArgumentNullException(nameof(summary));
            }

            string fileName = GetFileName(summary);
            return $"{Path.Combine(summary.ResultsDirectoryPath, fileName)}-report.xlsx";
        }

        public static string GetFileName(Summary summary)
        {
            if (summary is null)
            {
                throw new ArgumentNullException(nameof(summary));
            }

            var targets = summary.BenchmarksCases.Select(b => b.Descriptor.Type).Distinct().ToArray();

            if (targets.Length == 1)
                return FolderNameHelper.ToFolderName(targets.Single());

            return summary.Title;
        }
    }
}


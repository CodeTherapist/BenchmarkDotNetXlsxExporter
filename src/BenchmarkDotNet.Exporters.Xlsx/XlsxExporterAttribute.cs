using BenchmarkDotNet.Attributes;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    /// <summary>
    /// An attribute to enable the <see cref="XlsxExporter"/>.
    /// </summary>
    public sealed class XlsxExporterAttribute : ExporterConfigBaseAttribute
    {
        /// <summary>
        /// Initializes a new instance of <see cref="XlsxExporterAttribute"/>.
        /// </summary>
        public XlsxExporterAttribute() : base(XlsxExporter.Default)
        {
        }
    }
}


using System.Linq;
using BenchmarkDotNet.Exporters.Xlsx;
using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class XlsxExporterAttributeTests
    {
        [Fact]
        public void ConstructorDoesNotThrow()
        {
            Assert.Null(Record.Exception(() => new XlsxExporterAttribute()));
        }

        [Fact]
        public void AttributeConfigHasXlsxExporter()
        {
            var attribute = new XlsxExporterAttribute();
            attribute.Config.GetExporters().Any(a => string.Equals(a.Name, XlsxExporter.Default.Name));
        }
    }


        
}


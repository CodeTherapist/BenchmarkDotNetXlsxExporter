using BenchmarkDotNet.Exporters.Xlsx;
using Xunit;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    public class ColumnNameHelperTests
    {
        [Theory]
        [InlineData(1, "A")]
        [InlineData(2, "B")]
        [InlineData(17, "Q")]
        [InlineData(38, "AL")]
        [InlineData(384, "NT")]
        public void ColumnIndexTests(int index, string expected)
        {
            var actual = ColumnNameHelper.GetExcelColumnName(index);
            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-54)]
        [InlineData(-789)]
        [InlineData(int.MinValue)]
        public void EnsureIndexValueLessThanOneThrows(int index)
        {
            Record.Exception(() => {
                ColumnNameHelper.GetExcelColumnName(index);
            });
        }
    }
}

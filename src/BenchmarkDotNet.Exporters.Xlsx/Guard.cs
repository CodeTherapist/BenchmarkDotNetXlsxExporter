using System;
using System.Collections.Generic;
using System.Text;

namespace BenchmarkDotNet.Exporters.Xlsx
{
    public static class Guard
    {
        public static void EnsureNotEmpty(string value, string paramName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("Cannot be null or empty.", paramName);
            }
        }

    }
}

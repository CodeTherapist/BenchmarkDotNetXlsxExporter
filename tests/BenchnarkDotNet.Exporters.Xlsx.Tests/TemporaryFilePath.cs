using System;
using System.IO;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    /// <summary>
    /// A temporary file path that is deleted after disposing it.
    /// </summary>
    internal sealed class TemporaryFilePath : IDisposable
    {
        public TemporaryFilePath(string fullFileName)
        {
            if (string.IsNullOrWhiteSpace(fullFileName))
            {
                throw new ArgumentException("Cannot be null or empty.", nameof(fullFileName));
            }

            FullFileName = fullFileName;
            File.Delete(FullFileName);
        }

        public string FullFileName { get; }

        public void Dispose()
        {
            File.Delete(FullFileName);
        }

        public static implicit operator string(TemporaryFilePath t) => t.FullFileName;
    }
}


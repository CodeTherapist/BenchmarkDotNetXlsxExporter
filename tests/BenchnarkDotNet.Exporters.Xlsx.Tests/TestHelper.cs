using System.IO;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    /// <summary>
    /// A helper for unit testing purpose.
    /// </summary>
    internal static class TestHelper
    {
        /// <summary>
        /// Gets a temporary file path.
        /// </summary>
        /// <param name="fileName">An absolut or relative path to the current directory.</param>
        public static TemporaryFilePath GetTemporaryFilePath(string fileName)
        {
            if (Path.IsPathRooted(fileName))
                return new TemporaryFilePath(fileName);

            return new TemporaryFilePath(Path.Combine(Directory.GetCurrentDirectory(), fileName));
        }
    }
}


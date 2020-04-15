using System;
using System.Security.Cryptography;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Exporters.Xlsx;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

namespace ConsoleApp1
{ 
    namespace MyBenchmarks
    {
        [MemoryDiagnoser, SimpleJob(RuntimeMoniker.NetCoreApp31), SimpleJob(RuntimeMoniker.Net48), XlsxExporter]
        public class Md5VsSha256
        {
            private const int N = 10000;
            private readonly byte[] data;

            private readonly SHA256 sha256 = SHA256.Create();
            private readonly MD5 md5 = MD5.Create();

            public Md5VsSha256()
            {
                data = new byte[N];
                new Random(42).NextBytes(data);
            }

            [Benchmark]
            public byte[] Sha256() => sha256.ComputeHash(data);

            [Benchmark]
            public byte[] Md5() => md5.ComputeHash(data);
        }

        public class Program
        {
            public static void Main(string[] args)
            {   
                BenchmarkRunner.Run(typeof(Program).Assembly);
            }
        }
    }
}

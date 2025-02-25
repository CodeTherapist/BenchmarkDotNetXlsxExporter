﻿using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Globalization;
using System.Security.Cryptography;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Validators;

namespace BenchnarkDotNet.Exporters.Xlsx.Test
{
    /// <summary>
    /// Encapsulates summaries for testing purpose.
    /// </summary>
    public class TestBenchmarkRunner
    {
        /// <summary>
        /// Gets an empty Summary.
        /// </summary>
        public Summary EmptySummary => new Summary(string.Empty, new List<BenchmarkReport>().ToImmutableArray(),
                 BenchmarkDotNet.Environments.HostEnvironmentInfo.GetCurrent(),
                 string.Empty, string.Empty, TimeSpan.Zero, CultureInfo.CurrentCulture,
                 new List<ValidationError>().ToImmutableArray());

        /// <summary>
        /// Runs a test benchmark.
        /// </summary>
        /// <returns>The summary result of the benchmark.</returns>
        public Summary GetSummary()
        {
            return BenchmarkRunner.Run(typeof(Md5VsSha256), DefaultConfig.Instance.WithOptions(ConfigOptions.DisableOptimizationsValidator));
        }

        [MemoryDiagnoser, SimpleJob(RuntimeMoniker.NetCoreApp31), SimpleJob(RuntimeMoniker.Net48)]
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
    }
}


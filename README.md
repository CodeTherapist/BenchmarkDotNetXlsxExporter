# BenchmarkDotNet XLSX exporter

A [BenchmarkDotNet](https://benchmarkdotnet.org/) extension to export benchmark results as xlsx file.

## Installation

[BenchmarkDotNetXlsxExporter](https://www.nuget.org/packages/BenchmarkDotNetXlsxExporter/) is a [.NET Standard 2.0](https://docs.microsoft.com/en-us/dotnet/standard/net-standard) library.

Install the nuget package [BenchmarkDotNetXlsxExporter](https://www.nuget.org/packages/BenchmarkDotNetXlsxExporter/) by adding a `PackageReference`:

```html
<PackageReference Include="BenchmarkDotNetXlsxExporter" Version="1.0.0" />
```

**Platform Support**
* .NET Core 2.0, 3.0, 3.1
* .NET Framework 4.6.1
* Mono 5.4
* Xamarin.iOS 10.14
* Xamarin.Mac 3.8
* Xamarin.Android 8.0
* Universal Windows Platform 10.0.16299

### Usage examples

To enable the exporter, add the attribute `XlsxExporterAttribute` to the benchmark class:

```csharp
using System;
using System.Security.Cryptography;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Exporters.Xlsx;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

public class Program
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

    public static void Main(string[] args)
    {
        BenchmarkRunner.Run(typeof(Program).Assembly);
    }
}

```

For custom configurations of the `XlsxExporter` add it to the exporter list (without the attribute).

```csharp

public class Program
{
    [MemoryDiagnoser, SimpleJob(RuntimeMoniker.NetCoreApp31), SimpleJob(RuntimeMoniker.Net48)]
    public class Md5VsSha256
    {
        // omitted, see above.
    }

    public static void Main(string[] args)
    {
        var xlsxExporter = new XlsxExporter();
        BenchmarkRunner.Run(typeof(Program).Assembly, DefaultConfig.Instance.AddExporter(xlsxExporter));

        // or a minimal xlsx output:
        // BenchmarkRunner.Run(typeof(Program).Assembly, DefaultConfig.Instance.AddExporter(XlsxExporter.MinimalXlsxHandlers));
    }
}
```

**A example of the exporter output: [ConsoleApp.MyBenchmarks.Md5VsSha256-report.xlsx](ConsoleApp.MyBenchmarks.Md5VsSha256-report.xlsx).**

## Contribution - Any help is appreciated

Feel free to do pull requests for any kind of improvements, ask questions or just :star this repository when you find it useful.

### Solution

The solution `BenchmarkDotNetXlsxExporter` has three projects.

**Core** 

_BenchmarkDotNet.Exporters.Xlsx_ contains the core functionality, the exporter.

**Tests**

_BenchmarkDotNet.Exporters.Xlsx.Tests_ contains unit tests.

**Console App**

_ConsoleApp_ is an example/playground how the exporter can be used.

### Tests & Code Coverage

Unit Testing is done with [xUnit](https://github.com/xunit/xunit).

Code coverage is done with [Coverlet](https://github.com/tonerdo/coverlet).

Code coverage reports is done with [ReportGenerator](https://github.com/danielpalme/ReportGenerator).

**Code Coverage statistic**

| Module                         | Line   | Branch | Method |
|--------------------------------|--------|--------|--------|
| BenchmarkDotNet.Exporters.Xlsx | 82,44% | 73,86% | 86,53% |

Executing unit tests will trigger code coverage and report generation.

The code coverage result is stored in `.\artifacts\tests\coverage` (format is [OpenCover](https://github.com/OpenCover/opencover)).

The graphical code coverage reports are stored in `.\artifacts\tests\coverage\reports` and history in `.\artifacts\tests\coverage\reports\history`.

## Build, Test & Pack

The project separates the three concerns _Building_, _Testing_ and _Packaging_.
All of these steps could be executed individually.

**BuildTestPack.ps1**

This command calls internally **Build.ps1**, **Test.ps1** and **Pack.ps1** and supports following parameters:

Parameter|Description|Type|
---------|-----------|----|
Configuration|Can be set to "Release" or "Debug"|string
CollectCoverage|When set, code coverage is calculated|switch 
NoIntegrationTests|When set, integration tests are skipped|switch
Pack|When set, nuget packages are created (call to Pack.ps1)|switch

**Build.ps1**

Builds all projects. Supports the _Configuration_ parameter.

**Test.ps1**

Runs all tests. Supports _Configuration_, _CollectCoverage_, _NoIntegrationTests_ and _NoBuild_ parameters.

Parameter|Description|Type
---------|-----------|----
NoBuild|The projects are not re-build|switch

**Pack.ps1**

Creates the nuget packages into the `.\artifacts\packages` directory. Supports _Configuration_ and _NoBuild_ parameters.

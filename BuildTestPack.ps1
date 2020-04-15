

[CmdletBinding(PositionalBinding = $false)]
param(
    [ValidateSet('Debug', 'Release')]
    $Configuration = 'Release',
    [switch]
    $CollectCoverage,
    [switch]
    $NoIntegrationTests,
    [switch]
    $Pack
)

Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

.\Build.ps1 -Configuration $Configuration

.\Test.ps1 -Configuration $Configuration -NoBuild -CollectCoverage:$CollectCoverage -NoIntegrationTests:$NoIntegrationTests

if($Pack) {
    .\Pack.ps1 -Configuration $Configuration -NoBuild
}
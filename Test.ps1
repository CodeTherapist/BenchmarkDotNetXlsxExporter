

[CmdletBinding(PositionalBinding = $false)]
param(
    [ValidateSet('Debug', 'Release')]
    $Configuration = 'Release',
	[switch]
	$CollectCoverage,
	[switch]
	$NoBuild,
	[switch]
	$NoIntegrationTests
)

Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

. .\Functions.ps1

write-host -f Magenta '*** Test ***'

$arguments = New-Object System.Collections.ArrayList

if ($NoBuild) {
	[void]$arguments.Add('--no-build')
}

if($NoIntegrationTests) {
	[void]$arguments.Add('--filter="Category!=Integration"')
}

[void]$arguments.Add("-property:Configuration=$Configuration")

if ($CollectCoverage) { 
	[void]$arguments.Add("-p:CollectCoverage=true")
} 

exec dotnet test @arguments
   
write-host -f green 'Test Done!'


param()

Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

. .\Functions.ps1

write-host -f Magenta '*** Restore ***'

exec dotnet restore

write-host -f green 'Restore done!'


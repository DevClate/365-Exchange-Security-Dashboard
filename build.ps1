# Grab nuget bits, install modules, set build variables, start build.
Get-PackageProvider -Name NuGet -ForceBootstrap | Out-Null

import-Module Psake, PSDeploy, Pester, BuildHelpers

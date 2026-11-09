#Requires -Version 7.0
# Runs the Pester tests. The three page scenarios write their JSON output to examples/output/.
$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version -ge [version]'5.5.0' })) {
    throw 'Pester 5.5 or later is required: Install-Module Pester -MinimumVersion 5.5 -Scope CurrentUser'
}
Import-Module Pester -MinimumVersion 5.5
$config = New-PesterConfiguration
$config.Run.Path = Join-Path $PSScriptRoot 'tests'
$config.Run.Exit = $true
$config.Output.Verbosity = 'Detailed'
Invoke-Pester -Configuration $config

# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.

# Per-user shortcuts reuse the launcher rather than duplicating profile and environment logic.
[CmdletBinding()]
param([switch]$Optimized, [switch]$Primary)

$ErrorActionPreference = 'Stop'
if ($Optimized -and $Primary) {
	throw 'Choose either the optimized build or the primary checkout.'
}
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$launcher = Join-Path $PSScriptRoot 'code-agent-workspace.ts'
$icon = Join-Path $repoRoot 'resources/win32/code.ico'
$helper = Join-Path $repoRoot $(if ($Primary) { '.build/signal-c-primary-shortcut.vbs' } elseif ($Optimized) { '.build/signal-c-optimized-shortcut.vbs' } else { '.build/signal-c-shortcut.vbs' })
$nodeCommand = Get-Command node -CommandType Application -ErrorAction Stop | Select-Object -First 1
$scriptHost = Join-Path ([Environment]::GetFolderPath('System')) 'wscript.exe'
$description = "Signal-C $(if ($Primary) { 'primary checkout' } elseif ($Optimized) { 'optimized desktop' } else { 'development checkout' }): $repoRoot"
$helperHeader = "' $description"

foreach ($requiredFile in @($launcher, $icon, $scriptHost)) {
	if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
		throw "A required launcher file is missing: $requiredFile"
	}
}
if ((Test-Path -LiteralPath $helper) -and (Get-Content -LiteralPath $helper -TotalCount 1) -ne $helperHeader) {
	throw "The launcher helper already exists and does not belong to this shortcut: $helper"
}

$shell = New-Object -ComObject WScript.Shell
try {
	$programs = $shell.SpecialFolders.Item('Programs')
	$shortcutName = if ($Primary -or $Optimized) { 'Signal-C' } else { 'Signal-C (Development)' }
	$shortcutPath = Join-Path $programs "$shortcutName.lnk"
	$shortcut = $shell.CreateShortcut($shortcutPath)
	try {
		if ((Test-Path -LiteralPath $shortcutPath) -and $shortcut.Description -ne $description) {
			throw "An unrelated Start menu shortcut already uses this name: $shortcutPath"
		}
		# WScript hides only Node's console. The existing launcher opens the normal Agents window.
		$command = '"' + $nodeCommand.Source + '" "' + $launcher + '"'
		if ($Optimized) { $command += ' --optimized' }
		if ($Primary) { $command += ' --primary' }
		$escapedCommand = $command.Replace('"', '""')
		$helperSource = @"
$helperHeader
Option Explicit
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "$escapedCommand", 0, False
"@
		New-Item -ItemType Directory -Path (Split-Path -Parent $helper) -Force | Out-Null
		[IO.File]::WriteAllText($helper, $helperSource, [Text.UnicodeEncoding]::new($false, $true))
		$shortcut.TargetPath = $scriptHost
		$shortcut.Arguments = '"' + $helper + '"'
		$shortcut.WorkingDirectory = $repoRoot
		$shortcut.IconLocation = "$icon,0"
		$shortcut.Description = $description
		$shortcut.Save()
		if ($Primary) {
			$protocol = 'HKCU:\Software\Classes\signal-c'
			$openCommand = "$protocol\shell\open\command"
			New-Item -Path $openCommand -Force | Out-Null
			Set-Item -LiteralPath $protocol -Value 'URL:Signal-C'
			New-ItemProperty -LiteralPath $protocol -Name 'URL Protocol' -PropertyType String -Value '' -Force | Out-Null
			Set-Item -LiteralPath $openCommand -Value ($command + ' --open-url "%1"')
		}
		Write-Output "Created Start menu entry: $shortcutName"
		Write-Output $shortcutPath
	} finally {
		[Runtime.InteropServices.Marshal]::FinalReleaseComObject($shortcut) | Out-Null
	}
} finally {
	[Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell) | Out-Null
}

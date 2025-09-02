# Requires PowerShell 5+
param(
  [switch]$CloseApps = $true,
  [switch]$EnableTrust = $true,
  [switch]$WhatIf
)

Write-Host "Preparing Office environment for VBA automation..." -ForegroundColor Cyan

function Close-OfficeApps {
  param([string[]]$ProcessNames)
  foreach ($name in $ProcessNames) {
    $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
    if ($procs) {
      Write-Host ("Closing {0} ({1})" -f $name, ($procs | Measure-Object).Count)
      if ($WhatIf) { continue }
      $procs | Stop-Process -Force -ErrorAction SilentlyContinue
    }
  }
}

function Set-AccessVBOM {
  param(
    [Parameter(Mandatory)] [string]$Version,
    [Parameter(Mandatory)] [ValidateSet('Excel','Word','PowerPoint','Access')] [string]$App
  )
  $path = "HKCU:\Software\Microsoft\Office\$Version\$App\Security"
  if ($WhatIf) {
    Write-Host "Would set AccessVBOM=1 at $path" -ForegroundColor Yellow
    return
  }
  New-Item -Path $path -Force | Out-Null
  New-ItemProperty -Path $path -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null
  Write-Host "Set AccessVBOM=1 at $path" -ForegroundColor Green
}

if ($CloseApps) {
  Write-Host "Closing Office apps (unsaved work will be lost)." -ForegroundColor DarkYellow
  Close-OfficeApps -ProcessNames @('WINWORD','EXCEL','POWERPNT','MSACCESS','OUTLOOK')
}

if ($EnableTrust) {
  Write-Host "Enabling 'Trust access to the VBA project object model'..." -ForegroundColor DarkYellow
  $officeRoot = 'HKCU:\Software\Microsoft\Office'
  $versions = @()
  try {
    $versions = (Get-ChildItem $officeRoot -ErrorAction Stop | Where-Object { $_.PSChildName -match '^[0-9]+\.[0-9]+$' }).PSChildName
  } catch {
    Write-Host "No Office registry found under HKCU. You may need to run once per user." -ForegroundColor Red
  }
  if (-not $versions) {
    # Fallback to common versions
    $versions = @('16.0','15.0','14.0')
  }
  foreach ($v in $versions) {
    foreach ($app in @('Excel','Word','PowerPoint','Access')) {
      Set-AccessVBOM -Version $v -App $app
    }
  }
}

Write-Host "Done." -ForegroundColor Cyan

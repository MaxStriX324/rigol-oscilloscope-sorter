param(
  [Parameter(Mandatory=$true)][ValidatePattern('^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$')][string]$Repository,
  [ValidateSet('main','master')][string]$Branch='main',
  [ValidateSet('install.bat','install.cmd')][string]$Installer,
  [switch]$InstalledApp
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Find-Git {
  $cmd=Get-Command git.exe -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  $paths=@(
    (Join-Path $env:ProgramFiles 'Git\cmd\git.exe'),
    (Join-Path $env:LOCALAPPDATA 'Programs\Git\cmd\git.exe')
  )
  foreach($path in $paths) {
    if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
  }
  $winget=Get-Command winget.exe -ErrorAction SilentlyContinue
  if ($winget) {
    Write-Host 'Git is missing. Trying to install Git for the current Windows user...'
    & $winget.Source install --id Git.Git --exact --source winget --scope user --silent --accept-source-agreements --accept-package-agreements
    if ($LASTEXITCODE -ne 0) {
      Write-Warning ('winget Git installation exited with code ' + $LASTEXITCODE)
    }
    foreach($path in $paths) {
      if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
    }
    $cmd=Get-Command git.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
  }
  throw 'Git for Windows is not available. Install from https://git-scm.com/downloads/win and rerun update.bat.'
}

function Run-Git {
  param([string[]]$Arguments)
  & $script:GitExe @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw ('Git operation failed (exit ' + $LASTEXITCODE + '): git ' + ($Arguments -join ' '))
  }
}

function Check-Remote {
  param([string]$Directory)
  $remote=(& $script:GitExe -C $Directory remote get-url origin 2>$null)
  if ($LASTEXITCODE -ne 0 -or -not $remote) { throw ('Git remote origin is missing in ' + $Directory) }
  $remote=([string]($remote | Select-Object -First 1)).Trim().TrimEnd('/')
  $suffix=$Repository.ToLowerInvariant()
  $allowed=@(
    "https://github.com/$suffix",
    "https://github.com/$suffix.git",
    "git@github.com:$suffix",
    "git@github.com:$suffix.git",
    "ssh://git@github.com/$suffix",
    "ssh://git@github.com/$suffix.git"
  )
  if ($allowed -notcontains $remote.ToLowerInvariant()) {
    throw ("Unexpected Git origin: $remote. Expected $Repository. No files were changed.")
  }
}

function Update-Checkout {
  param([string]$Directory)
  Check-Remote $Directory
  $top=(& $script:GitExe -C $Directory rev-parse --show-toplevel 2>$null)
  if ($LASTEXITCODE -ne 0 -or -not $top) { throw 'Invalid Git checkout.' }
  if ([IO.Path]::GetFullPath(([string]$top).Trim()).TrimEnd('\') -ine [IO.Path]::GetFullPath($Directory).TrimEnd('\')) {
    throw 'The update folder is not the root of this repository.'
  }
  $current=(& $script:GitExe -C $Directory symbolic-ref --quiet --short HEAD 2>$null)
  if ($LASTEXITCODE -ne 0 -or ([string]$current).Trim() -cne $Branch) {
    throw ("This checkout is not on $Branch. Switch to $Branch before updating.")
  }
  $changes=(& $script:GitExe -C $Directory status --porcelain --untracked-files=no)
  if ($LASTEXITCODE -ne 0) { throw 'Cannot check local modifications.' }
  if ($changes) {
    throw 'Local changes are present. Save/commit/stash them first; update.bat will not overwrite them.'
  }
  Write-Host ('Fetching ' + $Repository + ' (' + $Branch + ') ...')
  Run-Git -Arguments @('-C',$Directory,'fetch','--prune','origin',$Branch)
  Run-Git -Arguments @('-C',$Directory,'merge','--ff-only','FETCH_HEAD')
  Write-Host 'Git update completed.' -ForegroundColor Green
}

function Install-Program {
  param([string]$Directory)
  if (-not $Installer) { return }
  $installerPath=Join-Path $Directory $Installer
  if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    throw ('Installer not found: ' + $installerPath)
  }
  Write-Host ('Running ' + $Installer + ' from ' + $Directory)
  if ($InstalledApp) {
    # ST-Link installer copies sources to LOCALAPPDATA and preserves settings/logs.
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Directory 'scripts\install.ps1') -NoLaunch
    if ($LASTEXITCODE -ne 0) { throw ('STM32 installer failed: ' + $LASTEXITCODE) }
  } else {
    # /quiet disables the installer's closing pause without skipping checks.
    $cmdline='/d /s /c ""' + $installerPath + '" /quiet"'
    $proc=Start-Process -FilePath $env:ComSpec -ArgumentList $cmdline -WorkingDirectory $Directory -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) { throw ('Installer failed: ' + $proc.ExitCode) }
  }
  Write-Host 'Dependencies and installed program checked.' -ForegroundColor Green
}

try {
  $GitExe=Find-Git
  $sourceRoot=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
  $repoName=($Repository -split '/')[1]
  $cloneUrl='https://github.com/' + $Repository + '.git'
  if (Test-Path -LiteralPath (Join-Path $sourceRoot '.git')) {
    Write-Host ('Git checkout found: ' + $sourceRoot)
    Update-Checkout $sourceRoot
    $workTree=$sourceRoot
  } else {
    # ZIP copies and the installed ST-Link application have no .git folder.
    # Never overwrite files there: clone into a separate, reusable checkout.
    if ($InstalledApp) {
      $workTree=Join-Path $env:LOCALAPPDATA ('Programs\' + $repoName + '-source')
    } else {
      $workTree=Join-Path (Split-Path -Parent $sourceRoot) ($repoName + '-git')
    }
    if (Test-Path -LiteralPath $workTree) {
      if (-not (Test-Path -LiteralPath (Join-Path $workTree '.git'))) {
        throw ('Target folder already exists and is not a clone: ' + $workTree)
      }
      Write-Host ('Existing clone: ' + $workTree)
      Update-Checkout $workTree
    } else {
      Write-Host ('Creating a new Git clone: ' + $workTree)
      Run-Git -Arguments @('clone','--branch',$Branch,'--single-branch','--',$cloneUrl,$workTree)
      Check-Remote $workTree
    }
    if (-not $InstalledApp) {
      Write-Host 'IMPORTANT: Original ZIP folder was not changed.' -ForegroundColor Yellow
      Write-Host ('Use the program from the new folder: ' + $workTree)
      Write-Host 'Copy your private settings/input files manually if needed; do not overwrite the cloned source files.'
    }
  }
  Install-Program $workTree
  Write-Host ('UPDATE COMPLETE: ' + $workTree) -ForegroundColor Green
  exit 0
} catch {
  Write-Host ('UPDATE FAILED: ' + $_.Exception.Message) -ForegroundColor Red
  exit 1
}

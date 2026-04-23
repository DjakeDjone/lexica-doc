Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoOwner = if ($env:REPO_OWNER) { $env:REPO_OWNER } else { "DjakeDjone" }
$repoName = if ($env:REPO_NAME) { $env:REPO_NAME } else { "lexica-doc" }
$branch = if ($env:BRANCH) { $env:BRANCH } else { "main" }
$installRoot = if ($env:INSTALL_ROOT) { $env:INSTALL_ROOT } else { Join-Path $env:USERPROFILE ".cargo" }
$binName = "wors"
$appName = "Wors"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$shortcutFile = Join-Path $startMenuDir "${appName}.lnk"

function Require-Command {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "error: $Name is required"
    }
}

function Add-UserPathEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Entry
    )

    $existingUserPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $segments = @()
    if ($existingUserPath) {
        $segments = $existingUserPath -split ";" | Where-Object { $_ }
    }

    $normalizedEntry = $Entry.TrimEnd("\")
    $alreadyPresent = $segments | Where-Object { $_.TrimEnd("\") -ieq $normalizedEntry }
    if ($alreadyPresent) {
        return $false
    }

    $newUserPath = if ($existingUserPath) { "$existingUserPath;$Entry" } else { $Entry }
    [Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")

    $currentProcessPath = $env:Path -split ";" | Where-Object { $_ }
    if (-not ($currentProcessPath | Where-Object { $_.TrimEnd("\") -ieq $normalizedEntry })) {
        $env:Path = if ($env:Path) { "$env:Path;$Entry" } else { $Entry }
    }

    return $true
}

Require-Command -Name "cargo"

$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("wors-install-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmpDir | Out-Null

try {
    $archiveUrl = "https://github.com/$repoOwner/$repoName/archive/refs/heads/$branch.zip"
    $archiveFile = Join-Path $tmpDir "source.zip"
    $sourceRoot = Join-Path $tmpDir "$repoName-$branch"
    $crateDir = $sourceRoot

    Write-Host "Downloading $archiveUrl"
    Invoke-WebRequest -Uri $archiveUrl -OutFile $archiveFile
    Expand-Archive -LiteralPath $archiveFile -DestinationPath $tmpDir

    if (-not (Test-Path (Join-Path $crateDir "Cargo.toml")) -and (Test-Path (Join-Path $sourceRoot "browser\Cargo.toml"))) {
        $crateDir = Join-Path $sourceRoot "browser"
    }

    if (-not (Test-Path (Join-Path $crateDir "Cargo.toml"))) {
        throw "error: Cargo.toml not found in downloaded source"
    }

    $installBinDir = Join-Path $installRoot "bin"
    New-Item -ItemType Directory -Force -Path $installBinDir | Out-Null

    Write-Host "Installing $binName to $installBinDir"
    & cargo install --path $crateDir --locked --force --root $installRoot
    if ($LASTEXITCODE -ne 0) {
        throw "error: cargo install failed"
    }

    $exePath = Join-Path $installBinDir "$binName.exe"
    if (-not (Test-Path $exePath)) {
        throw "error: installed executable not found at $exePath"
    }

    $pathUpdated = Add-UserPathEntry -Entry $installBinDir

    New-Item -ItemType Directory -Force -Path $startMenuDir | Out-Null
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutFile)
    $shortcut.TargetPath = $exePath
    $shortcut.WorkingDirectory = Split-Path -Parent $exePath
    $shortcut.Description = "Minimal desktop document editor"
    $shortcut.IconLocation = $exePath
    $shortcut.Save()

    Write-Host "Installed $binName and Start Menu shortcut."
    if ($pathUpdated) {
        Write-Host "Added $installBinDir to your user PATH. Open a new terminal before running $binName."
    }
}
finally {
    Remove-Item -Recurse -Force $tmpDir -ErrorAction SilentlyContinue
}

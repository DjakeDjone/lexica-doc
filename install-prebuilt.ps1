Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

$repoOwner = if ($env:REPO_OWNER) { $env:REPO_OWNER } else { "DjakeDjone" }
$repoName = if ($env:REPO_NAME) { $env:REPO_NAME } else { "lexica-doc" }
$branch = if ($env:BRANCH) { $env:BRANCH } else { "main" }
$workflowFile = if ($env:WORKFLOW_FILE) { $env:WORKFLOW_FILE } else { "build.yml" }
$installRoot = if ($env:INSTALL_ROOT) { $env:INSTALL_ROOT } else { Join-Path $env:USERPROFILE ".local" }
$artifactName = if ($env:ARTIFACT_NAME) { $env:ARTIFACT_NAME } else { "wors-windows-x86_64" }
$binName = "wors.exe"
$githubToken = if ($env:GH_TOKEN) { $env:GH_TOKEN } else { $env:GITHUB_TOKEN }

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

$headers = @{
    Accept = "application/vnd.github+json"
}
if ($githubToken) {
    $headers.Authorization = "Bearer $githubToken"
}

$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("wors-prebuilt-install-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmpDir | Out-Null

try {
    $extractDir = Join-Path $tmpDir "artifact"
    New-Item -ItemType Directory -Force -Path $extractDir | Out-Null

    if (Get-Command gh -ErrorAction SilentlyContinue) {
        Write-Host "Finding latest successful $workflowFile run on $branch"
        $runJson = & gh run list --repo "$repoOwner/$repoName" --workflow $workflowFile --branch $branch --status success --event push --limit 1 --json databaseId
        if ($LASTEXITCODE -ne 0) {
            throw "error: gh run list failed"
        }

        $run = $runJson | ConvertFrom-Json | Select-Object -First 1
        if (-not $run -or -not $run.databaseId) {
            throw "error: no successful workflow runs found"
        }

        $artifactZip = Join-Path $tmpDir "$artifactName.zip"
        Write-Host "Downloading $artifactName"
        $artifactJson = & gh api "repos/$repoOwner/$repoName/actions/runs/$($run.databaseId)/artifacts"
        if ($LASTEXITCODE -ne 0) {
            throw "error: gh api artifacts lookup failed"
        }

        $artifact = ($artifactJson | ConvertFrom-Json).artifacts | Where-Object { $_.name -eq $artifactName -and -not $_.expired } | Select-Object -First 1
        if (-not $artifact) {
            throw "error: artifact '$artifactName' not found or has expired"
        }

        $token = & gh auth token
        if ($LASTEXITCODE -ne 0 -or -not $token) {
            throw "error: gh auth token failed"
        }

        $downloadHeaders = @{
            Accept = "application/vnd.github+json"
            Authorization = "Bearer $token"
        }
        Invoke-WebRequest -Uri $artifact.archive_download_url -Headers $downloadHeaders -OutFile $artifactZip
        Expand-Archive -LiteralPath $artifactZip -DestinationPath $extractDir
    }
    else {
        if (-not $githubToken) {
            throw "error: install-prebuilt.ps1 requires gh or GH_TOKEN/GITHUB_TOKEN to download GitHub Actions artifacts"
        }

        $apiRoot = "https://api.github.com/repos/$repoOwner/$repoName"
        $runsUrl = "$apiRoot/actions/workflows/$workflowFile/runs?branch=$branch&status=success&event=push&per_page=1"

        Write-Host "Finding latest successful $workflowFile run on $branch"
        $runs = Invoke-RestMethod -Uri $runsUrl -Headers $headers
        if (-not $runs.workflow_runs -or $runs.workflow_runs.Count -eq 0) {
            throw "error: no successful workflow runs found"
        }

        $artifacts = Invoke-RestMethod -Uri $runs.workflow_runs[0].artifacts_url -Headers $headers
        $artifact = $artifacts.artifacts | Where-Object { $_.name -eq $artifactName -and -not $_.expired } | Select-Object -First 1
        if (-not $artifact) {
            throw "error: artifact '$artifactName' not found or has expired"
        }

        $artifactZip = Join-Path $tmpDir "$artifactName.zip"

        Write-Host "Downloading $artifactName"
        Invoke-WebRequest -Uri $artifact.archive_download_url -Headers $headers -OutFile $artifactZip
        Expand-Archive -LiteralPath $artifactZip -DestinationPath $extractDir
    }

    $binary = Get-ChildItem -Path $extractDir -Recurse -File -Filter $binName | Select-Object -First 1
    if (-not $binary) {
        throw "error: $binName not found in $artifactName"
    }

    $installBinDir = Join-Path $installRoot "bin"
    New-Item -ItemType Directory -Force -Path $installBinDir | Out-Null

    $exePath = Join-Path $installBinDir $binName
    Copy-Item -LiteralPath $binary.FullName -Destination $exePath -Force

    $pathUpdated = Add-UserPathEntry -Entry $installBinDir

    Write-Host "Installed $binName to $exePath."
    if ($pathUpdated) {
        Write-Host "Added $installBinDir to your user PATH. Open a new terminal before running wors."
    }
}
finally {
    Remove-Item -Recurse -Force $tmpDir -ErrorAction SilentlyContinue
}

$ErrorActionPreference = "Continue"

$RepoUrl = "https://github.com/the5kka/The5kka.git"
$BasePath = "D:\QC\8. Codex"
$SyncRoot = Join-Path $BasePath "0. The5kka_GitHub"
$LogPath = Join-Path $SyncRoot "git_auto_save.log"
$HadError = $false

$QuickAccessDir = Join-Path $BasePath "1. Quick Access Searcher"
$DncDir = Join-Path $BasePath "2. JIIN_DNC_Manager"
$IatfDir = Join-Path $BasePath "3. IATF 16949 Search"
$ShortcutDir = (Get-ChildItem -LiteralPath $BasePath -Directory | Where-Object { $_.Name -like "4. Window *" } | Select-Object -First 1).FullName
$CodexDir = (Get-ChildItem -LiteralPath $BasePath -Directory | Where-Object { $_.Name -like "5. Codex*" } | Select-Object -First 1).FullName
$OjtDir = (Get-ChildItem -LiteralPath $BasePath -Directory | Where-Object { $_.Name -like "6. OJT*" } | Select-Object -First 1).FullName

$SourceRepos = @($QuickAccessDir, $DncDir, $IatfDir, $ShortcutDir, $CodexDir, $OjtDir) | Where-Object { $_ }

function Write-Log($Message) {
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $LogPath -Value "[$time] $Message" -Encoding UTF8
}

function Save-LocalRepo($Repo) {
    try {
        if (-not (Test-Path -LiteralPath $Repo)) {
            Write-Log "SKIP missing folder: $Repo"
            return
        }
        if (-not (Test-Path -LiteralPath (Join-Path $Repo ".git"))) {
            git -C $Repo init | Out-Null
            git -C $Repo config user.name "Codex Auto Save"
            git -C $Repo config user.email "codex-auto-save@local"
        }
        git -C $Repo add -A | Out-Null
        $status = git -C $Repo status --porcelain
        if ($status) {
            $message = "Auto save " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            git -C $Repo commit -m $message | Out-Null
            Write-Log "LOCAL COMMIT $Repo - $message"
        } else {
            Write-Log "LOCAL NOCHANGE $Repo"
        }
    } catch {
        $script:HadError = $true
        Write-Log "LOCAL ERROR $Repo - $($_.Exception.Message)"
    }
}

function Prepare-GitHubRepo() {
    New-Item -ItemType Directory -Path $SyncRoot -Force | Out-Null
    if (-not (Test-Path -LiteralPath (Join-Path $SyncRoot ".git"))) {
        git -C $SyncRoot init | Out-Null
        git -C $SyncRoot remote add origin $RepoUrl
    } else {
        $remotes = git -C $SyncRoot remote
        if ($remotes -notcontains "origin") {
            git -C $SyncRoot remote add origin $RepoUrl
        } else {
            git -C $SyncRoot remote set-url origin $RepoUrl
        }
    }
    git -C $SyncRoot config user.name "Codex Auto Upload"
    git -C $SyncRoot config user.email "codex-auto-upload@local"
    git -C $SyncRoot fetch origin main
    if ($LASTEXITCODE -ne 0) { throw "GitHub fetch failed" }
    git -C $SyncRoot checkout -B main origin/main | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "main branch checkout failed" }
}

function Clear-SyncRoot() {
    $resolved = (Resolve-Path -LiteralPath $SyncRoot).Path
    $expected = (Resolve-Path -LiteralPath (Join-Path $BasePath "0. The5kka_GitHub")).Path
    if ($resolved -ne $expected) { throw "Unexpected sync root: $resolved" }

    Get-ChildItem -LiteralPath $SyncRoot -Force |
        Where-Object { $_.Name -ne ".git" } |
        Remove-Item -Recurse -Force
}

function Copy-Project($Source, $DestinationName) {
    if (-not (Test-Path -LiteralPath $Source)) {
        Write-Log "SKIP missing copy source: $Source"
        return
    }
    $destination = Join-Path $SyncRoot $DestinationName
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    robocopy $Source $destination /MIR /XD .git git-remotes logs build dist __pycache__ /XF desktop.ini git_auto_save.log *.log *.pyc *.pyo *.tmp *.bak | Out-Null
}

function Copy-SourcesToSyncRoot() {
    Copy-Project $QuickAccessDir "1. Quick Access Searcher"
    Copy-Project $DncDir "2. JIIN_DNC_Manager"
    Copy-Project $IatfDir "3. IATF 16949 Search"
    if ($ShortcutDir) {
        Copy-Project $ShortcutDir (Split-Path -Leaf $ShortcutDir)
    }
    Copy-Project $CodexDir (Split-Path -Leaf $CodexDir)
    if ($OjtDir) {
        Copy-Project $OjtDir (Split-Path -Leaf $OjtDir)
    }
}

function Write-Readme() {
    $shortcutName = if ($ShortcutDir) { Split-Path -Leaf $ShortcutDir } else { "4. Window Shortcuts" }
    $readme = @"
# The5kka

Auto upload folder for D:\QC\8. Codex.

- `1. Quick Access Searcher`
- `2. JIIN_DNC_Manager`
- `3. IATF 16949 Search`
- `$shortcutName`
- `$(Split-Path -Leaf $CodexDir)`
- `$(Split-Path -Leaf $OjtDir)`

This repository is synchronized from the working folders under D:\QC\8. Codex.
"@
    Set-Content -LiteralPath (Join-Path $SyncRoot "README.md") -Value $readme -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $SyncRoot ".gitignore") -Value "git_auto_save.log`r`n" -Encoding ASCII
}

function Sync-ToGitHub() {
    try {
        Prepare-GitHubRepo
        Clear-SyncRoot
        Copy-SourcesToSyncRoot
        Write-Readme

        git -C $SyncRoot add -A | Out-Null
        $status = git -C $SyncRoot status --porcelain
        if ($status) {
            $message = "Auto upload " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            git -C $SyncRoot commit -m $message | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "GitHub upload commit failed" }
            git -C $SyncRoot push origin main | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "GitHub push failed" }
            Write-Log "GITHUB PUSH OK - $message"
        } else {
            git -C $SyncRoot push origin main | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "GitHub push check failed" }
            Write-Log "GITHUB NOCHANGE - already up to date"
        }
    } catch {
        $script:HadError = $true
        Write-Log "GITHUB ERROR - $($_.Exception.Message)"
    }
}

foreach ($Repo in $SourceRepos) {
    Save-LocalRepo $Repo
}
Sync-ToGitHub

if ($HadError) { exit 1 }
exit 0

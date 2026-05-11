$ErrorActionPreference = "Continue"
$Repos = @(
    "D:\QC\8. Quick Access Searcher",
    "D:\QC\9. Codex"
)
$LogPath = "D:\QC\9. Codex\git_auto_save.log"
function Write-Log($Message) {
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $LogPath -Value "[$time] $Message" -Encoding UTF8
}
foreach ($Repo in $Repos) {
    try {
        if (-not (Test-Path -LiteralPath $Repo)) {
            Write-Log "SKIP missing folder: $Repo"
            continue
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
            Write-Log "COMMIT $Repo - $message"
        } else {
            Write-Log "NOCHANGE $Repo"
        }
        $remotes = git -C $Repo remote
        if ($remotes -contains "origin") {
            git -C $Repo push origin HEAD | Out-Null
            Write-Log "PUSH OK $Repo"
        } else {
            Write-Log "NO REMOTE $Repo - local commit only"
        }
    } catch {
        Write-Log "ERROR $Repo - $($_.Exception.Message)"
    }
}

$ErrorActionPreference = "Continue"
$RepoUrl = "https://github.com/the5kka/The5kka.git"
$BasePath = "D:\QC\8. The5kka"
$SourceRepos = @(
    "D:\QC\8. The5kka\1. Quick Access Searcher",
    "D:\QC\8. The5kka\2. Codex"
)
$SyncRoot = "D:\QC\8. The5kka\0. The5kka_GitHub"
$LogPath = "D:\QC\8. The5kka\2. Codex\git_auto_save.log"
$HadError = $false

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

function Sync-ToGitHub() {
    try {
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
        if ($LASTEXITCODE -ne 0) { throw "GitHub fetch 실패" }
        git -C $SyncRoot checkout -B main origin/main | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "main 브랜치 준비 실패" }

        $resolved = (Resolve-Path -LiteralPath $SyncRoot).Path
        if ($resolved -ne "D:\QC\8. The5kka\0. The5kka_GitHub") { throw "예상 동기화 폴더가 아닙니다: $resolved" }
        Get-ChildItem -LiteralPath $SyncRoot -Force | Where-Object { $_.Name -ne ".git" } | Remove-Item -Recurse -Force

        New-Item -ItemType Directory -Path (Join-Path $SyncRoot "1. Quick Access Searcher") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $SyncRoot "2. Codex") -Force | Out-Null
        robocopy "D:\QC\8. The5kka\1. Quick Access Searcher" (Join-Path $SyncRoot "1. Quick Access Searcher") /MIR /XD .git logs build dist __pycache__ /XF desktop.ini *.log *.pyc *.pyo *.tmp *.bak | Out-Null
        robocopy "D:\QC\8. The5kka\2. Codex" (Join-Path $SyncRoot "2. Codex") /MIR /XD .git logs build dist __pycache__ "0. The5kka_GitHub" "1. The5kka_GitHub" "10. The5kka_GitHub" /XF desktop.ini git_auto_save.log *.log *.pyc *.pyo *.tmp *.bak | Out-Null

        $readme = @"
# The5kka

자동 업로드 기준 폴더입니다.

- `1. Quick Access Searcher`
- `2. Codex`

PC 종료 전 `Ctrl + Alt + Q`를 누르면 D:\QC의 원본 폴더를 이 저장소로 동기화하고 GitHub에 업로드합니다.
"@
        Set-Content -LiteralPath (Join-Path $SyncRoot "README.md") -Value $readme -Encoding UTF8

        git -C $SyncRoot add -A | Out-Null
        $status = git -C $SyncRoot status --porcelain
        if ($status) {
            $message = "Auto upload " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            git -C $SyncRoot commit -m $message | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "GitHub 업로드 커밋 실패" }
            git -C $SyncRoot push origin main | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "GitHub push 실패" }
            Write-Log "GITHUB PUSH OK - $message"
        } else {
            git -C $SyncRoot push origin main | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "GitHub push 확인 실패" }
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

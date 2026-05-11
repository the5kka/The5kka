Add-Type -AssemblyName System.Windows.Forms
$result = [System.Windows.Forms.MessageBox]::Show(
    "Codex git 저장 후 GitHub에 업로드하고 PC를 종료할까요?",
    "Codex 업로드 후 종료",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)
if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
    exit
}
$AutoSave = "D:\QC\8. The5kka\3. Codex\git_auto_save.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $AutoSave
if ($LASTEXITCODE -ne 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "GitHub 업로드에 실패해서 PC 종료를 중단했습니다.`nD:\QC\9. Codex\git_auto_save.log 를 확인해주세요.",
        "Codex 업로드 실패",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}
Start-Sleep -Seconds 2
shutdown.exe /s /t 0


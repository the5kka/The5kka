Add-Type -AssemblyName System.Windows.Forms
$result = [System.Windows.Forms.MessageBox]::Show(
    "Codex git 저장 후 PC를 종료할까요?",
    "Codex 저장 후 종료",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)
if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
    exit
}
$AutoSave = "D:\QC\9. Codex\git_auto_save.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $AutoSave
Start-Sleep -Seconds 2
shutdown.exe /s /t 0

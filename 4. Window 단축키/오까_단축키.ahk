#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================
; 오까_단축키
; - Shift + F1: QC 루트 폴더 열기
; - Shift + F9 ~ F12: 자주 쓰는 프로그램 실행/활성화
; - 윈도우 시작 시 자동 실행 바로가기 등록
; ============================================================

; ------------------------------------------------------------
; [수정 영역] 기본 경로와 프로그램 경로
; 아래 값만 바꾸면 전체 단축키 경로를 쉽게 변경할 수 있습니다.
; ------------------------------------------------------------
QC_ROOT := "D:\QC"
SCRIPT_DIR := QC_ROOT "\8. Codex\4. Window 단축키"
SCRIPT_PATH := SCRIPT_DIR "\오까_단축키.ahk"
SETTINGS_PATH := SCRIPT_DIR "\오까_단축키.ini"

QUICK_ACCESS_SEARCHER_PATH := QC_ROOT "\8. Codex\1. Quick Access Searcher\Quick_Access_Searcher.exe"
KAKAOTALK_PATH := "C:\Program Files (x86)\Kakao\KakaoTalk\KakaoTalk.exe"

; ------------------------------------------------------------
; 스크립트 시작 시 스크립트 저장 위치를 준비하고 시작프로그램 바로가기를 자동 등록합니다.
; ------------------------------------------------------------
SetupStartupShortcut(SCRIPT_DIR, SCRIPT_PATH)

; ------------------------------------------------------------
; Windows 시작프로그램으로 실행된 경우, 저장된 설정에 따라 Shift + F11 매크로를 자동 실행합니다.
; ------------------------------------------------------------
if IsLaunchedFromStartup() {
    SetTimer(RunStartupMacroIfEnabled, -5000)
}

; ------------------------------------------------------------
; Shift + F1: QC 루트 폴더 열기
; ------------------------------------------------------------
+F1::OpenFolder(QC_ROOT)

; ------------------------------------------------------------
; Shift + F9: Excel 실행 또는 기존 창 활성화
; 창을 항상 최대 크기로 표시합니다.
; ------------------------------------------------------------
+F9::ActivateOrRunProgramMaximized("ahk_exe EXCEL.EXE", "excel.exe", "Excel")

; ------------------------------------------------------------
; Shift + F10: PowerPoint 실행 또는 기존 창 활성화
; ------------------------------------------------------------
+F10::ActivateOrRunProgram("ahk_exe POWERPNT.EXE", "powerpnt.exe", "PowerPoint")

; ------------------------------------------------------------
; Shift + F11:
; Windows 시작 시 Quick Access Searcher + Outlook 자동 실행 설정을 ON/OFF 합니다.
; ON으로 변경하면 현재도 바로 한 번 실행합니다.
; ------------------------------------------------------------
+F11::ToggleStartupMacro()

; ------------------------------------------------------------
; Shift + F12: KakaoTalk 실행 또는 기존 창 활성화
; ------------------------------------------------------------
+F12::ActivateOrRunFile("ahk_exe KakaoTalk.exe", KAKAOTALK_PATH, "KakaoTalk")

; ============================================================
; 함수 모음
; ============================================================

; ------------------------------------------------------------
; 폴더가 있으면 열고, 없으면 오류 메시지를 표시합니다.
; ------------------------------------------------------------
OpenFolder(folderPath) {
    if !DirExist(folderPath) {
        ShowPathNotFound(folderPath)
        return
    }

    TryRun(folderPath, "폴더를 열 수 없습니다.`n`n경로:`n" folderPath)
}

; ------------------------------------------------------------
; 파일이 있으면 실행하고, 없으면 오류 메시지를 표시합니다.
; ------------------------------------------------------------
RunFile(filePath) {
    if !FileExist(filePath) {
        ShowPathNotFound(filePath)
        return false
    }

    return TryRun(filePath, "파일을 실행할 수 없습니다.`n`n경로:`n" filePath)
}

; ------------------------------------------------------------
; 프로그램 창이 있으면 활성화하고, 없으면 실행 명령으로 실행합니다.
; 예: Excel, PowerPoint, Outlook처럼 PATH에 등록된 프로그램에 사용합니다.
; ------------------------------------------------------------
ActivateOrRunProgram(windowTitle, runCommand, programName) {
    if ActivateWindow(windowTitle) {
        return true
    }

    return TryRun(runCommand, programName "을(를) 실행할 수 없습니다.")
}

; ------------------------------------------------------------
; 프로그램 창이 있으면 활성화 후 최대화하고, 없으면 실행 후 최대화합니다.
; ------------------------------------------------------------
ActivateOrRunProgramMaximized(windowTitle, runCommand, programName) {
    if ActivateWindowMaximized(windowTitle) {
        return true
    }

    if !TryRun(runCommand, programName "을(를) 실행할 수 없습니다.") {
        return false
    }

    return WaitAndMaximizeWindow(windowTitle, programName)
}

; ------------------------------------------------------------
; 프로그램 창이 있으면 활성화하고, 없으면 지정된 파일 경로로 실행합니다.
; 예: KakaoTalk처럼 기본 설치 경로가 있는 프로그램에 사용합니다.
; ------------------------------------------------------------
ActivateOrRunFile(windowTitle, filePath, programName) {
    if ActivateWindow(windowTitle) {
        return true
    }

    if !FileExist(filePath) {
        ShowPathNotFound(filePath)
        return false
    }

    return TryRun(filePath, programName "을(를) 실행할 수 없습니다.`n`n경로:`n" filePath)
}

; ------------------------------------------------------------
; 프로그램 창이 있으면 활성화 후 최대화하고, 없으면 지정된 파일 실행 후 최대화합니다.
; ------------------------------------------------------------
ActivateOrRunFileMaximized(windowTitle, filePath, programName) {
    if ActivateWindowMaximized(windowTitle) {
        return true
    }

    if !FileExist(filePath) {
        ShowPathNotFound(filePath)
        return false
    }

    if !TryRun(filePath, programName "을(를) 실행할 수 없습니다.`n`n경로:`n" filePath) {
        return false
    }

    return WaitAndMaximizeWindow(windowTitle, programName)
}

; ------------------------------------------------------------
; Shift + F11 매크로 본문입니다.
; Quick Access Searcher와 Outlook을 실행 또는 활성화하고 최대화합니다.
; ------------------------------------------------------------
RunShiftF11Macro() {
    ActivateOrRunQuickAccessSearcher()
    ActivateOrRunProgramMaximized("ahk_exe OUTLOOK.EXE", "outlook.exe", "Outlook")
}

; ------------------------------------------------------------
; Windows 시작 시 자동 실행 설정이 켜져 있으면 Shift + F11 매크로를 실행합니다.
; ------------------------------------------------------------
RunStartupMacroIfEnabled() {
    if IsStartupMacroEnabled() {
        RunShiftF11Macro()
    }
}

; ------------------------------------------------------------
; Shift + F11로 다음 Windows 시작 시 자동 실행 설정을 ON/OFF 합니다.
; ------------------------------------------------------------
ToggleStartupMacro() {
    enabled := !IsStartupMacroEnabled()
    SetStartupMacroEnabled(enabled)

    if enabled {
        RunShiftF11Macro()
        MsgBox("[오까_단축키]`nPC 시작 시 자동 실행: ON`n`n다음 Windows 시작부터 Quick Access Searcher와 Outlook이 자동으로 실행됩니다.", "오까_단축키", "Iconi")
        return
    }

    MsgBox("[오까_단축키]`nPC 시작 시 자동 실행: OFF`n`n다음 Windows 시작부터 Quick Access Searcher와 Outlook을 자동으로 실행하지 않습니다.`n현재 열려 있는 창은 닫지 않습니다.", "오까_단축키", "Iconi")
}

; ------------------------------------------------------------
; INI 설정 파일에서 시작 자동 실행 여부를 읽습니다.
; 설정 파일이 없으면 기본값은 ON입니다.
; ------------------------------------------------------------
IsStartupMacroEnabled() {
    return IniRead(SETTINGS_PATH, "StartupMacro", "Enabled", "1") = "1"
}

; ------------------------------------------------------------
; 시작 자동 실행 여부를 INI 설정 파일에 저장합니다.
; ------------------------------------------------------------
SetStartupMacroEnabled(enabled) {
    IniWrite(enabled ? "1" : "0", SETTINGS_PATH, "StartupMacro", "Enabled")
}

; ------------------------------------------------------------
; 시작프로그램 바로가기에서 /startup 인수로 실행됐는지 확인합니다.
; ------------------------------------------------------------
IsLaunchedFromStartup() {
    for arg in A_Args {
        if StrLower(arg) = "/startup" {
            return true
        }
    }

    return false
}

; ------------------------------------------------------------
; Quick Access Searcher 전용 실행 함수입니다.
; 창 없는 잔여 프로세스가 있으면 새 창 중복 생성을 막기 위해 먼저 종료합니다.
; ------------------------------------------------------------
ActivateOrRunQuickAccessSearcher() {
    windowTitle := "ahk_exe Quick_Access_Searcher.exe"
    programName := "Quick Access Searcher"

    if ActivateWindowMaximized(windowTitle) {
        return true
    }

    closedAny := false
    while pid := ProcessExist("Quick_Access_Searcher.exe") {
        try {
            ProcessClose(pid)
            ProcessWaitClose(pid, 2)
            closedAny := true
        } catch as err {
            ShowError(programName " 잔여 프로세스를 정리할 수 없습니다.`n`n오류:`n" err.Message)
            return false
        }
    }

    if !FileExist(QUICK_ACCESS_SEARCHER_PATH) {
        ShowPathNotFound(QUICK_ACCESS_SEARCHER_PATH)
        return false
    }

    if !TryRunWithWorkDir(QUICK_ACCESS_SEARCHER_PATH, programName "을(를) 실행할 수 없습니다.`n`n경로:`n" QUICK_ACCESS_SEARCHER_PATH, DirName(QUICK_ACCESS_SEARCHER_PATH)) {
        return false
    }

    return WaitAndMaximizeWindow(windowTitle, programName, 20)
}

; ------------------------------------------------------------
; 창이 존재하면 활성화합니다.
; ------------------------------------------------------------
ActivateWindow(windowTitle) {
    if !WinExist(windowTitle) {
        return false
    }

    try {
        WinActivate(windowTitle)
        WinWaitActive(windowTitle, , 2)
        return true
    } catch as err {
        ShowError("창을 활성화할 수 없습니다.`n`n대상:`n" windowTitle "`n`n오류:`n" err.Message)
        return false
    }
}

; ------------------------------------------------------------
; 창이 존재하면 활성화한 뒤 최대화합니다.
; ------------------------------------------------------------
ActivateWindowMaximized(windowTitle) {
    if !ActivateWindow(windowTitle) {
        return false
    }

    try {
        WinMaximize(windowTitle)
        return true
    } catch as err {
        ShowError("창을 최대화할 수 없습니다.`n`n대상:`n" windowTitle "`n`n오류:`n" err.Message)
        return false
    }
}

; ------------------------------------------------------------
; 프로그램 실행 직후 창이 뜰 때까지 기다린 뒤 최대화합니다.
; ------------------------------------------------------------
WaitAndMaximizeWindow(windowTitle, programName, waitSeconds := 8) {
    try {
        if !WinWait(windowTitle, , waitSeconds) {
            ShowError(programName " 창을 찾을 수 없습니다.")
            return false
        }

        WinActivate(windowTitle)
        WinWaitActive(windowTitle, , 3)
        WinMaximize(windowTitle)
        return true
    } catch as err {
        ShowError(programName " 창을 최대화할 수 없습니다.`n`n오류:`n" err.Message)
        return false
    }
}

; ------------------------------------------------------------
; Run 명령 실행 중 오류가 발생하면 MsgBox로 표시합니다.
; ------------------------------------------------------------
TryRun(target, errorMessage) {
    try {
        Run(target)
        return true
    } catch as err {
        ShowError(errorMessage "`n`n오류:`n" err.Message)
        return false
    }
}

; ------------------------------------------------------------
; 지정한 작업 폴더에서 프로그램을 실행합니다.
; ------------------------------------------------------------
TryRunWithWorkDir(target, errorMessage, workingDir) {
    try {
        Run(target, workingDir)
        return true
    } catch as err {
        ShowError(errorMessage "`n`n오류:`n" err.Message)
        return false
    }
}

; ------------------------------------------------------------
; 파일 경로에서 폴더 경로만 반환합니다.
; ------------------------------------------------------------
DirName(filePath) {
    SplitPath(filePath, , &dir)
    return dir
}

; ------------------------------------------------------------
; 스크립트 저장 폴더를 준비한 뒤 시작프로그램 바로가기를 등록합니다.
; 다른 위치에서 실행해도 지정된 위치로 복사한 다음 그 파일을 등록합니다.
; ------------------------------------------------------------
SetupStartupShortcut(scriptDir, scriptPath) {
    if !DirExist(scriptDir) {
        try {
            DirCreate(scriptDir)
        } catch as err {
            ShowError("스크립트 저장 폴더를 만들 수 없습니다.`n`n경로:`n" scriptDir "`n`n오류:`n" err.Message)
            return
        }
    }

    if !FileExist(scriptPath) {
        try {
            FileCopy(A_ScriptFullPath, scriptPath, true)
        } catch as err {
            ShowError("스크립트 파일을 지정된 위치로 복사할 수 없습니다.`n`n경로:`n" scriptPath "`n`n오류:`n" err.Message)
            return
        }
    }

    startupFolder := A_Startup
    shortcutPath := startupFolder "\오까_단축키.lnk"

    try {
        ; 공백이 있는 경로도 안전하게 실행되도록 AutoHotkey 실행 파일에 스크립트 경로를 인수로 전달합니다.
        shortcutArgs := "`"" scriptPath "`" /startup"
        FileCreateShortcut(A_AhkPath, shortcutPath, scriptDir, shortcutArgs, "오까_단축키 자동 실행")
    } catch as err {
        ShowError("시작프로그램 바로가기를 만들 수 없습니다.`n`n경로:`n" shortcutPath "`n`n오류:`n" err.Message)
    }
}

; ------------------------------------------------------------
; 파일 또는 폴더 경로가 없을 때 표시하는 공통 메시지입니다.
; ------------------------------------------------------------
ShowPathNotFound(path) {
    MsgBox("[오까_단축키]`n경로를 찾을 수 없습니다.`n`n경로:`n" path, "오까_단축키", "Icon!")
}

; ------------------------------------------------------------
; 일반 오류 메시지를 표시합니다.
; ------------------------------------------------------------
ShowError(message) {
    MsgBox("[오까_단축키]`n" message, "오까_단축키", "Icon!")
}

#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================
; 오까_단축키
; - Shift + F1: QC 루트 폴더 열기
; - Shift + Esc: KakaoTalk 실행/맨앞으로 가져오기
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
KAKAOTALK_PATH := "C:\Program Files\Kakao\KakaoTalk\KakaoTalk.exe"

; ------------------------------------------------------------
; 스크립트 시작 시 스크립트 저장 위치를 준비하고 시작프로그램 바로가기를 자동 등록합니다.
; ------------------------------------------------------------
SetupStartupShortcut(SCRIPT_DIR, SCRIPT_PATH)

; ------------------------------------------------------------
; Windows 시작프로그램으로 실행되면 단축키만 등록합니다.
; ------------------------------------------------------------

; ------------------------------------------------------------
; Shift + F1: QC 루트 폴더 열기
; ------------------------------------------------------------
+F1::OpenFolder(QC_ROOT)

; ------------------------------------------------------------
; Shift + Esc: KakaoTalk 실행 또는 기존 창을 맨앞으로 가져오기
; ------------------------------------------------------------
+Esc::ActivateOrRunFile("ahk_exe KakaoTalk.exe", KAKAOTALK_PATH, "KakaoTalk")

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

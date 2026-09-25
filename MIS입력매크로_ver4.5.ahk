#NoEnv
#SingleInstance Force
SetWorkingDir %A_ScriptDir%

; 화면 전체 기준 좌표 설정
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen
CoordMode, ToolTip, Screen

; 기본 파일 경로 및 전역 설정
ImageFolder := A_ScriptDir "\img"
ExcelPath   := A_ScriptDir "\검사계획요청(매크로용).xlsx"
IniFile     := A_ScriptDir "\settings.ini"
Variation   := 30

; 전역 상태 변수
global CurrentVehicleIndex := 1, TotalCount := 1, CurrentStepText := "준비 중..."
global P_작업요청자 := "IT203094", P_검사자 := "IT203094", P_확인자 := "IT201751", P_담당자 := "IT201495"
global P_운전실기기 := "IT203094", P_객실기기 := "IT203094", P_대차및하부_옥상 := "IT203094", P_입환_면허자 := "IT203094"
global G_SavedPatternNum := 1, G_SavedRequester := "IT203094", P_TargetPatternNum := 9
global G_SavedStaff_1 := "IT203094", G_SavedStaff_2 := "IT203094", G_SavedStaff_3 := "IT203094", G_SavedStaff_4 := "IT203094"
global G_SavedInspector := "IT203094", G_SavedChecker := "IT201751", G_SavedManager := "IT201495"
global G_RecentStaffPipe := "IT203094|IT201751|IT201495"
global G_IsExcelSavedToday := False
global G_ActivePatternBtn := 0, G_BlinkState := False

; 🌟 접속 환경 설정 변수 ("GW": 그룹웨어 웹 / "APP": MIS앱) 및 Y좌표 오프셋
global G_EnvMode := "GW", G_YOffset := 0
global Var_DashProg, Var_DashStep, Var_DashStaff

; 🌟 [v4.4] MIS 앱 창 핸들 (로딩/응답없음 감지용)
global G_MisHwnd := 0, G_MisPid := 0

; 시작 시 이전 저장 설정 및 최근 사번 목록 자동 로드
LoadSettings()

; =================================================================
; 🌟 [1단계] 프로그램 실행 즉시 화면 중앙에 접속환경 선택창 표시
; =================================================================
ShowEnvSelectionGui()
return

ShowEnvSelectionGui()
{
    global
    Gui, EnvGui:New, +AlwaysOnTop -MaximizeBox -MinimizeBox, MIS 접속환경 선택
    Gui, EnvGui:Color, F0F4F8, FFFFFF

    Gui, EnvGui:Font, s14 Bold c1E3A8A, 맑은 고딕
    Gui, EnvGui:Add, Text, x20 y18 w480 Center, 🏢 MIS 접속 환경을 선택해 주세요

    Gui, EnvGui:Font, s9 Normal c555555, 맑은 고딕
    Gui, EnvGui:Add, Text, x20 y48 w480 Center, 접속 방식에 따라 클릭 좌표와 일상검사 표 이동이 자동 보정됩니다.

    ; 큼지막한 좌측 버튼 (그룹웨어 웹 접속)
    Gui, EnvGui:Font, s11 Bold c0D47A1, 맑은 고딕
    Gui, EnvGui:Add, Button, x25 y80 w230 h95 +0x2000 gSelectGroupware, 🌐 웹 브라우저`n[그룹웨어 접속]`n(기본 표준 좌표)

    ; 큼지막한 우측 버튼 (MIS앱 접속 - Y좌표 -10px, 전용 이미지, 일상검사 정밀 보정)
    Gui, EnvGui:Font, s11 Bold c1B5E20, 맑은 고딕
    Gui, EnvGui:Add, Button, x265 y80 w230 h95 +0x2000 gSelectMISApp, 💻 전용 클라이언트`n[MIS앱 접속]`n(좌표 및 이동 최적화)

    Gui, EnvGui:Show, Center w520 h195, MIS 접속환경 선택
}

EnvGuiClose:
    ExitApp
return

SelectGroupware:
    G_EnvMode := "GW"
    G_YOffset := 0
    IniWrite, %G_EnvMode%, %IniFile%, Config, EnvMode
    Gui, EnvGui:Destroy
    CreateAndShowMainGui()
return

SelectMISApp:
    G_EnvMode := "APP"
    G_YOffset := -10
    IniWrite, %G_EnvMode%, %IniFile%, Config, EnvMode
    Gui, EnvGui:Destroy
    CreateAndShowMainGui()
return

; =================================================================
; 🌟 [2단계] 메인 GUI 생성 및 표시
; =================================================================
CreateAndShowMainGui()
{
    global
    Gui, 1:Default
    Gui, Color, F0F4F8, FFFFFF

    Gui, Font, s13 Bold, 맑은 고딕
    Gui, Add, Text, x20 y8 w460 Center c1E3A8A, 🚆 인천교통공사 MIS입력 매크로 ver4.5-Dev 🚆

    ; 현재 접속 환경 안내 배지 및 변경 버튼
    ModeText := (G_EnvMode = "APP") ? "💻 접속환경: [MIS 앱 접속] (Y-10px / 전용 이미지 / 행 이동 최적화)" : "🌐 접속환경: [그룹웨어 웹 접속] (표준 좌표 / 웹 표준 설정)"
    ModeColor := (G_EnvMode = "APP") ? "1B5E20" : "0D47A1"

    Gui, Font, s9 Bold c%ModeColor%, 맑은 고딕
    Gui, Add, Text, vVar_EnvBadge x20 y34 w365 h22 +0x200, %ModeText%
    Gui, Font, s8 Normal c333333, 맑은 고딕
    Gui, Add, Button, x390 y33 w95 h24 gChangeEnvMode, 🔁 환경변경

    ; 시각적 시선 유도 대형 화살표 가이드 라벨
    Gui, Font, s10 Bold cB71C1C, 맑은 고딕
    Gui, Add, Text, vVar_ArrowGuide x15 y58 w470 h18 Center, 👇 👇 👇 [필수 1단계] 아래 엑셀 편성 버튼을 먼저 눌러주세요! 👇 👇 👇

    ; 1. [★선작업필수] 엑셀 차량 편성 입력 버튼
    Gui, Font, s10 Bold, 맑은 고딕
    Gui, Add, Button, vBtnExcel x15 y78 w470 h44 gOpenExcelMenu, 📊 [★선작업필수] 엑셀 차량 편성 입력 및 검사계획요청(매크로용) 생성 📑

    ; 엑셀 준비 상태 배지
    Gui, Font, s9 Bold cB71C1C, 맑은 고딕
    Gui, Add, Text, vVar_ExcelReadyStatus x15 y125 w470 h20 Center, 🔴 🔒 [잠금] 오늘자 엑셀 미작성 (위 화살표가 가리키는 버튼을 먼저 클릭하세요!)

    ; 2. 담당자 사번 입력 상자
    Gui, Font, s9 Normal cBlack, 맑은 고딕
    Gui, Add, GroupBox, x15 y146 w470 h122 c1E3A8A, 👨‍✈️ 운행/점검 담당자 사번 설정 (자주쓰는사번은 드롭다운클릭)

    Gui, Font, s9 Bold cB71C1C, 맑은 고딕
    Gui, Add, Text, x28 y170 w70 h20, ★ 검사자:
    Gui, Font, s9 Normal cBlack, 맑은 고딕
    Gui, Add, ComboBox, vVar_검사자 gOnStaffTextChange x98 y167 w135 h150, %G_RecentStaffPipe%

    Gui, Font, s9 Bold cB71C1C, 맑은 고딕
    Gui, Add, Text, x240 y170 w75 h20, ★ 확인자:
    Gui, Font, s9 Normal cBlack, 맑은 고딕
    Gui, Add, ComboBox, vVar_확인자 gOnStaffTextChange x315 y167 w135 h150, %G_RecentStaffPipe%

    Gui, Font, s9 Bold cB71C1C, 맑은 고딕
    Gui, Add, Text, x28 y203 w70 h20, ★ 담당자:
    Gui, Font, s9 Normal cBlack, 맑은 고딕
    Gui, Add, ComboBox, vVar_담당자 gOnStaffTextChange x98 y200 w135 h150, %G_RecentStaffPipe%

    ; 사번 정밀 규격 검증 상태 배지
    Gui, Font, s8 Bold c00897B, 맑은 고딕
    Gui, Add, Text, vVar_StaffValidationStatus x240 y203 w230 h20 Center, ✔ 사번 규격 검증 완료 (IT+6자리)

    ; 3. 입고검사 버튼 그룹 (1~5번)
    Gui, Font, s9 Bold c004D40, 맑은 고딕
    Gui, Add, GroupBox, x15 y275 w470 h215, 📥 [입고검사] 전동차 패턴 (1번~5번)
    Gui, Font, s9 Normal c004D40, 맑은 고딕
    Gui, Add, Button, vBtnTask1 x30 y297 w215 h42 gDoTask1, 🚇 1. 평일저녁 입고 1차 (16대)
    Gui, Add, Button, vBtnTask2 x255 y297 w215 h42 +0x2000 gDoTask2, 🚉 2. 평일저녁 입고 2차 (7대)`n(서구청 포함)
    Gui, Add, Button, vBtnTask3 x30 y335 w215 h42 gDoTask3, 🌙 3. 주말저녁 입고 1차 (4대)
    Gui, Add, Button, vBtnTask4 x255 y335 w215 h42 +0x2000 gDoTask4, 🚉 4. 주말저녁 입고 2차 (7대)`n(서구청 포함)
    Gui, Font, s10 Bold c4A148C, 맑은 고딕
    Gui, Add, Button, vBtnTask5 x30 y393 w440 h42 gDoTaskMorning, 🌅 5. 평일오전 입고(8~10대) (입고시각 입력필요)

    ; 4. 출고검사 버튼 그룹 (6~8번)
    Gui, Font, s9 Bold c880E4F, 맑은 고딕
    Gui, Add, GroupBox, x15 y497 w470 h165, 📤 [출고검사] 전동차 패턴 (6번~8번)
    Gui, Font, s9 Normal c880E4F, 맑은 고딕
    Gui, Add, Button, vBtnTask6 x30 y519 w215 h42 gDoTask6, ⚡ 6. 평일새벽 출고 (29대)
    Gui, Add, Button, vBtnTask7 x255 y519 w215 h42 gDoTask7, ☀️ 7. 평일오후 출고 (9대)
    Gui, Add, Button, vBtnTask8 x30 y567 w440 h42 +0x2000 gDoTask8, 🚉 8. 주말새벽 출고 (12대) (서구청 포함)

    ; 5. 일상검사 버튼 그룹 (9번 주간 / 10번 야간)
    Gui, Font, s9 Bold c1E3A8A, 맑은 고딕
    Gui, Add, GroupBox, x15 y669 w470 h100, 🔧 [일상검사] 전동차 패턴 (9번~10번)
    Gui, Font, s9 Bold c1E3A8A, 맑은 고딕
    Gui, Add, Button, vBtnTask9 x30 y692 w215 h45 gDoTask9, ☀️ 9. 일상검사 주간 (4대)
    Gui, Add, Button, vBtnTask10 x255 y692 w215 h45 gDoTask10, 🌙 10. 일상검사 야간 (2대)

    ; 하단 캡션
    Gui, Font, s9 c475569, 맑은 고딕
    Gui, Add, Text, x20 y778 w460 Center, 🚉 MIS입력 자동화  |  Esc: 강제종료  |  Shift+Enter: 일시정지

    ; 저장된 사번에 이름을 붙여 기본값 세팅
    GuiControl, 1:Text, Var_검사자, % GetStaffDisplayWithID(G_SavedInspector)
    GuiControl, 1:Text, Var_확인자, % GetStaffDisplayWithID(G_SavedChecker)
    GuiControl, 1:Text, Var_담당자, % GetStaffDisplayWithID(G_SavedManager)

    Gui, Show, Center w500 h805, MIS입력 매크로 ver4.5-Dev

    ValidateStaffInputs()
    UpdateExcelStatusAndButtons()
}

ChangeEnvMode:
    Gui, 1:Destroy
    ShowEnvSelectionGui()
return

GuiClose:
    ExitApp
return

; =================================================================
; 🌟 [고가시성 점멸] 활성화된 버튼 반짝반짝 애니메이션 타이머
; =================================================================
BlinkActiveButton:
    if (G_ActivePatternBtn = 0)
    {
        SetTimer, BlinkActiveButton, Off
        return
    }

    G_BlinkState := !G_BlinkState
    GetPatternData(G_ActivePatternBtn, ActiveTitle, YList, StartList, EndList, Code, EquipList)

    if (G_BlinkState)
    {
        GuiControl, 1:Text, BtnTask%G_ActivePatternBtn%, % "👉 ⭐ [" . G_ActivePatternBtn . "번] " . ActiveTitle . " 👈 (클릭하여 실행!)"
        GuiControl, 1:+c00897B, Var_ArrowGuide
        GuiControl, 1:, Var_ArrowGuide, % "✨✨✨ [실행 가능] " . G_ActivePatternBtn . "번 패턴 버튼이 활성화되었습니다! ✨✨✨"
    }
    else
    {
        GuiControl, 1:Text, BtnTask%G_ActivePatternBtn%, % "🚀 [" . G_ActivePatternBtn . "번] " . ActiveTitle . " (실행 준비 완료!)"
        GuiControl, 1:+c1E3A8A, Var_ArrowGuide
        GuiControl, 1:, Var_ArrowGuide, % "✔ [패턴 연동 완료] [" . ActiveTitle . "] 엑셀 준비됨!"
    }
return

; =================================================================
; 🌟 사번 정밀 규격 검증 함수 (IT + 정확히 숫자 6자리)
; =================================================================
IsValidStaffID(IDStr)
{
    IDStr := Trim(IDStr)
    return RegExMatch(IDStr, "i)^IT\d{6}$") ? True : False
}

OnStaffTextChange:
    ValidateStaffInputs()
return

ValidateStaffInputs()
{
    GuiControlGet, Insp, 1:, Var_검사자
    GuiControlGet, Chk,  1:, Var_확인자
    GuiControlGet, Mngr, 1:, Var_담당자

    Insp := ExtractPureStaffID(Insp)
    Chk  := ExtractPureStaffID(Chk)
    Mngr := ExtractPureStaffID(Mngr)

    if (IsValidStaffID(Insp) && IsValidStaffID(Chk) && IsValidStaffID(Mngr))
    {
        GuiControl, 1:+c00897B, Var_StaffValidationStatus
        GuiControl, 1:, Var_StaffValidationStatus, ✔ 사번 규격 검증 완료 (IT+6자리)
        GuiControl, 1:Enable, BtnExcel
        return True
    }
    else
    {
        GuiControl, 1:+cB71C1C, Var_StaffValidationStatus
        GuiControl, 1:, Var_StaffValidationStatus, ❌ 사번 오류 (예: IT203094)
        GuiControl, 1:Disable, BtnExcel
        return False
    }
}

; =================================================================
; 🌟 엑셀 준비 상태 검증 및 버튼 활성화 제어
; =================================================================
IsTodayExcelFile(FilePath)
{
    global G_IsExcelSavedToday
    if (G_IsExcelSavedToday)
        return True

    if (!FileExist(FilePath))
        return False

    FormatTime, TodayDate,, yyyyMMdd
    FileGetTime, FileModTime, %FilePath%, M
    FormatTime, FileModDate, %FileModTime%, yyyyMMdd

    return (TodayDate = FileModDate)
}

UpdateExcelStatusAndButtons()
{
    global
    TargetExcel := GetValidExcelPath()
    IsReady := IsTodayExcelFile(TargetExcel)

    if (IsReady)
    {
        ActivePattern := G_SavedPatternNum
        G_ActivePatternBtn := ActivePattern
        GetPatternData(ActivePattern, ActiveTitle, 차종List, 시작List, 종료List, SheetCode, DefaultEquipList)

        GuiControl, 1:+c00897B, Var_ExcelReadyStatus
        GuiControl, 1:, Var_ExcelReadyStatus, % "🟢 [준비완료] " . ActivePattern . "번 버튼만 실행 가능 (나머지 패턴 실수 방지 잠금)"

        Loop, 10
        {
            GetPatternData(A_Index, OrigTitle, YL, SL, EL, SC, EL2)
            if (A_Index = ActivePattern)
            {
                GuiControl, 1:Enable, BtnTask%A_Index%
                GuiControl, 1:Text, BtnTask%A_Index%, % "👉 ⭐ [" . A_Index . "번] " . OrigTitle . " 👈 (클릭하여 실행!)"
            }
            else
            {
                GuiControl, 1:Disable, BtnTask%A_Index%
                GuiControl, 1:Text, BtnTask%A_Index%, % A_Index . ". " . OrigTitle
            }
        }

        SetTimer, BlinkActiveButton, 600
    }
    else
    {
        G_ActivePatternBtn := 0
        SetTimer, BlinkActiveButton, Off

        GuiControl, 1:+cB71C1C, Var_ArrowGuide
        GuiControl, 1:, Var_ArrowGuide, 👇 👇 👇 [필수 1단계] 아래 엑셀 편성 버튼을 먼저 눌러주세요! 👇 👇 👇

        GuiControl, 1:+cB71C1C, Var_ExcelReadyStatus
        if (FileExist(TargetExcel))
            GuiControl, 1:, Var_ExcelReadyStatus, 🔴 🔒 [전체 잠금] 전날 이전 엑셀임 (위 버튼을 눌러 오늘자 엑셀을 생성하세요!)
        else
            GuiControl, 1:, Var_ExcelReadyStatus, 🔴 🔒 [전체 잠금] 오늘자 엑셀 미작성 (위 버튼을 눌러 엑셀을 먼저 생성하세요!)

        Loop, 10
        {
            GuiControl, 1:Disable, BtnTask%A_Index%
            GetPatternData(A_Index, OrigTitle, YL, SL, EL, SC, EL2)
            GuiControl, 1:Text, BtnTask%A_Index%, % A_Index . ". " . OrigTitle
        }
    }
}

; =================================================================
; [이벤트 핸들러]
; =================================================================
DoTask1:
DoTask2:
DoTask3:
DoTask4:
DoTask6:
DoTask7:
DoTask8:
    ExecuteTask(SubStr(A_ThisLabel, 7) + 0)
return
DoTask9:
DoTask10:
    OpenDailyStaffGui(SubStr(A_ThisLabel, 7) + 0)
return

SaveStaffSettings()
{
    global
    Gui, 1:Submit, NoHide

    PureInsp := ExtractPureStaffID(Var_검사자)
    PureChk  := ExtractPureStaffID(Var_확인자)
    PureMngr := ExtractPureStaffID(Var_담당자)

    P_검사자 := PureInsp, P_확인자 := PureChk, P_담당자 := PureMngr
    IniWrite, %P_검사자%, %IniFile%, MainStaff, Inspector
    IniWrite, %P_확인자%, %IniFile%, MainStaff, Checker
    IniWrite, %P_담당자%, %IniFile%, MainStaff, Manager
    G_SavedInspector := P_검사자, G_SavedChecker := P_확인자, G_SavedManager := P_담당자
}

CheckExcelFileFirst()
{
    global
    TargetExcel := GetValidExcelPath()
    if (!FileExist(TargetExcel) || !IsTodayExcelFile(TargetExcel))
    {
        SoundBeep, 750, 300
        MsgBox, 48, ⛔ 사전 필수 작업 누락!, % "⚠️ 오늘 자 [검사계획요청(매크로용).xlsx] 엑셀 파일이 존재하지 않거나 전날 이전 파일입니다!`n`n매크로를 시작하기 전, 최상단의`n[📊 [★선작업필수] 엑셀 차량 편성 입력 및 검사계획요청(매크로용) 생성] 버튼을 눌러 오늘 자 엑셀 파일을 새로 생성해 주세요!"
        Gosub, OpenExcelMenu
        return False
    }
    return True
}

; =================================================================
; 🌟 [일상검사 서브 GUI - 드롭다운 이름 완벽 표시]
; =================================================================
OpenDailyStaffGui(PatternNum)
{
    global
    SaveStaffSettings()
    if (!CheckExcelFileFirst())
        return

    P_TargetPatternNum := PatternNum

    Gui, DailyStaffGui:New, +Owner1, 🔧 일상검사 - 항목별 담당자 사번 입력
    Gui, DailyStaffGui:Color, F0F4F8, FFFFFF

    Gui, DailyStaffGui:Font, s11 Bold c1E3A8A, 맑은 고딕
    Gui, DailyStaffGui:Add, Text, x20 y12 w460 Center, 🔧 [일상검사] 점검항목별 담당자 사번 설정

    Gui, DailyStaffGui:Font, s9 Normal cBlack, 맑은 고딕
    Gui, DailyStaffGui:Add, GroupBox, x15 y45 w470 h180 c1E3A8A, 👨‍✈️ 항목별 검사자 사번 기입 (자주쓰는사번은 드롭다운클릭)

    StaffVars := [G_SavedStaff_1, G_SavedStaff_2, G_SavedStaff_3, G_SavedStaff_4]
    Labels    := ["운전실기기", "객실기기", "대차및하부, 옥상", "입환(면허자)"]

    Loop, 4
    {
        YPos := 72 + (A_Index - 1) * 35
        Val := (StaffVars[A_Index] != "") ? StaffVars[A_Index] : "IT203094"
        StringUpper, Val, Val

        DispVal := GetStaffDisplayWithID(Val)

        Gui, DailyStaffGui:Font, s9 Bold cB71C1C, 맑은 고딕
        Gui, DailyStaffGui:Add, Text, % "x30 y" YPos+3 " w130 h20", % "★ " Labels[A_Index] ":"
        Gui, DailyStaffGui:Font, s9 Normal cBlack, 맑은 고딕
        Gui, DailyStaffGui:Add, ComboBox, % "vVar_Staff_" A_Index " x165 y" YPos " w290 h150", %G_RecentStaffPipe%
        GuiControl, DailyStaffGui:Text, Var_Staff_%A_Index%, %DispVal%
    }

    Gui, DailyStaffGui:Font, s10 Bold c1E3A8A, 맑은 고딕
    Gui, DailyStaffGui:Add, Button, x15 y240 w470 h42 gStartDailyTask, 🚀 [일상검사] 매크로 실행하기

    Gui, DailyStaffGui:Show, w500 h295
}

StartDailyTask:
    Gui, DailyStaffGui:Submit, NoHide

    PureStaff1 := ExtractPureStaffID(Var_Staff_1)
    PureStaff2 := ExtractPureStaffID(Var_Staff_2)
    PureStaff3 := ExtractPureStaffID(Var_Staff_3)
    PureStaff4 := ExtractPureStaffID(Var_Staff_4)

    if (!IsValidStaffID(PureStaff1) || !IsValidStaffID(PureStaff2) || !IsValidStaffID(PureStaff3) || !IsValidStaffID(PureStaff4))
    {
        SoundBeep, 750, 300
        MsgBox, 48, 사번 규격 오류, ❌ 일상검사 4가지 사번이 모두 [IT + 숫자 6자리] 규격이어야 합니다! (예: IT203094)
        return
    }

    P_운전실기기 := PureStaff1, P_객실기기 := PureStaff2
    P_대차및하부_옥상 := PureStaff3, P_입환_면허자 := PureStaff4

    IniWrite, %PureStaff1%, %IniFile%, DailyStaff, Staff_1
    IniWrite, %PureStaff2%, %IniFile%, DailyStaff, Staff_2
    IniWrite, %PureStaff3%, %IniFile%, DailyStaff, Staff_3
    IniWrite, %PureStaff4%, %IniFile%, DailyStaff, Staff_4

    G_SavedStaff_1 := PureStaff1, G_SavedStaff_2 := PureStaff2
    G_SavedStaff_3 := PureStaff3, G_SavedStaff_4 := PureStaff4

    Gui, DailyStaffGui:Destroy
    ExecuteTask(P_TargetPatternNum)
return

; =================================================================
; [평일오전 입고 서브 GUI]
; =================================================================
DoTaskMorning:
    SaveStaffSettings()
    if (!CheckExcelFileFirst())
        return

    EquipList := GetCurrentEquipList()
    Count := EquipList.Length()

    Gui, MorningGui:New, +Owner1, 🌅 평일오전 입고 - 차량별 입고시각 입력
    Gui, MorningGui:Color, F0F4F8, FFFFFF

    Gui, MorningGui:Font, s11 Bold c1E3A8A, 맑은 고딕
    Gui, MorningGui:Add, Text, x20 y10 w610 Center, 🌅 [평일오전 입고] 차량별 입고시각(HHMM) 입력
    Gui, MorningGui:Font, s9 Normal c475569, 맑은 고딕
    Gui, MorningGui:Add, Text, x20 y32 w610 Center, 📌 입고시각만 입력해주세요. (예: 0830)

    Gui, MorningGui:Font, s9 Bold c004D40, 맑은 고딕
    Gui, MorningGui:Add, GroupBox, x15 y55 w620 h290, ⏰ 입고시각 입력

    Loop, %Count%
    {
        Col := Floor((A_Index - 1) / 5), Row := Mod(A_Index - 1, 5)
        X_Pos := 25 + (Col * 120), Y_Pos := 78 + (Row * 48)

        Gui, MorningGui:Font, s9 Bold cB71C1C, 맑은 고딕
        Gui, MorningGui:Add, Text, x%X_Pos% y%Y_Pos% w110 h18 +0x200, % Format("{:02d}.[{}]", A_Index, EquipList[A_Index])

        DefaultTime := G_SavedTime_%A_Index%
        Gui, MorningGui:Font, s9 Normal cBlack, 맑은 고딕
        Gui, MorningGui:Add, Edit, % "vVar_MorningStart_" A_Index " x" X_Pos " y" Y_Pos+18 " w105 h22 Limit4 Number Center", %DefaultTime%
    }

    Gui, MorningGui:Font, s10 Bold c1E3A8A, 맑은 고딕
    Gui, MorningGui:Add, Button, x15 y355 w620 h40 gStartMorningTask, 🚀 [평일오전 입고] 매크로 실행하기

    Gui, MorningGui:Show, w650 h408
return

StartMorningTask:
    Gui, MorningGui:Submit, NoHide
    EquipList := GetCurrentEquipList(), MaxCount := EquipList.Length()
    AllYList := [255, 277, 300, 320, 340, 360, 380, 400, 420, 445, 465, 485, 505, 530, 550, 570, 590, 610, 635, 655, 675, 695, 715, 740, 760, 780, 800, 820, 845]

    MorningYList := [], MorningStartList := [], MorningEndList := []
    ErrCnt := 0

    Loop, %MaxCount%
    {
        Val := Trim(Var_MorningStart_%A_Index%)
        if (Val = "")
            continue
        if (StrLen(Val) != 4)
        {
            ErrCnt++
            break
        }

        MorningYList.Push(AllYList[A_Index])
        MorningStartList.Push(Val)
        MorningEndList.Push(Add10Minutes(Val))

        G_SavedTime_%A_Index% := Val
        IniWrite, %Val%, %IniFile%, Times, Morning_%A_Index%
    }

    if (ErrCnt > 0 || MorningStartList.Length() = 0)
    {
        MsgBox, 48, 입력 오류, ❌ 입고시각은 4자리 숫자로 정확히 1개 이상 입력해 주세요! (예: 0830)
        return
    }

    Gui, MorningGui:Destroy
    Gui, 1:Hide
    ExecuteMacroEngine("평일오전 입고 (" . MorningStartList.Length() . "대)", MorningYList, MorningStartList, MorningEndList, "020")
return

Add10Minutes(HHMM)
{
    HHMM := Trim(HHMM)
    if (StrLen(HHMM) != 4)
        return HHMM

    H := SubStr(HHMM, 1, 2) + 0, M := SubStr(HHMM, 3, 2) + 10
    if (M >= 60)
    {
        M -= 60, H += 1
        if (H >= 24)
            H -= 24
    }
    return Format("{:02d}{:02d}", H, M)
}

GetCurrentEquipList()
{
    global
    List := []
    Loop, 29
    {
        Val := Trim(Var_EquipBox_%A_Index%)
        if (Val = "" && G_SavedEquip_%A_Index% != "")
            Val := G_SavedEquip_%A_Index%
        if (Val != "" && StrLen(Val) = 3)
            List.Push(Val)
    }
    return (List.Length() > 0) ? List : [202, 232, 230, 212, 241, 217, 205, 242, 213, 239, 226, 222, 216, 211, 229, 237]
}

; =================================================================
; 🌟 [서브 GUI - 엑셀 차량 편성 창 (작업요청자 드롭다운 이름 완벽 표시)]
; =================================================================
OpenExcelMenu:
    if (!ValidateStaffInputs())
    {
        SoundBeep, 750, 300
        MsgBox, 48, ⛔ 담당자 사번 누락/오류!, ⚠️ 검사자, 확인자, 담당자 사번 3가지가 모두 [IT + 숫자 6자리] 규격이어야 합니다! (예: IT203094)
        return
    }

    SaveStaffSettings()

    Gui, ExcelGui:New, +Owner1, 📊 검사계획요청 차량 편성 입력
    Gui, ExcelGui:Color, F0F4F8, FFFFFF

    Gui, ExcelGui:Font, s12 Bold c004D40, 맑은 고딕
    Gui, ExcelGui:Add, Text, x20 y12 w610 Center, 📊 검사계획요청 차량 편성 입력

    Gui, ExcelGui:Font, s10 Bold c1E3A8A, 맑은 고딕
    Gui, ExcelGui:Add, Text, x20 y42 w95 h22, 📌 패턴 선택:

    PatternChoiceIndex := (G_SavedPatternNum >= 1 && G_SavedPatternNum <= 10) ? G_SavedPatternNum : 1
    Gui, ExcelGui:Font, s9 Normal cBlack, 맑은 고딕
    Gui, ExcelGui:Add, DropDownList, vVar_ExcelPattern gOnExcelPatternChange x115 y39 w280 Choose%PatternChoiceIndex%, 1. 평일저녁 입고 1차 (16대)|2. 평일저녁 입고 2차 서구청 포함 (7대)|3. 주말저녁 입고 1차 (4대)|4. 주말저녁 입고 2차 서구청 포함 (7대)|5. 평일오전 입고(8대~10대)|6. 평일새벽 출고 (29대)|7. 평일오후 출고 (9대)|8. 주말새벽 출고 서구청 포함 (12대)|9. 일상검사 주간 (4대)|10. 일상검사 야간 (2대)

    Gui, ExcelGui:Font, s10 Bold cB71C1C, 맑은 고딕
    Gui, ExcelGui:Add, Text, x405 y42 w95 h22, ★ 작업요청자:
    Gui, ExcelGui:Font, s9 Normal cBlack, 맑은 고딕
    ReqVal := (G_SavedRequester != "") ? G_SavedRequester : "IT203094"
    StringUpper, ReqVal, ReqVal
    DispReqVal := GetStaffDisplayWithID(ReqVal)
    Gui, ExcelGui:Add, ComboBox, vVar_Sub_작업요청자 x500 y39 w135 h150, %G_RecentStaffPipe%
    GuiControl, ExcelGui:Text, Var_Sub_작업요청자, %DispReqVal%

    Gui, ExcelGui:Font, s9 Bold c004D40, 맑은 고딕
    Gui, ExcelGui:Add, GroupBox, x15 y68 w620 h265, 🚗 개별 차량 번호 입력 (Tab 키로 다음편성 입력)

    Loop, 29
    {
        Col := Floor((A_Index - 1) / 6), Row := Mod(A_Index - 1, 6)
        X_Pos := 30 + (Col * 120), Y_Pos := 92 + (Row * 35)

        Gui, ExcelGui:Font, s9 Bold c1E3A8A, 맑은 고딕
        Gui, ExcelGui:Add, Text, x%X_Pos% y%Y_Pos% w28 h22 +0x200, % Format("{:02d}:", A_Index)

        Gui, ExcelGui:Font, s9 Normal cBlack, 맑은 고딕
        Gui, ExcelGui:Add, Edit, % "vVar_EquipBox_" A_Index " gOnGridTextChange x" X_Pos+28 " y" Y_Pos " w70 h22 Limit3 Number Center", % G_SavedEquip_%A_Index%
    }

    Gui, ExcelGui:Font, s10 Bold cGreen, 맑은 고딕
    Gui, ExcelGui:Add, Text, vVar_CountStatus x20 y304 w610 h22 Center, ✔ 정상: 총 16대

    Gui, ExcelGui:Font, s10 Bold c004D40, 맑은 고딕
    Gui, ExcelGui:Add, Button, x15 y345 w620 h40 gSaveExcelData, 💾 엑셀 파일 생성 및 저장하기

    if (G_SavedEquip_1 = "")
        PopulateGridBoxes(PatternChoiceIndex)
    else
        UpdateGridCountStatus()

    Gui, ExcelGui:Show, w650 h400
return

#IfWinActive 📊 검사계획요청 차량 편성 입력
Enter::
NumpadEnter::
    ControlGetFocus, FocusedControl, A
    if (RegExMatch(FocusedControl, "i)Edit(\d+)", Match) && Match1 >= 2 && Match1 <= 29)
    {
        NextEditNum := Match1 + 1
        GuiControl, ExcelGui:Focus, Edit%NextEditNum%
    }
return
#IfWinActive

OnExcelPatternChange:
    GuiControlGet, SelectedPattern, ExcelGui:, Var_ExcelPattern
    RegExMatch(SelectedPattern, "^(\d+)", Match)
    PopulateGridBoxes(Match1 + 0)
return

PopulateGridBoxes(PatternNum)
{
    GetPatternData(PatternNum, TaskTitle, 차종List, 시작List, 종료List, SheetCode, 설비IDList)
    Loop, 29
    {
        Val := (A_Index <= 설비IDList.Length()) ? 설비IDList[A_Index] : ""
        GuiControl, ExcelGui:, Var_EquipBox_%A_Index%, %Val%
    }
    UpdateGridCountStatus()
}

OnGridTextChange:
    UpdateGridCountStatus()
return

UpdateGridCountStatus()
{
    GuiControlGet, SelectedPattern, ExcelGui:, Var_ExcelPattern
    RegExMatch(SelectedPattern, "^(\d+)", Match)
    PatternNum := Match1 + 0

    GetPatternData(PatternNum, TaskTitle, 차종List, 시작List, 종료List, SheetCode, DefaultEquipList)
    ExpectedCount := 차종List.Length()

    ValidCnt := 0, InvalidCnt := 0, ErrSample := ""
    Loop, 29
    {
        GuiControlGet, Val, ExcelGui:, Var_EquipBox_%A_Index%
        Val := Trim(Val)
        if (Val = "")
            continue

        Num := Val + 0
        if (StrLen(Val) = 3 && Num >= 201 && Num <= 243)
            ValidCnt++
        else
        {
            InvalidCnt++
            if (ErrSample = "")
                ErrSample := Val
        }
    }

    if (InvalidCnt > 0)
    {
        GuiControl, ExcelGui:+cRed, Var_CountStatus
        GuiControl, ExcelGui:, Var_CountStatus, % "❌ 잘못된 번호(" . ErrSample . ") 포함!"
    }
    else if (PatternNum = 5 && ValidCnt >= 8 && ValidCnt <= 10)
        || (PatternNum = 9 && ValidCnt = 4)
        || (PatternNum = 10 && ValidCnt = 2)
        || (PatternNum != 5 && PatternNum != 9 && PatternNum != 10 && ValidCnt = ExpectedCount)
    {
        GuiControl, ExcelGui:+cGreen, Var_CountStatus
        GuiControl, ExcelGui:, Var_CountStatus, % "✔ 정상: 총 " . ValidCnt . "대"
    }
    else
    {
        GuiControl, ExcelGui:+cOrange, Var_CountStatus
        GuiControl, ExcelGui:, Var_CountStatus, % "⚠️ " . ValidCnt . "대 입력됨 (기준 수량과 다름)"
    }
}

; =================================================================
; 🌟 [엑셀 COM 자동 작성 엔진 - 현황판 툴박스 미표시 최적화]
; =================================================================
SaveExcelData:
    Gui, ExcelGui:Submit, NoHide

    PureReq := ExtractPureStaffID(Var_Sub_작업요청자)
    StringUpper, PureReq, PureReq
    P_작업요청자 := PureReq

    RegExMatch(Var_ExcelPattern, "^(\d+)", Match)
    PatternNum := Match1 + 0
    GetPatternData(PatternNum, TaskTitle, 차종List, 시작List, 종료List, SheetCode, DefaultEquipList)

    ValidList := [], InvalidList := []
    Loop, 29
    {
        Val := Trim(Var_EquipBox_%A_Index%)
        if (Val = "")
            continue
        Num := Val + 0
        if (StrLen(Val) = 3 && Num >= 201 && Num <= 243)
            ValidList.Push(Val)
        else
            InvalidList.Push(Val)
    }

    if (InvalidList.Length() > 0 || ValidList.Length() = 0)
    {
        MsgBox, 48, 차량 번호 오류, ❌ 입력된 차량 번호가 없거나 잘못된 번호가 포함되어 있습니다!
        return
    }

    Count := ValidList.Length()
    IsPattern5Ok  := (PatternNum = 5 && Count >= 8 && Count <= 10)
    IsPattern9Ok  := (PatternNum = 9 && Count = 4)
    IsPattern10Ok := (PatternNum = 10 && Count = 2)

    if (Count != 차종List.Length() && !IsPattern5Ok && !IsPattern9Ok && !IsPattern10Ok)
    {
        MsgBox, 36, 차량 대수 변경 확인, % "⚠️ 입력된 대수(" Count "대)가 기본 패턴 기준과 다릅니다. 이대로 저장하시겠습니까?"
        IfMsgBox, No
            return
    }

    IniWrite, %PatternNum%, %IniFile%, Main, LastPatternNum
    IniWrite, %P_작업요청자%, %IniFile%, Main, Requester
    Loop, 29
    {
        EqVal := (A_Index <= ValidList.Length()) ? ValidList[A_Index] : ""
        G_SavedEquip_%A_Index% := EqVal
        IniWrite, %EqVal%, %IniFile%, Vehicles, Equip_%A_Index%
    }
    G_SavedPatternNum := PatternNum, G_SavedRequester := P_작업요청자
    Gui, ExcelGui:Destroy

    TargetExcel := A_ScriptDir "\검사계획요청(매크로용).xlsx"
    FormatTime, TodayDate,, yyyyMMdd
    TypeStr := (SheetCode = "020") ? "DYARR" : (SheetCode = "021") ? "DYDEP" : "CYDAY"
    TypeName := (SheetCode = "020") ? "입고검사" : (SheetCode = "021") ? "출고검사" : "일상검사"

    try
    {
        oExcel := ComObjCreate("Excel.Application")
        oExcel.Visible := False, oExcel.DisplayAlerts := False

        oWorkbook := FileExist(TargetExcel) ? oExcel.Workbooks.Open(TargetExcel) : oExcel.Workbooks.Add()
        try
            oSheet := oWorkbook.Sheets("작업계획요청")
        catch
        {
            oSheet := oWorkbook.Sheets(1)
            oSheet.Name := "작업계획요청"
        }

        Headers := ["설비ID", "검종", "검사계획시작일자", "검사계획종료일자", "작업요청자", "요청사유"]
        For idx, name in Headers
            oSheet.Cells(1, idx).Value := name

        oSheet.Cells(1, 8).Value := "검종"
        oSheet.Cells(1, 9).Value := "작업요청자"
        oSheet.Range("A2:F100").ClearContents()

        oSheet.Range("H2").Value := TypeName
        oSheet.Range("I2").Value := P_작업요청자

        Loop, %Count%
        {
            Row := A_Index + 1
            oSheet.Cells(Row, 1).Value := ValidList[A_Index]
            oSheet.Cells(Row, 2).Value := TypeStr
            oSheet.Cells(Row, 3).Value := TodayDate
            oSheet.Cells(Row, 4).Value := TodayDate
            oSheet.Cells(Row, 5).Value := P_작업요청자
        }

        if (FileExist(TargetExcel))
            oWorkbook.Save()
        else
            oWorkbook.SaveAs(TargetExcel)

        try oWorkbook.Close(False)
        try oExcel.Quit()

        G_IsExcelSavedToday := True
        UpdateExcelStatusAndButtons()

        MsgBox, 64, 엑셀 작성 성공, [%TaskTitle%]`n`n총 %Count%개 차량 데이터 저장 완료!`n· 구분: %TypeName% (%TypeStr%)`n· 파일: %TargetExcel%
    }
    catch err
    {
        if (IsObject(oExcel))
        {
            try oWorkbook.Close(False)
            try oExcel.Quit()
        }

        G_IsExcelSavedToday := True
        UpdateExcelStatusAndButtons()
        MsgBox, 64, 엑셀 작성 성공, [%TaskTitle%]`n`n총 %Count%개 차량 데이터 저장 완료!`n· 구분: %TypeName% (%TypeStr%)`n· 파일: %TargetExcel%
    }
return

; =================================================================
; 🌟 [설정 자동 저장 & 로드 함수]
; =================================================================
LoadSettings()
{
    global
    IniRead, LastEnv, %IniFile%, Config, EnvMode, GW
    G_EnvMode := LastEnv
    G_YOffset := (G_EnvMode = "APP") ? -10 : 0

    IniRead, PNum, %IniFile%, Main, LastPatternNum, 1
    IniRead, Req, %IniFile%, Main, Requester, IT203094
    StringUpper, Req, Req
    G_SavedPatternNum := PNum + 0, G_SavedRequester := Req

    IniRead, Insp, %IniFile%, MainStaff, Inspector, IT203094
    IniRead, Chk,  %IniFile%, MainStaff, Checker,   IT201751
    IniRead, Mngr, %IniFile%, MainStaff, Manager,   IT201495
    StringUpper, Insp, Insp
    StringUpper, Chk, Chk
    StringUpper, Mngr, Mngr

    G_SavedInspector := Insp, G_SavedChecker := Chk, G_SavedManager := Mngr
    P_검사자 := Insp, P_확인자 := Chk, P_담당자 := Mngr

    G_RecentStaffPipe := GetStaffDisplayList()

    Loop, 29
    {
        IniRead, EqVal, %IniFile%, Vehicles, Equip_%A_Index%, %A_Space%
        G_SavedEquip_%A_Index% := (EqVal != "ERROR" && EqVal != " ") ? EqVal : ""

        IniRead, TmVal, %IniFile%, Times, Morning_%A_Index%, %A_Space%
        G_SavedTime_%A_Index% := (TmVal != "ERROR" && TmVal != " ") ? TmVal : ""
    }

    Loop, 4
    {
        IniRead, SVal, %IniFile%, DailyStaff, % "Staff_" A_Index, IT203094
        StringUpper, SVal, SVal
        G_SavedStaff_%A_Index% := SVal
    }
}

GetStaffDisplayList()
{
    global IniFile
    IniRead, SectionText, %IniFile%, StaffList

    PipeStr := ""
    if (SectionText != "" && SectionText != "ERROR")
    {
        Loop, Parse, SectionText, `n, `r
        {
            Line := Trim(A_LoopField)
            if (Line = "" || SubStr(Line, 1, 1) = ";")
                continue

            Parts := StrSplit(Line, "=")
            if (Parts.Length() >= 2)
            {
                ID := Trim(Parts[1])
                Name := Trim(Parts[2])
                Item := (Name != "") ? ID . " (" . Name . ")" : ID
                PipeStr .= (PipeStr = "" ? "" : "|") . Item
            }
        }
    }

    return (PipeStr != "") ? PipeStr : "IT203094|IT201751|IT201495"
}

ExtractPureStaffID(RawStr)
{
    RawStr := Trim(RawStr)
    if RegExMatch(RawStr, "i)(IT\d{6})", Match)
        return Match1
    return RawStr
}

GetStaffDisplayWithID(StaffID)
{
    global IniFile
    StaffID := Trim(StaffID)
    if (StaffID = "")
        return ""

    IniRead, StaffName, %IniFile%, StaffList, %StaffID%, %A_Space%
    StaffName := Trim(StaffName)

    if (StaffName != "" && StaffName != "ERROR")
        return StaffID . " (" . StaffName . ")"
    return StaffID
}

ExecuteTask(PatternNum)
{
    global
    SaveStaffSettings()
    if (!CheckExcelFileFirst())
        return

    GetPatternData(PatternNum, TaskTitle, 차종List, 시작List, 종료List, SheetCode, 설비IDList)
    EquipList := GetCurrentEquipList(), ActualCount := EquipList.Length()

    if (ActualCount > 0 && ActualCount < 차종List.Length())
    {
        NewYList := [], NewStartList := [], NewEndList := []
        Loop, %ActualCount%
        {
            NewYList.Push(차종List[A_Index])
            NewStartList.Push(시작List[A_Index])
            NewEndList.Push(종료List[A_Index])
        }
        차종List := NewYList, 시작List := NewStartList, 종료List := NewEndList
        TaskTitle := SubStr(TaskTitle, 1, InStr(TaskTitle, "(") - 1) . "(" . ActualCount . "대)"
    }

    Gui, 1:Hide
    ExecuteMacroEngine(TaskTitle, 차종List, 시작List, 종료List, SheetCode)
}

GetValidExcelPath()
{
    Paths := [A_ScriptDir "\검사계획요청(매크로용).xlsx", A_ScriptDir "\검사계획요청 (mis 입출고 입력용).xlsx", A_ScriptDir "\검사계획요청.xlsx"]
    For _, p in Paths
        if (FileExist(p))
            return p
    return A_ScriptDir "\검사계획요청(매크로용).xlsx"
}

; =================================================================
; [10가지 정밀 패턴 데이터 정의]
; =================================================================
GetPatternData(PatternNum, ByRef Title, ByRef YList, ByRef StartList, ByRef EndList, ByRef Code, ByRef EquipList)
{
    Code := (PatternNum >= 1 && PatternNum <= 5) ? "020" : (PatternNum >= 6 && PatternNum <= 8) ? "021" : "018"

    if (PatternNum = 1) {
        Title := "평일저녁 입고 1차 (16대)"
        YList := [255, 277, 300, 320, 340, 360, 380, 400, 420, 445, 465, 485, 505, 530, 550, 570]
        StartList := ["1929", "1939", "1946", "1953", "1959", "2009", "2016", "2023", "2030", "2037", "2155", "2226", "2244", "2257", "2316", "2328"]
        EndList   := ["1939", "1949", "1956", "2003", "2009", "2019", "2026", "2033", "2040", "2047", "2205", "2236", "2254", "2307", "2326", "2338"]
        EquipList := [202, 232, 230, 212, 241, 217, 205, 242, 213, 239, 226, 222, 216, 211, 229, 237]
    } else if (PatternNum = 2) {
        Title := "평일저녁 입고 2차 서구청 포함 (7대)"
        YList := [255, 277, 300, 320, 340, 360, 380]
        StartList := ["0016", "0025", "0036", "0046", "0055", "0105", "0115"]
        EndList   := ["0026", "0035", "0046", "0056", "0105", "0115", "0125"]
        EquipList := [202, 232, 230, 212, 241, 217, 205]
    } else if (PatternNum = 3) {
        Title := "주말저녁 입고 1차 (4대)"
        YList := [255, 277, 300, 320], StartList := ["2229", "2240", "2258", "2316"], EndList := ["2239", "2250", "2308", "2326"], EquipList := [202, 232, 230, 212]
    } else if (PatternNum = 4) {
        Title := "주말저녁 입고 2차 서구청 포함 (7대)"
        YList := [255, 277, 300, 320, 340, 360, 380], StartList := ["0017", "0027", "0038", "0047", "0057", "0105", "0115"], EndList := ["0027", "0037", "0048", "0057", "0107", "0115", "0125"], EquipList := [202, 232, 230, 212, 241, 217, 205]
    } else if (PatternNum = 5) {
        Title := "평일오전 입고"
        YList := [255, 277, 300, 320, 340, 360, 380, 400, 420, 445], StartList := ["0900", "0910", "0920", "0930", "0940", "0950", "1000", "1010", "1020", "1030"], EndList := ["0910", "0920", "0930", "0940", "0950", "1000", "1010", "1020", "1030", "1040"], EquipList := [202, 232, 230, 212, 241, 217, 205, 242, 213, 239]
    } else if (PatternNum = 6) {
        Title := "평일새벽 출고 (29대)"
        YList := [255, 277, 300, 320, 340, 360, 380, 400, 420, 445, 465, 485, 505, 530, 550, 570, 590, 610, 635, 655, 675, 695, 715, 740, 760, 780, 800, 820, 845]
        StartList := ["0415", "0420", "0425", "0430", "0435", "0440", "0445", "0450", "0455", "0515", "0530", "0543", "0549", "0558", "0611", "0614", "0617", "0623", "0626", "0631", "0633", "0638", "0641", "0647", "0656", "0702", "0708", "0714", "0720"]
        EndList   := ["0420", "0425", "0430", "0435", "0440", "0445", "0450", "0455", "0500", "0520", "0535", "0548", "0554", "0603", "0616", "0619", "0622", "0628", "0631", "0636", "0638", "0643", "0646", "0652", "0701", "0707", "0713", "0719", "0725"]
        EquipList := [202, 232, 230, 212, 241, 217, 205, 242, 213, 239, 226, 222, 216, 211, 229, 237, 201, 203, 204, 206, 207, 208, 209, 210, 214, 215, 218, 219, 220]
    } else if (PatternNum = 7) {
        Title := "평일오후 출고 (9대)"
        YList := [255, 277, 300, 320, 340, 360, 380, 400, 420], StartList := ["1542", "1600", "1619", "1700", "1710", "1723", "1737", "1747", "1801"], EndList := ["1547", "1605", "1624", "1705", "1715", "1728", "1742", "1752", "1806"], EquipList := [202, 232, 230, 212, 241, 217, 205, 242, 213]
    } else if (PatternNum = 8) {
        Title := "주말새벽 출고 서구청 포함 (12대)"
        YList := [255, 277, 300, 320, 400, 420, 445, 465, 485, 505, 530, 550], StartList := ["0412", "0422", "0427", "0432", "0522", "0530", "0547", "0603", "0615", "0626", "0644", "0417"], EndList := ["0417", "0427", "0432", "0437", "0527", "0535", "0553", "0608", "0620", "0631", "0649", "0422"], EquipList := [202, 232, 230, 212, 241, 217, 205, 242, 213, 239, 226, 222]
    } else if (PatternNum = 9) {
        Title := "일상검사 주간 (4대)"
        YList := [255, 277, 300, 320], StartList := ["0930", "1030", "1330", "1430"], EndList := ["1030", "1130", "1430", "1530"], EquipList := [202, 232, 230, 212]
    } else if (PatternNum = 10) {
        Title := "일상검사 야간 (2대)"
        YList := [255, 277], StartList := ["1900", "2000"], EndList := ["2000", "2100"], EquipList := [202, 232]
    }
}

; =================================================================
; 🌟 [다크모드 오버레이 현황판 GUI 생성 및 갱신]
; =================================================================
UpdateDashboard(StepText := "")
{
    global CurrentStepText, CurrentVehicleIndex, TotalCount, P_검사자, P_확인자, P_담당자, G_EnvMode
    global Var_DashProg, Var_DashStep, Var_DashStaff
    static GuiCreated := False

    if (StepText != "")
    {
        CurrentStepText := StepText
        ; 🌟 [v4.5] 진행 단계 로그 (앱이 꺼진 시점 추적용)
        FormatTime, LogTime,, yyyy-MM-dd HH:mm:ss
        FileAppend, % LogTime . "  [" . CurrentVehicleIndex . "/" . TotalCount . "]  " . StepText . "`n", %A_ScriptDir%\macro_log.txt, UTF-8
    }

    Ratio := (TotalCount > 0) ? CurrentVehicleIndex / TotalCount : 0
    Filled := Round(Ratio * 10)

    BarStr := ""
    Loop, 10
        BarStr .= (A_Index <= Filled) ? "■" : "□"

    PctVal := Round(Ratio * 100)
    EnvTitle := (G_EnvMode = "APP") ? "💻 MIS앱 모드" : "🌐 그룹웨어 모드"

    ProgText := "📊 진행률: [" . BarStr . "] " . PctVal . "% (" . CurrentVehicleIndex . "/" . TotalCount . "대)"
    StepFullText := "📌 세부단계: " . CurrentStepText
    StaffText := "👨‍✈️ 검사자: " . P_검사자 . " | 확인자: " . P_확인자 . " | 담당자: " . P_담당자

    if (!GuiCreated)
    {
        Gui, DashboardGui:Destroy
        Gui, DashboardGui:New, +AlwaysOnTop -Caption +ToolWindow +E0x20, MIS 현황판
        Gui, DashboardGui:Color, 1E293B ; 다크 슬레이트 배경

        Gui, DashboardGui:Font, s10 Bold c38BDF8, 맑은 고딕
        Gui, DashboardGui:Add, Text, x12 y8 w400 h20, % "🚆 MIS 입력 자동화 현황판 🚩 [" . EnvTitle . "]"

        Gui, DashboardGui:Font, s9 Bold cFACC15, 맑은 고딕
        Gui, DashboardGui:Add, Text, vVar_DashProg x12 y30 w400 h20, %ProgText%

        Gui, DashboardGui:Font, s9 Bold cFFFFFF, 맑은 고딕
        Gui, DashboardGui:Add, Text, vVar_DashStep x12 y52 w400 h20, %StepFullText%

        Gui, DashboardGui:Font, s8 Normal c34D399, 맑은 고딕
        Gui, DashboardGui:Add, Text, vVar_DashStaff x12 y73 w400 h18, %StaffText%

        Gui, DashboardGui:Font, s8 Normal c94A3B8, 맑은 고딕
        Gui, DashboardGui:Add, Text, x12 y92 w400 h16, ⌨️ 일시정지: Shift+Enter  |  강제종료: Esc

        Gui, DashboardGui:Show, x1180 y140 NoActivate w420 h115, MIS 현황판
        GuiCreated := True
    }
    else
    {
        GuiControl, DashboardGui:Text, Var_DashProg, %ProgText%
        GuiControl, DashboardGui:Text, Var_DashStep, %StepFullText%
        GuiControl, DashboardGui:Text, Var_DashStaff, %StaffText%
    }
}

DestroyDashboard()
{
    Gui, DashboardGui:Destroy
}

; =================================================================
; 🌟 [v4.4 신규] MIS 앱 로딩(빙글빙글 커서 / 응답 없음) 완료 대기
;   - 로딩 중에 클릭/Enter가 들어가면 Windows가 "응답 없음 → 프로그램 닫기"
;     창을 띄우고, 매크로의 Enter가 그 창을 눌러 앱이 꺼지는 현상 방지
;   - 그룹웨어(웹) 모드에서는 아무것도 하지 않고 바로 통과
; =================================================================
WaitAppReady(MaxWaitSec := 90)
{
    global G_EnvMode, G_MisHwnd, G_MisPid, CurrentStepText
    if (G_EnvMode != "APP" || !G_MisHwnd)
        return True

    Start := A_TickCount, StableCnt := 0, Notified := False
    Loop
    {
        ; MIS 앱 창이 사라졌으면 즉시 매크로 중단 (빈 화면에 계속 입력하는 것 방지)
        if (!WinExist("ahk_id " G_MisHwnd))
        {
            ; 🌟 [v4.5] 창 핸들이 바뀐 것뿐인지(화면 전환 등) 프로세스로 재확인
            Process, Exist, %G_MisPid%
            NewHwnd := ErrorLevel ? WinExist("ahk_pid " G_MisPid) : 0
            if (!NewHwnd)
                AbortMacro("MIS 앱이 종료되었습니다. (단계: " . CurrentStepText . ")")
            G_MisHwnd := NewHwnd
        }

        Busy := False
        if DllCall("IsHungAppWindow", "Ptr", G_MisHwnd)       ; 응답 없음 상태
            Busy := True
        if (A_Cursor = "Wait" || A_Cursor = "AppStarting")    ; 빙글빙글 커서
            Busy := True
        if (IsHangDialogActive())                             ; "응답하지 않습니다" 창
            Busy := True

        if (Busy)
        {
            StableCnt := 0
            if (!Notified && (A_TickCount - Start) > 1500)
            {
                UpdateDashboard("⏳ MIS 앱 로딩 대기 중... (입력 일시 보류)")
                Notified := True
            }
        }
        else if (++StableCnt >= 5)                            ; 0.5초 연속 정상 → 준비 완료
            return True

        if ((A_TickCount - Start) > MaxWaitSec * 1000)
        {
            SoundBeep, 750, 400
            MsgBox, 48, MIS 앱 로딩 지연, % "MIS 앱이 " . MaxWaitSec . "초 넘게 로딩 중입니다.`n`n앱이 정상으로 돌아온 것을 확인한 뒤 [확인]을 누르면 매크로가 계속됩니다.`n(⚠️ '프로그램 닫기'는 절대 누르지 마세요!)"
            Start := A_TickCount, Notified := False
        }
        Sleep, 100
    }
}

; Windows "프로그램이 응답하지 않습니다" 창이 떠 있는지 확인
IsHangDialogActive()
{
    WinGet, ActProc, ProcessName, A
    if (ActProc = "WerFault.exe" || ActProc = "dwm.exe")
        return True
    WinGetTitle, ActTitle, A
    if InStr(ActTitle, "응답 없음") || InStr(ActTitle, "Not Responding")
        return True
    WinGetText, ActText, A
    if InStr(ActText, "응답하지 않") || InStr(ActText, "프로그램 닫기")
        return True
    return False
}

; 로딩이 끝난 것을 확인한 뒤에만 키 전송
SafeSend(Keys)
{
    WaitAppReady()
    Send, %Keys%
}

; 로딩이 끝난 것을 확인한 뒤에만 텍스트 입력
SafeSendRaw(Text)
{
    WaitAppReady()
    SendRaw, %Text%
}

; 로딩이 끝난 것을 확인한 뒤에만 좌표 클릭 (클릭 후 로딩도 끝까지 대기)
SafeClick(X, Y, ClickCount := 1)
{
    WaitAppReady()
    MouseClick, left, %X%, %Y%, %ClickCount%
    Sleep, 300
    WaitAppReady()
}

AbortMacro(Reason)
{
    DestroyDashboard()
    SoundBeep, 750, 500
    MsgBox, 16, 매크로 중단, % "🛑 " . Reason . "`n`n매크로를 중단합니다. MIS 앱을 다시 실행한 뒤 처음부터 진행해 주세요."
    Gui, 1:Show
    Exit
}

; =================================================================
; 🌟 [공용 화면 입력 자동화 엔진 - 병목 구간 전/후 대기 완벽 보강]
; =================================================================
ExecuteMacroEngine(TaskTitle, 차종List, 시작List, 종료List, SheetCode)
{
    global
    FormatTime, time,, yyyyMMdd

    MsgBox, 64, 업무 자동화 시작, % "[" . TaskTitle . "] 매크로를 실행합니다.`n`n📌 인천교통공사 MIS 창을 띄워놓아 주세요.`n(접속 환경: " . (G_EnvMode="APP" ? "MIS앱 (최적화 모드)" : "그룹웨어 웹") . ")`n`n(비상 강제종료: Esc / 일시정지: Shift + Enter)"
    Sleep, 1000

    UpdateDashboard("📌 [사전 1/15] 통합경영정보시스템 (MIS) 창 감지 및 전체화면 조절 중")
    G_MisHwnd := 0
    WinWait, 인천교통공사 통합경영정보시스템,, 15
    if (ErrorLevel = 0)
    {
        G_MisHwnd := WinExist("인천교통공사 통합경영정보시스템")
        WinGet, G_MisPid, PID, ahk_id %G_MisHwnd%
        WinMaximize, ahk_id %G_MisHwnd%
        WinActivate, ahk_id %G_MisHwnd%
    }
    Sleep, 1500
    WaitAppReady()

    UpdateDashboard("📌 [사전 2/15] 좌측 메뉴 [검사(&N)] 클릭 중")
    ClickImage("검사_N.png", 15, 1000)

    UpdateDashboard("📌 [사전 3/15] [2호선]검사계획관리 더블클릭 중")
    DoubleClickImage("2호선검사계획관리.png", 15, 1000)

    UpdateDashboard("📌 [사전 4/15] 세부메뉴 [2호선]검사계획관리* 더블클릭 중")
    DoubleClickImage("2호선검사계획관리_세부.png", 15, 1000)

    UpdateDashboard("📌 [사전 5/15] 상단 [작업요청] ➔ [Excel] 버튼 클릭 중")
    ClickImage("작업요청.png", 15, 1000)
    ClickImage("Excel.png", 15, 1000)

    UpdateDashboard("📌 [사전 6/15] Select File 창에서 엑셀 파일 첨부 중")
    WinWait, Select File,, 15
    if (ErrorLevel = 0)
    {
        WinActivate, Select File
        Sleep, 500
        SendRaw, % GetValidExcelPath()
        Sleep, 500
        Send, {Enter}
        Sleep, 1500
    }

    UpdateDashboard("📌 [사전 7/15] 업로드 완료 알림 확인 클릭 중")
    SafeSend("{Enter}")
    Sleep, 1000

    ; 🌟 [핵심 개선] 엑셀 대량 데이터 화면 표(그리드) 안착 대기 (3.5초)
    UpdateDashboard("📌 [사전 7.5/15] 엑셀 데이터 그리드 안착 대기 중 (3.5초)...")
    Sleep, 3500

    UpdateDashboard("📌 [사전 8/15] 우측 툴바 [저장] 버튼 클릭 중")
    ClickImage("저장.png", 15, 1000)
    Sleep, 800
    SafeSend("{Enter}")
    Sleep, 2000

    UpdateDashboard("📌 [사전 9/15] 연속 중복 팝업 감지 및 전체 확인 처리 중")
    HandleAllRepeatedPopups()

    ; 엑셀 대량 저장 후 DB 커밋 및 내부 메모리 정리 완전 안정화 대기 (3.0초)
    UpdateDashboard("📌 [사전 9.5/15] 데이터 저장 후처리 안정화 대기 중 (3초)...")
    Sleep, 3000

    UpdateDashboard("📌 [사전 10/15] [종료] 버튼 ➔ [검종] 버튼 클릭 중")
    ClickImage("종료.png", 30, 2000)
    Sleep, 1500

    GumjongImg := (G_EnvMode = "APP") ? "검종(MIS앱 환경).png" : "검종.png"
    if (G_EnvMode = "APP" && !FileExist(ImageFolder "\" GumjongImg))
        GumjongImg := "검종.png"
    ClickImage(GumjongImg, 15, 1000)

    UpdateDashboard("📌 [사전 11/15] 패턴 맞춤 검종 방향키 선택 중")
    DownCount := (SheetCode = "020") ? 2 : (SheetCode = "021") ? 3 : 4
    SafeSend("{Down " DownCount "}")
    Sleep, 400
    SafeSend("{Enter}")
    Sleep, 1500

    UpdateDashboard("📌 [사전 12/15] [조회] 버튼 클릭 중")
    ClickImage("조회.png", 15, 2000)

    UpdateDashboard("📌 [사전 13/15] [전체선택] 버튼 클릭 중")
    ClickImage("전체선택.png", 15, 1000)

    UpdateDashboard("📌 [사전 14/15] [계획확정] 버튼 클릭 중")
    ClickImage("계획확정.png", 15, 1000)

    UpdateDashboard("📌 [사전 15/15] 계획확정 승인 팝업 확인 클릭 중")
    Sleep, 800
    SafeSend("{Enter}")
    Sleep, 1500

    ; -------------------------------------------------------------
    ; [반복 자동화 단계] N개 전동차 데이터 연속 기입
    ; -------------------------------------------------------------
    FocusY := 277 + G_YOffset
    UpdateDashboard("📌 [준비] 목록 포커스 초기화 클릭 (Y: " . FocusY . ")")
    SafeClick(355, FocusY)
    Sleep, 600

    TotalCount := 차종List.Length()

    Loop, %TotalCount%
    {
        CurrentVehicleIndex := A_Index
        TargetY := 차종List[A_Index] + G_YOffset
        CurrentStart := 시작List[A_Index], CurrentEnd := 종료List[A_Index]

        UpdateDashboard("📌 [1/8] 차종 선택 진행 중 (Y: " . TargetY . ")")
        SafeClick(355, TargetY)
        Sleep, 600

        UpdateDashboard("📌 [2/8] 검사작업내역 이동 중")
        ClickImage("검사작업내역관리.png", 15, 2000)
        ClickImage("공정작업일반.png", 15, 1200)

        ; 🌟 [수정 반영] 공정작업일반 ➔ 입력 직전 1초 숨고르기
        Sleep, 1000
        ClickImage("입력.png", 20, 1500)
        Sleep, 1500

        UpdateDashboard("📌 [3/8] 시간/담당자 입력 중")
        SafeSend("{Tab}")
        Sleep, 400
        SafeSendRaw(CurrentStart)
        Sleep, 400
        SafeSendRaw(CurrentEnd)
        Sleep, 400
        SafeSendRaw(P_담당자)
        Sleep, 300
        SafeSend("{enter}")
        Sleep, 400
        SafeSend("100")

        ; 🌟 [핵심 개선] 진도율(100) 입력 후 테이블 내부 계산 및 사번 유효성 검증 완전 완료 대기 (4초)
        UpdateDashboard("📌 [3.5/8] 작업내역 수치 계산 및 데이터 안착 대기 중 (4초)...")
        Sleep, 4000

        ; [저장 단계]
        UpdateDashboard("📌 [4/8] 기본 작업내역 저장 중")
        ClickImage("저장.png", 15, 1000)
        SafeSend("{enter}")
        Sleep, 2000 ; 저장 후 2.0초 대기

        UpdateDashboard("📌 [5/8] 검사표 데이터 작성 중")
        ClickImage("검사표.png", 15, 1200)

        ; 검사표 탭 로딩 대기 1.5초
        Sleep, 1500
        ClickImage("입력.png", 20, 1500)
        Sleep, 1500

        SafeSend(SheetCode)
        Sleep, 400
        SafeSend("{tab 4}")
        Sleep, 400
        SafeSend(time)
        SafeSend("{tab}")
        Sleep, 400
        SafeSend(time)
        SafeSend("{tab}")
        Sleep, 400
        SafeSendRaw(P_검사자)
        SafeSend("{tab}")
        Sleep, 400
        SafeSendRaw(P_확인자)
        SafeSend("{tab}")
        Sleep, 400
        SafeSendRaw(P_담당자)
        Sleep, 300

        ; [저장 단계]
        UpdateDashboard("📌 [6/8] 검사표 입력 저장 중")
        Sleep, 500
        ClickImage("저장.png", 15, 1000)
        SafeSend("{enter}")

        ; 🌟 [수정 반영] 검사표 대량 항목 DB 저장 및 백엔드 트랜잭션 완료 대기 (10초 대기!)
        UpdateDashboard("📌 [6.5/8] 검사표 대량 데이터 DB 커밋 대기 중 (10초)...")
        Sleep, 10000

        UpdateDashboard("📌 [7/8] 검사표 항목 복사 및 수정 중")
        ClickImage("일반검사표.png", 15, 2000)

        ; [표준점검항목복사 직전 숨고르기]
        Sleep, 2500
        ClickImage("표준점검항목복사.png", 25, 1500)
        SafeSend("{Left}{enter}")

        ; 복사 승인 후 대량 그리드 행 로딩 대기 (2.5초)
        Sleep, 2500
        ClickImage("수정.png", 15, 2000)
        Sleep, 1000

        if (SheetCode = "018") ; ★ 일상검사 (018) 사번 입력 정밀 루프
        {
            UpdateDashboard("📌 [7/8-1] 일상검사 테이블 로딩 대기 중 (3초)...")
            Sleep, 3000

            DailyCellY := 300 + G_YOffset
            SafeClick(740, DailyCellY)
            Sleep, 500
            SafeClick(740, DailyCellY)
            Sleep, 500

            if (G_EnvMode = "APP")
            {
                DailySteps := [ [0, P_운전실기기], [5, P_운전실기기], [3, P_운전실기기], [6, P_운전실기기], [15, P_운전실기기]
                              , [4, P_객실기기], [5, P_객실기기], [5, P_객실기기]
                              , [19, P_대차및하부_옥상], [22, P_대차및하부_옥상], [27, P_대차및하부_옥상], [15, P_대차및하부_옥상]
                              , [7, P_입환_면허자] ]
            }
            else
            {
                DailySteps := [ [0, P_운전실기기], [6, P_운전실기기], [4, P_운전실기기], [7, P_운전실기기], [16, P_운전실기기]
                              , [6, P_객실기기], [6, P_객실기기], [6, P_객실기기]
                              , [21, P_대차및하부_옥상], [23, P_대차및하부_옥상], [28, P_대차및하부_옥상], [16, P_대차및하부_옥상]
                              , [8, P_입환_면허자] ]
            }

            For _, Step in DailySteps
            {
                if (Step[1] > 0)
                    SafeSend("{Down " Step[1] "}")
                Sleep, 300
                SafeSendRaw(Step[2])
                Sleep, 300
            }
        }
        else if (SheetCode = "020") ; 입고검사 (020)
        {
            SafeSend("{tab 3}")
            Sleep, 400
            SafeSendRaw(P_검사자)
            Sleep, 400
            SafeSend("{enter 5}")
            Sleep, 400
            SafeSend("{tab 3}")
            Sleep, 400
            SafeSendRaw(P_검사자)
            Sleep, 400
            SafeSend("{enter 7}")
            Sleep, 400
            SafeSend("{tab 6}")
            Sleep, 400
            SafeSendRaw(P_검사자)
            Sleep, 400
        }
        else ; 출고검사 (021)
        {
            SafeSend("{tab 3}")
            Sleep, 400
            SafeSendRaw(P_검사자)
            Sleep, 400
            SafeSend("{enter 6}")
            Sleep, 400
            SafeSend("{tab 3}")
            Sleep, 400
            SafeSendRaw(P_검사자)
            Sleep, 400
        }

        ; 🌟 [저장 단계]
        Sleep, 500
        ClickImage("저장.png", 15, 1500)
        SafeSend("{enter}")

        ; 🌟 [수정 반영] 검사표 대량 항목 최종 저장 후에도 10초 넉넉히 대기!
        UpdateDashboard("📌 [7/8-2] 검사표 데이터 DB 커밋 안정화 대기 중 (10초)...")
        Sleep, 10000

        ; 닫기 직전 숨고르기 대기
        Sleep, 1500

        ; 종료 버튼 탐색 타임아웃 60초로 대폭 상향
        UpdateDashboard("📌 [7/8-3] 검사표 닫기 [종료] 버튼 클릭 중...")
        ClickImage("종료.png", 60, 2000)

        ; 창이 완전히 닫히고 이전 메뉴로 복귀하는 안전 대기 2.0초
        Sleep, 2000

        UpdateDashboard("📌 [8/8] 승인요청 및 최종 마무리 중")
        ClickImage("검사표승인요청.png", 15, 1000)
        SafeSend("{enter}")
        Sleep, 1000

        ; 🌟 [수정 반영] 검사작업일반 클릭 후 다음 입력/조회 전 1초 숨고르기
        ClickImage("검사작업일반.png", 15, 1000)
        Sleep, 1000
        ClickImage("조회.png", 15, 2000)

        ; 검사작업완료 클릭 전 숨고르기
        Sleep, 1500
        ClickImage("검사작업완료.png", 20, 1500)

        ; 확인용 팝업창이 화면에 확실히 뜰 때까지 1.5초 대기한 후 승인 키 전송
        Sleep, 1500
        SafeSend("{Left}{enter 2}")

        ; 작업완료 DB 트랜잭션 반영 대기 2.0초
        Sleep, 2000

        ; 최종 종료 직전 숨고르기 대기 1.0초
        Sleep, 1000
        ClickImage("종료.png", 15, 2000)
        Sleep, 1500
    }

    DestroyDashboard()
    MsgBox, 64, 완료, [%TaskTitle%] 작업이 성공적으로 완료되었습니다!
    Gui, 1:Show
}

HandleAllRepeatedPopups()
{
    QuietCount := 0
    Loop
    {
        WaitAppReady()

        ; 🌟 [v4.4] "응답하지 않습니다 / 프로그램 닫기" 창에는 절대 Enter 금지
        if (IsHangDialogActive())
        {
            Sleep, 1000
            continue
        }

        if (WinActive("알림") || WinActive("확인") || WinActive("ahk_class #32770"))
        {
            Sleep, 300
            SafeSend("{Enter}")
            Sleep, 600
            QuietCount := 0
        }
        else if (++QuietCount >= 5)
            break
        Sleep, 300
    }
}

ClickImage(ImageName, MaxWaitSec := 15, PostSleep := 500) {
    return SearchAndClickImage(ImageName, MaxWaitSec, PostSleep, 1)
}

DoubleClickImage(ImageName, MaxWaitSec := 15, PostSleep := 500) {
    return SearchAndClickImage(ImageName, MaxWaitSec, PostSleep, 2)
}

; =================================================================
; [고도화 엔진] 단계별 오차 범위 에스컬레이션 이미지 탐색 함수
; =================================================================
SearchAndClickImage(ImageName, MaxWaitSec := 15, PostSleep := 500, ClickCount := 1)
{
    global ImageFolder, Variation, G_EnvMode
    ImagePath := ImageFolder "\" ImageName

    ; 🌟 [v4.5] MIS앱 모드: "이름(MIS앱 환경).png" 전용 이미지가 있으면 우선 사용
    if (G_EnvMode = "APP")
    {
        AppPath := RegExReplace(ImagePath, "i)\.png$", "(MIS앱 환경).png")
        if (FileExist(AppPath))
            ImagePath := AppPath
    }

    if (!FileExist(ImagePath))
    {
        UpdateDashboard("⚠️ [파일 없음] " ImageName)
        MsgBox, 48, 이미지 파일 누락, % "[" . ImageName . "] 이미지 파일이 img 폴더에 없습니다!`n매크로를 일시정지합니다."
        Pause
        return False
    }

    ; 🌟 [v4.4] 이전 동작의 로딩이 끝난 뒤에 탐색 시작
    WaitAppReady()

    StartTime := A_TickCount, MaxWaitMs := MaxWaitSec * 1000
    ; 🌟 [v4.5] [종료] 버튼은 앱 전체 종료 버튼과 헷갈리지 않도록 오차 확대 금지
    VarTolerances := InStr(ImageName, "종료") ? [30, 30, 30] : [30, 60, 90]

    Loop
    {
        Elapsed := A_TickCount - StartTime
        CurVar := (Elapsed < 3000) ? VarTolerances[1] : (Elapsed < 7000) ? VarTolerances[2] : VarTolerances[3]

        ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *%CurVar% %ImagePath%
        if (ErrorLevel != 0)
        {
            AltPath := RegExReplace(ImagePath, "i)\.png$", "_alt.png")
            if (FileExist(AltPath))
                ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *%CurVar% %AltPath%
        }

        if (ErrorLevel = 0)
        {
            ; 🌟 [v4.4] 클릭 직전에도 로딩 확인 → 클릭 → 클릭으로 시작된 로딩 끝날 때까지 대기
            WaitAppReady()
            MouseClick, left, % FoundX + 10, % FoundY + 10, %ClickCount%
            Sleep, %PostSleep%
            WaitAppReady()
            return True
        }

        if (Elapsed > MaxWaitMs)
        {
            SoundBeep, 750, 400
            UpdateDashboard("⚠️ [일시정지] " ImageName " 탐색 실패")
            MsgBox, 48, 이미지 탐색 실패 및 일시정지, % "[" . ImageName . "] 이미지 버튼을 화면에서 찾을 수 없습니다!`n`n🔍 대표적 원인 및 해결 방법:`n1. 윈도우 디스플레이 배율이 125%/150%인 경우 ➔ 100% 설정 권장`n2. 브라우저 화면 축소/확대가 100%가 아닌 경우 ➔ 브라우저에서 Ctrl + 0 키 누름`n`n👉 화면에서 해당 버튼을 마우스로 직접 클릭하신 후 [확인]을 누르시면 다음 단계부터 매크로가 정상 계속됩니다."
            WaitAppReady()
            return False
        }
        Sleep, 150
    }
}

Esc::
    DestroyDashboard()
    MsgBox, 48, 매크로 강제 종료 안내, 🛑 사용자에 의해 매크로가 강제 종료되었습니다.
    ExitApp
return

+enter::Pause

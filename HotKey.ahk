#Requires AutoHotkey v2.0
#SingleInstance Force

; =================================================================
; 软件切换配置区 
; Alt + N + 代号: 开启新窗口
; Alt+代号: 已激活则隐藏 / 未激活则呼出 / 未启动则启动
; =================================================================

; 终端
!t:: GetKeyState("n", "P") ? RunNewInstance("wt.exe") : SmartActivate("ahk_class CASCADIA_HOSTING_WINDOW_CLASS", "wt.exe")

; VS Code
!v:: {
    path := "C:\Users\24051\Desktop\Code\Code.lnk"
    GetKeyState("n", "P") ? RunNewInstance(path) : SmartActivate("ahk_exe Code.exe", path)
}

; 文件资源管理器
; !e:: SmartActivate("ahk_class CabinetWClass", "explorer.exe")

; Everything
!s:: {
    path := "C:\Users\24051\Desktop\Tool\Everything.lnk"
    GetKeyState("n", "P") ? RunNewInstance(path) : SmartActivate("ahk_exe Everything.exe", path)
}

; Chrome 
!g:: {
    path := "C:\Users\24051\Desktop\Search\Chrome.lnk"
    GetKeyState("n", "P") ? RunNewInstance(path) : SmartActivate("ahk_exe chrome.exe", path)
}

; QQ 
!q:: SmartActivate("ahk_exe QQ.exe", "C:\Users\24051\Desktop\Social\QQ.lnk")

; 微信 
!w:: SmartActivate("ahk_exe WeChat.exe", "C:\Users\24051\Desktop\Social\Weixin.lnk")

; 豆包
!d:: SmartActivate("ahk_exe 豆包.exe", "C:\Users\24051\Desktop\ai\豆包.lnk")

; =================================================================
; 脚本文件创建配置区 
; Ctrl + Alt + 后缀: 在当前资源管理器目录下新建带模版的文件
; =================================================================

; Markdown
^!m:: NewFileFromExplorer(".md", "# ")

; C++
^!c:: NewFileFromExplorer(".cpp", "#include <iostream>`n`nint main() {`n    std::cout << `"Hello World`" << std::endl;`n    return 0;`n}")

; Python
^!p:: NewFileFromExplorer(".py", "import os`n`ndef main():`n    print(`"Hello World`")`n`nif __name__ == '__main__':`n    main()")

; Header
^!h:: NewFileFromExplorer(".h", "#ifndef HEADER_H`n#define HEADER_H`n`n#endif")

; AHK Script
^!k:: NewFileFromExplorer(".ahk", "; AutoHotkey v2 Script`n`n!j::MsgBox(`"Hello`")")

; =================================================================
; 窗口管理与系统增强区
; Alt + 功能: 针对当前可见窗口进行焦点调度或生命周期管理(关闭)
; =================================================================

; 遍历切换变量初始化
global lastSwitchTime := 0
global switchIndex := 1

; 转移焦点到下一个未隐藏且未最小化的窗口 (Alt + F)
!f::
{
    global lastSwitchTime, switchIndex
    currentTime := A_TickCount
    
    ; 如果距离上次按键超过 800ms，重置索引，从头开始找
    if (currentTime - lastSwitchTime > 800) {
        switchIndex := 1
    }
    
    allWindows := WinGetList()
    activeHWnd := WinExist("A")
    validWindows := []

    ; 1. 筛选符合条件的窗口
    for hwnd in allWindows {
        style := WinGetStyle(hwnd)
        exStyle := WinGetExStyle(hwnd)
        title := WinGetTitle(hwnd)
        
        ; 核心逻辑改进：
        ; - (style & 0x10000000): 窗口必须是可见状态
        ; - WinGetMinMax(hwnd) != -1: 窗口不能是最小化状态 (重点)
        if (title != "" && (style & 0x10000000) && !(exStyle & 0x80)) {
            try {
                if (WinGetMinMax(hwnd) != -1 && !(title ~= "Program Manager|Taskbar")) {
                    validWindows.Push(hwnd)
                }
            }
        }
    }

    ; 2. 执行循环切换逻辑
    if (validWindows.Length > 0) {
        if (switchIndex > validWindows.Length) {
            switchIndex := 1
        }
        
        targetHwnd := validWindows[switchIndex]
        
        ; 如果目标是当前窗口且还有别的可选，跳过它
        if (targetHwnd == activeHWnd && validWindows.Length > 1) {
            switchIndex += 1
            if (switchIndex > validWindows.Length) {
                switchIndex := 1
            }
            targetHwnd := validWindows[switchIndex]
        }

        try {
            WinActivate(targetHwnd)
            switchIndex += 1
            lastSwitchTime := currentTime
        }
    }
}

!Left::CycleSnap("Left")    ; Alt + 左箭头 → 左侧分屏
!Right::CycleSnap("Right")  ; Alt + 右箭头 → 右侧分屏

; 关闭当前活动窗口 (Alt + C)
!c::
{
    activeHWnd := WinExist("A")
    if (activeHWnd)
    {
        WinClose(activeHWnd)
    }
}

; =================================================================
; 核心功能函数区
; =================================================================

; --- 软件切换函数 ---
/**
 * 智能激活/循环/隐藏函数
 */
SmartActivate(TargetIdentifier, PathOrEXE := "") {
    ; 使用 Map 对象存储每个 TargetIdentifier 对应的最后一个窗口 ID
    static LastIDMap := Map() 
    
    ; 初始化该标识的记录
    if !LastIDMap.Has(TargetIdentifier)
        LastIDMap[TargetIdentifier] := 0

    originalMode := A_TitleMatchMode
    SetTitleMatchMode(2)

    try {
        fullList := WinGetList(TargetIdentifier)

        ; --- 场景 1: 没有窗口 -> 启动程序 ---
        if (fullList.Length = 0) {
            if (PathOrEXE = "") 
                throw Error("未提供程序路径")
            RunNewInstance(PathOrEXE)
            return
        }

        ; --- 场景 2: 逻辑判断 ---
        activeID := WinActive("A")
        isActive := false
        
        for id in fullList {
            if (id = activeID) {
                isActive := true
                break
            }
        }

        if (isActive) {
            ; 当前窗口已置顶 -> 最小化
            WinMinimize(activeID)
            LastIDMap[TargetIdentifier] := activeID 
        } else {
            ; 寻找循环起始点
            nextIndex := 1 
            for index, id in fullList {
                if (id = LastIDMap[TargetIdentifier]) {
                    nextIndex := index + 1
                    break
                }
            }
            
            ; 越界重置
            if (nextIndex > fullList.Length) 
                nextIndex := 1

            targetID := fullList[nextIndex]
            
            ; 状态恢复与激活
            if WinGetMinMax(targetID) = -1 
                WinRestore(targetID)
            
            WinActivate(targetID)
            LastIDMap[TargetIdentifier] := targetID
        }
    }
    catch Any as e {
        ToolTip "执行出错: " e.Message
        SetTimer () => ToolTip(), -3000
    }
    finally {
        SetTitleMatchMode(originalMode)
    }
}

/**
 * 强制运行新实例函数
 */
RunNewInstance(PathOrEXE) {
    if (PathOrEXE = "") 
    return
    try {
        Run('"' PathOrEXE '"')
    } catch Any as e {
        MsgBox("启动失败: " e.Message)
    }
}

; --- 资源管理器新建文件函数 ---
NewFileFromExplorer(Extension, TemplateContent := "") {
    try {
        ; --- 以下是你提供的获取路径核心逻辑 ---
        shellApp := ComObject("Shell.Application")
        activeWindow := ""
        currentHwnd := WinExist("A")
        
        for window in shellApp.Windows {
            if (window.HWND = currentHwnd) {
                activeWindow := window
                break
            }
        }
        
        if (activeWindow = "") {
            MsgBox("请在资源管理器窗口中使用此快捷键", "提示", "0x40")
            return
        }

        ; 获取路径 (添加异常捕获防止在“此电脑”等特殊路径崩溃)
        try {
            targetPath := activeWindow.Document.Folder.Self.Path
        } catch {
            MsgBox("无法在此窗口创建文件（可能是特殊系统目录）", "提示")
            return
        }

        ; --- 弹出输入框 ---
        myGui := InputBox("请输入文件名 (无需后缀):", "新建 " Extension, "w300 h130")
        if (myGui.Result = "Cancel")
            return

        fileName := (myGui.Value = "") ? "NewFile" : myGui.Value
        fullPath := targetPath . "\" . fileName . Extension

        if FileExist(fullPath) {
            MsgBox("文件已存在！", "警告", "0x30")
            return
        }

        ; --- 写入文件 ---
        ; 如果是 Markdown 且没有传入自定义模板，使用你要求的格式
        if (Extension = ".md" && TemplateContent = "# ") {
            TemplateContent := "# " . fileName . "`n`n创建时间: " . FormatTime(, "yyyy-MM-dd HH:mm") . "`n`n"
        }

        FileAppend(TemplateContent, fullPath, "UTF-8")

        ; --- 修改点：改为用 VS Code 打开 ---
        ; 使用 code -r 确保在当前 VS Code 窗口打开，不新开窗口
        try {
            Run('code -r "' . fullPath . '"')
        } catch {
            ; 如果 code 命令不可用，则用系统默认程序打开
            Run(fullPath)
        }
        
    } catch Error as e {
        MsgBox("发生错误: " e.Message)
    }
}

/**
 * 循环分屏函数 (一步到位版)
 * 极致性能，无抖动，支持多显示器独立记忆
 */
CycleSnap(Side) {
    static RATIOS := [0.618, 0.50, 0.382]
    static POS_MAP := Map("Left", 0, "Right", 0)

    if !(Side == "Left" || Side == "Right")
        return

    hwnd := WinExist("A")
    if !hwnd 
    return

    try {
        ; 1. 状态处理：如果窗口最大化，先还原
        if WinGetMinMax(hwnd) != 0 
            WinRestore(hwnd)

        ; 2. 获取当前窗口位置以确定所在显示器
        WinGetPos(&startX, &startY, &startW, &startH, hwnd)
        midX := startX + startW/2
        midY := startY + startH/2
        
        targetMon := 1
        loop MonitorGetCount() {
            MonitorGetWorkArea(A_Index, &mL, &mT, &mR, &mB)
            if (midX >= mL && midX <= mR && midY >= mT && midY <= mB) {
                targetMon := A_Index
                break
            }
        }
        MonitorGetWorkArea(targetMon, &L, &T, &R, &B)

        ; 3. 计算目标尺寸
        POS_MAP[Side] := Mod(POS_MAP[Side], RATIOS.Length) + 1
        targetW := (R - L) * RATIOS[POS_MAP[Side]]
        targetH := B - T
        targetY := T
        targetX := (Side == "Left") ? L : R - targetW

        ; 4. 一步到位移动窗口
        WinMove(Floor(targetX), Floor(targetY), Floor(targetW), Floor(targetH), hwnd)

    } catch Error {
        return
    }
}
#Requires AutoHotkey v2.0
#SingleInstance Force

; =================================================================
; 1. 全局配置 (数据驱动层)
;    在这里集中管理所有应用路径、快捷键和模板内容
; =================================================================

global GlobalConfig := {
    ; 应用切换配置
    Apps: [
        { Key: "!t", ID: "ahk_class CASCADIA_HOSTING_WINDOW_CLASS", Path: "wt.exe" },
        { Key: "!v", ID: "ahk_exe Code.exe", Path: "C:\Users\24051\Desktop\Code\Code.lnk" },
        { Key: "!s", ID: "ahk_exe Everything.exe", Path: "C:\Users\24051\Desktop\Tool\Everything.lnk" },
        { Key: "!g", ID: "ahk_exe chrome.exe", Path: "C:\Users\24051\Desktop\Search\Chrome.lnk" },
        { Key: "!q", ID: "ahk_exe QQ.exe", Path: "C:\Users\24051\Desktop\Social\QQ.lnk" },
        { Key: "!w", ID: "ahk_exe WeChat.exe", Path: "C:\Users\24051\Desktop\Social\Weixin.lnk" },
        { Key: "!d", ID: "ahk_exe 豆包.exe", Path: "C:\Users\24051\Desktop\ai\豆包.lnk" }
    ],
    ; 文件模板配置
    Templates: [
        { Key: "^!m", Ext: ".md",  Tpl: "# {{Name}}`n`n创建时间: {{Time}}`n`n" },
        { Key: "^!c", Ext: ".cpp", Tpl: "#include <iostream>`n`nint main() {`n    std::cout << `"Hello World`" << std::endl;`n    return 0;`n}" },
        { Key: "^!p", Ext: ".py",  Tpl: "import os`n`ndef main():`n    print(`"Hello World`")`n`nif __name__ == '__main__':`n    main()" },
        { Key: "^!h", Ext: ".h",   Tpl: "#ifndef HEADER_H`n#define HEADER_H`n`n#endif" },
        { Key: "^!k", Ext: ".ahk", Tpl: "; AutoHotkey v2 Script`n`n!j::MsgBox(`"Hello`")" }
    ],
    ; 窗口管理参数
    Settings: {
        SwitchThreshold: 800,    ; Alt+F 连续切换间隔 (ms)
        SnapRatios: [0.618, 0.50, 0.382]
    }
}

; =================================================================
; 2. 动态热键绑定 (初始化层)
;    脚本启动时自动执行，将配置转换为可用的快捷键
; =================================================================

; 绑定应用热键
for app in GlobalConfig.Apps {
    ; 闭包处理防止变量捕获冲突
    ( (a) => Hotkey(a.Key, (k) => GetKeyState("n", "P") ? RunNewInstance(a.Path) : SmartActivate(a.ID, a.Path)) )(app)
}

; 绑定模板热键
for tpl in GlobalConfig.Templates {
    ( (t) => Hotkey(t.Key, (k) => NewFileFromExplorer(t.Ext, t.Tpl)) )(tpl)
}

; =================================================================
; 3. 静态热键定义 (交互层)
;    常规的快捷键定义，如分屏、关闭窗口、遍历切换等
; =================================================================

;窗口快速关闭 Alt + C
!c:: {
    if (active := WinExist("A"))
        WinClose(active)
}

; 循环分屏 (Alt + Left/Right)
!Left::CycleSnap("Left")
!Right::CycleSnap("Right")

; 全局窗口循环切换 Alt + F
global lastSwitchTime := 0
global switchIndex := 1

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
        ; - WinGetMinMax(hwnd) != -1: 窗口不能是最小化状态
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
; =================================================================
; 4. 核心功能函数 (业务逻辑层)
;    被热键调用的复杂逻辑实现
; =================================================================

;智能激活/最小化/启动函数
SmartActivate(TargetIdentifier, PathOrEXE := "") {
    static LastIDMap := Map()
    if !LastIDMap.Has(TargetIdentifier)
        LastIDMap[TargetIdentifier] := 0

    try {
        fullList := WinGetList(TargetIdentifier)
        
        if (fullList.Length = 0) {
            if (PathOrEXE = "") 
                throw Error("未提供路径")
            return RunNewInstance(PathOrEXE)
        }

        activeID := WinExist("A")
        isCurrentActive := false
        for id in fullList {
            if (id = activeID) {
                isCurrentActive := true
                break
            }
        }

        if (isCurrentActive) {
            WinMinimize(activeID)
            LastIDMap[TargetIdentifier] := activeID
        } else {
            ; 找到下个窗口的逻辑
            nextIndex := 1
            for index, id in fullList {
                if (id = LastIDMap[TargetIdentifier]) {
                    nextIndex := index + 1
                    break
                }
            }
            if (nextIndex > fullList.Length) 
                nextIndex := 1

            targetID := fullList[nextIndex]
            if WinGetMinMax(targetID) = -1 
                WinRestore(targetID)
            
            WinActivate(targetID)
            LastIDMap[TargetIdentifier] := targetID
        }
    } catch Any as e {
        NotifyError("激活失败: " e.Message)
    }
}


; 资源管理器新建文件
NewFileFromExplorer(Extension, TemplateContent := "") {
    try {
        ; 获取当前资源管理器路径
        shellApp := ComObject("Shell.Application")
        activeHwnd := WinExist("A")
        targetPath := ""
        
        for window in shellApp.Windows {
            if (window.HWND = activeHwnd) {
                targetPath := window.Document.Folder.Self.Path
                break
            }
        }

        if (targetPath = "") {
            MsgBox("请在有效的资源管理器窗口中使用此功能。", "提示")
            return
        }

        ; 输入框
        userInput := InputBox("请输入文件名 (无需后缀):", "新建 " Extension, "w300 h130")
        if (userInput.Result = "Cancel") 
        return
        
        fileName := (userInput.Value = "") ? "NewFile" : userInput.Value
        
        ; 防御性编程：非法字符过滤
        if (fileName ~= '[\\/:*?"<>|]') {
            MsgBox("文件名包含非法字符！", "错误")
            return
        }

        fullPath := targetPath . "\" . fileName . Extension
        if FileExist(fullPath) {
            MsgBox("文件已存在！", "警告")
            return
        }

        ; 模板变量替换
        content := StrReplace(TemplateContent, "{{Name}}", fileName)
        content := StrReplace(content, "{{Time}}", FormatTime(, "yyyy-MM-dd HH:mm"))

        FileAppend(content, fullPath, "UTF-8")

        ; 尝试用 VS Code 打开，失败则默认打开
        try {
            Run('code -r "' . fullPath . '"')
        } catch {
            Run(fullPath)
        }
        
    } catch Any as e {
        NotifyError("新建文件失败: " e.Message)
    }
}

; 循环分屏函数
CycleSnap(Side) {
    static POS_MAP := Map("Left", 0, "Right", 0)
    ratios := GlobalConfig.Settings.SnapRatios
    
    hwnd := WinExist("A")
    if !hwnd 
    return

    try {
        if WinGetMinMax(hwnd) != 0 
            WinRestore(hwnd)

        ; 获取显示器工作区
        MonitorGetWorkArea(GetMonitorIndexFromWindow(hwnd), &L, &T, &R, &B)

        POS_MAP[Side] := Mod(POS_MAP[Side], ratios.Length) + 1
        
        targetW := (R - L) * ratios[POS_MAP[Side]]
        targetH := B - T
        targetX := (Side == "Left") ? L : R - targetW
        
        WinMove(Floor(targetX), Floor(T), Floor(targetW), Floor(targetH), hwnd)
    } catch Error as e {
        NotifyError("分屏失败: " e.Message)
    }
}

; =================================================================
; 5. 底层辅助工具 (系统调用层)
;    通用的、不涉及业务逻辑的辅助函数
; =================================================================

; 启动新实例函数
RunNewInstance(Path) {
    try {
        Run('"' Path '"')
    } catch Any as e {
        NotifyError("启动失败: " e.Message)
    }
}

; 根据窗口句柄获取所在显示器索引
GetMonitorIndexFromWindow(hwnd) {
    WinGetPos(&x, &y, &w, &h, hwnd)
    midX := x + w/2
    midY := y + h/2
    loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &mL, &mT, &mR, &mB)
        if (midX >= mL && midX <= mR && midY >= mT && midY <= mB)
            return A_Index
    }
    return 1
}

; 错误通知函数
NotifyError(msg) {
    ToolTip("Error: " msg)
    SetTimer () => ToolTip(), -3000
}
#Requires AutoHotkey v2.0
#SingleInstance Force

; =================================================================
; 1. 软件切换配置区 (Alt + 字母)
; 逻辑：已激活则隐藏 / 未激活则呼出 / 未启动则启动
; =================================================================

; 终端
!t:: SmartActivate("ahk_class CASCADIA_HOSTING_WINDOW_CLASS", "wt.exe")

; VS Code (新增)
!v:: SmartActivate("ahk_exe Code.exe", "C:\Users\24051\Desktop\Code\Code.lnk")

; 文件资源管理器
; !e:: SmartActivate("ahk_class CabinetWClass", "explorer.exe")

; Everything
!s:: SmartActivate("ahk_exe Everything.exe", "C:\Users\24051\Desktop\Tool\Everything.lnk")

; Chrome 
!g:: SmartActivate("ahk_exe chrome.exe", "C:\Users\24051\Desktop\Search\Chrome.lnk")

; Clash 
; !c:: SmartActivate("ahk_exe Clash for Windows.exe", "C:\Users\24051\Desktop\Tool\ClashforWindows.lnk")

; QQ 
!q:: SmartActivate("ahk_exe QQ.exe", "C:\Users\24051\Desktop\Social\QQ.lnk")

; 微信 
!w:: SmartActivate("ahk_exe WeChat.exe", "C:\Users\24051\Desktop\Social\Weixin.lnk")

; 豆包
!d:: SmartActivate("ahk_exe 豆包.exe", "C:\Users\24051\Desktop\ai\豆包.lnk")

; =================================================================
; 2. 脚本创建配置区 (Ctrl + Alt + 字母)
; 逻辑：在当前资源管理器目录下新建带模版的文件
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
; 3. 窗口管理与系统增强区 (Alt + 功能)
; 逻辑：针对当前可见窗口进行焦点调度或生命周期管理(关闭)
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
; 4. 核心功能函数区
; =================================================================

; --- 软件智能切换函数 ---
/**
 * 智能激活/循环/隐藏函数
 * @param TargetIdentifier 窗口标识 (如 "ahk_class CabinetWClass")
 * @param PathOrEXE 程序路径
 */
SmartActivate(TargetIdentifier, PathOrEXE := "") {
    static lastShownID := 0  ; 静态变量保持记忆
    originalMode := A_TitleMatchMode
    SetTitleMatchMode(2)

    try {
        ; v2 中 WinGetList 返回的是一个数组对象
        fullList := WinGetList(TargetIdentifier)

        ; --- 场景 1: 没有窗口 -> 启动程序 ---
        if (fullList.Length = 0) {
            if (PathOrEXE = "") {
                throw Error("未提供程序路径，无法启动。")
            }
            Run('"' PathOrEXE '"') ; v2 推荐用单引号包裹双引号
            if WinWaitActive(TargetIdentifier, , 5) {
                lastShownID := WinActive("A")
            }
            return true
        }

        ; --- 场景 2: 逻辑判断 ---
        activeID := WinActive("A")
        
        ; 检查当前窗口是否属于目标程序 (v2 简洁写法)
        isActive := false
        for id in fullList {
            if (id = activeID) {
                isActive := true
                break
            }
        }

        if (isActive) {
            ; 当前活动窗口就是目标之一 -> 最小化隐藏
            WinMinimize(activeID)
            ; 记录最后操作的 ID，下次唤醒时可能想先唤醒这一个，或者跳过它
            lastShownID := activeID 
        } else {
            ; 当前不活动 -> 准备唤醒循环中的下一个
            nextIndex := 1 ; 默认第一个
            
            for index, id in fullList {
                if (id = lastShownID) {
                    nextIndex := index + 1
                    break
                }
            }

            ; 越界检查 (v2 数组索引从 1 开始)
            if (nextIndex > fullList.Length) {
                nextIndex := 1
            }

            targetID := fullList[nextIndex]
            WinRestore(targetID)
            WinActivate(targetID)
            lastShownID := targetID
        }
    }
    catch Error as e {
        MsgBox("执行出错：`n" e.Message)
    }
    finally {
        SetTitleMatchMode(originalMode)
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

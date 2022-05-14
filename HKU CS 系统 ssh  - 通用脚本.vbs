Dim cancel_or_not 

cancel_or_not = MsgBox("确定要打开吗，请调成英文输入法；或者是在等待的1秒钟内按下shift键", 1, "确认框")

rem 如果是按确定，返回1，取消，返回2，关闭（没试验）
If cancel_or_not = 2 Then
WScript.Quit
End If

rem Windows 是一种脚本宿主 ，WshShell 对象,提供对本地 Windows 外壳程序的访问。
Dim WshShell
Set WshShell = Wscript.CreateObject("WScript.Shell")

rem 区别？
Wshshell.Run "cmd"
rem Wshshell.Run "cmd.exe"
' CreateObject("Shell.Application").ShellExecute "cmd.exe","pause","","runas",1

WScript.Sleep 1000

rem 表示不想 以这种方式暂停程序
rem WScript.Sleep 1500

WScript.Sleep 300

' 这里写你的账号（第一道门）
WshShell.SendKeys "ssh @gatekeeper.cs.hku.hk"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 300

' 写密码
WshShell.SendKeys ""
WshShell.SendKeys "{ENTER}"

WScript.Sleep 300

' 第二道门（账号）
WshShell.SendKeys "ssh @academy.cs.hku.hk"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 300

' 写密码
WshShell.SendKeys ""

WshShell.SendKeys "{ENTER}"



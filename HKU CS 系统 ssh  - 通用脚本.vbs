Dim cancel_or_not 

cancel_or_not = MsgBox("ȷ��Ҫ���������Ӣ�����뷨���������ڵȴ���1�����ڰ���shift��", 1, "ȷ�Ͽ�")

rem ����ǰ�ȷ��������1��ȡ��������2���رգ�û���飩
If cancel_or_not = 2 Then
WScript.Quit
End If

rem Windows ��һ�ֽű����� ��WshShell ����,�ṩ�Ա��� Windows ��ǳ���ķ��ʡ�
Dim WshShell
Set WshShell = Wscript.CreateObject("WScript.Shell")

rem ����
Wshshell.Run "cmd"
rem Wshshell.Run "cmd.exe"
' CreateObject("Shell.Application").ShellExecute "cmd.exe","pause","","runas",1

WScript.Sleep 1000

rem ��ʾ���� �����ַ�ʽ��ͣ����
rem WScript.Sleep 1500

WScript.Sleep 300

' ����д����˺ţ���һ���ţ�
WshShell.SendKeys "ssh @gatekeeper.cs.hku.hk"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 300

' д����
WshShell.SendKeys ""
WshShell.SendKeys "{ENTER}"

WScript.Sleep 300

' �ڶ����ţ��˺ţ�
WshShell.SendKeys "ssh @academy.cs.hku.hk"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 300

' д����
WshShell.SendKeys ""

WshShell.SendKeys "{ENTER}"



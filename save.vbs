Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate "notepad"
WshShell.SendKeys "^a"
WshShell.SendKeys "^s"
WshShell.SendKeys "passwords"
WshShell.SendKeys "{ENTER}"
WshShell.SendKeys "{LEFT}"
WshShell.SendKeys "{ENTER}"
WshShell.SendKeys "%{F4}"
WScript.Sleep 200
WshShell.SendKeys "%{F4}"








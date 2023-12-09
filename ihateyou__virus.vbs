Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 2000 ' Initial delay

Do
    WshShell.SendKeys "I HATE YOU"
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 500 ' Wait for 0.5 seconds
Loop

Dim objResult

Set objShell = WScript.CreateObject("Wscript.Shell")

Do While True
    objResult = objShell.sendKeys("{NUMLOCK}{NUMLOCK}")
    Wscript.Sleep(15000)
Loop
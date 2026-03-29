Dim shell, scriptPath, command, i

scriptPath = Replace(WScript.ScriptFullName, "launch.vbs", "ps-toolbox.ps1")
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & QuoteArg(scriptPath)

For i = 0 To WScript.Arguments.Count - 1
    command = command & " " & QuoteArg(WScript.Arguments(i))
Next

Set shell = CreateObject("WScript.Shell")
shell.Run command, 0, False

Function QuoteArg(value)
    QuoteArg = Chr(34) & Replace(value, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Dim fso, JSON, jsonStr, o, i, obj
Dim wsh, AppPath, code

Include("vbsJson.vbs")
Set JSON = New VbsJson
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
jsonStr = fso.OpenTextFile("cmdsConfig.js").ReadAll
Set obj = JSON.Decode(jsonStr)

AppPathDefault = "C:/Windows/System32/cmd.exe"
AppPath = obj("AppPath")
commonSleep = obj("sleep")

If commonSleep = "" Then
    commonSleep = 300
End If

If AppPath = "" Then
	AppPath = AppPathDefault
End If
AppPath = "" + AppPath + ""



' WScript.Echo AppPath
Set wsh=WScript.CreateObject("WScript.Shell")
wsh.Run AppPath
WScript.Sleep 1000

For Each sendKey In obj("cmds")
    key = sendKey("cmd")
    sleep = sendKey("sleep")
    If sleep = "" Then
    	sleep = commonSleep
    End If
    wsh.SendKeys key
    wsh.SendKeys " {ENTER} "
    ' WScript.Echo sleep
    WScript.Sleep sleep
Next


Sub Include(sInstFile) 
    Dim oFSO, f, s 
    Set oFSO = CreateObject("Scripting.FileSystemObject") 
    Set f = oFSO.OpenTextFile(sInstFile) 
    s = f.ReadAll 
    f.Close
    ExecuteGlobal s 
End Sub
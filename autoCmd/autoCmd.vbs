Dim fso, JSON, jsonStr, o, i, obj
Dim wsh, AppPath, code

cmdsConfigFile = "cmdsConfig.json"

Include("vbsJson.vbs")
Set JSON = New VbsJson
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
if fso.FileExists(cmdsConfigFile) then
    
else
    msgbox "���ڵ�ǰĿ¼��������������ļ�:"&cmdsConfigFile
    Wscript.Quit
end if

jsonStr = fso.OpenTextFile(cmdsConfigFile).ReadAll
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
' wsh.Run "notepad"
WScript.Sleep 1000

For Each sendKey In obj("cmds")
    key = sendKey("cmd")
    sleep = sendKey("sleep")
    isPrompt = sendKey("isPrompt")
    title = sendKey("title")
    If sleep = "" Then
        sleep = commonSleep
    End If
    If title = "" Then
        title = "please input" 
    End If

    If isPrompt Then
        key = Inputbox(title, "", "")
    End If
    
    key = keyMapTrans(key)
    ' msgBox key

    wsh.SendKeys key
    wsh.SendKeys "{ENTER}"
    ' WScript.Echo sleep
    WScript.Sleep sleep
Next

Function keyMapTrans(ByVal str)
    Dim keysMapStr, keysMap, transKey
    keysMapStr = fso.OpenTextFile("./keysMap.json").ReadAll
    Set keysMap = JSON.Decode(keysMapStr)
    For Each key In keysMap
        transKey = keysMap.Item(key)
        ' msgbox str + key + transKey
        str = Replace(str, key, transKey)
    Next
    keyMapTrans = str
End Function

Sub keyReplace(str, find, replaceWith)
    a =  Replace(str, find, replaceWith)
End Sub


Sub Include(sInstFile) 
    Dim oFSO, f, s 
    Set oFSO = CreateObject("Scripting.FileSystemObject") 
    Set f = oFSO.OpenTextFile(sInstFile) 
    s = f.ReadAll 
    f.Close
    ExecuteGlobal s 
End Sub
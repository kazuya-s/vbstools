'*********************************************************
' Purpose: Display first lines of a file
' Copyright (c) 2014 Kazuya Shindome
' This software is released under the MIT License
'
' Usage: cscript head.vbs file [count]
'*********************************************************
Option Explicit

Dim objArgs
Dim strFilePath
Dim lngMaxCount

lngMaxCount = 10

Set objArgs = WScript.Arguments

If objArgs.Count >= 1 Then
    strFilePath = objArgs(0)
Else
    showUsage
End If

If objArgs.Count >= 2 Then
    lngMaxCount = CLng(objArgs(1))
End If

Main strFilePath, lngMaxCount

Sub Main(ByVal strFilePath, ByVal lngMaxCount)
    Dim objFileSys
    Dim objRead
    Dim objWrite
    Dim lngCount

    Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
    Set objRead = objFileSys.OpenTextFile(strFilePath)
    Set objWrite = WScript.CreateObject("ADODB.Stream")

    objWrite.Open

    lngCount = 0
    Do Until objRead.AtEndOfStream
        If lngCount >= lngMaxCount Then
            Exit Do
        End If
        
        objWrite.WriteText objRead.ReadLine & vbCrLf
        
        lngCount = lngCount + 1
    Loop
    objRead.Close

    objWrite.Position = 0
    WScript.Echo objWrite.ReadText
    objWrite.Close

End Sub

Sub ShowUsage()
    If blnIsWScript() Then
        WScript.Echo "Error: Could not determine the target of the Drag&Drop operation."
    Else
        WScript.Echo "Usage:"
        WScript.Echo "  cscript head.vbs file [count]"
    End If

    WScript.Quit 0
End Sub


Function blnIsWScript()
    blnIsWScript = InStr(1, LCase(WScript.FullName), "wscript.exe")
End Function

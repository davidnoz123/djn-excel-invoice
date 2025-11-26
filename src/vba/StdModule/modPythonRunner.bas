Attribute VB_Name = "modPythonRunner"
Option Explicit

' Run a Python script and stream stdout/stderr to a callback object.
'
' pythonExePath   - full path to python.exe, or "python"
' scriptPath      - full path to script.py
' args            - optional command line arguments
' stdoutCallback  - object with Public Sub OnStdout(line As String)
' exitCode        - returned exit code of the Python process
Public Sub RunPythonNoStdin( _
        ByVal pythonExePath As String, _
        ByVal scriptPath As String, _
        Optional ByVal args As String = "", _
        Optional ByVal stdoutCallback As Object = Nothing, _
        Optional ByVal envDict As Scripting.Dictionary = Nothing, _
        Optional ByRef exitCode As Long)

    Dim shellObj As Object
    Dim proc As Object
    Dim cmd As String
    Dim line As String

    ' Build command
    cmd = """" & pythonExePath & """ """ & scriptPath & """"
    If Len(args) > 0 Then cmd = cmd & " " & args

    Set shellObj = CreateObject("WScript.Shell")
    
    If Not envDict Is Nothing Then
        Dim env As Object
        Set env = shellObj.Environment("Process")
        Dim key As Variant
        For Each key In envDict.Keys
            env(CStr(key)) = CStr(envDict(key))
        Next key
    End If
    
    Set proc = shellObj.Exec(cmd)

    ' ----- read stdout -----
    Do While Not proc.StdOut.AtEndOfStream
        line = proc.StdOut.ReadLine
        If Not stdoutCallback Is Nothing Then
            On Error Resume Next
            stdoutCallback.OnStdout line
            On Error GoTo 0
        End If
    Loop

    ' ----- read stderr (optional but prevents blocking) -----
    Do While Not proc.StdErr.AtEndOfStream
        line = proc.StdErr.ReadLine
        If Not stdoutCallback Is Nothing Then
            On Error Resume Next
            stdoutCallback.OnStdout "[stderr] " & line
            On Error GoTo 0
        End If
    Loop

    exitCode = proc.exitCode
End Sub



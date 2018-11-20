Attribute VB_Name = "modLog"
Option Explicit

Dim FileRegister As New Collection

''----------------''
Private Declare Function GetFileAttributes Lib "kernel32" Alias _
                        "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Function FileExists(ByVal strPathName As String) As Boolean
  Dim af As Long
    af = GetFileAttributes(strPathName)
    FileExists = ((af <> -1) And af <> vbDirectory)
End Function

Public Sub DGPSLog(data1 As String, sName As String)
    
Dim f1
Dim f2
Dim Fname     As String
Dim str1      As String
Dim str2      As String
Dim SaveDir   As String
 
''//Date-LotNo.log '' filename
    
    str1 = Format(Now, "YYYYMMDD")
    str2 = Format(Now, "YYYYMMDD-hh:mm:ss")
    
    ''SaveDir = App.Path & "\" & str1 & ".Log"
    SaveDir = "C:\BIN_LOG" ''& "\" & str1 & ".Log"
    

On Error GoTo errFile1

    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
    End If
    
    Fname = SaveDir & "\" & sName & "_" & str1 & ".log"  ''"_DataFile.log"
    
    If Not FileExists(Fname) Then
        f1 = FreeFile
        Open Fname For Binary Access Write As #f1
            ''Put #f1, , "DAC-LOG :: " + Fname + vbCrLf + vbCrLf
            Put #f1, , str2 & " " & data1$ & vbCrLf
        Close #f1
        DoEvents
        'Sleep 10
    Else
    
         f2 = FreeFile
        Open Fname For Binary Access Write As #f2
            Seek #f2, LOF(f2) + 1
            Put #f2, , str2 & " " & data1$ & vbCrLf
            ''Put #f2, , vbCrLf & data1$
        Close #f2
        DoEvents
    
    End If

errFile1:
    SaveDir = ""
    ''''''''''''(just-cancle~)
    
End Sub

Public Sub SaveBuffer2File(FileNamePrefix As String, buffer() As Byte, size As Long)
    Dim dirName         As String
    Dim fileName        As String
    Dim FileNumber
    Dim i As Long
'
    dirName = "C:\BIN_LOG\"
    fileName = _
        dirName & FileNamePrefix _
        & Format(Now, "YYYYMMDD_hhmmss") _
        & "_" & Format(GetTickCount() Mod 1000, "000") & ".dat"
'
On Error GoTo errFile1
'
    If Dir(dirName, vbDirectory) = "" Then
        MkDir dirName
    End If
'
    If FileExists(fileName) Then
        Exit Sub
    End If
'
    FileNumber = FreeFile
    Open fileName For Binary Access Write As #FileNumber
'
    For i = 0 To size - 1
        Put #FileNumber, , buffer(i)
    Next i
'
    Close #FileNumber
'
    DoEvents
'
errFile1:
    dirName = ""
    ''''''''''''(just-cancle~)
'
End Sub

Public Function Tilt3Dlog_start(FileNamePrefix As String, Message As String) As Integer

    Dim dirName         As String
    Dim fileName        As String
    Dim FileNumber
'
    dirName = "C:\BIN_LOG\"
    fileName = _
        dirName & FileNamePrefix & "_3D_" _
        & Format(Now, "YYYYMMDD_hhmmss") _
        & "_" & Format(GetTickCount() Mod 1000, "000") & ".dat"
'
On Error GoTo errFile1
'
    If Dir(dirName, vbDirectory) = "" Then
        MkDir dirName
    End If
'
    FileNumber = FreeFile
    Open fileName For Binary Access Write As #FileNumber
    If FileExists(fileName) Then
        Seek #FileNumber, LOF(FileNumber) + 1
    End If
    Put #FileNumber, , Message & vbCrLf
'
On Error Resume Next
'
    FileRegister.Add fileName, Str(FileNumber)
'
errFile1:
    dirName = ""
    ''''''''''''(just-cancle~)
'
    Tilt3Dlog_start = FileNumber
'
End Function

Public Sub Tilt3Dlog_add(FileNumber As Integer, Message As String)
    Dim fileName        As String
'
On Error Resume Next
'
    fileName = FileRegister.Item(Str(FileNumber))
'
On Error GoTo errFile1
'
    If FileExists(fileName) Then
        Put #FileNumber, , Message & vbCrLf
    End If
'
errFile1:
    ''''''''''''(just-cancle~)
'
End Sub

Public Sub Tilt3Dlog_end(FileNumber As Integer)
'
On Error GoTo errFile1
'
    Close #FileNumber
'
errFile1:
    ''''''''''''(just-cancle~)
'
On Error Resume Next
    FileRegister.Remove Str(FileNumber)
'
End Sub

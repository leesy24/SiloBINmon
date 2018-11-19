Attribute VB_Name = "modLog"
Option Explicit

''----------------''
Public Declare Function GetFileAttributes Lib "kernel32" Alias _
                        "GetFileAttributesA" (ByVal lpFileName As String) As Long

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

Public Sub SaveBuffer2File(Index As Integer, buffer() As Byte, size As Integer)
    Dim dirName         As String
    Dim fileName        As String
    Dim fileNumber
    Dim i As Integer
    
    dirName = "C:\BIN_LOG\"
    fileName = dirName & Index & "_" & Format(Now, "YYYYMMDDhhmmss") & ".dat"

On Error GoTo errFile1

    If Dir(dirName, vbDirectory) = "" Then
        MkDir dirName
    End If
    
    If FileExists(fileName) Then
        Exit Sub
    End If
    
    fileNumber = FreeFile
    Open fileName For Binary Access Write As #fileNumber
       
    For i = 0 To size - 1
        Put #fileNumber, , buffer(i)
    Next i
    
    Close #fileNumber
    
    DoEvents

errFile1:
    dirName = ""
    ''''''''''''(just-cancle~)
    
End Sub

Public Sub Tilt3Dlog(FileNamePrefix As String, str As String)
    Dim dirName         As String
    Dim fileName        As String
    Dim fileNumber
    Dim i As Integer
    
    dirName = "C:\BIN_LOG\"
    fileName = dirName & FileNamePrefix & "_3D_" & Format(Now, "YYYYMMDD") & ".dat"

On Error GoTo errFile1

    If Dir(dirName, vbDirectory) = "" Then
        MkDir dirName
    End If
    
    If Not FileExists(fileName) Then
        fileNumber = FreeFile
        Open fileName For Binary Access Write As #fileNumber
            Put #fileNumber, , str & vbCrLf
        Close #fileNumber
        DoEvents
    Else
    
        fileNumber = FreeFile
        Open fileName For Binary Access Write As #fileNumber
            Seek #fileNumber, LOF(fileNumber) + 1
            Put #fileNumber, , str & vbCrLf
        Close #fileNumber
        DoEvents
    End If
    
errFile1:
    dirName = ""
    ''''''''''''(just-cancle~)
    
End Sub











































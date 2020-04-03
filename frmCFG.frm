VERSION 5.00
Begin VB.Form frmCFG 
   BorderStyle     =   1  '단일 고정
   Caption         =   "BIN-LEVEL CONFIG"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin VB.Frame frTypes 
      Caption         =   "센서 종류 설정"
      Height          =   2355
      Left            =   180
      TabIndex        =   17
      Top             =   4200
      Width           =   7875
      Begin VB.TextBox txtCtypes 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   360
         Width           =   555
      End
      Begin VB.CommandButton cmdSetTYPE 
         BackColor       =   &H8000000A&
         Caption         =   "적 용"
         Height          =   375
         Left            =   6120
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lbBinNO2 
         Caption         =   "SILO-"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   420
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "누적횟수"
      Height          =   1215
      Left            =   3360
      TabIndex        =   16
      Top             =   2760
      Width           =   975
      Begin VB.CommandButton cmdDeepMAX 
         BackColor       =   &H8000000A&
         Caption         =   "적 용"
         Height          =   315
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtAVRcnt 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H000040C0&
         Height          =   270
         Left            =   180
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "99"
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCFGexit 
      BackColor       =   &H8000000A&
      Caption         =   "닫 기"
      Height          =   555
      Left            =   6240
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   3420
      Width           =   1575
   End
   Begin VB.Frame frOpset 
      Caption         =   "기준높이 설정"
      Height          =   1215
      Left            =   180
      TabIndex        =   13
      Top             =   2760
      Width           =   3015
      Begin VB.CommandButton cmdMinMax 
         BackColor       =   &H8000000A&
         Caption         =   "적 용"
         Height          =   615
         Left            =   1920
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   420
         Width           =   915
      End
      Begin VB.TextBox txtMaxHH 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H000040C0&
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Text            =   "5000"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox txtBaseHH 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H000040C0&
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Text            =   "100"
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lbBaseHH_ 
         BackStyle       =   0  '투명
         Caption         =   "cm   0%"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   780
         TabIndex        =   15
         Top             =   780
         Width           =   855
      End
      Begin VB.Label lbMaxHH_ 
         BackStyle       =   0  '투명
         Caption         =   "cm 100%"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   780
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame frAngle 
      Caption         =   "센서 기울기 설정"
      Height          =   2355
      Left            =   180
      TabIndex        =   11
      Top             =   240
      Width           =   7875
      Begin VB.CommandButton cmdLoadANG 
         BackColor       =   &H8000000A&
         Caption         =   "기본값 읽기"
         Height          =   375
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdSetANG 
         BackColor       =   &H8000000A&
         Caption         =   "적 용"
         Height          =   375
         Left            =   6120
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtCangle 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Text            =   "0"
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lbBinNO 
         Caption         =   "SILO-"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   1035
      End
   End
   Begin VB.Timer tmrCFG 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   7140
      Top             =   2580
   End
End
Attribute VB_Name = "frmCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const TIMEOUT = 60000 ' 60secs

Private Sub cmdCFGexit_Click()
    frmSettings.Visible = False
    frmCFG.Visible = False
End Sub


Private Sub cmdLoadANG_Click()

    ''  SaveSetting App.Title, "Settings", "SILOang_01", 2
    ''  SaveSetting App.Title, "Settings", "SILOang_02", 0
    ''  SaveSetting App.Title, "Settings", "SILOang_03", 1
    ''  SaveSetting App.Title, "Settings", "SILOang_04", 2
    ''  SaveSetting App.Title, "Settings", "SILOang_05", -2
    ''  SaveSetting App.Title, "Settings", "SILOang_06", -1
    ''  SaveSetting App.Title, "Settings", "SILOang_07", 0
    ''  SaveSetting App.Title, "Settings", "SILOang_08", 1
    ''  SaveSetting App.Title, "Settings", "SILOang_09", 0
    ''  SaveSetting App.Title, "Settings", "SILOang_10", 1
    ''  SaveSetting App.Title, "Settings", "SILOang_11", 2
    ''  SaveSetting App.Title, "Settings", "SILOang_12", 1
    ''  SaveSetting App.Title, "Settings", "SILOang_13", 0
    ''  SaveSetting App.Title, "Settings", "SILOang_14", 0
    ''  SaveSetting App.Title, "Settings", "SILOang_15", 0
    
    ''    ucSilo1(0).set_Angle 2
    ''    ucSilo1(1).set_Angle 0
    ''    ucSilo1(2).set_Angle 1
    ''    ucSilo1(3).set_Angle 2
    ''    ucSilo1(4).set_Angle -2
    ''    ucSilo1(5).set_Angle -1
    ''    ucSilo1(6).set_Angle 0
    ''    ucSilo1(7).set_Angle 1
    ''    ucSilo1(8).set_Angle 0
    ''    ucSilo1(9).set_Angle 1
    ''    ucSilo1(10).set_Angle 2
    ''    ucSilo1(11).set_Angle 1
    ''    ucSilo1(12).set_Angle 0
    ''    ucSilo1(13).set_Angle 0
    ''    ucSilo1(14).set_Angle 0

    txtCangle(0) = "2"
    txtCangle(1) = "0"
    txtCangle(2) = "1"
    txtCangle(3) = "2"
    txtCangle(4) = "-2"
    txtCangle(5) = "-1"
    txtCangle(6) = "0"
    txtCangle(7) = "1"
    txtCangle(8) = "0"
    txtCangle(9) = "1"
    txtCangle(10) = "2"
    txtCangle(11) = "1"
    txtCangle(12) = "0"
    txtCangle(13) = "0"
    txtCangle(14) = "0"
    
End Sub



Private Sub cmdMinMax_Click()

    txtMaxHH = Trim(Val(txtMaxHH))
    frmMain.txtMaxHH = txtMaxHH
    
    txtBaseHH = Trim(Val(txtBaseHH))
    frmMain.txtBaseHH = txtBaseHH

    SaveSetting App.Title, "Settings", "MaxHH", Trim(txtMaxHH.Text)
    SaveSetting App.Title, "Settings", "BaseHH", Trim(txtBaseHH.Text)
    
    Dim i
    For i = 0 To 14  ''3 ''10
        frmMain.ucSilo1(i).set_maxHH CLng(txtMaxHH)
        frmMain.ucSilo1(i).set_baseHH CLng(txtBaseHH)
    Next i
    
    For i = 15 To 18  ''New-CTS-SILO
        frmMain.ucSilo1(i).set_maxHH CLng(txtMaxHH)
        frmMain.ucSilo1(i).set_baseHH CLng(txtBaseHH)
    Next i

    tmrCFG_update

End Sub

Private Sub cmdSetANG_Click()
Dim i
    For i = 0 To 14
        frmMain.ucSilo1(i).set_Angle CDbl(txtCangle(i))
    Next i
    
    tmrCFG_update
    
End Sub


Private Sub cmdDeepMAX_Click()

    If (Val(txtAVRcnt) < 10) Or (Val(txtAVRcnt) > 99) Then
            txtAVRcnt = frmMain.AOdeepMAX
                MsgBox "누적횟수는 10 이상  99 이하입니다.", vbOKOnly
            Exit Sub
    End If
    
    SaveSetting App.Title, "Settings", "DeepMax", Val(txtAVRcnt)
    frmMain.AOdeepCNT = 1
    frmMain.AOdeepMAX = Val(txtAVRcnt)
    frmMain.txtAVRcnt = Val(txtAVRcnt)

    tmrCFG_update
End Sub


Private Sub cmdSetTYPE_Click()
Dim i
    For i = 0 To 14
        frmMain.ucSilo1(i).setScanTYPE CInt(txtCtypes(i))
    Next i
    
    tmrCFG_update
    
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim iLeft As Long
Dim iTop As Long
Dim TapIndex_base

    TapIndex_base = txtCangle(0).TabIndex

    For i = 1 To 14
        Load lbBinNO(i)
        Load txtCangle(i)
        
        iLeft = lbBinNO(0).Left + ((i) \ 3) * 1550
        iTop = lbBinNO(0).Top + ((i) Mod 3) * 350
        
        lbBinNO(i).Left = iLeft
        lbBinNO(i).Top = iTop
        
        txtCangle(i).TabIndex = TapIndex_base + i
    Next i

    For i = 0 To 14
        lbBinNO(i).Caption = "SILO-" & Format(i + 1, "00")
        
        txtCangle(i) = frmMain.ucSilo1(i).get_Angle
        
        iLeft = lbBinNO(0).Left + ((i) \ 3) * 1550
        iTop = lbBinNO(0).Top + ((i) Mod 3) * 350
        
        txtCangle(i).Left = iLeft + 740
        txtCangle(i).Top = iTop - 60
        
        lbBinNO(i).Visible = True
        txtCangle(i).Visible = True
    Next i

    For i = 15 To 18
        Load lbBinNO(i)
        Load txtCangle(i)
        
        iLeft = lbBinNO(0).Left + (i - 15) * 1550
        iTop = lbBinNO(0).Top + 3 * 350
        
        lbBinNO(i).Left = iLeft - 180
        lbBinNO(i).Top = iTop
        
        txtCangle(i).TabIndex = TapIndex_base + i
    Next i

    For i = 15 To 18
        lbBinNO(i).Caption = "S" & Format(19 + 15 - i, "00") & "-CTS" & Format(i - 15 + 1, "0")
        
        txtCangle(i) = frmMain.ucSilo1(i).get_Angle
        
        iLeft = lbBinNO(0).Left + (i - 15) * 1550
        iTop = lbBinNO(0).Top + 3 * 350
        
        txtCangle(i).Left = iLeft + 740
        txtCangle(i).Top = iTop - 60
        
        lbBinNO(i).Visible = True
        txtCangle(i).Visible = True
    Next i

    txtMaxHH = frmMain.txtMaxHH
    txtBaseHH = frmMain.txtBaseHH
    
    txtAVRcnt = frmMain.AOdeepMAX
    
    TapIndex_base = txtCtypes(0).TabIndex
    
    For i = 1 To 14
        Load lbBinNO2(i)
        Load txtCtypes(i)
        
        iLeft = lbBinNO2(0).Left + ((i) \ 3) * 1550
        iTop = lbBinNO2(0).Top + ((i) Mod 3) * 350
        
        lbBinNO2(i).Left = iLeft
        lbBinNO2(i).Top = iTop
        
        txtCtypes(i).TabIndex = TapIndex_base + i
    Next i

    For i = 0 To 14
        lbBinNO2(i).Caption = "SILO-" & Format(i + 1, "00")
        
        txtCtypes(i) = frmMain.ucSilo1(i).getScanTYPE
        
        iLeft = lbBinNO2(0).Left + ((i) \ 3) * 1550
        iTop = lbBinNO2(0).Top + ((i) Mod 3) * 350
        
        txtCtypes(i).Left = iLeft + 740
        txtCtypes(i).Top = iTop - 60
        
        lbBinNO2(i).Visible = True
        txtCtypes(i).Visible = True
    Next i
    
    For i = 15 To 18
        Load lbBinNO2(i)
        Load txtCtypes(i)
        
        iLeft = lbBinNO2(0).Left + (i - 15) * 1550
        iTop = lbBinNO2(0).Top + 3 * 350
        
        lbBinNO2(i).Left = iLeft - 180
        lbBinNO2(i).Top = iTop
        
        txtCtypes(i).TabIndex = TapIndex_base + i
    Next i

    For i = 15 To 18
        lbBinNO2(i).Caption = "S" & Format(19 + 15 - i, "00") & "-CTS" & Format(i - 15 + 1, "0")
        
        txtCtypes(i) = frmMain.ucSilo1(i).getScanTYPE
        
        iLeft = lbBinNO2(0).Left + (i - 15) * 1550
        iTop = lbBinNO2(0).Top + 3 * 350
        
        txtCtypes(i).Left = iLeft + 740
        txtCtypes(i).Top = iTop - 60
        
        lbBinNO2(i).Visible = True
        txtCtypes(i).Visible = True
    Next i
'
End Sub

Private Sub lbBaseHH__Click()
    tmrCFG_update
End Sub

Private Sub lbBinNO_Click(Index As Integer)
    tmrCFG_update
End Sub

Private Sub lbBinNO2_Click(Index As Integer)
'
    tmrCFG_update
    
    If frmSettings.Visible = True Then
        frmSettings.Show
    End If
'
    frmSettings.Init _
        Index _
        , frmMain.ucSilo1(Index).ipAddr _
        , frmMain.ucSilo1(Index).ipPort _
        , frmMain.ucSilo1(Index).CenterX _
        , frmMain.ucSilo1(Index).CenterY _
        , frmMain.ucSilo1(Index).Radius _
        , frmMain.ucSilo1(Index).TiltDefault _
        , frmMain.ucSilo1(Index).TiltMax _
        , frmMain.ucSilo1(Index).TiltMin _
        , frmMain.ucSilo1(Index).TiltStep
'
    frmSettings.Visible = True
'
End Sub

Private Sub lbMaxHH__Click()
    tmrCFG_update
End Sub

''Private Sub Form_LostFocus()
''        tmrCFG_update
''End Sub


Private Sub tmrCFG_Timer()

    tmrCFG.Enabled = False
    
    frmSettings.Visible = False
    frmCFG.Visible = False
    
End Sub

Private Sub txtAVRcnt_GotFocus()
    tmrCFG_update
End Sub

Private Sub txtBaseHH_GotFocus()
    tmrCFG_update
End Sub

Private Sub txtCangle_GotFocus(Index As Integer)
    tmrCFG_update
End Sub

Private Sub txtCtypes_GotFocus(Index As Integer)
    tmrCFG_update
End Sub

Private Sub txtMaxHH_GotFocus()
    tmrCFG_update
End Sub

Public Sub tmrCFG_update()
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
End Sub


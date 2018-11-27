VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin VB.TextBox txtBinTiltDefault 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   20
      Text            =   "0"
      Top             =   1275
      Width           =   555
   End
   Begin VB.TextBox txtBinTiltStep 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   17
      Text            =   "0"
      Top             =   2340
      Width           =   555
   End
   Begin VB.TextBox txtBinTiltMax 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   12
      Text            =   "0"
      Top             =   1620
      Width           =   555
   End
   Begin VB.TextBox txtBinTiltMin 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   11
      Text            =   "0"
      Top             =   1980
      Width           =   555
   End
   Begin VB.CommandButton cmdSettingsApply 
      BackColor       =   &H8000000A&
      Caption         =   "적 용"
      Height          =   375
      Left            =   3000
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdSettingsExit 
      BackColor       =   &H8000000A&
      Caption         =   "닫 기"
      Height          =   555
      Left            =   3000
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtBinRadius 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   5
      Text            =   "0"
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtBinCenterY 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   4
      Text            =   "0"
      Top             =   525
      Width           =   555
   End
   Begin VB.TextBox txtBinCenterX 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Text            =   "0"
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label8 
      Caption         =   "°, 48~-48"
      Height          =   195
      Left            =   1680
      TabIndex        =   22
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lbBinTiltDefault 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt default:"
      Height          =   195
      Left            =   0
      TabIndex        =   21
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label6 
      Caption         =   "°, 0.5~5.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   19
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label lbBinTiltStep 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt step:"
      Height          =   195
      Left            =   0
      TabIndex        =   18
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label lbBinTiltMax 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt max.:"
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label lbBinTiltMin 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt min.:"
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "°, 48.0~1.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "°, -48.0~-1.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   13
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "meter, 1.0~25.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "meter, -25.0~25.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "meter, -25.0~25.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbBinRadius 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Radius:"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lbBinCenterY 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Center Y:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lbBinCenterX 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Center X:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim Index%

Public Sub Init(Index_I%, CenterX_I!, CenterY_I!, Radius_I!, _
    TiltDefault_I%, TiltMax_I!, TiltMin_I!, TiltStep_I!)
'
    Index = Index_I
'
    frmSettings.Caption = "SILO-" & Format(Index + 1, "00") & " Settings"
'
    txtBinCenterX = CenterX_I
    txtBinCenterY = CenterY_I
    txtBinRadius = Radius_I
    txtBinTiltDefault = TiltDefault_I
    txtBinTiltMax = TiltMax_I
    txtBinTiltMin = TiltMin_I
    txtBinTiltStep = TiltStep_I
'
End Sub

Private Sub cmdSettingsApply_Click()
'
    If IsNumeric(txtBinCenterX) = False _
        Or Abs(CSng(Val(txtBinCenterX))) > 25! _
        Then
        MsgBox lbBinCenterX & "는 -25.0 ~ 25.0 사이의 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
    If IsNumeric(txtBinCenterY) = False _
        Or Abs(CSng(Val(txtBinCenterY))) > 25! _
        Then
        MsgBox lbBinCenterY & "는 -25.0 ~ 25.0 사이의 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
    If IsNumeric(txtBinRadius) = False _
        Or CSng(Val(txtBinRadius)) < 1! _
        Or CSng(Val(txtBinRadius)) > 25! _
        Then
        MsgBox lbBinRadius & "는 1.0 ~ 25.0 사이의 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
    If IsNumeric(txtBinTiltDefault) = False _
        Or CSng(CInt(Val(txtBinTiltDefault))) <> CSng(Val(txtBinTiltDefault)) _
        Or CInt(Val(txtBinTiltDefault)) > 48! Or CInt(Val(txtBinTiltDefault)) < -48! _
        Then
        MsgBox lbBinTiltDefault & "는 48 ~ -48 사이의 정수 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
    If IsNumeric(txtBinTiltMax) = False _
        Or CSng(Val(txtBinTiltMax)) > 48! Or CSng(Val(txtBinTiltMax)) < 1! _
        Then
        MsgBox lbBinTiltMax & "는 48.0 ~ 1.0 사이의 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
    If IsNumeric(txtBinTiltMin) = False _
        Or CSng(Val(txtBinTiltMin)) < -48! Or CSng(Val(txtBinTiltMin)) > -1! _
        Then
        MsgBox lbBinTiltMin & "는 -48.0 ~ -1.0 사이의 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
    If CSng(Val(txtBinTiltMax)) <= CSng(Val(txtBinTiltMin)) Then
        MsgBox lbBinTiltMax & "는 " & lbBinTiltMin & "보다 큰 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
    If IsNumeric(txtBinTiltStep) = False _
        Or CSng(Val(txtBinTiltStep)) > 5! Or CSng(Val(txtBinTiltStep)) < 0.5! _
        Then
        MsgBox lbBinTiltStep & "는 0.5 ~ 5.0 사이의 값 이어야 합니다.", vbOKOnly
        Exit Sub
    End If
'
    SaveSetting App.Title, "Settings", "SILOcenterX_" & Format(Index + 1, "00") _
        , txtBinCenterX
    SaveSetting App.Title, "Settings", "SILOcenterY_" & Format(Index + 1, "00") _
        , txtBinCenterY
    SaveSetting App.Title, "Settings", "SILOradius_" & Format(Index + 1, "00") _
        , txtBinRadius
    SaveSetting App.Title, "Settings", "SILOtiltDefault_" & Format(Index + 1, "00") _
        , txtBinTiltDefault
    SaveSetting App.Title, "Settings", "SILOtiltMax_" & Format(Index + 1, "00") _
        , txtBinTiltMax
    SaveSetting App.Title, "Settings", "SILOtiltMin_" & Format(Index + 1, "00") _
        , txtBinTiltMin
    SaveSetting App.Title, "Settings", "SILOtiltStep_" & Format(Index + 1, "00") _
        , txtBinTiltStep
'
    frmMain.ucSilo1(Index).setBinSettings _
        txtBinCenterX, txtBinCenterY, txtBinRadius _
        , txtBinTiltDefault, txtBinTiltMax, txtBinTiltMin, txtBinTiltStep
'
End Sub

Private Sub cmdSettingsExit_Click()
'
    frmSettings.Visible = False
'
End Sub

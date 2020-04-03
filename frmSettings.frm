VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin VB.TextBox txtBinIPAddr 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Text            =   "255.255.255.255"
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox txtBinIPPort 
      Height          =   270
      Left            =   1080
      TabIndex        =   1
      Text            =   "99999"
      Top             =   465
      Width           =   555
   End
   Begin VB.TextBox txtBinTiltDefault 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   5
      Text            =   "0"
      Top             =   1950
      Width           =   555
   End
   Begin VB.TextBox txtBinTiltStep 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   8
      Text            =   "0"
      Top             =   3015
      Width           =   555
   End
   Begin VB.TextBox txtBinTiltMax 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   6
      Text            =   "0"
      Top             =   2295
      Width           =   555
   End
   Begin VB.TextBox txtBinTiltMin 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   7
      Text            =   "0"
      Top             =   2655
      Width           =   555
   End
   Begin VB.CommandButton cmdSettingsApply 
      BackColor       =   &H8000000A&
      Caption         =   "적 용"
      Height          =   375
      Left            =   3000
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   2235
      Width           =   1575
   End
   Begin VB.CommandButton cmdSettingsExit 
      BackColor       =   &H8000000A&
      Caption         =   "닫 기"
      Height          =   555
      Left            =   3000
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   2715
      Width           =   1575
   End
   Begin VB.TextBox txtBinRadius 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   4
      Text            =   "0"
      Top             =   1575
      Width           =   555
   End
   Begin VB.TextBox txtBinCenterY 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Text            =   "0"
      Top             =   1200
      Width           =   555
   End
   Begin VB.TextBox txtBinCenterX 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Text            =   "0"
      Top             =   855
      Width           =   555
   End
   Begin VB.Label lbBinIPAddr 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "IP addr:"
      Height          =   195
      Left            =   0
      TabIndex        =   28
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lbBinIPPort 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "IP port:"
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lbBinIPAddr_ 
      Caption         =   "Serial2Net의 IP address"
      Height          =   195
      Left            =   2520
      TabIndex        =   26
      Top             =   180
      Width           =   2295
   End
   Begin VB.Label lbBinIPPort_ 
      Caption         =   "Serial2Net의 IP port number"
      Height          =   195
      Left            =   1680
      TabIndex        =   25
      Top             =   540
      Width           =   3075
   End
   Begin VB.Label lbBinTiltDefault_ 
      Caption         =   "°, 48~-48"
      Height          =   195
      Left            =   1680
      TabIndex        =   24
      Top             =   1995
      Width           =   1275
   End
   Begin VB.Label lbBinTiltDefault 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt default:"
      Height          =   195
      Left            =   0
      TabIndex        =   23
      Top             =   1995
      Width           =   1035
   End
   Begin VB.Label lbBinTiltStep_ 
      Caption         =   "°, 0.5~5.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   22
      Top             =   3075
      Width           =   1275
   End
   Begin VB.Label lbBinTiltStep 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt step:"
      Height          =   195
      Left            =   0
      TabIndex        =   21
      Top             =   3075
      Width           =   1035
   End
   Begin VB.Label lbBinTiltMax 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt max.:"
      Height          =   195
      Left            =   0
      TabIndex        =   20
      Top             =   2355
      Width           =   1035
   End
   Begin VB.Label lbBinTiltMin 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Tilt min.:"
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   2715
      Width           =   1035
   End
   Begin VB.Label lbBinTiltMax_ 
      Caption         =   "°, 48.0~1.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   18
      Top             =   2355
      Width           =   1275
   End
   Begin VB.Label lbBinTiltMin_ 
      Caption         =   "°, -48.0~-1.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   2715
      Width           =   1275
   End
   Begin VB.Label lbBinRadius_ 
      Caption         =   "meter, 1.0~25.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   1635
      Width           =   1635
   End
   Begin VB.Label lbBinCenterY_ 
      Caption         =   "meter, -25.0~25.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   15
      Top             =   1275
      Width           =   1755
   End
   Begin VB.Label lbBinCenterX_ 
      Caption         =   "meter, -25.0~25.0"
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   915
      Width           =   1695
   End
   Begin VB.Label lbBinRadius 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Radius:"
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   1635
      Width           =   1035
   End
   Begin VB.Label lbBinCenterY 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Center Y:"
      Height          =   195
      Left            =   0
      TabIndex        =   12
      Top             =   1275
      Width           =   1035
   End
   Begin VB.Label lbBinCenterX 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Center X:"
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   915
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
Dim orgBinIPAddr$, orgBinIPPort$
Dim orgBinCenterX$, orgBinCenterY$, orgBinRadius$
Dim orgBinTiltDefault$, orgBinTiltMax$, orgBinTiltMin$, orgBinTiltStep$

Public Sub Init(Index_I%, BinIPAddr_I$, BinIPPort_I$, CenterX_I!, CenterY_I!, Radius_I!, _
    TiltDefault_I%, TiltMax_I!, TiltMin_I!, TiltStep_I!)
'
    Index = Index_I
'
    If Index < 15 Then
        frmSettings.Caption = "SILO-" & Format(Index + 1, "00") & " Settings"
    Else
        frmSettings.Caption = "S" & Format(19 + 15 - Index, "00") & "-CTS" & Format(Index - 15 + 1, "0") & " Settings"
    End If
'
    orgBinIPAddr = BinIPAddr_I
    txtBinIPAddr = BinIPAddr_I
'
    orgBinIPPort = BinIPPort_I
    txtBinIPPort = BinIPPort_I
'
    orgBinCenterX = CenterX_I
    txtBinCenterX = CenterX_I
'
    orgBinCenterY = CenterY_I
    txtBinCenterY = CenterY_I
'
    orgBinRadius = Radius_I
    txtBinRadius = Radius_I
'
    orgBinTiltDefault = TiltDefault_I
    txtBinTiltDefault = TiltDefault_I
'
    orgBinTiltMax = TiltMax_I
    txtBinTiltMax = TiltMax_I
'
    orgBinTiltMin = TiltMin_I
    txtBinTiltMin = TiltMin_I
'
    orgBinTiltStep = TiltStep_I
    txtBinTiltStep = TiltStep_I
'
End Sub

Private Sub cmdSettingsApply_Click()
'
    Dim IsValid As Boolean
'
    frmCFG.tmrCFG_update
'
    IsValid = False
'
    If txtBinIPAddr <> orgBinIPAddr Then
        If IsValidIPAddress(txtBinIPAddr) = False Then
            MsgBox lbBinIPAddr & "는 IPv4(ex. 192.168.0.1) 형태의 값 이어야 합니다.", vbOKOnly
        Else
            orgBinIPAddr = txtBinIPAddr
            SaveSetting App.Title, "Settings", "BinIPAddr_" & Index, txtBinIPAddr
            IsValid = True
        End If
    End If
    If txtBinIPPort <> orgBinIPPort Then
        If IsValidIPPort(txtBinIPPort) = False Then
            MsgBox lbBinIPPort & "는 1024 ~ 65535 사이의 정수 값 이어야 합니다.", vbOKOnly
        Else
            orgBinIPPort = txtBinIPPort
            SaveSetting App.Title, "Settings", "BinIPPort_" & Index, txtBinIPPort
            IsValid = True
        End If
    End If
'
    If (IsValid = True) Then
        frmMain.ucSilo1(Index).setIDX Index, txtBinIPAddr, txtBinIPPort
        frmMain.ucSilo1(Index).initStart
        'IsApplied = True
    End If
'
    IsValid = False
'
    If txtBinCenterX <> orgBinCenterX Then
        If IsNumeric(txtBinCenterX) = False _
            Or Abs(CSng(Val(txtBinCenterX))) > 25! _
            Then
            MsgBox lbBinCenterX & "는 -25.0 ~ 25.0 사이의 값 이어야 합니다.", vbOKOnly
        Else
            orgBinCenterX = txtBinCenterX
            SaveSetting App.Title, "Settings", "SILOcenterX_" & Format(Index + 1, "00") _
                , txtBinCenterX
            IsValid = True
        End If
    End If
    If txtBinCenterY <> orgBinCenterY Then
        If IsNumeric(txtBinCenterY) = False _
            Or Abs(CSng(Val(txtBinCenterY))) > 25! _
            Then
            MsgBox lbBinCenterY & "는 -25.0 ~ 25.0 사이의 값 이어야 합니다.", vbOKOnly
        Else
            orgBinCenterY = txtBinCenterY
            SaveSetting App.Title, "Settings", "SILOcenterY_" & Format(Index + 1, "00") _
                , txtBinCenterY
            IsValid = True
        End If
    End If
    If txtBinRadius <> orgBinRadius Then
        If IsNumeric(txtBinRadius) = False _
            Or CSng(Val(txtBinRadius)) < 1! _
            Or CSng(Val(txtBinRadius)) > 25! _
            Then
            MsgBox lbBinRadius & "는 1.0 ~ 25.0 사이의 값 이어야 합니다.", vbOKOnly
        Else
            orgBinRadius = txtBinRadius
            SaveSetting App.Title, "Settings", "SILOradius_" & Format(Index + 1, "00") _
                , txtBinRadius
            IsValid = True
        End If
    End If
    If txtBinTiltDefault <> orgBinTiltDefault Then
        If IsNumeric(txtBinTiltDefault) = False _
            Or CSng(CInt(Val(txtBinTiltDefault))) <> CSng(Val(txtBinTiltDefault)) _
            Or CInt(Val(txtBinTiltDefault)) > 48! Or CInt(Val(txtBinTiltDefault)) < -48! _
            Then
            MsgBox lbBinTiltDefault & "는 48 ~ -48 사이의 정수 값 이어야 합니다.", vbOKOnly
        Else
            orgBinTiltDefault = txtBinTiltDefault
            SaveSetting App.Title, "Settings", "SILOtiltDefault_" & Format(Index + 1, "00") _
                , txtBinTiltDefault
            IsValid = True
        End If
    End If
    If txtBinTiltMax <> orgBinTiltMax Or txtBinTiltMin <> orgBinTiltMin Then
        If IsNumeric(txtBinTiltMax) = False _
            Or CSng(Val(txtBinTiltMax)) > 48! Or CSng(Val(txtBinTiltMax)) < 1! _
            Then
            MsgBox lbBinTiltMax & "는 48.0 ~ 1.0 사이의 값 이어야 합니다.", vbOKOnly
        ElseIf IsNumeric(txtBinTiltMin) = False _
            Or CSng(Val(txtBinTiltMin)) < -48! Or CSng(Val(txtBinTiltMin)) > -1! _
            Then
            MsgBox lbBinTiltMin & "는 -48.0 ~ -1.0 사이의 값 이어야 합니다.", vbOKOnly
        ElseIf CSng(Val(txtBinTiltMax)) <= CSng(Val(txtBinTiltMin)) Then
            MsgBox lbBinTiltMax & "는 " & lbBinTiltMin & "보다 큰 값 이어야 합니다.", vbOKOnly
        Else
            orgBinTiltMax = txtBinTiltMax
            orgBinTiltMin = txtBinTiltMin
            SaveSetting App.Title, "Settings", "SILOtiltMax_" & Format(Index + 1, "00") _
                , txtBinTiltMax
            SaveSetting App.Title, "Settings", "SILOtiltMin_" & Format(Index + 1, "00") _
                , txtBinTiltMin
            IsValid = True
        End If
    End If
    If txtBinTiltStep <> orgBinTiltStep Then
        If IsNumeric(txtBinTiltStep) = False _
            Or CSng(Val(txtBinTiltStep)) > 5! Or CSng(Val(txtBinTiltStep)) < 0.5! _
            Then
            MsgBox lbBinTiltStep & "는 0.5 ~ 5.0 사이의 값 이어야 합니다.", vbOKOnly
        Else
            orgBinTiltStep = txtBinTiltStep
            SaveSetting App.Title, "Settings", "SILOtiltStep_" & Format(Index + 1, "00") _
                , txtBinTiltStep
            IsValid = True
        End If
    End If
'
    If (IsValid = True) Then
        frmMain.ucSilo1(Index).setBinSettings _
            txtBinCenterX, txtBinCenterY, txtBinRadius _
            , txtBinTiltDefault, txtBinTiltMax, txtBinTiltMin, txtBinTiltStep
    End If
'
End Sub

Private Sub cmdSettingsExit_Click()
'
    frmCFG.tmrCFG_update
'
    frmSettings.Visible = False
'
End Sub

Private Sub Form_Load()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinCenterX__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinCenterX_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinCenterY__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinCenterY_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinIPAddr__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinIPAddr_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinIPPort__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinIPPort_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinRadius__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinRadius_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltDefault__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltDefault_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltMax__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltMax_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltMin__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltMin_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltStep__Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub lbBinTiltStep_Click()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinCenterX_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinCenterY_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinIPAddr_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinIPPort_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinRadius_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinTiltDefault_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinTiltMax_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinTiltMin_GotFocus()
    frmCFG.tmrCFG_update
End Sub

Private Sub txtBinTiltStep_GotFocus()
    frmCFG.tmrCFG_update
End Sub

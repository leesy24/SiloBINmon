VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{40DD8EA0-284B-11D0-A7B0-0020AFF929F4}#2.3#0"; "Adsocx.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   BorderStyle     =   0  '없음
   Caption         =   "BIN5_Monitor"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16230
   FillStyle       =   0  '단색
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   16230
   ShowInTaskbar   =   0   'False
   Begin ADSOCXLib.AdsOcx AdsOcx1 
      Left            =   960
      Top             =   1320
      _Version        =   131074
      _ExtentX        =   900
      _ExtentY        =   953
      _StockProps     =   0
      AdsAmsServerNetId=   ""
      AdsAmsClientPort=   33017
      AdsClientType   =   ""
      AdsClientAdsState=   ""
      AdsClientAdsControl=   ""
   End
   Begin ADSOCXLib.AdsOcx AdsOcx2 
      Left            =   1860
      Top             =   1380
      _Version        =   131074
      _ExtentX        =   900
      _ExtentY        =   953
      _StockProps     =   0
      AdsAmsServerNetId=   ""
      AdsAmsClientPort=   33019
      AdsClientType   =   ""
      AdsClientAdsState=   ""
      AdsClientAdsControl=   ""
   End
   Begin VB.TextBox txtPcsPort2 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   4200
      TabIndex        =   31
      Text            =   "8009"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtWSpcs2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3960
      TabIndex        =   30
      Top             =   840
      Width           =   180
   End
   Begin VB.Frame frScale 
      Appearance      =   0  '평면
      BackColor       =   &H00008000&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11040
      TabIndex        =   22
      Top             =   840
      Width           =   3495
      Begin VB.TextBox txtBaseHH 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H000040C0&
         Height          =   270
         Left            =   1500
         TabIndex        =   26
         Text            =   "100"
         Top             =   60
         Width           =   495
      End
      Begin VB.TextBox txtMaxHH 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H000040C0&
         Height          =   270
         Left            =   2760
         TabIndex        =   25
         Text            =   "5000"
         Top             =   60
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   195
         Left            =   2220
         TabIndex        =   24
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "기준높이: 0%"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   195
         Left            =   60
         TabIndex        =   23
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Timer tmrDRAWmode 
      Interval        =   5000
      Left            =   6960
      Top             =   840
   End
   Begin VB.CommandButton cmdDRAWmode 
      BackColor       =   &H0000C000&
      Caption         =   "3D/2D 보기"
      Height          =   300
      Left            =   14640
      MaskColor       =   &H00E0E0E0&
      Style           =   1  '그래픽
      TabIndex        =   21
      Top             =   840
      Width           =   1395
   End
   Begin prjBIN5mon.ucSilo ucSilo1 
      Height          =   4215
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7435
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H00008000&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9000
      TabIndex        =   16
      Top             =   840
      Width           =   1875
      Begin VB.TextBox txtAVRcnt 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H000040C0&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "99"
         Top             =   60
         Width           =   555
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "누적횟수:"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdPcsRun 
      BackColor       =   &H00008000&
      Caption         =   "PcsRUN"
      Height          =   255
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtPcsPort 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1800
      TabIndex        =   14
      Text            =   "8005"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtPcsIP 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   480
      TabIndex        =   13
      Text            =   "127.0.0.1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtWSpcs 
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   180
   End
   Begin VB.CommandButton cmdDmon 
      Caption         =   "dMon"
      Height          =   255
      Left            =   3540
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboIDX 
      Height          =   300
      Left            =   2580
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSD1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   6.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   8
      Top             =   6900
      Width           =   10575
   End
   Begin VB.CommandButton cmdADSclr 
      Caption         =   "ADSclr"
      Height          =   255
      Left            =   5100
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdADS1 
      Caption         =   "ADS1"
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   15315
      TabIndex        =   0
      Top             =   120
      Width           =   15375
      Begin VB.Timer tmrPcs2 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   14400
         Top             =   0
      End
      Begin VB.Timer tmrAoDo2 
         Interval        =   1000
         Left            =   13920
         Top             =   0
      End
      Begin VB.CommandButton cmdCFG 
         BackColor       =   &H00008000&
         Caption         =   "설정"
         Height          =   375
         Left            =   10560
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   240
         Width           =   915
      End
      Begin VB.Timer tmrSinit 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   5760
         Top             =   360
      End
      Begin VB.Timer tmrPcs 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   7800
         Top             =   360
      End
      Begin MSWinsockLib.Winsock wsPcs 
         Left            =   8220
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrAoDo 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7320
         Top             =   360
      End
      Begin VB.CommandButton cmdRunStop 
         BackColor       =   &H00008000&
         Caption         =   "RUN/STOP"
         Height          =   375
         Left            =   9120
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Timer tmrINIT 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   6240
         Top             =   360
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00808080&
         Caption         =   "종 료"
         Height          =   375
         Left            =   12840
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdHide 
         BackColor       =   &H00808080&
         Caption         =   "화면감추기"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11520
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin MSWinsockLib.Winsock wsPcs2 
         Left            =   14880
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lbRelVersion 
         BackStyle       =   0  '투명
         Caption         =   "Release version"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   1440
         TabIndex        =   33
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label lbRelDate 
         BackStyle       =   0  '투명
         Caption         =   "Release date"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label lbUpTime 
         BackStyle       =   0  '투명
         Caption         =   "UpTime"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label lbTeam2 
         BackColor       =   &H00808080&
         Caption         =   "DASAN-InfoTek"
         BeginProperty Font 
            Name            =   "바탕체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lbNow 
         BackStyle       =   0  '투명
         Caption         =   "UpTime"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4380
         TabIndex        =   20
         Top             =   420
         Width           =   3015
      End
      Begin VB.Label lbTitle 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "[SILO] BIN LEVEL MONITORING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   705
         Left            =   4080
         TabIndex        =   4
         Top             =   0
         Width           =   9195
      End
      Begin VB.Image imgLogo1 
         BorderStyle     =   1  '단일 고정
         Height          =   495
         Left            =   120
         Picture         =   "frmMain.frx":16AC2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label lbTeam 
         BackColor       =   &H00808080&
         Caption         =   "(주)제일시스템"
         BeginProperty Font 
            Name            =   "바탕체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label lbVS1 
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===========================================================================================
'
'                       2D LEVEL Monitoring System
'                       for BIN5 with SICK LMS-211
'
'                                   BIN5mon V1.00
'
'===========================================================================================


Option Explicit

Private Const relVersion = "v2.00.02_3D"
Private Const relDate = "2018-11-26"

Dim d1 As Single


Dim ipAddr(20) As String  ''19  ''11  '''''''''''''''''New-CTS-Silo(15+4)!!  ''4x2==8''???
Dim ipPort(20) As String  ''19  ''11

Dim AOdata(33) As Integer       ''미분광: use only 0~7

Dim AOdeep(20, 100) As Integer   ''''''''''''''''''''''New-CTS-Silo(15+4)!!
'''Dim AOdeep(15, 100) As Integer   ''미분광!!
Public AOdeepCNT As Integer
Public AOdeepMAX As Integer        ''<=MAX:99
Dim AOdeepFull As Boolean

Public AOdeepCNT2 As Integer  ''''''''''''''''''''''New-CTS-Silo(15+4)!!
Dim AOdeepFull2 As Boolean

Private Declare Function GetModuleFileNameW Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long

Private Function GetEXEName() As String
    Const MAX_PATH = 260&
    
    GetEXEName = Space$(MAX_PATH - 1&)
    GetEXEName = Left$(GetEXEName, GetModuleFileNameW(0&, StrPtr(GetEXEName), MAX_PATH))
    GetEXEName = Right$(GetEXEName, Len(GetEXEName) - InStrRev(GetEXEName, "\"))
End Function

Private Sub cmdADS1_Click()

Dim ioD(33) As Integer
Dim i As Long
    
    tmrAoDo.Enabled = False  '''for Test!!

    For i = 0 To 32
        ioD(i) = 0
    Next i

'''SILO'''
''''''''''[BackHoff]-IO-MAP
''''-----------------------------------------------
''''    0(1)    4(10)   8(5)    12(14)  16(9)   20
''''                            ------
''''    1(4)    5(13)   9(8)    13(3)   17(12)  21
''''           ------
''''    2(7)    6(2)    10(11)  14(6)   18(15)  22
''''                                    ------
''''    3       7       11      15      19      23
''''-----------------------------------------------
''''

    ioD(0) = 1 * 2048
    ioD(1) = 2 * 2048
    ioD(2) = 4 * 2048
    ioD(3) = 8 * 2048

    ioD(4) = 10 * 2048
    ioD(5) = 12 * 2048
    ioD(6) = 14 * 2048
    ioD(7) = 15 * 2048

    ioD(8) = 1 * 2048
    ioD(9) = 2 * 2048
    ioD(10) = 4 * 2048
    ioD(11) = 8 * 2048

    ioD(12) = 10 * 2048
    ioD(13) = 12 * 2048
    ioD(14) = 14 * 2048
    ioD(15) = 15 * 2048

    ioD(16) = 0
    ioD(17) = 0
    ioD(18) = 0
    ioD(19) = 0


''    ioD(0) = 1 * 2048
''    ioD(1) = 32767      ''(4)
''    ioD(2) = 32768 / 2  ''(7)
''    ioD(3) = 0

''Port: 800, IGrp: 0xF020, IOffs: 0x64, Len: 2


    ''AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 40, ioD   ''SILO:[40]==4*5=20channel!
    
    AdsOcx2.AdsSyncWriteReq &HF020&, &H64&, 32, ioD  '''4x4==16x2==32

End Sub


Private Sub cmdADSclr_Click()
Dim i As Integer
Dim d As Integer

   
    Dim ioD(33) As Integer  ''(0~31)
    For i = 0 To 31
        ioD(i) = 0
    Next i
    
'    AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 40, ioD   ''SILO:[40]==4*5=20channel!
'
'    tmrAoDo.Enabled = True  ''for test!!
    
End Sub


Private Sub cmdCFG_Click()

    frmCFG.txtMaxHH = frmMain.txtMaxHH
    frmCFG.txtBaseHH = frmMain.txtBaseHH

    If frmCFG.Visible = True Then
        frmCFG.Show
    Else
        frmCFG.Visible = True
    End If
    
''    frmCFG.tmrCFG.Interval = 5000
''    frmCFG.tmrCFG.Enabled = True

End Sub


Private Sub Form_Click()
        frmCFG.tmrCFG.Enabled = False
        frmCFG.tmrCFG.Interval = 5000
        frmCFG.tmrCFG.Enabled = True
End Sub

Private Sub Form_DblClick()
        frmCFG.tmrCFG.Enabled = False
        frmCFG.tmrCFG.Interval = 5000
        frmCFG.tmrCFG.Enabled = True
End Sub

Private Sub tmrDRAWmode_Timer()
    tmrDRAWmode.Enabled = False
    
    cmdDRAWmode.BackColor = vbGreen
    
    Dim i As Integer
    For i = 0 To 18  '''14  '''New-CTS-Silo(15+4)!
        ucSilo1(i).set_DRAWmode 0
    Next i
    
End Sub

Private Sub cmdDRAWmode_Click()

Dim i As Integer
        
    If cmdDRAWmode.BackColor = vbGreen Then
        cmdDRAWmode.BackColor = &H808080
        ''
        tmrDRAWmode.Interval = 5000
        tmrDRAWmode.Enabled = True
        
        For i = 0 To 18  '''14  '''New-CTS-Silo(15+4)!
            ucSilo1(i).set_DRAWmode 1
        Next i
    Else
            ''cmdFilt.BackColor = vbGreen
        If tmrDRAWmode.Enabled = True Then
            tmrDRAWmode.Enabled = False
            tmrDRAWmode.Interval = 10000
            tmrDRAWmode.Enabled = True
            
            For i = 0 To 18  '''14  '''New-CTS-Silo(15+4)!
                ucSilo1(i).set_DRAWmode 1
            Next i
        End If
    End If
End Sub


Private Sub cmdExit_Click()

Dim ret1
    ret1 = MsgBox("종료하면 모든 기능이 정지됩니다." & vbCrLf & "정말 종료 하시겠습니까?", vbYesNo)

    If ret1 <> vbYes Then
        Exit Sub
    End If

    End

End Sub



Private Sub cmdHide_Click()

    ''frmMain.Visible = False
    frmMain.Hide
    
    
End Sub

Private Sub cmdPcsRun_Click()
    ''&H00008000& ''G
    ''&H00000080& ''R
    ''QBColor
  Dim i As Integer
  
    If cmdPcsRun.BackColor = &H8000& Then  ''run
        wsPcs.Close
        wsPcs2.Close
        '''
        tmrPcs.Enabled = False
        tmrPcs2.Enabled = False
        '''
        cmdPcsRun.BackColor = &H80&        ''stop
    Else  ''stop
        tmrPcs.Enabled = True
        tmrPcs2.Enabled = True
        '''
        cmdPcsRun.BackColor = &H8000&        ''run
    End If

End Sub

Private Sub cmdRunStop_Click()

    ''&H00008000& ''G
    ''&H00000080& ''R
    ''QBColor
    
  Dim i As Integer

    If cmdRunStop.BackColor = &H8000& Then  ''run
        For i = 0 To 18  '''14  '''New-CTS-Silo(15+4)!
            ucSilo1(i).scan_STOP:   DoEvents
            ucSilo1(i).scan_STOP:   DoEvents
            ucSilo1(i).scan_STOP:   DoEvents
        Next i
        cmdRunStop.BackColor = &H80&        ''stop
        
        txtMaxHH.Enabled = True
        txtBaseHH.Enabled = True
        
    Else  ''stop
        For i = 0 To 18  '''14  '''New-CTS-Silo(15+4)!
            ucSilo1(i).set_maxHH CLng(txtMaxHH)
            ucSilo1(i).set_baseHH CLng(txtBaseHH)

            ucSilo1(i).scan_RUN:   DoEvents
        Next i
        cmdRunStop.BackColor = &H8000&        ''run
        
        txtMaxHH.Enabled = False
        txtBaseHH.Enabled = False

        SaveSetting App.Title, "Settings", "MaxHH", Trim(txtMaxHH.Text)
        SaveSetting App.Title, "Settings", "BaseHH", Trim(txtBaseHH.Text)


        tmrINIT.Interval = 1000  ''5000
        tmrINIT.Enabled = True

    End If

        
End Sub



Private Sub Form_Load()

Dim i As Integer
Dim j As Integer

    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행되었습니다."
       Unload Me
       End
    End If
    
    If (Screen.Width <> 28800) Or (Screen.Height <> 16200) Then
       MsgBox "화면해상도 [1920x1080]WIDE 이상에서만 실행합니다."
       Unload Me
       End
    End If
    
    lbUpTime.Caption = "UpTime: " & Format(Now, "YYYY-MM-DD h:m:s")
    
    frmMain.AutoRedraw = True

'    Me.Width = Screen.Width * (1280 / 1400)
'    Me.Height = Screen.Height * (1024 / 1050)

'    Me.Left = Screen.Width - Width
'    Me.Top = 0
'    frmMain.Move Screen.Width - Width, 0
    
    
    frmMain.Move 0, 0, Screen.Width, Screen.Height
    
    Debug.Print Screen.Width, Screen.Height

    DGPSLog vbCrLf, "SILO"
    DGPSLog " ====[SILO BIN-LEVEL START]=== " & GetEXEName() & vbCrLf, "SILO"

    AOdeepMAX = GetSetting(App.Title, "Settings", "DeepMax", 60)
    If AOdeepMAX < 10 Then AOdeepMAX = 10
    If AOdeepMAX > 99 Then AOdeepMAX = 99


    AOdeepFull = False
    ''AOdeepMAX = 60  ''30  ''''''''MAX:99
    
    AOdeepCNT = 0
    AOdeepCNT2 = 0
    
    For i = 0 To 18  '''''''''''''''''''''''''''''''''''''''''''''14   '''New-CTS
        For j = 0 To 99  ''AOdeepMAX
            AOdeep(i, j) = 0
        Next j
    Next i


    txtSD1.Left = (Width / 5) * 4 + 100  '''Width - 5000   '''100
    txtSD1.Top = Height - 3500
    txtSD1.Width = 5500   '''Width - 200
    txtSD1.Height = 3300


    picTop.Left = 100
    picTop.Top = 100
    picTop.Height = 700   '''Height * 0.05 + 100
    picTop.Width = Width - 200
    ''''
        imgLogo1.Left = 100
        imgLogo1.Top = 100 ''100
        lbTitle.Left = (Width * 0.32)    ''+ 200  ''frTop.Width * 0.3
        lbTitle.Top = 50
        lbTitle.Height = 600
        lbTitle.Width = (Width * 0.5) - 500
        ''
        cmdExit.Top = 200
        cmdExit.Left = picTop.Width - 1200
        cmdHide.Top = 200
        cmdHide.Left = picTop.Width - 2600
        cmdRunStop.Top = 200
        cmdRunStop.Left = picTop.Width - 4000
        
        cmdCFG.Top = 200
        cmdCFG.Left = picTop.Width - 5000
        
        lbRelVersion.Top = 200
        lbRelVersion.Left = picTop.Width - 6050
        lbRelVersion = relVersion
        lbRelDate.Top = 400
        lbRelDate.Left = picTop.Width - 6050
        lbRelDate = relDate
        
    For i = 0 To 32
        AOdata(i) = 0
    Next i


    For i = 1 To 14  ''3  '''10
        Load ucSilo1(i)
    Next i

    For i = 0 To 14  ''"SILO*15"

        ucSilo1(i).Width = 5500 ''3500  ''Width / 11 - 30
        ucSilo1(i).Height = 3700  ''4000 ''3500

        ucSilo1(i).Left = ((i \ 3) * (Width / 5)) + 150  ''(i * (Width / 11)) + 20
        ucSilo1(i).Top = ((i Mod 3) * 3700) + 1200  ''(i * (Width / 11)) + 20

        ucSilo1(i).setIDX i, "", ""
        ucSilo1(i).Visible = True
    Next i

    ipAddr(0) = "192.168.0.71": ipPort(0) = "7001"
    ipAddr(1) = "192.168.0.71": ipPort(1) = "7002"
    ipAddr(2) = "192.168.0.71": ipPort(2) = "9003"  '''"7003"
    ipAddr(3) = "192.168.0.71": ipPort(3) = "7004"
    ipAddr(4) = "192.168.0.71": ipPort(4) = "7005"
    ipAddr(5) = "192.168.0.71": ipPort(5) = "7006"
    ipAddr(6) = "192.168.0.71": ipPort(6) = "7007"
    ipAddr(7) = "192.168.0.71": ipPort(7) = "7008"
    ''
    ipAddr(8) = "192.168.0.72": ipPort(8) = "7001"
    ipAddr(9) = "192.168.0.72": ipPort(9) = "7002"
    ipAddr(10) = "192.168.0.72": ipPort(10) = "7003"
    ipAddr(11) = "192.168.0.72": ipPort(11) = "7004"
    ipAddr(12) = "192.168.0.72": ipPort(12) = "7005"
    ipAddr(13) = "192.168.0.72": ipPort(13) = "7006"
    ipAddr(14) = "192.168.0.72": ipPort(14) = "7007"
    
    Dim typeTmp As Integer
    Dim centerXTmp$, centerYTmp$, radiusTmp$
    Dim tiltDefaultTmp$, tiltMaxTmp$, tiltMinTmp$, tiltStepTmp$
    
    For i = 0 To 14
        ucSilo1(i).setIDX i, ipAddr(i), ipPort(i)
        ''
        typeTmp = Trim(Str(GetSetting(App.Title, "Settings", "SILOtypes_" & Format(i + 1, "00"), 3100)))
        ucSilo1(i).setScanTYPE typeTmp  ''3100  '''LD-LRS-3100,, DPS-2590
        centerXTmp = _
            GetSetting(App.Title, "Settings", "SILOcenterX_" & Format(i + 1, "00"), "Fail")
        centerYTmp = _
            GetSetting(App.Title, "Settings", "SILOcenterY_" & Format(i + 1, "00"), "Fail")
        radiusTmp = _
            GetSetting(App.Title, "Settings", "SILOradius_" & Format(i + 1, "00"), "Fail")
        tiltDefaultTmp = _
            GetSetting(App.Title, "Settings", "SILOtiltDefault_" & Format(i + 1, "00"), "Fail")
        tiltMaxTmp = _
            GetSetting(App.Title, "Settings", "SILOtiltMax_" & Format(i + 1, "00"), "Fail")
        tiltMinTmp = _
            GetSetting(App.Title, "Settings", "SILOtiltMin_" & Format(i + 1, "00"), "Fail")
        tiltStepTmp = _
            GetSetting(App.Title, "Settings", "SILOtiltStep_" & Format(i + 1, "00"), "Fail")
        If IsNumeric(centerXTmp) = False _
            Or Abs(CSng(Val(centerXTmp))) > 25! _
            Then
            centerXTmp = "0.0"
            SaveSetting App.Title, "Settings", "SILOcenterX_" & Format(i + 1, "00") _
                , centerXTmp
        End If
        If IsNumeric(centerYTmp) = False _
            Or Abs(CSng(Val(centerYTmp))) > 25! _
            Then
            centerYTmp = "0.0"
            SaveSetting App.Title, "Settings", "SILOcenterY_" & Format(i + 1, "00") _
                , centerYTmp
        End If
        If IsNumeric(radiusTmp) = False _
            Or CSng(Val(radiusTmp)) < 1! Or CSng(Val(radiusTmp)) > 25! _
            Then
            radiusTmp = "19.0"
            SaveSetting App.Title, "Settings", "SILOradius_" & Format(i + 1, "00") _
                , radiusTmp
        End If
        If IsNumeric(tiltDefaultTmp) = False _
            Or CSng(CInt(Val(tiltDefaultTmp))) <> CSng(Val(tiltDefaultTmp)) _
            Or CInt(Val(tiltDefaultTmp)) > 48! Or CInt(Val(tiltDefaultTmp)) < -48! _
            Then
            tiltDefaultTmp = "-1"
            SaveSetting App.Title, "Settings", "SILOtiltDefault_" & Format(i + 1, "00") _
                , tiltDefaultTmp
        End If
        If IsNumeric(tiltMaxTmp) = False _
            Or CSng(Val(tiltMaxTmp)) > 48! Or CSng(Val(tiltMaxTmp)) < 1! _
            Then
            tiltMaxTmp = "48.0"
            SaveSetting App.Title, "Settings", "SILOtiltMax_" & Format(i + 1, "00") _
                , tiltMaxTmp
        End If
        If IsNumeric(tiltMinTmp) = False _
            Or CSng(Val(tiltMinTmp)) < -48! Or CSng(Val(tiltMinTmp)) > -1! _
            Then
            tiltMinTmp = "-48.0"
            SaveSetting App.Title, "Settings", "SILOtiltMin_" & Format(i + 1, "00") _
                , tiltMinTmp
        End If
        If CSng(Val(tiltMaxTmp)) <= CSng(Val(tiltMinTmp)) Then
            tiltMaxTmp = "48.0"
            tiltMinTmp = "-48.0"
            SaveSetting App.Title, "Settings", "SILOtiltMax_" & Format(i + 1, "00") _
                , tiltMaxTmp
            SaveSetting App.Title, "Settings", "SILOtiltMin_" & Format(i + 1, "00") _
                , tiltMinTmp
        End If
        If IsNumeric(tiltStepTmp) = False _
            Or CSng(Val(tiltStepTmp)) > 5! Or CSng(Val(tiltStepTmp)) < 0.5! _
            Then
            tiltStepTmp = "2.0"
            SaveSetting App.Title, "Settings", "SILOtiltStep_" & Format(i + 1, "00") _
                , tiltStepTmp
        End If
        ucSilo1(i).setBinSettings _
            CSng(centerXTmp), CSng(centerYTmp), CSng(radiusTmp) _
            , CInt(tiltDefaultTmp), CSng(tiltMaxTmp), CSng(tiltMinTmp), CSng(tiltStepTmp)
    Next i

''    ucSilo1(2).setScanTYPE 2590  '''''LD-LRS-3100,, DPS-2590 ==> CONSOLE-Mode!
''
''    ucSilo1(1).setScanTYPE 22590  '''''201810~
''    ucSilo1(4).setScanTYPE 22590  '''''201810~
''    ucSilo1(7).setScanTYPE 22590  '''''201810~
''    ucSilo1(10).setScanTYPE 22590  '''''201810~


    Dim TiltStr1 As String
    TiltStr1 = Trim(Str(GetSetting(App.Title, "Settings", "TiltStrBase", "-1")))
        


    '''''''''''''''''''''''''"New-CTS-SILO*4x2==8"
    For i = 15 To 18
        Load ucSilo1(i)
    Next i

    For i = 15 To 18

        ucSilo1(i).Width = 5500 ''3500  ''Width / 11 - 30
        ucSilo1(i).Height = 3700  ''4000 ''3500

        ucSilo1(i).Left = ((i - 15) * (Width / 5)) + 150  ''(i * (Width / 11)) + 20
        ucSilo1(i).Top = ((3) * 3700) + 1200  ''(i * (Width / 11)) + 20

        ucSilo1(i).setIDX i, "", ""
        ucSilo1(i).Visible = True
    Next i

    ipAddr(15) = "192.168.0.171": ipPort(15) = "7002"  ''"7001"  '''
    ipAddr(16) = "192.168.0.171": ipPort(16) = "7004"  '''"7003"
    ipAddr(17) = "192.168.0.171": ipPort(17) = "7006"  '''"7005"
    ipAddr(18) = "192.168.0.171": ipPort(18) = "7007"  '''"7007"

    For i = 15 To 18
        ucSilo1(i).setIDX i, ipAddr(i), ipPort(i)
        ''
        ''ucSilo1(i).setScanTYPE 12590  '''LD-LRS-3100,, DPS-2590  ==> UDP-Mode!(Rx_8056bytes)
        ''
        typeTmp = Trim(Str(GetSetting(App.Title, "Settings", "SILOtypes_" & Format(i + 1, "00"), 12590)))
        ucSilo1(i).setScanTYPE typeTmp  ''3100  '''LD-LRS-3100,, DPS-2590
        
    Next i

'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''((Test))
'    ipAddr(2) = "192.168.0.171": ipPort(2) = "7001"
'    ipAddr(5) = "192.168.0.171": ipPort(5) = "7003"
'    ipAddr(8) = "192.168.0.171": ipPort(8) = "7005"
'    ipAddr(11) = "192.168.0.171": ipPort(11) = "7007"
'    ''
'    ucSilo1(2).setIDX 2, ipAddr(2), ipPort(2):      ucSilo1(2).setScanTYPE 12590
'    ucSilo1(5).setIDX 5, ipAddr(5), ipPort(5):      ucSilo1(5).setScanTYPE 12590
'    ucSilo1(8).setIDX 8, ipAddr(8), ipPort(8):      ucSilo1(8).setScanTYPE 12590
'    ucSilo1(11).setIDX 11, ipAddr(11), ipPort(11):  ucSilo1(11).setScanTYPE 12590


    
'    cboIDX.ListIndex = 0
'    cboIDX.Refresh


    txtPcsPort.Text = Trim(Str(GetSetting(App.Title, "Settings", "PcsPORT", "8005")))
    ''txtPcsIP.Text = Trim(Str(GetSetting(App.Title, "Settings", "PcsIP", "172.24.55.27")))
    txtPcsIP.Text = "172.24.55.27"  ''"127.0.0.1"  '''"172.24.55.27"
    
    txtPcsPort2.Text = Trim(Str(GetSetting(App.Title, "Settings", "PcsPORT2", "8009")))  '''NewCTS-Silo
    
    
    txtMaxHH.Text = Trim(Str(GetSetting(App.Title, "Settings", "MaxHH", "5000")))
    txtBaseHH.Text = Trim(Str(GetSetting(App.Title, "Settings", "BaseHH", "100")))
    
    txtMaxHH.Enabled = False
    txtBaseHH.Enabled = False


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
   
    For i = 0 To 18  '''14  '''New-CTS-Silo(15+4)!
        ucSilo1(i).set_Angle CDbl(GetSetting(App.Title, "Settings", "SILOang_" & Format(i + 1, "00"), 0))
    Next i
    
    
    cmdDRAWmode.BackColor = vbGreen
    ''
    For i = 0 To 18  '''14  '''New-CTS-Silo(15+4)!
        
        ucSilo1(i).set_DRAWmode 0
        
        ucSilo1(i).set_maxHH CLng(txtMaxHH)
        ucSilo1(i).set_baseHH CLng(txtBaseHH)
            
    Next i

''Port: 800, IGrp: 0xF020, IOffs: 0x64, Len: 2

    AdsOcx1.AdsAmsServerNetId = "192.168.0.73.1.1"   '''SILO--15 '''AdsOcx1.AdsAmsClientNetId
    AdsOcx1.AdsAmsServerPort = 800
    AdsOcx1.EnableErrorHandling = False

    AdsOcx2.AdsAmsServerNetId = "192.168.0.173.1.1"  '''New-CTS-SILO--4(x2) '''AdsOcx1.AdsAmsClientNetId
    AdsOcx2.AdsAmsServerPort = 800
    AdsOcx2.EnableErrorHandling = False
    
    
    tmrINIT.Interval = 3000  ''5000
    tmrINIT.Enabled = True  ''=========>> Run: <<tmrAoDo>> <<tmrPcs>>


    tmrSinit.Interval = 3000  ''9000
    tmrSinit.Enabled = True  ''=========>> Run: << ucSilo1(i).initStart >>

End Sub


Private Sub Form_Terminate()
    ''Return
End Sub



Private Sub tmrAoDo2_Timer()  '''New-CTS-SILO--4(x2)

Dim i As Integer
Dim j As Integer

Dim ioD(33) As Integer
Dim str1 As String

Dim aaD(20) As Integer

Dim avrD(20) As Integer
Dim avrDsum(20) As Long



    For i = 15 To 18
        aaD(i) = ucSilo1(i).ret_AOd
        '''''''''''''''''''''''''''''
    Next i


    ''SAVE--First!!
    For i = 15 To 18
        If (aaD(i) > 0) And (aaD(i) < 32768) Then
            SaveSetting App.Title, "Settings", "AV_" & Trim(i), aaD(i)
        Else
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 0)
        End If
    Next i


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''<AVR)
    For i = 15 To 18
        AOdeep(i, AOdeepCNT2) = aaD(i)
    Next i
    ''
    AOdeepCNT2 = AOdeepCNT2 + 1
    ''
    If AOdeepCNT2 >= AOdeepMAX Then  ''99
        AOdeepFull = True
        AOdeepCNT2 = 0       ''''Loop!
    End If


    For i = 15 To 18
        avrDsum(i) = 0
    Next i


    ''//??????????
    If AOdeepFull = True Then
    ''
        For i = 15 To 18
            For j = 0 To AOdeepMAX - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepMAX)
        Next i
    ''
    ElseIf AOdeepCNT2 > 1 Then
    ''
      txtAVRcnt = Trim(AOdeepCNT2 + 1)
        For i = 15 To 18
            For j = 0 To AOdeepCNT2 - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepCNT2)
        Next i
    ''
    Else
        txtAVRcnt = Trim(AOdeepCNT2 + 1)
        For i = 15 To 18
            avrD(i) = aaD(i)
        Next i
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>AVR)

    ''set_avrHH for View
    For i = 15 To 18
        ucSilo1(i).set_avrHH avrD(i)
    Next i


    ''''===[ LOG Save aaD(i),avrD(i) ]==='''
    str1 = ""  ''"BIN> "
    For i = 15 To 18
    ''  str1 = str1 & Trim(i + 1) & ")" & Format(aaD(i), "00000") & "," & Format(avrD(i), "00000") & "," & Format((avrD(i) / 32768 * 100), "00.0") & "% "
        str1 = str1 & Trim(i + 1) & ")" & Format((avrD(i) / 32768 * 100), "00.0") & "% "
    Next i
    ''
    DGPSLog str1, "SILO"


    ''Replace!!
    For i = 15 To 18
        aaD(i) = avrD(i)
    Next i


    ''SAVE--Replace!!
    For i = 15 To 18
        If (aaD(i) > 0) And (aaD(i) < 32768) Then
            SaveSetting App.Title, "Settings", "AV_" & Trim(i), aaD(i)
        Else
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 0)
        End If
    Next i



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''SILO'''
''''''''''[BackHoff]-IO-MAP
''''-----------------------------------------------
''''    4 x 4 = 16 ::: 8+8  ????????????????????????????????
''''-----------------------------------------------
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ioD(0) = aaD(15)
    ioD(1) = aaD(15)
    ioD(2) = aaD(16)
    ioD(3) = aaD(16)

    ioD(4) = aaD(17)
    ioD(5) = aaD(17)
    ioD(6) = aaD(18)
    ioD(7) = aaD(18)

    ioD(8) = aaD(15)
    ioD(9) = aaD(15)
    ioD(10) = aaD(16)
    ioD(11) = aaD(16)

    ioD(12) = aaD(17)
    ioD(13) = aaD(17)
    ioD(14) = aaD(18)
    ioD(15) = aaD(18)

    ioD(16) = 1
    ioD(17) = 1
    ioD(18) = 1
    ioD(19) = 1 ''0


    For i = 0 To 19  ''31
    ''--------------------------------------------------------(Temp)
''        If (ioD(i) > 0) And (ioD(i) <= 32767) Then
''            AOdata(i) = ioD(i)
''        Else
''            txtSD1 = txtSD1 & " 12-" & (i + 1) & "?"
''            txtSD1.SelStart = Len(txtSD1)
''            ''
''            Exit Sub
''            ''=========>> Cancle for Next~~ /(protect_Zero_send)
''        End If
    ''--------------------------------------------------------(Temp)
        AOdata(i) = ioD(i)
        ''''''''''''''''''
    Next i
    

On Error GoTo wsErrADS2
    AdsOcx2.EnableErrorHandling = True
    AdsOcx2.AdsSyncWriteReq &HF020&, &H64&, 32, AOdata   ''New-CTS:4x4:: ''SILO:[40]==4*5=20channel!
    ''''''''''''''''''''''''''''''''''''''''''''''''''

wsErrADS2:
    AdsOcx2.EnableErrorHandling = False
    '''''Just-Cancle...for next


    txtSD1 = txtSD1 & vbCrLf & str1
    txtSD1.SelStart = Len(txtSD1)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    txtWSpcs2 = wsPcs2.State

    If wsPcs2.State = sckConnected Then
        '''''''''''''
        EditPcsData2 17  ''''5
        '''''''''''''''
    End If

End Sub



Private Sub tmrAoDo_Timer()

Dim i As Integer
Dim j As Integer

Dim ioD(33) As Integer
Dim str1 As String

Dim aaD(15) As Integer

Dim avrD(15) As Integer
Dim avrDsum(15) As Long


    lbNow.Caption = Format(Now, "YYYY-MM-DD h:m:s")


''    txtSD1 = ucSilo1(0).getTXT
''    DoEvents


    If Len(txtSD1) > 9000 Then
        txtSD1 = Mid(txtSD1, 5000)
    End If

    For i = 0 To 14
        aaD(i) = ucSilo1(i).ret_AOd
        '''''''''''''''''''''''''''''
    Next i

    ''SAVE--First!!
    For i = 0 To 14
        If (aaD(i) > 0) And (aaD(i) < 32768) Then
            SaveSetting App.Title, "Settings", "AV_" & Trim(i), aaD(i)
        Else
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 0)
        End If
    Next i
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''<AVR)
    For i = 0 To 14
        AOdeep(i, AOdeepCNT) = aaD(i)
    Next i
    ''
    AOdeepCNT = AOdeepCNT + 1
    ''
    If AOdeepCNT >= AOdeepMAX Then  ''99
        AOdeepFull = True
        AOdeepCNT = 0       ''''Loop!
    End If


    For i = 0 To 14
        avrDsum(i) = 0
    Next i
    
    ''//??????????
    If AOdeepFull = True Then
    ''
        For i = 0 To 14
            For j = 0 To AOdeepMAX - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepMAX)
        Next i
    ''
    ElseIf AOdeepCNT > 1 Then
    ''
      txtAVRcnt = Trim(AOdeepCNT + 1)
        For i = 0 To 14
            For j = 0 To AOdeepCNT - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepCNT)
        Next i
    ''
    Else
        txtAVRcnt = Trim(AOdeepCNT + 1)
        For i = 0 To 14
            avrD(i) = aaD(i)
        Next i
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>AVR)

    ''set_avrHH for View
    For i = 0 To 14
        ucSilo1(i).set_avrHH avrD(i)
    Next i



    ''''===[ LOG Save aaD(i),avrD(i) ]==='''
    str1 = ""  ''"BIN> "
    For i = 0 To 14
    ''  str1 = str1 & Trim(i + 1) & ")" & Format(aaD(i), "00000") & "," & Format(avrD(i), "00000") & "," & Format((avrD(i) / 32768 * 100), "00.0") & "% "
        str1 = str1 & Trim(i + 1) & ")" & Format((avrD(i) / 32768 * 100), "00.0") & "% "
    Next i
    ''
    DGPSLog str1, "SILO"


    ''Replace!!
    For i = 0 To 14
        aaD(i) = avrD(i)
    Next i


    ''SAVE--Replace!!
    For i = 0 To 14
        If (aaD(i) > 0) And (aaD(i) < 32768) Then
            SaveSetting App.Title, "Settings", "AV_" & Trim(i), aaD(i)
        Else
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 0)
        End If
    Next i




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''SILO'''
''''''''''[BackHoff]-IO-MAP
''''-----------------------------------------------
''''    0(1)    4(10)   8(5)    12(14)  16(9)   20
''''                            ------
''''    1(4)    5(13)   9(8)    13(3)   17(12)  21
''''           ------
''''    2(7)    6(2)    10(11)  14(6)   18(15)  22
''''                                    ------
''''    3       7       11      15      19      23
''''-----------------------------------------------
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ioD(0) = aaD(1 - 1)
    ioD(1) = aaD(4 - 1)
    ioD(2) = aaD(7 - 1)
    ioD(3) = 1 ''0

    ioD(4) = aaD(10 - 1)
    ioD(5) = aaD(13 - 1)
    ioD(6) = aaD(2 - 1)
    ioD(7) = 1 ''0

    ioD(8) = aaD(5 - 1)
    ioD(9) = aaD(8 - 1)
    ioD(10) = aaD(11 - 1)
    ioD(11) = 1 ''0

    ioD(12) = aaD(14 - 1)
    ioD(13) = aaD(3 - 1)
    ioD(14) = aaD(6 - 1)
    ioD(15) = 1 ''0

    ioD(16) = aaD(9 - 1)
    ioD(17) = aaD(12 - 1)
    ioD(18) = aaD(15 - 1)
    ioD(19) = 1 ''0


    For i = 0 To 19  ''31
    ''--------------------------------------------------------(Temp)
''        If (ioD(i) > 0) And (ioD(i) <= 32767) Then
''            AOdata(i) = ioD(i)
''        Else
''            txtSD1 = txtSD1 & " 12-" & (i + 1) & "?"
''            txtSD1.SelStart = Len(txtSD1)
''            ''
''            Exit Sub
''            ''=========>> Cancle for Next~~ /(protect_Zero_send)
''        End If
    ''--------------------------------------------------------(Temp)
        AOdata(i) = ioD(i)
        ''''''''''''''''''
    Next i
    

On Error GoTo wsErrADS
    AdsOcx1.EnableErrorHandling = True
    AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 40, AOdata   ''SILO:[40]==4*5=20channel!
    ''''''''''''''''''''''''''''''''''''''''''''''''''

wsErrADS:
    AdsOcx1.EnableErrorHandling = False
    '''''Just-Cancle...for next


    txtSD1 = txtSD1 & vbCrLf & str1
    txtSD1.SelStart = Len(txtSD1)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    txtWSpcs = wsPcs.State

    If wsPcs.State = sckConnected Then
        '''''''''''''
        EditPcsData 5
        '''''''''''''
    End If
    
End Sub



Private Sub EditPcsData2(Pno As Integer)

''Dim sendbuf(3128) As Byte  ''Variant  ''1672!! '''2295
''
Dim sendbuf(840) As Byte  ''NewCTS-Silo((201801))
''
Dim i As Integer
Dim j As Integer
Dim cnt1 As Integer
Dim ret1 As Integer
Dim str1 As String
Dim L8 As Byte
Dim H8 As Byte

''struct SENDBUF
''{
''   short   head;           //0x1122 고정
''   short   size;           //Buffer 전체 Size
''   short   plant;          //0015:<SILO>
''   short   spare;
''
''   short   linkstat[15];   //통신 상태  1:정상 0:이상
''   short   height[15];     //평균 높이
''   short   volume[15];     //용적 m3
''   short   data[15][101];  //BIN Level Data
''};
'''''0x0C38 <= 3128 <= 2+2+2+2+30+30+30+(2*15*101) :: SLIO:15ea

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''((SNG-201409)) // NewCTS-Silo((201801))
''struct SENDBUF
''{
''   short   head;           //0x1122 고정
''   short   size;           //Buffer 전체 Size
''   short   plant;          //0015:<SILO>  //0017:<CTS-SILO>
''   short   spare;
''
''   short   linkstat[4];   //통신 상태  1:정상 0:이상
''   short   height[4];     //평균 높이
''   short   volume[4];     //용적 m3
''   short   data[4][101];  //BIN Level Data
''};
'''''0x0348 <= 840 =(32+808)= <= 2+2+2+2+8+8+8+(2*4*101) :: ''''''''SNG:4ea


    sendbuf(0) = &H22:    sendbuf(1) = &H11     ''Header
    sendbuf(2) = &H48:    sendbuf(3) = &H3      ''Size  ''0x0C38 <= 3128
    sendbuf(4) = &H11:    sendbuf(5) = &H0      ''Plant-No  ''15  ''17(0x11) <==<CTS-SILO>
    sendbuf(6) = &H0:     sendbuf(7) = &H0      ''spare

'<08>''
    For i = 0 To 3  ''14  ''SILO
        sendbuf(8 + i * 2) = CByte(ucSilo1(i + 15).ret_Act) ''BIN_Comm_Act
        sendbuf(8 + i * 2 + 1) = &H0
    Next i

'<16>''
    For i = 0 To 3  ''14  ''SILO
        ret1 = ucSilo1(i + 15).ret_HH
        sendbuf(16 + i * 2) = CByte(ret1 Mod 256)
        sendbuf(16 + i * 2 + 1) = CByte(ret1 \ 256) ''Height...AVR
        
    Next i
    
'<22>''
    For i = 0 To 3  ''14  ''SILO
        ret1 = ucSilo1(i + 15).ret_VV
        sendbuf(22 + i * 2) = CByte(ret1 Mod 256)
        sendbuf(22 + i * 2 + 1) = CByte(ret1 \ 256) ''VVV...AVR
                
    Next i

'<30>''
    ''''74''((+(11*202)==2222==>((2296))  ''5소결
    ''''56''((+(08*202)==1616==>((1672))   ''1234미분광
    For j = 0 To 3  ''14  ''SILO
    
      If ucSilo1(j + 15).ret_Act > 0 Then
      ''''''''''''''''''''''''''''''''
        For i = 0 To 100
            ret1 = ucSilo1(j + 15).GETscanD(i)
            ''''''''''''''''''''''''''''''''''scan_Data''
            If ret1 < 0 Then ret1 = 0
            
            cnt1 = 30 + (j * 202) + (i * 2)
            sendbuf(cnt1) = ret1 Mod 256  ''L8
            sendbuf(cnt1 + 1) = ret1 \ 256  ''H8
            
        Next i
        
      Else
        For i = 0 To 100
            cnt1 = 30 + (j * 202) + (i * 2)
            sendbuf(cnt1) = 0
            sendbuf(cnt1 + 1) = 0
        Next i
      End If
      
    Next j

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo wsErrPcs2

    txtWSpcs2 = wsPcs2.State
    
    If wsPcs2.State = sckConnected Then
        wsPcs2.SendData sendbuf
        ''''''''''''''''''''''
    End If

    Exit Sub
    
wsErrPcs2:
    wsPcs2.Close
    DoEvents

End Sub


Private Sub EditPcsData(Pno As Integer)

Dim sendbuf(3128) As Byte  ''Variant  ''1672!! '''2295
Dim i As Integer
Dim j As Integer
Dim cnt1 As Integer
Dim ret1 As Integer
Dim str1 As String
Dim L8 As Byte
Dim H8 As Byte


''    ''struct SENDBUF
''    ''{
''    ''   short   head;           //0x1122 고정
''    ''   short   size;           //Buffer 전체 Size
''    ''   short   plant;          //1:1소결 2:2소결 3:3소결 4:4소결
''    ''   short   spare;
''    ''
''    ''   short   linkstat[10];   //통신 상태  1:정상 0:이상
''    ''   short   height[10];     //평균 높이
''    ''   short   volume[10];     //용적       10-2 m3
''    ''   short   data[10][101];  //BIN Level
''    ''};
''    '''Connect( "172.24.55.27", 8001);
''
''    ''0x0828 <= 2088 <= 2+2+2+2+20+20+20++2020 ::(1234)
''    '''''''''''''''''''''''''''''''''''''''''
''    ''0x08f8 <= 2296 <= 2+2+2+2+22+22+22++2222 ::(5)
''
''
''    '   sendbuf.head = 0x1122;
''    '   sendbuf.size = sizeof(sendbuf);
''    '   sendbuf.plant = Plant;


'''''''''''''''''''''''''''''''''''''''''[미분광: 12~4EA, 34~4EA]'''''''''''''''''''
''    ''struct SENDBUF
''    ''{
''    ''   short   head;           //0x1122 고정
''    ''   short   size;           //Buffer 전체 Size
''    ''   short   plant;          //1:1소결 2:2소결 3:3소결 4:4소결 5:5소결 6:1234미분광
''    ''   short   spare;
''    ''
''    ''   short   linkstat[8];   //통신 상태  1:정상 0:이상
''    ''   short   height[8];     //평균 높이 <------------------((Base-Leve)??)
''    ''   short   volume[8];     //용적       10-2 m3
''    ''   short   data[8][101];  //BIN Level
''    ''};
''    '''Connect( "172.24.55.27", 8004);  ''1234미분광
''
''    ''0x0828 <= 2088 <= 2+2+2+2+20+20+20++2020 ::(1234)---10ea
''    '''''''''''''''''''''''''''''''''''''''''
''    ''0x08f8 <= 2296 <= 2+2+2+2+22+22+22++2222 ::(5)---11ea
''
''    ''0x0688 <= 1672 <= 2+2+2+2+16+16+16++1616 ::(1234미분광)---8ea


''struct SENDBUF
''{
''   short   head;           //0x1122 고정
''   short   size;           //Buffer 전체 Size
''   short   plant;          //0015:<SILO>
''   short   spare;
''
''   short   linkstat[15];   //통신 상태  1:정상 0:이상
''   short   height[15];     //평균 높이
''   short   volume[15];     //용적 m3
''   short   data[15][121];  //BIN Level Data
''};
''
''0x0E90 <= 3728 <= 2+2+2+2+30+30+30+(30*121) :: SLIO:15ea

''struct SENDBUF
''{
''   short   head;           //0x1122 고정
''   short   size;           //Buffer 전체 Size
''   short   plant;          //0015:<SILO>
''   short   spare;
''
''   short   linkstat[15];   //통신 상태  1:정상 0:이상
''   short   height[15];     //평균 높이
''   short   volume[15];     //용적 m3
''   short   data[15][101];  //BIN Level Data
''};
'''''0x0C38 <= 3128 <= 2+2+2+2+30+30+30+(2*15*101) :: SLIO:15ea



    sendbuf(0) = &H22:    sendbuf(1) = &H11     ''Header
    sendbuf(2) = &H38:    sendbuf(3) = &HC      ''Size  ''0x0C38 <= 3128
    sendbuf(4) = &HF:     sendbuf(5) = &H0      ''Plant-No  ''15
    sendbuf(6) = &H0:     sendbuf(7) = &H0      ''spare

'<08>''
    For i = 0 To 14  ''SILO
        sendbuf(8 + i * 2) = CByte(ucSilo1(i).ret_Act)   ''BIN_Comm_Act
        sendbuf(8 + i * 2 + 1) = &H0
    Next i

'<38>''
    For i = 0 To 14  ''SILO
        ret1 = ucSilo1(i).ret_HH
        sendbuf(38 + i * 2) = CByte(ret1 Mod 256)
        sendbuf(38 + i * 2 + 1) = CByte(ret1 \ 256) ''Height...AVR
        
    Next i
    
'<68>''
    For i = 0 To 14  ''SILO
        ret1 = ucSilo1(i).ret_VV
        sendbuf(68 + i * 2) = CByte(ret1 Mod 256)
        sendbuf(68 + i * 2 + 1) = CByte(ret1 \ 256) ''VVV...AVR
                
    Next i

'<98>''
    ''''74''((+(11*202)==2222==>((2296))  ''5소결
    ''''56''((+(08*202)==1616==>((1672))   ''1234미분광
    For j = 0 To 14  ''SILO
    
      If ucSilo1(j).ret_Act > 0 Then
      ''''''''''''''''''''''''''''''''
        For i = 0 To 100
            ret1 = ucSilo1(j).GETscanD(i)
            ''''''''''''''''''''''''''''''''''scan_Data''
            If ret1 < 0 Then ret1 = 0
            
            cnt1 = 98 + (j * 202) + (i * 2)
            sendbuf(cnt1) = ret1 Mod 256  ''L8
            sendbuf(cnt1 + 1) = ret1 \ 256  ''H8
        Next i
        
''        If j = 4 Then  '''Debug-Draw
''            str1 = "SendDRAW: "
''            For i = 0 To 100
''                cnt1 = 98 + (j * 202) + (i * 2)
''                ret1 = (sendbuf(cnt1 + 1) * 256) + sendbuf(cnt1)
''
''                str1 = Format((i), "000") & " " & Format((ret1), "0000") & " : "
''                For cnt1 = 0 To ((ret1 - 1000) / 10)
''                    str1 = str1 & "."
''                Next cnt1
''                Debug.Print str1
''
''            Next i
''        End If
        
      Else
        For i = 0 To 100
            cnt1 = 98 + (j * 202) + (i * 2)
            sendbuf(cnt1) = 0
            sendbuf(cnt1 + 1) = 0
        Next i
      End If
      
    Next j



On Error GoTo wsErrPcs

    txtWSpcs = wsPcs.State
    
    If wsPcs.State = sckConnected Then
        wsPcs.SendData sendbuf
        ''''''''''''''''''''''
        
''            str1 = "(" & Hex(Val(UBound(sendbuf))) & ") "
''            For i = 0 To UBound(sendbuf)
''                str1 = str1 & Format(Hex(sendbuf(i)), "00") & " "
''            Next i
''            ''txtSD1 = str1
''            Debug.Print str1
            
    End If

    Exit Sub
    
wsErrPcs:
    wsPcs.Close
    DoEvents

End Sub



Private Sub tmrINIT_Timer()
    tmrINIT.Enabled = False
    
    tmrAoDo.Interval = 2000  ''1000
    tmrAoDo.Enabled = True

    tmrPcs.Interval = 3000
    tmrPcs.Enabled = True
    
    
    tmrAoDo2.Interval = 2000  ''1000
    tmrAoDo2.Enabled = True
    
    tmrPcs2.Interval = 3000
    tmrPcs2.Enabled = True
    
End Sub


Private Sub tmrPcs2_Timer()

    If wsPcs2.State <> sckConnected Then

        wsPcs2.Close
    
        wsPcs2.RemoteHost = txtPcsIP.Text   ''"172.24.55.27"
        wsPcs2.RemotePort = txtPcsPort2.Text  '''8009 (201801--NewCTS-Silo) ''"8005"  "8004"   ''"8003"

        ''SaveSetting App.Title, "Settings", "PcsIP", Trim(txtPcsIP.Text)
        SaveSetting App.Title, "Settings", "PcsPORT2", Trim(txtPcsPort2.Text)

        wsPcs2.Connect
    
    End If

    txtWSpcs2 = wsPcs2.State

End Sub


Private Sub tmrPcs_Timer()

    If wsPcs.State <> sckConnected Then

        wsPcs.Close
    
        wsPcs.RemoteHost = txtPcsIP.Text   ''"172.24.55.27"
        wsPcs.RemotePort = txtPcsPort.Text  ''"8005"  "8004"   ''"8003"

        SaveSetting App.Title, "Settings", "PcsIP", Trim(txtPcsIP.Text)
        SaveSetting App.Title, "Settings", "PcsPORT", Trim(txtPcsPort.Text)

        wsPcs.Connect
    
    End If

    txtWSpcs = wsPcs.State

End Sub


Private Sub wsPcs2_DataArrival(ByVal bytesTotal As Long)
Dim rBuf As Variant
    wsPcs2.GetData rBuf  ''''null...
    
End Sub

Private Sub wsPcs2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsPcs2.Close
    DoEvents
End Sub


Private Sub wsPcs_DataArrival(ByVal bytesTotal As Long)
Dim rBuf As Variant
    wsPcs.GetData rBuf  ''''null...
    
End Sub

Private Sub wsPcs_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsPcs.Close
    DoEvents
End Sub



'''
'''  SILO
'''

Private Sub tmrSinit_Timer()
Dim i

    '======================!
    tmrSinit.Enabled = False
    '''''''''''''''''''''''!
    '''''''''''''''''''''''!
    
    For i = 0 To 18   '''14  '''New-CTS-Silo(15+4)!
        ucSilo1(i).initStart
    Next i

End Sub





VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ucSilo 
   Appearance      =   0  '평면
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   FillStyle       =   0  '단색
   ScaleHeight     =   4455
   ScaleWidth      =   6345
   Begin VB.TextBox txtRxS 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   31
      Text            =   ".."
      Top             =   0
      Width           =   3255
   End
   Begin VB.Timer tmrReStart 
      Enabled         =   0   'False
      Interval        =   9000
      Left            =   3240
      Top             =   3600
   End
   Begin VB.CommandButton cmdFilt 
      BackColor       =   &H00808080&
      Caption         =   "FILT"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   420
      Width           =   975
   End
   Begin VB.TextBox ldWidth 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   3720
      TabIndex        =   18
      Text            =   "0"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtAVRheight 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   3720
      TabIndex        =   12
      Top             =   1380
      Width           =   975
   End
   Begin VB.TextBox txtAcnt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4320
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtAsum 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   3720
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtScaleHight 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   3720
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtAOd 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   3720
      TabIndex        =   8
      Top             =   1980
      Width           =   975
   End
   Begin VB.Timer tmrTRX 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1800
      Top             =   3660
   End
   Begin MSWinsockLib.Winsock wsockLD 
      Left            =   1080
      Top             =   3660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrSrun 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   3660
   End
   Begin VB.CommandButton cmdCONN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BIN1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   45
      Width           =   1635
   End
   Begin VB.PictureBox picSilo 
      BackColor       =   &H00404040&
      Height          =   3435
      Left            =   120
      Picture         =   "ucSILO.ctx":0000
      ScaleHeight     =   3375
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   60
      Width           =   3555
      Begin VB.Label lbTiltTX 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   840
         TabIndex        =   36
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lbTiltV 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2880
         TabIndex        =   34
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lbTiltRX 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1800
         TabIndex        =   33
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lbRxHead 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   0
         TabIndex        =   32
         Top             =   3240
         Width           =   495
      End
   End
   Begin VB.Label lbRadius 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "00.0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   5040
      TabIndex        =   39
      Top             =   3495
      Width           =   375
   End
   Begin VB.Label lbCenterY 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "-00.0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4440
      TabIndex        =   38
      Top             =   3495
      Width           =   495
   End
   Begin VB.Label lbCenterX 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "-00.0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3840
      TabIndex        =   37
      Top             =   3495
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Caption         =   "비례높이"
      Height          =   195
      Left            =   4740
      TabIndex        =   30
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Caption         =   "측정높이"
      Height          =   195
      Left            =   4740
      TabIndex        =   29
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "[On/Off]"
      Height          =   195
      Left            =   4680
      TabIndex        =   28
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      Top             =   2820
      Width           =   315
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "mA"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   2580
      Width           =   315
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "전류Data"
      Height          =   195
      Left            =   4740
      TabIndex        =   23
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "체적"
      Height          =   195
      Left            =   5100
      TabIndex        =   22
      Top             =   2880
      Width           =   435
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "전류"
      Height          =   195
      Left            =   5100
      TabIndex        =   21
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "높이"
      Height          =   195
      Left            =   5100
      TabIndex        =   20
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "Width"
      Height          =   195
      Left            =   4740
      TabIndex        =   19
      Top             =   780
      Width           =   675
   End
   Begin VB.Label lbHP 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lbAO 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H0000C000&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "바탕체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lbVVV 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   2820
      Width           =   975
   End
   Begin VB.Label lbHH 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4680
      TabIndex        =   13
      Top             =   3180
      Width           =   375
   End
   Begin VB.Label lbMode 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   7
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label lbRXerr 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2820
      TabIndex        =   6
      Top             =   4140
      Width           =   855
   End
   Begin VB.Label lbRXcnt 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1920
      TabIndex        =   5
      Top             =   4140
      Width           =   855
   End
   Begin VB.Label lbXC 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1200
      TabIndex        =   4
      Top             =   4140
      Width           =   282
   End
   Begin VB.Label lbAngle 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1560
      TabIndex        =   3
      Top             =   4140
      Width           =   255
   End
   Begin VB.Label lbPointErrCnt 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   840
      TabIndex        =   35
      Top             =   4140
      Width           =   282
   End
   Begin VB.Label lbCnt 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   360
      TabIndex        =   2
      Top             =   4140
      Width           =   375
   End
End
Attribute VB_Name = "ucSilo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Event Resize()

Public Event upDXY()

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

''''''''''''''''''''''''''''''''[ DPS2590-EX :: Tilt-Motion ] 201809~DPSex
Private AutoTiltON As Boolean
Private AutoTiltStarted As Boolean
Private AutoTiltOffDelayCnt As Integer
Private AutoTiltErrorCnt As Integer
Private AutoTiltCnt As Integer
Private AutoTiltStep As Double
Private AutoTiltNow As Double
Private AutoTiltMax As Double
Private AutoTiltMin As Double
''''''''''''''''''''''''''''''''

Private ScanTYPE As Integer  '''LD-LRS-3100,, DPS-2590
Public CenterX!, CenterY!, Radius!
Public TiltDefault%, TiltMax!, TiltMin!, TiltStep!
''
Private inBUF2590 As String   '''inBUF2590(100000) As Byte
''
Private Scan2590Echo(1001) As Integer
Private Scan2590Direc(1001) As Double '''Long
Private Scan2590Dist(1001) As Double '''Long
Private Scan2590Pulse(1001) As Double '''Long
''______________________________________________________________________________
''  Point ;    Echo ;   Direction ;    Distance ; Pulse width ;
''        ;         ;       [deg] ;         [m] ;        [ps] ;
''______________________________________________________________________________


Private UCindex As Integer


Private ipAddr As String
Private ipPort As String

Const PI = 3.14159265359   '''3.14159265358979  ''3.1415926535897932384626433832795

Private tSrunMode As Integer


Private inBUF(20000) As Byte   ''<==(8056)== DPS-12590  ''inBUF(2000)
Private inCNT As Long
''
Private xcMax As Integer


Private rxBcnt As Integer
Private rxBYTE(2000) As Byte        ''(120~240)=>121*0.5degree: 121*2*2==484==242word
Private rxWORD(1001) As Long  '''1000


Private rxSTOP As Integer

''-------------------------------------
Private rxWdeep(5, 300) As Long
Private rxWdeepSum(300) As Long
Private rxWdeepCnt(300) As Integer
Private cnWdeep As Integer
Private cnWring As Integer
''-------------------------------------



Private RxMSG As Variant


'''Private LD_sBUF(50)
Private LD_sBUF(60)   ''[ DPS2590-EX :: Tilt-Motion ] 201809~DPSex


Private txtRx1 As String

Dim TxHeader As Variant  ''HDmsg = Chr(2) + Chr(2) + Chr(2) + Chr(2) + Chr(0) + Chr(0)
Dim RxHeader As Variant  ''HDmsg = Chr(2) + Chr(2) + Chr(2) + Chr(2) + Chr(0) + Chr(0)


''''(for picSiloDRAW)''
    Dim maxyrange As Double                     'Sets max y range of Scan
    Dim minyrange As Double                     'Sets min y range of Scan
    Dim maxxrange As Double                     'sets max x range of scan
    Dim minxrange As Double                     'sets min x range of scan

    Dim r(0 To 2000) As Double                  'radius data
    Dim X(0 To 2000) As Double                  'x - cartesian coordinate
    Dim Y(0 To 2000) As Double                  'y - cartesian coordinate
    Dim n As Integer                            'number of data values

    
    Dim minXL As Double
    Dim minXR As Double
''''

Private avrHH As Integer


Private scanDfilt(101) As Long  ''101
Private scanDfiltX(101) As Long  ''101

Private DRAWmode As Integer

Private maxHH As Long
Private baseHH As Long
''
Private scaleHH As Long

Private Enum eSrunMode
    InitConn = 0
    CheckConn
    Init1Device
    Check1Device
    Init2Device
    Check2Device
    SendCmd
    ReceiveData
End Enum

Private rxWaitTime As Integer

Private tilt3Dlog_fn As Integer

''Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long


Public Sub setScanTYPE(iScan As Integer)  '''LD-LRS-3100,, DPS-2590

    ScanTYPE = iScan
    
    SaveSetting App.Title, "Settings", "SILOtypes_" & Format(UCindex + 1, "00"), CInt(ScanTYPE)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If ScanTYPE = 2590 Then
            txtRxS.Visible = True
    Else
            txtRxS.Visible = False
    End If
    
    
    If ScanTYPE = 22590 Then
            lbTiltV.Visible = True
            lbTiltRX.Visible = True
            lbTiltTX.Visible = True
            lbRxHead.Visible = True
            UserControl.BackColor = &HFF8080     '''&HFFC0C0
    ElseIf ScanTYPE = 12590 Then
            lbRxHead.Visible = True
            lbTiltV.Visible = False
            lbTiltRX.Visible = False
            lbTiltTX.Visible = False
            UserControl.BackColor = &HFF8080     '''&HFFC0C0
    Else
            lbRxHead.Visible = False
            lbTiltV.Visible = False
            lbTiltRX.Visible = False
            lbTiltTX.Visible = False
    End If
    
    lbTiltTX.BackColor = &HFF8080
    lbTiltRX.BackColor = &HFF8080
    lbTiltV.BackColor = &HFF8080
    
End Sub

Public Sub setBinSettings(CenterX_I!, CenterY_I!, Radius_I! _
    , TiltDefault_I%, TiltMax_I!, TiltMin_I!, TiltStep_I!)
'
    CenterX = CenterX_I
    CenterY = CenterY_I
    Radius = Radius_I
    TiltDefault = TiltDefault_I
    TiltMax = TiltMax_I
    TiltMin = TiltMin_I
    TiltStep = TiltStep_I
'
    lbCenterX = Format(CenterX, "0.0")
    lbCenterY = Format(CenterY, "0.0")
    lbRadius = Format(Radius, "0.0")
'
End Sub

Public Function getScanTYPE() As Integer
    getScanTYPE = ScanTYPE
End Function


Public Sub set_Angle(hh As Double)
    lbAngle = Trim(hh)
    SaveSetting App.Title, "Settings", "SILOang_" & Format(UCindex + 1, "00"), CInt(hh)
End Sub

Public Function get_Angle() As Integer
    get_Angle = Val(lbAngle)
End Function

Public Sub set_DRAWmode(mode As Integer)
    
    DRAWmode = mode
    
End Sub


Public Sub set_maxHH(hh As Long)
    maxHH = hh
End Sub

Public Sub set_baseHH(hh As Long)
    baseHH = hh
End Sub


Public Sub set_avrHH(hh As Integer)  '''from AVR for txtAOd

    ''avrHH = (CLng(hh) * (5000) / 32768)
    avrHH = (CLng(hh) * (maxHH - baseHH) / 32768)

End Sub


Public Sub scan_STOP()
''    If wsock1.State = sckConnected Then
''        wsock1.SendData stopString
''    End If
    If AutoTiltON = True Then
        If AutoTiltStarted = False Then
            autoTilt_off
        Else
            autoTilt_stop
        End If
    End If
    
    tmrSrun.Enabled = False
End Sub

Public Sub scan_RUN()

''    Dim i As Integer
''    For i = 0 To 100
''        scanD(i) = 0
''    Next i
''
''    If wsock1.State = sckConnected Then
''        wsock1.SendData startString
''    End If
''
''    wsPause = False
''    '''''''''''''''
    
    initStart  ''tmrSrun.Enabled = True
    
End Sub



Private Sub cmdCONN_Click()
    
    txtRx1 = ""
    RxMSG = ""
    
''    inCNT = 0
''
''    LDtxDATA 39
''    Sleep (10)
''    rxWaitTime = 10
''    LDrxDATA 39
''
    
End Sub


Private Sub cmdFilt_Click()

    If cmdFilt.BackColor = vbGreen Then
        cmdFilt.BackColor = &H808080
        ''
        tmrReStart.Interval = 5000
        tmrReStart.Enabled = True
    Else
            ''cmdFilt.BackColor = vbGreen
        If tmrReStart.Enabled = True Then
            If (tmrReStart.Interval <> 30000) Then
                tmrReStart.Enabled = False
                If (tmrReStart.Interval = 5000) Then
                    tmrReStart.Interval = 10000
                Else
                    tmrReStart.Interval = 30000
                End If
                tmrReStart.Enabled = True
            End If
        End If
    End If

End Sub


Private Sub tmrReStart_Timer()

    tmrReStart.Enabled = False
    
    cmdFilt.BackColor = vbGreen

End Sub


Private Sub UserControl_Initialize()

    '''ScanTYPE = 3100  '''LD-LRS-3100,, DPS-2590

    tSrunMode = eSrunMode.InitConn
    
    DRAWmode = 0
    
    cmdFilt.BackColor = vbGreen
    
    rxBcnt = 0
    
    rxSTOP = 0
    
    cnWdeep = 0
    cnWring = 0
    
    TxHeader = Chr(2) + Chr(2) + Chr(2) + Chr(2) + Chr(0) + Chr(0)
    RxHeader = Chr(2) + Chr(2) + Chr(2) + Chr(2) + Chr(0) + Chr(0)
    
    LDinitVAR
    '''''''''
    
    lbMode = tSrunMode
    lbRXcnt = 0
    lbRXerr = 0

    lbCnt.Top = UserControl.Height - 950  ''650
    lbPointErrCnt.Top = Height - 950  ''650
    lbXC.Top = Height - 950  ''650
    lbAngle.Top = Height - 950  ''650
    
    lbRXcnt.Top = Height - 950  ''650
    lbRXerr.Top = Height - 950  ''650
    
    lbMode.Top = UserControl.Height - 950  ''650
    
    ''lbRxHead.Top = UserControl.Height - 1050  ''<-----DPS_GSCN??
    

''    Private maxHH As Long
''    Private baseHH As Long


End Sub

Private Sub autoTilt_on()
    AutoTiltON = True
    AutoTiltOffDelayCnt = 0
    AutoTiltCnt = 0
    AutoTiltErrorCnt = 0
    AutoTiltStep = TiltStep * (-1)
    AutoTiltNow = TiltMax
    AutoTiltMax = TiltMax
    AutoTiltMin = TiltMin
    lbTiltTX.BackColor = &HFF00&
    lbTiltRX.BackColor = &HFF00&
    lbTiltV.BackColor = &HFF00&
End Sub


Private Sub autoTilt_start()
End Sub

Private Sub autoTilt_stop()
    AutoTiltStarted = False
    Tilt3Dlog_end tilt3Dlog_fn
    AutoTiltOffDelayCnt = 6
    lbTiltTX.BackColor = &HC000&
    lbTiltRX.BackColor = &HC000&
    lbTiltV.BackColor = &HC000&
End Sub

Private Sub autoTilt_off()
    AutoTiltON = False
    lbTiltTX.BackColor = &HFF8080
    lbTiltRX.BackColor = &HFF8080
    lbTiltV.BackColor = &HFF8080
End Sub

Private Sub picSiloDrawInit()
Dim i As Integer
    
    minxrange = 0
    maxxrange = 6000  ''20000
    minyrange = 0
    maxyrange = 6000

    picSilo.Scale (minxrange, maxyrange)-(maxxrange, minyrange)

    picSilo.Cls

    picSilo.ForeColor = vbCyan
    picSilo.FillStyle = vbFSSolid
    picSilo.DrawWidth = 1

    picSilo.FillColor = &H404040
    Ellipse 3000, 500, 2500, 250

    picSilo.FillColor = vbCyan
    Ellipse 3000, 5500, 2500, 250

    picSilo.Line (500, 500)-(500, 5500)
    picSilo.Line (5500, 500)-(5500, 5500)

    '''''''''''''''''''''''''''''''''''''''''''''''''(원기둥)''
    If (lbHH.Caption <> "") And (DRAWmode = 0) Then
      If (CInt(lbHH.Caption) > 0) Then
        picSilo.FillStyle = vbFSSolid  ''vbFSTransparent  ''vbHorizontalLine  ''vbFSSolid
        picSilo.FillColor = &H202020
        picSilo.DrawWidth = 1
        
        If baseHH > 50 Then
            picSilo.ForeColor = &H808080               ''vbCyan  ''&H404040  ''vbMagenta  '' vbCyan  ''&H707000
            For i = 25 To (baseHH) Step 25
                Ellipse 3000, 500 + i, 2050, 205  ''2250, 225    ''2480, 248
            Next i
        End If
        
        picSilo.ForeColor = &H4080               ''vbCyan  ''&H404040  ''vbMagenta  '' vbCyan  ''&H707000
        For i = baseHH To baseHH + (CLng(lbHH) * 100) Step 25
            Ellipse 3000, 500 + i, 2050, 205  ''2250, 225    ''2480, 248
        Next i

        DoEvents
      End If
    End If
End Sub


Public Sub picSiloDRAW()

Dim k As Double
Dim s As Double
Dim cut As Integer

Dim i As Integer
Dim j As Integer
Dim d As Integer

Dim SideD As Integer  '';20160617~



    n = xcMax  ''119
    
'    x(1) = 0
'    Y(1) = 0

    minXL = 3000  ''0
    minXR = 3000  ''0
    
    cut = 24   ''(angle*2);;+-12degree==>120~240-->[132~228]=96degree
      
    
    If ScanTYPE = 2590 Then
        cut = 30 '''60
    End If
    
    If (ScanTYPE = 12590) Or (ScanTYPE = 22590) Then
        cut = 30 '''60
    End If
    
    For k = cut To n - cut + 2   '''{0 to n}'''

            s = k / 2#

            X(k) = (rxWORD(k) / 10) * Cos(((s) + 30 + lbAngle) * (PI / 180)) + 3000
    
            ''x(k) = x(k) + Val(txtOpX.Text)
    
            ''y(k) = r(k) * Sin((angle(k) + 40) * (3.14159 / 180)) ''180
            
            ''Y(k) = maxyrange - 500 - (((rxWORD(k) / 10) * Sin(((s) + 30 + lbAngle) * (PI / 180))) * 0.97)
            
            Y(k) = (rxWORD(k) / 10) * Sin(((s) + 30 + lbAngle) * (PI / 180))
''            Y(k) = Y(k) * 0.97
''            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''';;HHHHHHH<---(50M+@):SILO
            
            
            ''Y(k) = maxyrange - 500 - Y(k)   ''; Y(x)=0 ==> 5500
''20171229
            Y(k) = maxyrange - 500 - Y(k)   ''; Y(x)=0 ==> 5500

            
            
            
            If (X(k) < 500) Or (X(k) > 5500) Then GoTo cancleDRAW
            If (Y(k) < 10) Or (Y(k) > 5500) Then GoTo cancleDRAW
            
            
            If (X(k) < minXL) Then
                                    minXL = X(k)
            End If
            If (X(k) > minXR) Then
                                    minXR = X(k)
            End If
            
            If DRAWmode = 0 Then
            
                If k = cut Then
                    X(k - 1) = X(k)
                    Y(k - 1) = Y(k)
                        picSilo.ForeColor = vbRed
                        picSilo.Circle (X(k), Y(k)), 60
                End If
    
                'Draw lines between data points
                picSilo.ForeColor = vbMagenta  ''vbBlue  ''vbRed  ''vbCyan  ''vbBlack
                If k > 0 Then
                    picSilo.Line (X(k - 1), Y(k - 1))-(X(k), Y(k))
                End If
        
                'Plot the data points as circles
                picSilo.ForeColor = vbMagenta  ''vbCyan  ''vbYellow  ''vbMagenta  ''vbBlack
                picSilo.Circle (X(k), Y(k)), 30
            
            End If
            

cancleDRAW:
            ''DoEvents
    Next k
    

    ldWidth = CInt(minXR - minXL)
    picSilo.ForeColor = vbBlue  ''vbMagenta  ''vbBlack
    picSilo.Line (minXL, 5500)-(minXR, 5500)


    If (lbRXcnt > 2) And (AutoTiltON = False) And (cmdFilt.BackColor = vbGreen) Then
    
        Dim YY(300) As Long
        Dim CC As Long
        Dim SS As Long
        Dim s1 As String
        
        '''
        For i = 0 To 300
            YY(i) = 0
        Next i
        
        For i = 0 To 100
            scanDfilt(i) = 0
        Next i
        
        SS = 0
        CC = 0
        
        For i = cut To n - cut + 2   '''{0 to n}'''
    
            If (minXR - minXL) < 3000 Then   ''''';20160617~  3000<--3300

                If (Y(i) < 5500) And (Y(i) > 0) Then
                    
                    YY(i) = Y(i)
                    SS = SS + Y(i)
                    CC = CC + 1
                                        
                    '';; x(k)=(500~5500) y(k)=(500~5500)  ==> x(k)=(0~5000) ; 40M:(1000~5000) => (0~4000)
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 35M:(1250~4750) => (0~3500)
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 34M:(1300~4700) => (0~3400)
''''                    If DRAWmode > 0 Then
                        k = (X(i) - 1300) / 34   ''k = (x(i) - 1250) / 35    ''k = (x(i) - 1000) / 40
                        If k < 0 Then k = 0
                        If k > 100 Then k = 100
                        scanDfilt(CInt(k)) = Y(i)
''''                    End If
                    ''''''''''''''''''''''''''''''''''''''
                    
                    
                End If
                
            Else
                '';Not-Use--Side-1meterDATA
                ''If (Y(i) < 5500) And (Y(i) > 0) And (X(i) > (minXL + 100)) And (X(i) < (minXR - 100)) Then
                
                ''(2012-12-28::ZeroFilt...Null-Center!!)
                ''If (Y(i) < 5500) And (Y(i) > 0) And (X(i) > (minXL + 10)) And (X(i) < (minXR - 10)) Then
                '';??????---(10)????
                
                '';Not-Use--Side-0.5meterDATA '';20160617~
                SideD = 50 ''50
                If ScanTYPE = 2590 Then
                    SideD = 400 ''50
                End If
''''20171229~
                If (ScanTYPE = 12590) Or (ScanTYPE = 22590) Then
                    SideD = 200 '''''''''20180105
                End If
                ''
                If (Y(i) < 5500) And (Y(i) > 0) And (X(i) > (minXL + SideD)) And (X(i) < (minXR - SideD)) Then
                
                    
                    YY(i) = Y(i)
                    SS = SS + Y(i)
                    CC = CC + 1
                    
                    ''''''''''''''''''''''''''''''''''''''
''''                    If DRAWmode > 0 Then
                        k = (X(i) - 1300) / 34   ''k = (x(i) - 1250) / 35    ''k = (x(i) - 1000) / 40
                        If k < 0 Then k = 0
                        If k > 100 Then k = 100
                        scanDfilt(CInt(k)) = Y(i)
''''                    End If
                    ''''''''''''''''''''''''''''''''''''''
                    
                End If
                
            End If
            
            ''DoEvents
        Next i


''''        If DRAWmode > 0 Then
            
            For i = 51 To 95   '''0~100
                If scanDfilt(i) = 0 Then
                    k = 0
                    s = 0
                    For j = 1 To 5
                        If scanDfilt(i + j) > 0 Then
                            s = s + scanDfilt(i + j):   k = k + 1
                        End If
                        If scanDfilt(i - j) > 0 Then
                            s = s + scanDfilt(i - j):   k = k + 1
                        End If
                    Next j
                    
                    If k > 0 Then scanDfilt(i) = s / k
                    
                End If
            Next i
            
            For i = 50 To 5 Step -1  '''0~100
                If scanDfilt(i) = 0 Then
                    k = 0
                    s = 0
                    For j = 1 To 5
                        If scanDfilt(i + j) > 0 Then
                            s = s + scanDfilt(i + j):   k = k + 1
                        End If
                        If scanDfilt(i - j) > 0 Then
                            s = s + scanDfilt(i - j):   k = k + 1
                        End If
                    Next j
                    
                    If k > 0 Then scanDfilt(i) = s / k
                    
                End If
            Next i
            
            If scanDfilt(4) = 0 Then scanDfilt(4) = scanDfilt(5)
            If scanDfilt(3) = 0 Then scanDfilt(3) = scanDfilt(4)
            If scanDfilt(2) = 0 Then scanDfilt(2) = scanDfilt(3)
            If scanDfilt(1) = 0 Then scanDfilt(1) = scanDfilt(2)
            If scanDfilt(0) = 0 Then scanDfilt(0) = scanDfilt(1)
            
            If scanDfilt(96) = 0 Then scanDfilt(96) = scanDfilt(95)
            If scanDfilt(97) = 0 Then scanDfilt(97) = scanDfilt(96)
            If scanDfilt(98) = 0 Then scanDfilt(98) = scanDfilt(97)
            If scanDfilt(99) = 0 Then scanDfilt(99) = scanDfilt(98)
            If scanDfilt(100) = 0 Then scanDfilt(100) = scanDfilt(99)
            
            
            If DRAWmode > 0 Then
            
                For i = 0 To 100
''                    If scanDfilt(i) = 0 Then
''                        '''scanDfilt(i) = s
''                    End If
                    
                    If scanDfilt(i) > 0 Then
                    
                        picSilo.ForeColor = &H4080   ''vbGreen  ''vbMagenta  ''vbBlue  ''vbRed  ''vbCyan  ''vbBlack
                        picSilo.DrawWidth = 2
                        picSilo.Line (i * 40 + 1000, 500)-(i * 40 + 1000, scanDfilt(i))
                    Else
                        Debug.Print "scanDfilt---Error: ", UCindex, i
                    End If
                    
                    scanDfiltX(i) = scanDfilt(i)
                    ''''''''''''''''''''''''''''
                Next i
            
            End If

            For i = 0 To 100
                scanDfiltX(i) = scanDfilt(i)
            Next i

''''        End If
        

        ''Debug.Print s1
        txtAVRheight = 0
        
        If (CC > 0) And (SS > 0) Then
        
            txtAsum = SS
            txtAcnt = CC
            If CInt(SS / CC) < 500 Then
                txtAVRheight = 0
            Else
                txtAVRheight = CInt(SS / CC) - 500
            End If
        
            picSilo.ForeColor = vbBlue  ''vbYellow  ''vbMagenta  ''vbBlack
            picSilo.Line (2500, txtAVRheight + 500)-(3500, txtAVRheight + 500)
            
            picSilo.ForeColor = vbRed
            picSilo.Line (2500, avrHH + 500 + baseHH)-(3500, avrHH + 500 + baseHH)
        
        
        End If
    
        
        txtScaleHight = txtAVRheight - baseHH
        '''''''''''''''''''''''''''''''''''''
        If txtScaleHight < 0 Then txtScaleHight = 0
        
        
''        ''txtAOd = CLng((txtAVRheight / (maxHH - baseHH)) * 32767)
''        txtAOd = CLng((txtAVRheight / 5000) * 32767)
        txtAOd = CLng((txtScaleHight / (maxHH - baseHH)) * 32768)
        ''''''
        If txtAOd < 1 Then txtAOd = 1
        If txtAOd > 32767 Then txtAOd = 32767


        ''=================================================================USE:avrHH===!!

        If avrHH >= (maxHH - baseHH) Then ''Silo
            lbHP.Caption = "100"  ''//20120730
        Else
            lbHP.Caption = Format((avrHH / (maxHH - baseHH)) * 100, "#0.0")  ''//20120813
        End If
        
        lbHH = Format((avrHH / 100), "#0.00")
        
        lbAO = Format((avrHH / (maxHH - baseHH)) * 16 + 4, "#0.00") ''SILO
        
        ''VV==PI*(R*R)*Hight ''원기둥''
        ''SS = CDbl(PI) * (20.5 * 20.5) * (baseHH / 100#)
        lbVVV = Format(CDbl(PI) * (20.5 * 20.5) * ((baseHH + avrHH) / 100#), "#0")   ''SILO
        
        
    End If


    picSilo.DrawWidth = 1
    
    picSilo.ForeColor = vbWhite  ''vbBlack
    picSilo.Line (2500, 500)-(3500, 500)  ''picSilo.Line (500, 500)-(5500, 500)
    picSilo.Line (500, 5500)-(5500, 5500)

    picSilo.Line (3000, 500)-(3000, 5500)


    picSilo.ForeColor = vbWhite  ''vbBlack
    picSilo.Line (2000, 500)-(4000, 500)
    For i = 1 To 4
        picSilo.Line (2800, 500 + (1000 * i))-(3100, 500 + (1000 * i))

        'Labeling the yaxis
        picSilo.CurrentX = 3120
        picSilo.CurrentY = 500 + (1000 * i) + 100
        picSilo.Print Trim(i * 10) & "M"
    Next i

    For i = 1 To 9
        picSilo.Line (2900, 500 + (500 * i))-(3100, 500 + (500 * i))
    Next i


    picSilo.ForeColor = vbYellow  ''vbWhite  ''vbBlack  ''vbRed
    picSilo.Line (1000, 500 + maxHH)-(2200, 500 + maxHH)
    picSilo.Line (3800, 500 + maxHH)-(5000, 500 + maxHH)
    picSilo.Line (1000, 500 + baseHH)-(2200, 500 + baseHH)
    picSilo.Line (3800, 500 + baseHH)-(5000, 500 + baseHH)


    '';; {100~4500} => {0~4400}
    scaleHH = CLng(maxHH - baseHH)    ''''NotUse-Yet~
    If scaleHH < 0 Then scaleHH = 0

End Sub



Public Function ret_AOd() As Integer
    ret_AOd = Val(txtAOd)
End Function


Public Function ret_Act() As Integer
    
    ''If (wsACT = True) And (tSrunMode >= eSrunMode.SendCmd) Then
    If (tSrunMode >= eSrunMode.SendCmd) Then
        ret_Act = 1
    Else
        ret_Act = 0
    End If
End Function


Public Function ret_HH() As Integer  ''[0-50.000M]==>((0~5000))
    If lbHH <> "" Then
        ret_HH = CInt(Val(lbHH) * 100)  ''CInt(Val(lbHH) * 1000)
    Else
        ret_HH = 0
    End If
End Function

''''체적:[0-100,000VVV]==>((0~10000))
Public Function ret_VV() As Integer  ''Long  ''Integer
    If lbVVV <> "" Then
        If CInt(Val(lbVVV) / 10) < 0 Then
            ret_VV = 0
        Else
            ret_VV = CInt(Val(lbVVV) / 10)
        End If
    Else
        ret_VV = 0
    End If
End Function


Public Function GETscanD(ang As Integer) As Integer
    GETscanD = CInt(scanDfiltX(ang))   ''/ 10) '' / 10)
End Function



Public Sub setIDX(id As Integer, ip As String, port As String)
    
    UCindex = id

    cmdCONN.Caption = "S" & Format(id + 1, "00")
    If (id > 14) Then
    
        cmdCONN.Caption = "S" & Format(34 - id, "00")  ''20180222 add edite
    
        cmdCONN.Caption = cmdCONN.Caption & "-CTS" & Trim(id - 14)    ''New-CTS-Silo(15+4)!!  ''4x2==8''??
        
        
    End If

    ipAddr = ip
    ipPort = port
    
    If ip <> "" Then wsockLD.RemoteHost = ip
    If port <> "" Then wsockLD.RemotePort = port
    
    ''tmrRun.Enabled = True
    '''''''''''''''''''''
End Sub


Public Function getTXT() As String
    
    getTXT = txtRx1
    
End Function

Sub Ellipse(X As Single, Y As Single, RadiusX As Single, RadiusY As Single)
  Dim ratio As Single, Radius As Single
    ratio = RadiusY / RadiusX
    If ratio < 1 Then
        Radius = RadiusX
    Else
        Radius = RadiusY
    End If
    picSilo.Circle (X, Y), Radius, , , , ratio
End Sub


Public Sub initStart()
'
    Dim i As Integer
'
    RX_filt_Init
'
    tSrunMode = eSrunMode.InitConn
'
    tmrSrun.Interval = 1000
    tmrSrun.Enabled = True
'
End Sub


Private Sub tmrSrun_Timer()
Dim ret As Integer
Dim strA As String
Dim bb() As Byte
'Dim t As Long
    
    't = GetTickCount
    'DGPSLog "tmrSrun_Timer(" & UCindex & ") START " & tSrunMode & "", "SILO"

    tmrSrun.Enabled = False
    '''''''''''''''''''''''
    
    lbMode = tSrunMode

    Select Case tSrunMode
        Case InitConn
            lbRXerr = 0
            lbRXcnt = 0
        
            cmdCONN.BackColor = vbRed
            
            CONN_wsockLD
            
            tSrunMode = eSrunMode.CheckConn
            
            tmrSrun.Interval = 2000
            tmrSrun.Enabled = True
            
            picSiloDrawInit  ''Blank
            
        Case CheckConn
            tmrSrun.Interval = 2000
            
            If wsockLD.State <> sckConnected Then
                tSrunMode = eSrunMode.InitConn
            Else
                cmdCONN.BackColor = &HFF80FF   ''<핑크  ''vbBlue
                tSrunMode = Init1Device
            End If

            tmrSrun.Enabled = True
            
        Case Init1Device
            If ScanTYPE = 22590 Then
                ''ret = LDtx12590(53) ''' Red laser marker status temporary on
                ''ret = LDtx12590(55) ''' Red laser marker status temporary off
                ret = LDtx12590(51)   '''DPS_SPRM_SDC4 SDC(Scan Data Content) to 4(=distances olny)
            
                tSrunMode = eSrunMode.Check1Device  '''==>
                rxWaitTime = 0

                tmrSrun.Interval = 100

            ElseIf ScanTYPE = 12590 Then  '''DPS-2590::12590
                ret = LDtx12590(45)   '''DPS-2590:BIN-Mode!!  "SCAN"--Run!!
            
                tSrunMode = eSrunMode.Check1Device
                rxWaitTime = 0
            
                tmrSrun.Interval = 100

            ElseIf ScanTYPE = 2590 Then  '''DPS-2590
                ret = LDtx2590(41)   '''DPS-2590:Terminal-Mode!!
            
                tSrunMode = eSrunMode.Check1Device
                rxWaitTime = 0
                
                tmrSrun.Interval = 100

            Else
                If LDinitTXs <> 0 Then  '''LD-LRS-3100
                ''''''''''''''''''''''
                    tSrunMode = eSrunMode.InitConn
                    
                    tmrSrun.Interval = 2000

                Else
                    tSrunMode = eSrunMode.Check1Device
                    rxWaitTime = 0
                    
                    tmrSrun.Interval = 2000

                End If
            
            End If
                        
            tmrSrun.Enabled = True

        Case Check1Device
            rxWaitTime = rxWaitTime + tmrSrun.Interval
            If ScanTYPE = 22590 Then
                ret = LDrx12590(51)   '''DPS_SPRM_SDC4 SDC(Scan Data Content) to 4(=distances olny)
                
                If ret = 0 Or ret < 0 Then
                    lbTiltV = "0"
                    SEND_wsickLD "AutoTiltON"  '''[[ "AutoTiltON"=="TorqueMax(200)"&"SetAngle[0]" ]]
                
                    tSrunMode = eSrunMode.Init2Device  '''==>
                    
                    tmrSrun.Interval = 1000
                End If

            ElseIf ScanTYPE = 12590 Then  '''DPS-2590::12590
                ret = LDrx12590(45)   '''DPS-2590:BIN-Mode!!  "SCAN"--Run!!
            
                If ret = 0 Or ret < 0 Then
                    tSrunMode = eSrunMode.SendCmd
                
                    tmrSrun.Interval = 1000
                End If
            
            ElseIf ScanTYPE = 2590 Then  '''DPS-2590
                ret = LDrx2590(41)   '''DPS-2590:Terminal-Mode!!
            
                If ret = 0 Or ret < 0 Then
                    tSrunMode = eSrunMode.SendCmd
                    
                    tmrSrun.Interval = 2000
                End If

            Else
                tSrunMode = eSrunMode.SendCmd
                tmrSrun.Interval = 2000
                
            End If
                        
            tmrSrun.Enabled = True

        Case Init2Device
            If ScanTYPE = 22590 Then
                ret = LDtx12590(45)   '''DPS-2590:BIN-Mode!!  "SCAN"--Run!!
                
                tSrunMode = eSrunMode.Check2Device
                rxWaitTime = 0
                
                tmrSrun.Interval = 100
            
            Else
                tSrunMode = eSrunMode.InitConn
                
                tmrSrun.Interval = 2000
            
            End If
    
            ''''
            tmrSrun.Enabled = True

        Case Check2Device
            rxWaitTime = rxWaitTime + tmrSrun.Interval
            If ScanTYPE = 22590 Then
                ret = LDrx12590(45)   '''DPS-2590:BIN-Mode!!  "SCAN"--Run!!
                
                If ret = 0 Or ret < 0 Then
                    lbTiltV = Str(TiltDefault)
                    strA = "SetAngle[" & TiltDefault & "]"
    
                    bb = StrConv(strA, vbFromUnicode)
                    ''
                    SEND_wsickLD bb
                    '''''''''''''''
                    
                    tSrunMode = eSrunMode.SendCmd
                    
                    tmrSrun.Interval = 2000 '''500
                End If
            
            Else
                tSrunMode = eSrunMode.InitConn
                
                tmrSrun.Interval = 2000
            
            End If
    
            ''''
            tmrSrun.Enabled = True

        Case SendCmd
            lbCnt = "TX"

            ''picSiloDrawInit

            txtRx1 = ""

            If ScanTYPE = 22590 Then

                '''''''''
                ret = LDtx12590(47)   ''''''DPS-2590:BIN-Mode!!  "GSCN"--Scan#1 !!

                tSrunMode = eSrunMode.ReceiveData
                rxWaitTime = 0

                tmrSrun.Interval = 100

            ElseIf ScanTYPE = 12590 Then  '''DPS-2590::12590
                ret = LDtx12590(47)   ''''''DPS-2590:BIN-Mode!!  "GSCN"--Scan#1 !!

                tSrunMode = eSrunMode.ReceiveData
                rxWaitTime = 0

                tmrSrun.Interval = 100

            ElseIf ScanTYPE = 2590 Then   '''''''''''''''''''' DPS-2590
                ret = LDtx2590(43)  '''43<--49 for DPS2590--12590
                                    '''Console: "s\r\n"  : Array(Asc("s"), &HD, &HA)

                tSrunMode = eSrunMode.ReceiveData
                rxWaitTime = 0

                tmrSrun.Interval = 100

            Else  ''''''''''''''''''''LD-LRS-3100
                ret = LDtxDATA(39)
                ''''''''''''''''''
                    '            If ret < 0 Then
                    '                ret = LDtxDATA(39)
                    '            End If

                If ret < 0 Then
                    tmrSrun.Interval = 1000

                    tSrunMode = eSrunMode.InitConn

                Else
                    tmrSrun.Interval = 10

                    tSrunMode = eSrunMode.ReceiveData
                    rxWaitTime = 0

                End If
            
            End If
                            
            tmrSrun.Enabled = True
        
        Case ReceiveData
            rxWaitTime = rxWaitTime + tmrSrun.Interval
'
            If ScanTYPE = 22590 Then

                '''''''''
                ret = LDrx12590(47)   ''''''DPS-2590:BIN-Mode!!  "GSCN"--Scan#1 !!
'
                If ret = 0 Then
                    If AutoTiltON = True Then
                        If AutoTiltOffDelayCnt <> 0 Then
                            AutoTiltOffDelayCnt = AutoTiltOffDelayCnt - 1
                            If AutoTiltOffDelayCnt = 0 Then
                                autoTilt_off
                            End If
'
                            lbTiltV = Str(TiltDefault)
                            strA = "SetAngle[" & TiltDefault & "]"
                            tmrSrun.Interval = 1000
                        ElseIf AutoTiltStarted = False Then
                            RX_filt_Init
                            tilt3Dlog_fn = _
                                Tilt3Dlog_start( _
                                    cmdCONN.Caption, _
                                    " " & lbCenterX _
                                    & " " & lbCenterY _
                                    & " " & lbRadius)
'
                            AutoTiltStarted = True
'
                            strA = Trim(Str(CInt(AutoTiltNow * 100) / 100))
                            lbTiltV.Caption = strA
                            strA = "SetAngle[" & strA & "]"
                            ''
                            tmrSrun.Interval = 2000
                        Else
                            AutoTiltErrorCnt = 0
                            AutoTiltNow = AutoTiltNow + AutoTiltStep
                            AutoTiltCnt = AutoTiltCnt + 1
                            ''''''''''''''''''''''''''''''''''''''''
                            If AutoTiltNow > AutoTiltMax Or AutoTiltNow < AutoTiltMin Then
                                autoTilt_stop
'
                                lbTiltV = Str(TiltDefault)
                                strA = "SetAngle[" & TiltDefault & "]"
                                tmrSrun.Interval = 1000
                            Else
                                strA = Trim(Str(CInt(AutoTiltNow * 100) / 100))
                                lbTiltV = strA
                                strA = "SetAngle[" & strA & "]"
                                ''
                                tmrSrun.Interval = 1000
                            End If
                        End If
                    Else
                        lbTiltV = Str(TiltDefault)
                        strA = "SetAngle[" & TiltDefault & "]"
                        tmrSrun.Interval = 2000
                    End If
'
                    bb = StrConv(strA, vbFromUnicode)
                    ''
                    SEND_wsickLD bb
                    '''''''''''''''
'
                    tSrunMode = eSrunMode.SendCmd
'
                    'tmrSrun.Interval = 2000  '''100  ''1000 ''after-Done!! ''2000  ''<===!
                ElseIf ret < 0 Then
                    If AutoTiltStarted = True Then
                        AutoTiltErrorCnt = AutoTiltErrorCnt + 1
                        If AutoTiltErrorCnt > 5 Then
                            autoTilt_stop
'
                            lbTiltV = Str(TiltDefault)
                            strA = "SetAngle[" & TiltDefault & "]"
                            tmrSrun.Interval = 1000
                        Else
                            strA = Trim(Str(CInt(AutoTiltNow * 100) / 100))
                            lbTiltV.Caption = strA
                            strA = "SetAngle[" & strA & "]"
                            ''
                        End If
                        tmrSrun.Interval = 1000
                    Else
                        lbTiltV = Str(TiltDefault)
                        strA = "SetAngle[" & TiltDefault & "]"
                        tmrSrun.Interval = 2000
                    End If
'
                    bb = StrConv(strA, vbFromUnicode)
                    ''
                    SEND_wsickLD bb
                    '''''''''''''''
'
                    tSrunMode = eSrunMode.SendCmd
                    
                    'tmrSrun.Interval = 2000  '''100  ''1000 ''after-Done!! ''2000  ''<===!
                End If
'
            ElseIf ScanTYPE = 12590 Then  '''DPS-2590::12590
                ret = LDrx12590(47)   ''''''DPS-2590:BIN-Mode!!  "GSCN"--Scan#1 !!
                
                If ret = 0 Or ret < 0 Then
                    tSrunMode = eSrunMode.SendCmd
                    
                    tmrSrun.Interval = 2000  ''1000 ''after-Done!! ''2000  ''<===!
                End If
                
            ElseIf ScanTYPE = 2590 Then   '''''''''''''''''''' DPS-2590
                ret = LDrx2590(43)  '''43<--49 for DPS2590--12590
                                    '''Console: "s\r\n"  : Array(Asc("s"), &HD, &HA)
            
                If ret = 0 Or ret < 0 Then
                    tSrunMode = eSrunMode.SendCmd
                    
                    tmrSrun.Interval = 2000  ''8000  ''9000
                End If

            Else  ''''''''''''''''''''LD-LRS-3100
                ret = LDrxDATA(39)
                ''''''''''''''''''
                    '            If ret < 0 Then
                    '                ret = LDrxDATA(39)
                    '            End If

                If ret < 0 Then
                    tSrunMode = eSrunMode.InitConn
                
                    tmrSrun.Interval = 1000
                
                Else
                    tSrunMode = eSrunMode.SendCmd
                
                    tmrSrun.Interval = 1000  ''400
                
                End If
            
            End If
                            
            tmrSrun.Enabled = True
        
        Case Else
            tSrunMode = eSrunMode.InitConn
            tmrSrun.Interval = 2000
            tmrSrun.Enabled = True
    
    End Select
    
    't = GetTickCount - t
    'DGPSLog "tmrSrun_Timer(" & UCindex & ") END " & tSrunMode & ", " & t & "", "SILO"

End Sub


Private Sub tmrTRX_Timer()
    tmrTRX.Enabled = False
End Sub


Private Sub CONN_wsockLD()

    If wsockLD.State <> sckConnected Then

        wsockLD.Close

        wsockLD.Connect
    
    End If

End Sub



Private Sub wsockLD_DataArrival(ByVal bytesTotal As Long)

    Dim buffData As Variant     ''This stores the incoming data from the buffer
    Dim i, j, c As Integer      ''These are general counters
    
    Dim cCNT As Long
    Dim buff2590 As String
    
    
    
    If (ScanTYPE = 12590) Or (ScanTYPE = 22590) Then
    
        wsockLD.GetData buffData  ''';BIN-Mode!
        '''''''''''''''''''''''''
        ''DoEvents   ''' Comment out to avoid out of stack space error.
        
        If rxSTOP > 0 Then
            Exit Sub  ''''===============>>
        End If
        
        '''If (bytesTotal + inCNT) > 9999 Then  ''1990<--1999 ;Mask...Error!
        If (bytesTotal + inCNT) > 8086 Then  '' ;Mask...Error!  8056+30 (201809~DPSex)
            inCNT = 0
            Exit Sub  ''''===============>>
        End If
        
        For i = 0 To bytesTotal - 1
            inBUF(inCNT + i) = buffData(i)
        Next i
        
        inCNT = inCNT + bytesTotal
        ''''''''''''''''''''''''''
        lbCnt = inCNT
        
        '''If (inCNT = 4056) Or (inCNT >= 8056) Then   ''''''<=="GSCN"
        If (inCNT = 4056) Or (inCNT = 4086) Or (inCNT >= 8056) Then  '''(201809~DPSex)
            rxSTOP = 1          ''''''=======>
            
            Exit Sub  ''''===============>>
            
        End If
    
        ''DoEvents   ''' Comment out to avoid out of stack space error.
        
        Exit Sub  ''''===============>>

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf ScanTYPE = 2590 Then  '''DPS-2590

        wsockLD.GetData buff2590$
        '''''''''''''''''''''''''

        If rxSTOP > 0 Then
            Exit Sub  ''''''''''''''''':===>
        End If
        

        If (bytesTotal + inCNT) > 90000 Then  ''1990<--1999 ;Mask...Error!
            inCNT = 0
            Exit Sub  '''''''===>
        End If
        
        
        cCNT = bytesTotal
        
        ''If InStr(buff2590, "      1 ;       1 ;") > 0 Then
        If InStr(buff2590, "      1 ;       ") > 0 Then
            
            buff2590 = Mid(buff2590, InStr(buff2590, "      1 ;       "))
            cCNT = Len(buff2590)
            inBUF2590 = ""
            inCNT = 0
            
        End If
        
        inBUF2590 = inBUF2590 & buff2590
        ''''''''''''''''''''''''''''''''
        inCNT = inCNT + cCNT
        lbCnt = inCNT
        
        If InStr(inBUF2590, "   1000 ;       ") > 0 Then
            
            If inCNT > (InStr(inBUF2590, "   1000 ;       ") + 62) Then
            
                inBUF2590 = Left(inBUF2590, InStr(inBUF2590, "   1000 ;       ") + 62)
                
                inCNT = Len(inBUF2590)
                lbCnt = inCNT
            
            End If
            
        End If
        
        ''txtRxS.Text = inBUF2590  ''buff2590
''        If Len(inBUF2590) > 62 Then
''            txtRxS.Text = Mid(inBUF2590, Len(inBUF2590) - 62)
''        End If
        
        If inCNT >= 62998 Then  ''''''<==63000=[63*1000]
            rxSTOP = 1  ''''''===>
        End If
    
        DoEvents
        Exit Sub  '''''''===>
        
    Else
    ''''LD-LRS-3100
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        wsockLD.GetData buffData
        ''DoEvents
    
    End If
    


    If rxSTOP > 0 Then
        Exit Sub
        '''''''''''''''''===>
    End If


    If (bytesTotal + inCNT) > 1990 Then  ''1990<--1999 ;Mask...Error!
        inCNT = 0
        Exit Sub ''GoTo exit1  ''===============>>
    End If

''    If inCNT = 0 Then
''        txtRx1 = txtRx1 & vbCrLf & "RX:"
''    End If

    For i = 0 To bytesTotal - 1
        inBUF(inCNT + i) = buffData(i)

        ''txtRx1 = txtRx1 & Format(Trim(Hex(buffData(i))), "##") & " "
        ''
''        If buffData(i) < 16 Then
''            txtRx1 = txtRx1 & "0" & Hex(buffData(i)) & " "
''        Else
''            txtRx1 = txtRx1 & Hex(buffData(i)) & " "
''        End If

    Next i
    
    inCNT = inCNT + bytesTotal
    ''''''''''''''''''''''''''
    lbCnt = inCNT


    If inCNT >= 516 Then    '''If inCNT = 516 Then
        rxSTOP = 1
        '''''''''''''''''===>
    End If

    ''DoEvents   ''' Comment out to avoid out of stack space error.
    
    
'            c = bytesTotal  ''UBound(buffData)
'
'
'            txtRx1 = txtRx1 & vbCrLf & "RX:"
'            For i = 0 To c - 1
'
'                ''RxMSG = RxMSG & Chr(buffData(i))  ''chr()?
'
'                rxBYTE(rxBcnt + i) = buffData(i)
'
'                txtRx1 = txtRx1 & Hex(buffData(i)) & " "
'
'            Next i
'
'            rxBcnt = rxBcnt + c
    
End Sub

Private Sub wsockLD_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    cmdCONN.BackColor = vbRed
    tSrunMode = eSrunMode.InitConn
End Sub


Private Sub SEND_wsickLD(SD As Variant)
    If wsockLD.State = sckConnected Then
        wsockLD.SendData SD
    Else
        cmdCONN.BackColor = vbRed
        tSrunMode = eSrunMode.InitConn
    End If
End Sub



Private Function LDinitTXs() As Integer

Dim i As Integer

    LDinitVAR

    txtRx1 = ""
    
    For i = 0 To 3
        LDtxDATA (0)
        Sleep (10)
        rxWaitTime = 10
        If LDrxDATA(0) = 0 Then Exit For
        ''''''''''''''''''''''''''''''''
    Next i
    
    If i > 2 Then
        LDinitTXs = -1
        Exit Function ''---->TX-RX:ERROR!
    End If
    
    
    For i = 0 To 29  ''DWONLOAD-ALL!!
        LDtxDATA i
        Sleep (10)
        rxWaitTime = 10
        LDrxDATA i
    Next i
    
    
    ''LDtxDATA 29
    ''Sleep (10)
    ''rxWaitTime = 10
    ''LDrxDATA 29

    LDinitTXs = 0
    
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''DPS-2590::BIN-Mode!
Private Function LDtx12590(ix As Integer) As Integer

Dim i As Integer
Dim tBuf As Variant

    ReDim tBuf(UBound(LD_sBUF(ix))) As Byte
    ''
    For i = 0 To UBound(LD_sBUF(ix))
        tBuf(i) = (LD_sBUF(ix)(i))
    Next i
    
    RxMSG = ""
    
    inCNT = 0   ''ws
    xcMax = 0   ''scanData!!
    
    rxSTOP = 0  '''''''''''''''''''''''''''''''''<==!!
    
    '''inBUF2590 = ""
    
    
    SEND_wsickLD tBuf
    '''''''''''''''''

    LDtx12590 = 0

End Function


Private Function LDrx12590(ix As Integer) As Integer

Dim i As Integer
Dim pointErr As Integer
Dim rxAngleOK As Boolean
Dim rxAngle#
'Dim t As Long

    If rxWaitTime < 2500 And rxSTOP = 0 Then
        LDrx12590 = 1
        Exit Function  ''===>
    End If
    
''    For i = 0 To 1000  '''1 To 239
''        rxWORD(i) = 0
''    Next i


'''    ''char DPS_SCAN [16] ; LD_sBUF(45) = Array(&H53, &H43, &H41, &H4E, &H0, &H0, &H0, &H4, &H0, &H0, &H0, &H10, &HF0, &H85, &H33, &HD2)
'''    ''char DPS_GSCN [16] ; LD_sBUF(47) = Array(&H47, &H53, &H43, &H4E, &H0, &H0, &H0, &H4, &H0, &H0, &H0, &H0, &H48, &H2F, &HE1, &HC3)
'''
'''    '''''''''''''ERR???? ; LD_sBUF(49) = Array(&H45, &H52, &H52, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    'DGPSLog "LDrx12590(" & UCindex & ") rxWaitTime=" & rxWaitTime & ", rxSTOP=" & rxSTOP & "", "SILO"
    'DGPSLog "LDrx12590(" & UCindex & ") inCNT=" & inCNT & "", "SILO"

    If inCNT < 22 Then   ''''''<=="ERR///SCAN"---"GSCN"
    
        If (inBUF(0) = LD_sBUF(49)(0)) And (inBUF(1) = LD_sBUF(49)(1)) And (inBUF(2) = LD_sBUF(49)(2)) Then
        ''''ERR????;;
            lbRxHead.Caption = "ERR"
            lbRxHead.BackColor = vbRed
            
            LDrx12590 = -1
            Exit Function  ''===>
        
        ElseIf (inBUF(0) = LD_sBUF(45)(0)) And (inBUF(1) = LD_sBUF(45)(1)) And (inBUF(2) = LD_sBUF(45)(2)) Then
        ''''SCAN---;;
            lbRxHead.Caption = "SCAN"
            lbRxHead.BackColor = &HFFC0C0
            
        ElseIf (inBUF(0) = LD_sBUF(47)(0)) And (inBUF(1) = LD_sBUF(47)(1)) And (inBUF(2) = LD_sBUF(47)(2)) Then
        ''''GSCN---;;
            lbRxHead.Caption = "GSCN"
            lbRxHead.BackColor = &HFFC0C0
            
        ElseIf (inBUF(0) = LD_sBUF(51)(0)) And (inBUF(1) = LD_sBUF(51)(1)) And (inBUF(2) = LD_sBUF(51)(2)) Then
        ''''SPRM---;;
            lbRxHead.Caption = "SPRM"
            lbRxHead.BackColor = vbYellow  ''&HFFC0C0
            
        End If
        
        '''?'' inCNT = 0   ''ws
        '''?'' xcMax = 0   ''scanData!!
    End If
    

    If (inCNT < 30) Then
        lbCnt.BackColor = vbYellow  ''&HC0C0C0
            LDrx12590 = -2
            Exit Function  ''===>
    ElseIf (inCNT < 4056) Or ((inCNT > 4086) And (inCNT < 8056)) Then
        lbCnt.BackColor = vbRed  ''&HC0C0C0
            LDrx12590 = -3
            Exit Function  ''===>
    End If

    't = GetTickCount
    'DGPSLog "LDrx12590(" & UCindex & ") START " & inCNT & "", "SILO"
    
    If (inCNT = 4056) Or (inCNT = 4086) Then   ''''''<=="GSCN"  ''(201809~DPSex::++30)
        Dim dataCRC32 As Long
        Dim calcCRC32 As Long
        
        dataCRC32 = GetLong4Bytes(inBUF, 4056 - 4)
        calcCRC32 = GetCrc32(inBUF, 0, 4056 - 4)

        If (calcCRC32 <> dataCRC32) Then
            lbRxHead.Caption = "ECRC"
            lbRxHead.BackColor = vbRed
            lbRXerr = lbRXerr + 1
            LDrx12590 = -4
            Exit Function  ''===>
        End If
        
        ''rxSTOP = 1          ''''''=======>
        
        lbCnt.BackColor = &HC0C0C0
        
        If (inBUF(0) = LD_sBUF(47)(0)) And (inBUF(1) = LD_sBUF(47)(1)) And (inBUF(2) = LD_sBUF(47)(2)) Then
        ''''GSCN---;;
            lbRxHead.Caption = "GSCN"
            lbRxHead.BackColor = &HFF8080
            
            ''''cmdCONN.BackColor = vbGreen
            
            If (cmdFilt.BackColor <> vbGreen) And (AutoTiltStarted <> True) Then
                SaveBuffer2File cmdCONN.Caption & "_raw_", inBUF, inCNT
            End If
            
            Dim angleN As Integer
            '''
            pointErr = 0
            
            For i = 0 To 999  '''1000
                Scan2590Direc(i) = 45# + (i * 0.09)
                ''
                angleN = (i * 4) + 52  '''<==(i * 8)
                Scan2590Dist(i) = (inBUF(angleN) * 2 ^ 24) + (inBUF(angleN + 1) * 2 ^ 16) + (inBUF(angleN + 2) * 2 ^ 8) + inBUF(angleN + 3)
                'Scan2590Dist(i) = Scan2590Dist(i) * 10
                ''angleN = angleN + 4
                ''Scan2590Pulse(i) = (inBUF(angleN) * 2 ^ 24) + (inBUF(angleN + 1) * 2 ^ 16) + (inBUF(angleN + 2) * 2 ^ 8) + inBUF(angleN + 3)
                ''
                '' Check the distance value is -2147483648 (0x80000000) in case that the echo signal was too low.
                '' or, the distance value is 2147483647 (0x7FFFFFFF) in case that the echo signal was noisy.
                '' And Check the distance is near to 5meters.
                '' 2147483646 <= 2147483647 - 1
                If (Scan2590Dist(i) > 2147483646) _
                    Or (Scan2590Dist(i) < 50000) _
                        Then
                    Scan2590Dist(i) = 0
                    pointErr = pointErr + 1
                End If
                'If (UCindex = 0) Then
                '    Scan2590Dist(i) = Scan2590Dist(i) * 10
                'ElseIf (UCindex >= 4) Then
                '    Scan2590Dist(i) = Scan2590Dist(i) * 4
                'End If
            Next i
            
            lbPointErrCnt = pointErr
            
            ''                주소    크기    명칭                의미                    접근    기본값
            ''                --------------------------------------------------------------
            ''        <RX-Read>29Bytes!
            ''        0 0xff
            ''        1 0xff
            ''        2 0x01(id)
            ''        3 0x19(25)
            ''        4 Err -Stat
            ''        5   0   24  1   Torque Enable   토크 켜기               RW  0
            ''        6   1   25  1   LED             LED On/Off          RW  0
            ''        7   2   26  1   D Gain          Derivative Gain     RW  0
            ''        8   3   27  1   I Gain          Integral Gain       RW  0
            ''        9   4   28  1   P Gain          Proportional Gain   RW  32
            ''        10  5   29 ???
            ''        11  6   30  2   Goal Position   목표 위치 값의 바이트   RW  -
            ''        13  8   32  2   Moving Speed    목표 속도 값의 바이트   RW  -
            ''        15  10  34  2   Torque Limit    토크 한계 값의 바이트   RW  ADD 14\&15
            ''        17  12  36  2   Present Positio 현재 위치 값의 바이트   R   -
            ''        19  14  38  2   Present Speed   현재 속도 값의 바이트   R   -
            ''        21  16  40  2   Present Load    현재 하중 값의 바이트   R   -
            ''        23  18  42  1   Present Voltage 현재 전압               R   -
            ''        24  19  43  1   Present Temper  현재 온도               R   -
            ''        25  20  44  1   Registered      Instruction의 등록여부  R   0
            ''        26  21  45 ???
            ''        27  22  46  1   Moving          움직임 유무         R   0
            ''        28 CHSUM
            ''    --------------------------------------------------------------
            
            If (inCNT = 4086) Then  '''Do:8086??
            ''
                ''
                ''
                ''Do:Max~??
                
                rxAngle = ((CDbl(inBUF(4056 + 12)) * 256) + (CDbl(inBUF(4056 + 11)))) / 11.3777778
                rxAngle = rxAngle - 180#
                ''If (rxAngle > (-50)) And (rxAngle < (50)) Then
                    ''lbTiltTX.Caption = Trim(Str(CLng(rxAngle * 100) / 100))
                    lbTiltTX = CLng(rxAngle * 100#) / 100
                ''End If
                ''
                rxAngle = ((CDbl(inBUF(4056 + 18)) * 256) + (CDbl(inBUF(4056 + 17)))) / 11.3777778
                rxAngle = rxAngle - 180#
                ''If (rxAngle > (-50)) And (rxAngle < (50)) Then
                    ''lbTiltRX.Caption = Trim(Str(CLng(rxAngle * 100) / 100))
                    lbTiltRX = CLng(rxAngle * 100#) / 100
                ''End If
                rxAngleOK = True
            ''
            End If
            
        End If
        
    ElseIf inCNT >= 8056 Then   ''''''<=="GSCN"
        ''rxSTOP = 1          ''''''=======>
        
        If (inBUF(0) = LD_sBUF(47)(0)) And (inBUF(1) = LD_sBUF(47)(1)) And (inBUF(2) = LD_sBUF(47)(2)) Then
        ''''GSCN---;;
            lbRxHead.Caption = "GSCN"
            lbRxHead.BackColor = &HFF8080
            
            ''''cmdCONN.BackColor = vbGreen
            
            
            '' BytesToLong = Byte1 + Byte2 * 2 ^ 8 + Byte3 * 2 ^ 16 + Byte4 * 2 ^ 24
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''            for ( angle = 0; angle < 1000 ; angle++ )
''''            {
''''                //memcpy (&DPS_Vdist, &WORK[(angle * 8) + 52], 4);
''''                //memcpy (&DPS_Vpw, &WORK[(angle * 8) + 56], 4);
''''
''''                ////
''''                    angleN = (angle * 8) + 52;
''''                    //
''''                    DPS_Vdist = (WORK[angleN] & 0x000000ff ) << 24;
''''                    DPS_Vdist += (WORK[angleN+1] & 0x000000ff ) << 16;
''''                    DPS_Vdist += (WORK[angleN+2] & 0x000000ff ) << 8;
''''                    DPS_Vdist += (WORK[angleN+3] & 0x000000ff );
''''                    //
''''                    DPS_Vpw = (WORK[angleN+4] & 0x000000ff ) << 24;
''''                    DPS_Vpw += (WORK[angleN+5] & 0x000000ff ) << 16;
''''                    DPS_Vpw += (WORK[angleN+6] & 0x000000ff ) << 8;
''''                    DPS_Vpw += (WORK[angleN+7] & 0x000000ff );
''''                ////
''''                //
''''                if ( (DPS_Vdist > 3000000) || (DPS_Vdist < 0) )  DPS_Vdist = 0;
''''
''''                DPSautoD[angle] = (int)(DPS_Vdist / 100);  // 0~30000cm
''''
''''                // sprintf (WORK2, "$GSCN,%04d,%d,%d\n\r", angle, DPS_Vdist, DPS_Vpw );
''''                // SB_SendSerial (SYS.sfd, WORK2, strlen(WORK2));
''''                // SB_msleep (2);
''''            }
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''Dim angleN As Integer
            '''
            For i = 0 To 999  '''1000
                Scan2590Direc(i) = 45# + (i * 0.09)
                ''
                angleN = (i * 8) + 52
                Scan2590Dist(i) = (inBUF(angleN) * 2 ^ 24) + (inBUF(angleN + 1) * 2 ^ 16) + (inBUF(angleN + 2) * 2 ^ 8) + inBUF(angleN + 3)
                angleN = angleN + 4
                Scan2590Pulse(i) = (inBUF(angleN) * 2 ^ 24) + (inBUF(angleN + 1) * 2 ^ 16) + (inBUF(angleN + 2) * 2 ^ 8) + inBUF(angleN + 3)
                ''
                '' Check the distance value is -2147483648 (0x80000000) in case that the echo signal was too low.
                '' or, the distance value is 2147483647 (0x7FFFFFFF) in case that the echo signal was noisy.
                '' And Check the distance is near to 5meters.
                '' 2147483646 <= 2147483647 - 1
                If (Scan2590Dist(i) > 2147483646) _
                    Or (Scan2590Dist(i) < 50000) _
                        Then
                    Scan2590Dist(i) = 0
                    pointErr = pointErr + 1
                End If
            Next i
            
        End If
        
    End If





    For i = 0 To 239
        rxWORD(i) = 0
    Next i
    

    Dim tAng As Double
    Dim tCnt As Integer
    Dim tSum As Double
    ''''''''''''''''''''''''''''
    
    tAng = 45#  ''angle
    tCnt = 0
    tSum = 0
    ''
    For i = 0 To 999  ''(1~1000)

        If (Scan2590Dist(i) > 0) Then
        
            If (Scan2590Direc(i) < tAng + 0.25) Then
                tSum = tSum + Scan2590Dist(i)
                tCnt = tCnt + 1
            Else
                If tCnt > 0 Then
                    tSum = (tSum / tCnt) / 10   ''* 1000  '';;;[cm]distance
                    '''
                    If tSum > 80000 Then
                        tSum = 0
                    End If
                    '''
                    
                    rxWORD((tAng - 45#) * 2 + 30) = CLng(tSum)
                    
                    ''[90::{45.0~135.0}]=(-45):{0~90)==[x2]=>0~180==[+30]==>{~(30~210)~} // [120::{0~120}]==>[240::{0~238}]
                    
                End If
                
                tCnt = 0
                tSum = 0
                
AngleUp:
                tAng = tAng + 0.5
                
                If (Scan2590Direc(i) < tAng + 0.25) Then
                    tSum = tSum + Scan2590Dist(i)
                    tCnt = tCnt + 1
                Else
                    GoTo AngleUp
                End If
            End If

        End If
            
    Next i


    If tCnt > 0 Then
        tSum = (tSum / tCnt) / 10   ''* 1000  '';;;[cm]distance
        '''
        If tSum > 80000 Then
            tSum = 0
        End If
        '''
        rxWORD((tAng - 45#) * 2 + 30) = CLng(tSum)
        
        ''[90::{45.0~135.0}]=(-45):{0~90)==[x2]=>0~180==[+30]==>{~(30~210)~} // [120::{0~120}]==>[240::{0~238}]
        
    End If
    
    xcMax = 238
    ''''''''''''''''''''''''''''
    ''xcMax = xcMax + 1
    lbXC = xcMax

    ''''''''
            ''lbRXerr = 0
            lbRXcnt = lbRXcnt + 1
            If lbRXcnt > 999999999 Then
                lbRXcnt = lbRXcnt / 10
                lbRXerr = lbRXerr / 10
            End If

            If (cmdFilt.BackColor = vbGreen) Then
                RX_filt_DEEP
                ''''''''''''
            End If

            RX_filt
            '''''''

            picSiloDrawInit

            picSiloDRAW

            cmdCONN.BackColor = vbGreen
    ''''''''

    If (ScanTYPE = 22590) And (AutoTiltStarted = True) And (rxAngleOK = False) Then
        LDrx12590 = -5
        Exit Function  ''===>
    End If
    
    If (AutoTiltStarted = True) And (rxAngleOK = True) Then
        If (Abs(rxAngle - AutoTiltNow) > 1.5) Then
            LDrx12590 = -6
            Exit Function  ''===>
        End If
'
        If (cmdFilt.BackColor = vbGreen) And (AutoTiltStarted = True) Then
            SaveBuffer2File cmdCONN.Caption & "_" & Format(AutoTiltCnt, "00") & "_" & Format(rxAngle, "0.00") & "_", inBUF, inCNT
        End If
'
        Dim X As Double
        Dim Y As Double
        Dim z As Double
        Dim r As Double
'
        i = 30
        r = (rxWORD(i) + rxWORD(i + 1) + rxWORD(i + 2) / 2#) / 2.5
        X = r * Cos(((i / 2) + 30) * (PI / 180)) / 1000#
        Y = r * Sin(((i / 2) + 30) * (PI / 180)) * Sin(lbTiltRX * (PI / 180)) / 1000#
        z = 50# - (r * Sin(((i / 2) + 30) * (PI / 180)) * Cos(lbTiltRX * (PI / 180)) / 1000#)
        Tilt3Dlog_add tilt3Dlog_fn, _
            " " & lbTiltRX _
            & vbTab & (i / 2) + 30 _
            & vbTab & vbTab & X _
            & " " & vbTab & Y _
            & " " & vbTab & z
        For i = 34 To xcMax - 30 - 2 Step 4
            r = ( _
                    rxWORD(i - 2) / 2# + rxWORD(i - 1) _
                    + rxWORD(i) _
                    + rxWORD(i + 1) + rxWORD(i + 2) / 2# _
                ) _
                / 4#
            X = r * Cos(((i / 2) + 30) * (PI / 180)) / 1000#
            Y = r * Sin(((i / 2) + 30) * (PI / 180)) * Sin(lbTiltRX * (PI / 180)) / 1000#
            z = 50# - (r * Sin(((i / 2) + 30) * (PI / 180)) * Cos(lbTiltRX * (PI / 180)) / 1000#)
            Tilt3Dlog_add tilt3Dlog_fn, _
                " " & lbTiltRX _
                & vbTab & (i / 2) + 30 _
                & vbTab & vbTab & X _
                & " " & vbTab & Y _
                & " " & vbTab & z
        Next i
        'i = xcMax - 30 - 2
        r = (rxWORD(i - 2) / 2# + rxWORD(i - 1) + rxWORD(i)) / 2.5
        X = r * Cos(((i / 2) + 30) * (PI / 180)) / 1000#
        Y = r * Sin(((i / 2) + 30) * (PI / 180)) * Sin(lbTiltRX * (PI / 180)) / 1000#
        z = 50# - (r * Sin(((i / 2) + 30) * (PI / 180)) * Cos(lbTiltRX * (PI / 180)) / 1000#)
        Tilt3Dlog_add tilt3Dlog_fn, _
            " " & lbTiltRX _
            & vbTab & (i / 2) + 30 _
            & vbTab & vbTab & X _
            & " " & vbTab & Y _
            & " " & vbTab & z
    End If
'
    LDrx12590 = 0
'
    't = GetTickCount - t
    'DGPSLog "LDrx12590(" & UCindex & ") END " & t & "", "SILO"
'
End Function





''''LD-LRS-3100,, DPS-2590

Private Function LDtx2590(ix As Integer) As Integer

Dim i As Integer
Dim tBuf As Variant

    ReDim tBuf(UBound(LD_sBUF(ix))) As Byte
    ''
    For i = 0 To UBound(LD_sBUF(ix))
        tBuf(i) = (LD_sBUF(ix)(i))
    Next i
    
    
    RxMSG = ""
    
    inCNT = 0   ''ws
    xcMax = 0   ''scanData!!
    
    rxSTOP = 0
    ''''''''''''''''''''''''''''''''''''''''''''''<==!!
    
    inBUF2590 = ""
    
    
    SEND_wsickLD tBuf

    LDtx2590 = 0
    
End Function


Private Function LDrx2590(ix As Integer) As Integer

Dim i As Integer

    If rxWaitTime < 7000 And rxSTOP = 0 Then
        LDrx2590 = 1
        Exit Function  ''===>
    End If
    

''______________________________________________________________________________
''  Point ;    Echo ;   Direction ;    Distance ; Pulse width ;
''        ;         ;       [deg] ;         [m] ;        [ps] ;
''______________________________________________________________________________
''0000000001111111111222222222233333333333444444444455555555566666666667
''1234567890123456789012345678901234567890123456789012345678901234567890
''______________________________________________________________________________
''      1 ;       1 ;      45.000 ;     23.6961 ;        8694 ;
''      2 ;       1 ;      45.090 ;     23.7613 ;        8854 ;
''      3 ;       1 ;      45.180 ;     23.7934 ;        8780 ;
''      4 ;       1 ;      45.270 ;     23.8172 ;        8360 ;
''      5 ;       1 ;      45.360 ;     23.8558 ;        8335 ;
''      6 ;       1 ;      45.450 ;     23.9122 ;        8408 ;
''      7 ;       1 ;      45.540 ;     23.9379 ;        8186 ;
''      8 ;       1 ;      45.630 ;     23.9610 ;        8373 ;
''      9 ;       1 ;      45.720 ;     23.9845 ;        8087 ;
''     10 ;       1 ;      45.810 ;     24.0410 ;        8137 ;
''     11 ;       1 ;      45.900 ;     24.0792 ;        8297 ;
''     12 ;       1 ;      45.990 ;     24.1157 ;        8099 ;
''     13 ;       1 ;      46.080 ;     24.1456 ;        7889 ;
''
''    494 ;       1 ;      89.370 ;     42.3705 ;        2315 ;
''    495 ;       1 ;      89.460 ;     42.3681 ;        2699 ;
''    496 ;       1 ;      89.550 ;     42.4362 ;        3703 ;
''    497 ;       1 ;      89.640 ;     42.5219 ;        2749 ;
''    498 ;       1 ;      89.730 ;     42.4190 ;        5264 ;
''    499 ;       1 ;      89.820 ;    Low echo ;        1028 ;
''    500 ;       0 ;      89.910 ;    No echo! ;           0 ;
''    501 ;       1 ;      90.000 ;     42.4158 ;        3134 ;
''    502 ;       1 ;      90.090 ;     42.4278 ;        2575 ;
''    503 ;       1 ;      90.180 ;     42.4776 ;        2440 ;
''    504 ;       0 ;      90.270 ;    No echo! ;           0 ;
''    505 ;       1 ;      90.360 ;     42.1684 ;        2526 ;
''    506 ;       1 ;      90.450 ;     42.4352 ;        3889 ;
''
''    997 ;       1 ;     134.640 ;     25.9198 ;        7950 ;
''    998 ;       1 ;     134.730 ;     25.8690 ;        8000 ;
''    999 ;       1 ;     134.820 ;     25.8224 ;        7864 ;
''   1000 ;       1 ;     134.910 ;     25.7942 ;        7939 ;
''
''______________________________________________________________________________
'' Terminal mode
    

  Dim strLine As String
  Dim strNo As String
  
    xcMax = 0
    
    For i = 1 To 1000
        strNo = "   "
        If i < 1000 Then strNo = strNo & " "
        If i < 100 Then strNo = strNo & " "
        If i < 10 Then strNo = strNo & " "
        ''
        strNo = strNo & Trim(Str(i)) & " ;       "   '';''strNo = strNo & Trim(Str(i))
        
        ''Debug.Print strNo & " "
        
        
        Scan2590Echo(i) = 0
        Scan2590Direc(i) = 0
        ''
        Scan2590Dist(i) = 0
        Scan2590Pulse(i) = 0
        
        If InStr(inBUF2590, strNo) > 0 Then
        
            strLine = Mid(inBUF2590, InStr(inBUF2590, strNo), 60)

            Scan2590Echo(i) = Val(Mid(strLine, 17, 1))
            Scan2590Direc(i) = Val(Mid(strLine, 25, 7))
            ''
            If Scan2590Echo(i) = 1 Then
                Scan2590Dist(i) = Val(Mid(strLine, 38, 8))
                Scan2590Pulse(i) = Val(Mid(strLine, 55, 5))
            End If
            
            xcMax = xcMax + 1
            lbXC = xcMax
        End If
        
    ''Debug.Print vbCrLf & i & vbTab & Scan2590Echo(i) & vbTab & Scan2590Direc(i) & vbTab & Scan2590Dist(i) & vbTab & Scan2590Pulse(i)
        
        
    Next i


    '' xcMax = 238 Then '''> 230 Then                 ''[238]max....120~240:0.5degree


    ''For i = 1 To 180  ''(45~135: 90 / 0.5) ''<==''0~120/0.5::0~238::240::::30~150degree




    For i = 1 To 239
        rxWORD(i) = 0
    Next i
    

  Dim tAng As Double
  Dim tCnt As Integer
  Dim tSum As Double
  Dim tCntH As Integer
  Dim tSumH As Double
    
    tAng = 45#  ''angle
    tCnt = 0
    tSum = 0
    tCntH = 0
    tSumH = 0
    ''
    For i = 2 To 999  ''(1~1000)

        If (Scan2590Dist(i) > 0) Then
        
            If (Scan2590Direc(i) < tAng + 0.5) Then
                tSum = tSum + Scan2590Dist(i)
                tCnt = tCnt + 1
            ElseIf (Scan2590Direc(i) < tAng + 1#) Then
                tSumH = tSumH + Scan2590Dist(i)
                tCntH = tCntH + 1
            Else
            
''                    rxWORD(xc) = CLng(inBUF(X)) * 1000 + (CLng(inBUF(X + 1)) * 1000 / 256)
''                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If tCnt > 0 Then
                    tSum = (tSum / tCnt) * 1000  '';;;[cm]distance
                    '''
                    If tSum > 50000 Then
                        tSum = 0
                    End If
                    '''
                    rxWORD((tAng) * 2 - 30) = CLng(tSum)
                End If
                
                If tCntH > 0 Then
                    tSumH = (tSumH / tCntH) * 1000  '';;;[cm]distance
                    '''
                    If tSumH > 50000 Then
                        tSumH = 0
                    End If
                    '''
                    rxWORD((tAng) * 2 - 30 + 1) = CLng(tSumH)
                    
                    'Debug.Print tSum, tSumH
                    
                End If
            
            
                tAng = tAng + 1#
                tCnt = 0
                tSum = 0
                tCntH = 0
                tSumH = 0
            End If

        End If
            
    Next i

    xcMax = 238
    ''''''''''''''''''''''''''''
    ''xcMax = xcMax + 1
    lbXC = xcMax
            
    ''''''''
            lbRXerr = 0
            lbRXcnt = lbRXcnt + 1
            If lbRXcnt > 999999999 Then
                lbRXcnt = lbRXcnt / 10
                lbRXerr = lbRXerr / 10
            End If
            
            If cmdFilt.BackColor = vbGreen Then
                RX_filt_DEEP
                ''''''''''''
            End If
            
            RX_filt
            '''''''
            
            picSiloDrawInit
            
            picSiloDRAW
            
            cmdCONN.BackColor = vbGreen
    ''''''''

    LDrx2590 = 0
    
End Function


Private Function LDtxDATA(ix As Integer) As Integer    ''(sBUF())

Dim i As Integer
Dim tBuf As Variant

    ReDim tBuf(UBound(LD_sBUF(ix))) As Byte
    ''
    For i = 0 To UBound(LD_sBUF(ix))
        tBuf(i) = (LD_sBUF(ix)(i))
    Next i
    

'    txtRx1 = txtRx1 & vbCrLf & "TX:"
'    For i = 0 To UBound(tBuf)
'        txtRx1 = txtRx1 & Hex(tBuf(i)) & " "
'    Next i
    
    ''''''''
    ''DoEvents
    
    
    
    RxMSG = ""
    
    inCNT = 0   ''ws
    xcMax = 0   ''scanData!!
    
    rxSTOP = 0
    ''''''''''''''''''''''''''''''''''''''''''''''<==!!
    
    SEND_wsickLD tBuf
    '''''''''''''''''

    LDtxDATA = 0

End Function


Private Function LDrxDATA(ix As Integer) As Integer    ''(sBUF())

Dim i As Integer
Dim c As Integer
Dim z As Integer

Dim strHD(8) As Integer

Dim tBuf As Variant

Dim ONEcnt As Integer

    If rxWaitTime < 10 Then
        LDrxDATA = 1
        Exit Function  ''===>
    End If
    
    ONEcnt = 0

    strHD(0) = 2
    strHD(1) = 2
    strHD(2) = 2
    strHD(3) = 0
    strHD(4) = 0

    If tSrunMode >= eSrunMode.SendCmd Then
        tmrTRX.Interval = 200  ''300
    Else
        tmrTRX.Interval = 300
    End If

ONEmore:

    tmrTRX.Enabled = True
    
    i = 0

    Do While tmrTRX.Enabled = True
    
    
        If (tSrunMode < eSrunMode.SendCmd) And (inCNT > 10) Then   ''Command-RX

            i = 1

            If (inBUF(i) = strHD(0)) And _
                (inBUF(i + 1) = strHD(1)) And _
                (inBUF(i + 2) = strHD(2)) And _
                (inBUF(i + 3) = strHD(3)) And _
                (inBUF(i + 4) = strHD(4)) _
                Then
    
                z = inBUF(i + 5) * 256 + inBUF(i + 6)

                If (inCNT - i - 7) >= z Then
                
                        inCNT = 0
                        
                        GoTo RxOK  ''Exit Do  ''-->
                        '''''''''''''''''''''''''''
                End If
                
            End If
            
        ElseIf (tSrunMode >= eSrunMode.SendCmd) And (inCNT >= 516) Then  ''DATA-RX
    
            i = 1

            If (inBUF(i) = strHD(0)) And _
                (inBUF(i + 1) = strHD(1)) And _
                (inBUF(i + 2) = strHD(2)) And _
                (inBUF(i + 3) = strHD(3)) And _
                (inBUF(i + 4) = strHD(4)) _
                Then

                z = inBUF(i + 5) * 256 + inBUF(i + 6)

                If (inCNT - i - 7) >= z Then  ''''''''(516-8==508)

                            '''[SCAN-DATA]'''(RX:516bytes==HD+507)
                            '''RX:02 02 02 02 00 00 01 FB 73 52 41 00 62 00 FA 81 BA 01 02 00 4F 00 00 00 08 00 EF 07 88 03 CC 03 BF ~~ 17 B9 00 01 00 08 00 00 0F 00 CA
                            '''RX:02 02 02 02 00 00 01 FB 73 52 41 00 62 00 FA 81 BA 01 02 00 AF 00 00 00 08 00 EF 07 88 03 C6 03 C0 ~~ 03 E1 00 01 00 08 00 00 0F 00 F5
                            '''##-00-01-02-03-04-05-06--7--8--9-10-11-12-13-14-15-16-17-18-19-20-21-22-23-24-25-26-27-hh-ll-hh-ll-
                            ''':::##-00-01-02-03-04-05-06--7--8--9-10-11-12-13-14-15-16-17-18-19-20-21-22-23-24-25-26-27-hh-ll-hh-ll-
                            '''''''''''''''''''''''':::##-00-01-02-03-04-05-06--7--8--9-10-11-12-13-14-15-16-17-18-19-20-21-22-23-24-25-26-27-hh-ll-hh-ll-
                            Dim X As Integer
                            Dim xc As Integer
                            Dim xs As String
                            xc = 0
                            xs = vbCrLf & inCNT & ", " & i & ", " & z & vbCrLf
                            For X = i + 28 To inCNT - 10 Step 2  ''!!
                            '''''''''''''''''''''''''''''''''''''''!!
                            
                                '' rxWORD(xc) = inBUF(x) * 256 + inBUF(x + 1)  ''X''
                                ''
                                rxWORD(xc) = CLng(inBUF(X)) * 1000 + (CLng(inBUF(X + 1)) * 1000 / 256)
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                
                                If (xc Mod 10) = 0 Then
                                    xs = xs & vbCrLf
                                End If
                                xs = xs & Trim(xc) & "-" & Trim(rxWORD(xc)) & vbTab
                                
                                xcMax = xc
                                xc = xc + 1
                                
                                If xc > 240 Then Exit For       ''[238]max....120~240:0.5degree
                                
                                ''DoEvents
                                ''''''''
                            Next X
                            
                            ''Debug.Print xs
                            
                            lbXC = xcMax
                            
                            If xcMax = 238 Then '''> 230 Then                 ''[238]max....120~240:0.5degree
                                
                                lbRXerr = 0
                                lbRXcnt = lbRXcnt + 1
                                If lbRXcnt > 999999999 Then
                                    lbRXcnt = lbRXcnt / 10
                                    lbRXerr = lbRXerr / 10
                                End If
                                
                                If cmdFilt.BackColor = vbGreen Then
                                    RX_filt_DEEP
                                    ''''''''''''
                                End If
                                
                                RX_filt
                                '''''''
                                
                                picSiloDrawInit
                                
                                picSiloDRAW
                                
                                cmdCONN.BackColor = vbGreen
                                
                                
                                inCNT = 0
                                
                                GoTo RxOK  ''Exit Do  ''-->
                                '''''''''''''''''''''''''''
                                
                            Else
                                lbRXerr = lbRXerr + 1
                                cmdCONN.BackColor = vbBlue  ''vbRed
                            End If

                End If

            End If
    
        End If


        DoEvents
        ''''''''
    
        Sleep 1

    Loop


    If (ONEcnt = 0) And (inCNT > 10) Then  '''(inCNT = 516) Then

        ONEcnt = 1
        tmrTRX.Interval = 100

        GoTo ONEmore
        ''''========>
    End If


'    If inCNT > 0 Then
'        Debug.Print "RX-Error:", UCindex + 1, inCNT, i, xcMax
'
'        Dim s1 As String
'        s1 = ""
'        For i = 0 To 20
'            If inBUF(i) < 16 Then
'                s1 = s1 & "0" & Hex(inBUF(i)) & " "
'            Else
'                s1 = s1 & Hex(inBUF(i)) & " "
'            End If
'        Next i
'        Debug.Print "RX: " & s1
'
'    End If


    lbRXerr = lbRXerr + 1
 
    If (lbRXcnt > 9) And (lbRXerr > 1) And (lbRXerr < 9) Then  '''ERR+2~
        cmdCONN.BackColor = &H8000&       ''vbRed
    End If

    If lbRXerr > 9 Then
    
        tSrunMode = eSrunMode.InitConn  ''<<RESTART>>''
        '''''''''''''
        
        tmrTRX.Enabled = False
        
        inCNT = 0 ''Time-OUT??
        
    End If

    tmrTRX.Enabled = False

    inCNT = 0 ''Time-OUT??

    LDrxDATA = -1
    
    Exit Function  ''-------->
    
RxOK:
    tmrTRX.Enabled = False
    
    LDrxDATA = 0

End Function

Private Sub RX_filt_Init()
'
    Dim i%
'
    For i = 0 To 300 - 1
        rxWdeepSum(i) = 0
        rxWdeepCnt(i) = 0
    Next i
'
    cnWdeep = 0
    cnWring = 0
'
End Sub

Private Sub RX_filt_DEEP()

''xcMax
'''-------------------------------------
'Private rxWdeep(5, 300) As Long
'Private cnWdeep As Integer
'''-------------------------------------

Dim i As Integer
Dim j As Integer
    
    For i = 0 To xcMax - 1  ''Just~238
    
        If (rxWORD(i) < 2000) Or (rxWORD(i) > 80000) Then
            rxWORD(i) = 0 ''''''''''''''''''''''''''''''''''Miss-Value!
        End If
        If cnWdeep > 4 Then
            If rxWdeep(cnWring, i) <> 0 Then
                rxWdeepSum(i) = rxWdeepSum(i) - rxWdeep(cnWring, i)
                rxWdeepCnt(i) = rxWdeepCnt(i) - 1
            End If
        End If
        If rxWORD(i) <> 0 Then
            rxWdeepSum(i) = rxWdeepSum(i) + rxWORD(i)
            rxWdeepCnt(i) = rxWdeepCnt(i) + 1
        End If
        rxWdeep(cnWring, i) = rxWORD(i)
    
    Next i

    cnWring = cnWring + 1
    If cnWring > 4 Then
        cnWring = 0
    End If
    
    cnWdeep = cnWdeep + 1
    If cnWdeep > 5 Then
        cnWdeep = 5  ''''''''MAX!
    End If


''    If cnWdeep > 4 Then
''
''    End If


    ''(Miss Proce)''
    Dim c As Integer
    Dim s As Long
    Dim minVal As Long
    Dim maxVal As Long
  
    If cnWdeep >= 3 Then
    
        For i = 10 To xcMax - 11
    
            If rxWORD(i) = 0 Then
                '''''''''''''''''''''''''''''''''(Time-Filt)!
                s = rxWdeepSum(i - 1) + rxWdeepSum(i) + rxWdeepSum(i + 1)
                c = rxWdeepCnt(i - 1) + rxWdeepCnt(i) + rxWdeepCnt(i + 1)
                '''''''''''''''''''''''''''''''''
                ''If UCindex = 9 And c > 0 Then
                ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":i=" & i & ",c=" & c & ",s/c=" & CLng(s / c) & "", "SILO"
                ''End If
                If c >= 5 Then
                    rxWORD(i) = s / c
                    ''If UCindex = 5 Then
                    ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":i=" & i & ",c=" & c & ",s/c=" & CLng(s / c) & ",maxVal=" & maxVal & ",minVal=" & minVal & "", "SILO"
                    ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":zero" & ",rxWORD(" & i & ")=" & rxWORD(i) & "", "SILO"
                    ''End If
                End If
            Else
                c = 0
                s = 0
                maxVal = &H80000000 ' min of Long
                minVal = &H7FFFFFFF ' max of Long
                '''''''''''''''''''''''''''''''''(Time-Filt)!
                For j = 0 To 4
                    If rxWdeep(j, i) > 0 Then
                        If rxWdeep(j, i) > maxVal Then
                            maxVal = rxWdeep(j, i)
                        End If
                        If rxWdeep(j, i) < minVal Then
                            minVal = rxWdeep(j, i)
                        End If
                        s = s + rxWdeep(j, i)
                        c = c + 1
                    End If
                Next j
                s = s + rxWdeepSum(i - 1) + rxWdeepSum(i + 1)
                c = c + rxWdeepCnt(i - 1) + rxWdeepCnt(i + 1)
                ''If UCindex = 9 And c > 0 Then
                ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":i=" & i & ",c=" & c & ",s/c=" & CLng(s / c) & ",maxVal=" & maxVal & ",minVal=" & minVal & "", "SILO"
                ''End If
                '''''''''''''''''''''''''''''''''
                If c >= 7 Then
                    '' Check max & min value over 5meters with average
                    If (maxVal > s / c + 5000) And (minVal < s / c - 5000) Then
                        rxWORD(i) = (s - maxVal - minVal) / (c - 2)
                        ''If UCindex = 5 Then
                        ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":i=" & i & ",c=" & c & ",s/c=" & CLng(s / c) & ",maxVal=" & maxVal & ",minVal=" & minVal & "", "SILO"
                        ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":max." & ",rxWORD(" & i & ")=" & rxWORD(i) & "", "SILO"
                        ''End If
                    '' Check max value over 5meters with average
                    ElseIf maxVal > s / c + 5000 Then
                        rxWORD(i) = (s - maxVal) / (c - 1)
                        ''If UCindex = 5 Then
                        ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":i=" & i & ",c=" & c & ",s/c=" & CLng(s / c) & ",maxVal=" & maxVal & ",minVal=" & minVal & "", "SILO"
                        ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":max." & ",rxWORD(" & i & ")=" & rxWORD(i) & "", "SILO"
                        ''End If
                    '' Check min value over 5meters with average
                    ElseIf minVal < s / c - 5000 Then
                        rxWORD(i) = (s - minVal) / (c - 1)
                        ''If UCindex = 5 Then
                        ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":i=" & i & ",c=" & c & ",s/c=" & CLng(s / c) & ",maxVal=" & maxVal & ",minVal=" & minVal & "", "SILO"
                        ''    DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":min." & ",rxWORD(" & i & ")=" & rxWORD(i) & "", "SILO"
                        ''End If
                    End If
                End If
            End If
        Next i
    
        ''If UCindex = 9 Then
        ''    For i = 10 To xcMax - 11
        ''        DGPSLog "RX_filt_DEEP(" & UCindex & ")" & ":rxWORD(" & i & ")=" & rxWORD(i) & "", "SILO"
        ''    Next i
        ''End If

    End If

End Sub


Private Sub RX_filt()

Dim i, j As Integer
Dim Dsum As Long
Dim Dcnt As Integer
Dim X1 As Integer
Dim Y1&, Y2&
Dim s As Integer

    If cmdFilt.BackColor = vbGreen Then
        'Filter by equation of a line from 2 points.
        X1 = 0
        Y1 = 0
        s = 0

        For i = 30 To (xcMax - 30) Step 1
            If rxWORD(i) >= 5000 Then '' 5meter
                If s = 0 Then
                    s = 1
                    X1 = i
                ElseIf s = 1 Then
                    X1 = i
                ElseIf s = 2 Then
                    Dsum = 0
                    Dcnt = 0
                    For j = X1 To X1 - 5 Step -1
                        If rxWORD(j) >= 5000 Then
                            Dsum = Dsum + rxWORD(j)
                            Dcnt = Dcnt + 1
                        End If
                        Y1 = Dsum / Dcnt
                    Next j
                    If Abs(Y1 - rxWORD(X1)) < 5000 Then
                        Y1 = rxWORD(X1)
                    End If
                    Dsum = 0
                    Dcnt = 0
                    For j = i To i + 5 Step 1
                        If rxWORD(j) >= 5000 Then
                            Dsum = Dsum + rxWORD(j)
                            Dcnt = Dcnt + 1
                        End If
                        Y2 = Dsum / Dcnt
                    Next j
                    If Abs(Y2 - rxWORD(i)) < 5000 Then
                        Y2 = rxWORD(i)
                    End If
                    If (Y1 / 2) <= 5000 Or Abs(Y2 - Y1) < (Y1 / 2) Then
                        For j = X1 + 1 To i - 1 Step 1
                            rxWORD(j) = (Y2 - Y1) * (j - X1) / (i - X1) + Y1
                        Next j
                        s = 1
                        X1 = i
                    Else
                    End If
                End If
            Else '' under 5meter
                If s = 1 Then
                    s = 2
                End If
            End If
        Next i
        
        For i = 30 To (xcMax - 30) Step 1     '''(Right--to--Left)''   '';<--[238]word
            Dsum = 0
            Dcnt = 0
        
            If rxWORD(i) < 5000 Then  ''2001
                For j = 1 To 5  ''3ea
                    If rxWORD(i - j) >= 5000 Then
                        Dsum = Dsum + rxWORD(i - j)
                        Dcnt = Dcnt + 1
                    End If
                Next j
        
                If Dcnt > 0 Then
                    rxWORD(i) = Dsum / Dcnt   '''3
                End If
            End If
        Next i

        For i = (xcMax - 30) To 30 Step -1   '''(Left--to--Right)''   '';<--[238]word
            Dsum = 0
            Dcnt = 0
        
            If rxWORD(i) < 5000 Then  ''2001
                For j = 1 To 5  ''3ea
                    If rxWORD(i + j) >= 5000 Then
                        Dsum = Dsum + rxWORD(i + j)
                        Dcnt = Dcnt + 1
                    End If
                Next j
        
                If Dcnt > 0 Then
                    rxWORD(i) = Dsum / Dcnt   '''3
                End If
            End If
        Next i

        For i = 29 To 0 Step -1     '''(Right--to--End)''   '';<--[238]word
            Dsum = 0
            Dcnt = 0
            
            If rxWORD(i) < 5000 Then  ''2001
                For j = 1 To 5  ''3ea
                    If rxWORD(i + j) >= 5000 Then
                        Dsum = Dsum + rxWORD(i + j)
                        Dcnt = Dcnt + 1
                    End If
                Next j
        
                If Dcnt > 0 Then
                    rxWORD(i) = Dsum / Dcnt * 0.955   '''3
                End If
            End If
        Next i
    
        For i = (xcMax - 29) To xcMax Step 1     '''(Left--to--End)''   '';<--[238]word
            Dsum = 0
            Dcnt = 0
            
            If rxWORD(i) < 5000 Then  ''2001
                For j = 1 To 5  ''3ea
                    If rxWORD(i - j) >= 5000 Then
                        Dsum = Dsum + rxWORD(i - j)
                        Dcnt = Dcnt + 1
                    End If
                Next j
        
                If Dcnt > 0 Then
                    rxWORD(i) = Dsum / Dcnt * 0.955   '''3
                End If
            End If
        Next i
    End If
End Sub






Private Sub LDinitVAR_Error1()

''    SEND    2012-04-06 22:15:32.723 - <02><02><02><02><00><00><00><03>sPEf
LD_sBUF(0) = Array(2, 2, 2, 2, 0, 0, 0, 3, Asc("s"), Asc("P"), Asc("E"), Asc("f"))
''    RECEIVE 2012-04-06 22:15:32.729 - <02><02><02><02><00><00><00><03>sPX{


''    ''    SEND    2012-04-06 22:15:43.946 - <02><02><02><02><00><00><00><0a>sMI<00><00><03><0f4>rGD<0f1>
''    LD_sBUF(1) = Array(2, 2, 2, 2, 0, 0, 0, &HA, Asc("s"), Asc("M"), Asc("I"), 0, 0, 3, &HF4, Asc("r"), Asc("G"), Asc("D"), &HF1)
''    ''    RECEIVE 2012-04-06 22:15:43.953 - <02><02><02><02><00><00><00><06>sAI<00><00><01>z

''    SEND    2012-04-06 22:15:35.015 - <02><02><02><02><00><00><00><0a>sMI<00><05><03><0f4>rGD<0f4>
''    RECEIVE 2012-04-06 22:15:35.021 - <02><02><02><02><00><00><00><06>sAI<00><05><01><7f>
LD_sBUF(1) = Array(2, 2, 2, 2, 0, 0, 0, &HA, Asc("s"), Asc("M"), Asc("I"), 0, 5, 3, &HF4, Asc("r"), Asc("G"), Asc("D"), &HF4)



''    SEND    2012-07-18 17:10:35.156 - <02><02><02><02><00><00><00><05>sRI<00><00>h
''    SEND    2012-04-06 22:15:43.865 - <02><02><02><02><00><00><00><05>sRI<00><00>h
LD_sBUF(2) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, 0, Asc("h"))
''    RECEIVE 2012-04-06 22:15:43.874 - <02><02><02><02><00><00><00>$sRA<00><00><00><07>LD_XXXX<00><14>V01.00.00-07.05.2008V



'''''    SEND    2012-04-06 22:16:01.163 - <02><02><02><02><00><00><00><05>sRI<00>Z2
'''LD_sBUF(2) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("Z"), Asc("2"))
'''''    RECEIVE 2012-04-06 22:16:01.169 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;




''    SEND    2012-04-06 22:15:43.946 - <02><02><02><02><00><00><00><0a>sMI<00><00><03><0f4>rGD<0f1>
LD_sBUF(3) = Array(2, 2, 2, 2, 0, 0, 0, &HA, Asc("s"), Asc("M"), Asc("I"), 0, 0, 3, &HF4, Asc("r"), Asc("G"), Asc("D"), &HF1)
''    RECEIVE 2012-04-06 22:15:43.953 - <02><02><02><02><00><00><00><06>sAI<00><00><01>z

''    SEND    2012-04-06 22:15:43.953 - <02><02><02><02><00><00><00><0d>sWI<00><02><00><06>NoNameo
LD_sBUF(4) = Array(2, 2, 2, 2, 0, 0, 0, &HD, Asc("s"), Asc("W"), Asc("I"), 0, 2, 0, 6, Asc("N"), Asc("o"), Asc("N"), Asc("a"), Asc("m"), Asc("e"), Asc("o"))
''    RECEIVE 2012-04-06 22:15:43.960 - <02><02><02><02><00><00><00><05>sWA<00><02>g



''    SEND    2012-04-06 22:15:43.960 - <02><02><02><02><00><00><00><06>sWI<00>U<01>9                           ''<<===!! {-0-U-0-8}
LD_sBUF(5) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("U"), 1, Asc("9"))
''    RECEIVE 2012-04-06 22:15:43.966 - <02><02><02><02><00><00><00><05>sWA<00>U0



''    SEND    2012-04-06 22:15:43.966 - <02><02><02><02><00><00><00><06>sWI<00>V<00>;
LD_sBUF(6) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("V"), 0, Asc(";"))
''    RECEIVE 2012-04-06 22:15:43.972 - <02><02><02><02><00><00><00><05>sWA<00>V3
''
''    SEND    2012-04-06 22:15:43.972 - <02><02><02><02><00><00><00><07>sWI<00>W<00><00>:
LD_sBUF(7) = Array(2, 2, 2, 2, 0, 0, 0, 7, Asc("s"), Asc("W"), Asc("I"), 0, Asc("W"), 0, 0, Asc(":"))
''    RECEIVE 2012-04-06 22:15:43.979 - <02><02><02><02><00><00><00><05>sWA<00>W2
''
''    SEND    2012-04-06 22:15:43.980 - <02><02><02><02><00><00><00><06>sWI<00>X<0ff><0ca>
LD_sBUF(8) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("X"), &HFF, &HCA)
''    RECEIVE 2012-04-06 22:15:43.987 - <02><02><02><02><00><00><00><05>sWA<00>X=
''
''    SEND    2012-04-06 22:15:43.988 - <02><02><02><02><00><00><00><06>sWI<00>Y<00>4
LD_sBUF(9) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("Y"), 0, Asc("4"))
''    RECEIVE 2012-04-06 22:15:43.995 - <02><02><02><02><00><00><00><05>sWA<00>Y<
''
''    SEND    2012-04-06 22:15:43.995 - <02><02><02><02><00><00><00><07>sWI<00>\<081><0ba><0a>                  ''<<===!! {-0-\-0-0-1}
LD_sBUF(10) = Array(2, 2, 2, 2, 0, 0, 0, 7, Asc("s"), Asc("W"), Asc("I"), 0, Asc("\"), &H81, &HBA, &HA)
''    RECEIVE 2012-04-06 22:15:44.001 - <02><02><02><02><00><00><00><05>sWA<00>\9
''
''    SEND    2012-04-06 22:15:44.001 - <02><02><02><02><00><00><00><07>sWI<00>]<00><00>0
LD_sBUF(11) = Array(2, 2, 2, 2, 0, 0, 0, 7, Asc("s"), Asc("W"), Asc("I"), 0, Asc("]"), 0, 0, Asc("0"))
''    RECEIVE 2012-04-06 22:15:44.008 - <02><02><02><02><00><00><00><05>sWA<00>]8
''
''    SEND    2012-04-06 22:15:44.008 - <02><02><02><02><00><00><00><05>sMI<00><02>u
LD_sBUF(12) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, 2, Asc("u"))
''    RECEIVE 2012-04-06 22:15:44.014 - <02><02><02><02><00><00><00><06>sAI<00><02><01>x

''

''    SEND    2012-07-18 17:10:35.406 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-07-18 17:10:35.421 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>


''    SEND    2012-04-06 22:15:44.015 - <02><02><02><02><00><00><00><0a>sMI<00><00><03><0f4>rGD<0f1>
LD_sBUF(13) = Array(2, 2, 2, 2, 0, 0, 0, &HA, Asc("s"), Asc("M"), Asc("I"), 0, 0, 3, &HF4, Asc("r"), Asc("G"), Asc("D"), &HF1)
''    RECEIVE 2012-04-06 22:15:44.021 - <02><02><02><02><00><00><00><06>sAI<00><00><01>z




''
''    SEND    2012-04-06 22:15:44.021 - <02><02><02><02><00><00><00><0d>sWI<00>N<080><02><00><10><00><03><00><08><0ba>
LD_sBUF(14) = Array(2, 2, 2, 2, 0, 0, 0, &HD, Asc("s"), Asc("W"), Asc("I"), 0, Asc("N"), &H80, 2, 0, &H10, 0, 3, 0, 8, &HBA)
''    RECEIVE 2012-04-06 22:15:44.028 - <02><02><02><02><00><00><00><05>sWA<00>N+
''
''    SEND    2012-04-06 22:15:44.028 - <02><02><02><02><00><00><00><12>sWI<00>O<00><00><05><00><00><00><00><00><00><00><00><00><00>'
LD_sBUF(15) = Array(2, 2, 2, 2, 0, 0, 0, &H12, Asc("s"), Asc("W"), Asc("I"), 0, Asc("O"), 0, 0, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Asc("'"))
''    RECEIVE 2012-04-06 22:15:44.035 - <02><02><02><02><00><00><00><05>sWA<00>O*
''
''    SEND    2012-04-06 22:15:44.035 - <02><02><02><02><00><00><00><18>sWI<00>P<06><02><0a8><00><03><07><00><0ff><0ff><00><00><00><00><00><00><00><00><00><00><095>
LD_sBUF(16) = Array(2, 2, 2, 2, 0, 0, 0, &H18, Asc("s"), Asc("W"), Asc("I"), 0, Asc("P"), 6, 2, &HA8, 0, 3, 7, 0, &HFF, &HFF, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, &H95)
''    RECEIVE 2012-04-06 22:15:44.043 - <02><02><02><02><00><00><00><05>sWA<00>P5
''
''    SEND    2012-04-06 22:15:44.043 - <02><02><02><02><00><00><00><16>sWI<00>Q<06><00><00><00><01><00><08><00><00><00><00><00><00><00><00><00><00>3
LD_sBUF(17) = Array(2, 2, 2, 2, 0, 0, 0, &H16, Asc("s"), Asc("W"), Asc("I"), 0, Asc("Q"), 6, 0, 0, 0, 1, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Asc("3"))
''    RECEIVE 2012-04-06 22:15:44.051 - <02><02><02><02><00><00><00><05>sWA<00>Q4
''
''    SEND    2012-04-06 22:15:44.051 - <02><02><02><02><00><00><00><15>sWI<00>R<0c0><0a8><01><15><0ff><0ff><0ff><00><0c0><0a8><01><01><00><10><00><00><0c4>
LD_sBUF(18) = Array(2, 2, 2, 2, 0, 0, 0, &H15, Asc("s"), Asc("W"), Asc("I"), 0, Asc("R"), &HC0, &HA8, 1, &H15, &HFF, &HFF, &HFF, 0, &HC0, &HA8, 1, 1, 0, &H10, 0, 0, &HC4)
''    RECEIVE 2012-04-06 22:15:44.058 - <02><02><02><02><00><00><00><05>sWA<00>R7
''
''    SEND    2012-04-06 22:15:44.058 - <02><02><02><02><00><00><00><15>sWI<00>S<0e><0f8><07><080><00><00><00><00><00><00><00><00><00><00><00><00>O
LD_sBUF(19) = Array(2, 2, 2, 2, 0, 0, 0, &H15, Asc("s"), Asc("W"), Asc("I"), 0, Asc("S"), &HE, &HF8, 7, &H80, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Asc("O"))
''    RECEIVE 2012-04-06 22:15:44.065 - <02><02><02><02><00><00><00><05>sWA<00>S6
''
''    SEND    2012-04-06 22:15:44.065 - <02><02><02><02><00><00><00><0e>sWI<00>T<03><01><00><00><00><00><00><00><00>;
LD_sBUF(20) = Array(2, 2, 2, 2, 0, 0, 0, &HE, Asc("s"), Asc("W"), Asc("I"), 0, Asc("T"), 3, 1, 0, 0, 0, 0, 0, 0, 0, Asc(";"))
''    RECEIVE 2012-04-06 22:15:44.072 - <02><02><02><02><00><00><00><05>sWA<00>T1
''
''    SEND    2012-04-06 22:15:44.072 - <02><02><02><02><00><00><00><05>sMI<00><02>u
LD_sBUF(21) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, 2, Asc("u"))
''    RECEIVE 2012-04-06 22:15:44.078 - <02><02><02><02><00><00><00><06>sAI<00><02><01>x



''    <<<<RUN>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>0
''
''    SEND    2012-04-06 22:16:01.089 - <02><02><02><02><00><00><00><09>sMI<00><0b><00><01><05><0ba><0c2>
LD_sBUF(22) = Array(2, 2, 2, 2, 0, 0, 0, 9, Asc("s"), Asc("M"), Asc("I"), 0, &HB, 0, 1, 5, &HBA, &HC2)
''    RECEIVE 2012-04-06 22:16:01.096 - <02><02><02><02><00><00><00><05>sAI<00><0b>p
''
''
''    SEND    2012-04-06 22:16:01.163 - <02><02><02><02><00><00><00><05>sRI<00>Z2
LD_sBUF(23) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("Z"), Asc("2"))
''    RECEIVE 2012-04-06 22:16:01.169 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-04-06 22:16:01.174 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
LD_sBUF(24) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("c"), &HB)
''    RECEIVE 2012-04-06 22:16:01.180 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''    SEND    2012-04-06 22:16:01.268 - <02><02><02><02><00><00><00><05>sRI<00>Z2
LD_sBUF(25) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("Z"), Asc("2"))   '';LD_sBUF(22)
''    RECEIVE 2012-04-06 22:16:01.274 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-04-06 22:16:01.278 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
LD_sBUF(26) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("c"), &HB)        '';LD_sBUF(23)
''    RECEIVE 2012-04-06 22:16:01.284 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>

''    SEND    2012-04-06 22:16:04.214 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-04-06 22:16:04.220 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-04-06 22:16:04.223 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-04-06 22:16:04.229 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''
''    <<<<RUN>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>1
''
''    SEND    2012-04-06 22:16:04.276 - <02><02><02><02><00><00><00><05>sMI<00><09>~
LD_sBUF(27) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, 9, Asc("~"))
''    RECEIVE 2012-04-06 22:16:08.232 - <02><02><02><02><00><00><00><06>sAI<00><09><00>r
''
''    SEND    2012-04-06 22:16:08.232 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-04-06 22:16:08.238 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><03>9


''    SEND    2012-04-06 22:16:08.238 - <02><02><02><02><00><00><00><03>sPEf
''    RECEIVE 2012-04-06 22:16:08.243 - <02><02><02><02><00><00><00><03>sPX{
''
''    SEND    2012-04-06 22:16:08.243 - <02><02><02><02><00><00><00><05>sMI<00><0c>{
LD_sBUF(28) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, &HC, Asc("{"))
''    RECEIVE 2012-04-06 22:16:08.249 - <02><02><02><02><00><00><00><09>sAI<00><0c><00><00><00><03>t
''
''    SEND    2012-04-06 22:16:08.368 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-04-06 22:16:08.373 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''
''    SEND    2012-04-06 22:16:08.376 - <02><02><02><02><00><00><00><09>sMI<00><0b><00><01><01><0ba><0c6>
LD_sBUF(29) = Array(2, 2, 2, 2, 0, 0, 0, 9, Asc("s"), Asc("M"), Asc("I"), 0, &HB, 0, 1, 1, &HBA, &HC6)
''    RECEIVE 2012-04-06 22:16:08.382 - <02><02><02><02><00><00><00><05>sAI<00><0b>p
''
''
''    SEND    2012-04-06 22:16:08.383 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-04-06 22:16:08.388 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><03>9



''
''    SEND    2012-04-06 22:16:08.727 - <02><02><02><02><00><00><00><05>sRI<00>b<0a>
LD_sBUF(39) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("b"), &HA)
''    RECEIVE 2012-04-06 22:16:08.732 - <02><02><02><02><00><00><00><07>sRA<00>b<00><00><02>
''
''SSget(1) = startString + Chr(0) + Chr(5) + "s" + "R" + "I" + Chr(0) + "b" + Chr(&HA)


End Sub





Private Sub LDinitVAR()


''    SEND    2012-04-06 22:15:32.723 - <02><02><02><02><00><00><00><03>sPEf
''    RECEIVE 2012-04-06 22:15:32.729 - <02><02><02><02><00><00><00><03>sPX{
LD_sBUF(0) = Array(2, 2, 2, 2, 0, 0, 0, 3, Asc("s"), Asc("P"), Asc("E"), Asc("f"))
    
''    SEND    2012-04-06 22:15:35.015 - <02><02><02><02><00><00><00><0a>sMI<00><05><03><0f4>rGD<0f4>
''    RECEIVE 2012-04-06 22:15:35.021 - <02><02><02><02><00><00><00><06>sAI<00><05><01><7f>
LD_sBUF(1) = Array(2, 2, 2, 2, 0, 0, 0, &HA, Asc("s"), Asc("M"), Asc("I"), 0, 5, 3, &HF4, Asc("r"), Asc("G"), Asc("D"), &HF4)
    
''    SEND    2012-04-06 22:15:43.865 - <02><02><02><02><00><00><00><05>sRI<00><00>h
''    RECEIVE 2012-04-06 22:15:43.874 - <02><02><02><02><00><00><00>$sRA<00><00><00><07>LD_XXXX<00><14>V01.00.00-07.05.2008V
LD_sBUF(2) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, 0, Asc("h"))

''    SEND    2012-04-06 22:15:43.946 - <02><02><02><02><00><00><00><0a>sMI<00><00><03><0f4>rGD<0f1>
LD_sBUF(3) = Array(2, 2, 2, 2, 0, 0, 0, &HA, Asc("s"), Asc("M"), Asc("I"), 0, 0, 3, &HF4, Asc("r"), Asc("G"), Asc("D"), &HF1)
''    RECEIVE 2012-04-06 22:15:43.953 - <02><02><02><02><00><00><00><06>sAI<00><00><01>z

''    SEND    2012-04-06 22:15:43.953 - <02><02><02><02><00><00><00><0d>sWI<00><02><00><06>NoNameo
LD_sBUF(4) = Array(2, 2, 2, 2, 0, 0, 0, &HD, Asc("s"), Asc("W"), Asc("I"), 0, 2, 0, 6, Asc("N"), Asc("o"), Asc("N"), Asc("a"), Asc("m"), Asc("e"), Asc("o"))
''    RECEIVE 2012-04-06 22:15:43.960 - <02><02><02><02><00><00><00><05>sWA<00><02>g

''    SEND    2012-04-06 22:15:43.960 - <02><02><02><02><00><00><00><06>sWI<00>U<00>8
''    SEND    2012-07-18 17:10:35.265 - <02><02><02><02><00><00><00><06>sWI<00>U<01>9  ''(New)
LD_sBUF(5) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("U"), 1, Asc("9"))
''    RECEIVE 2012-04-06 22:15:43.966 - <02><02><02><02><00><00><00><05>sWA<00>U0
''
''    SEND    2012-07-18 17:10:35.281 - <02><02><02><02><00><00><00><06>sWI<00>V<00>;
LD_sBUF(6) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("V"), 0, Asc(";"))
''    RECEIVE 2012-04-06 22:15:43.972 - <02><02><02><02><00><00><00><05>sWA<00>V3
''
''    SEND    2012-04-06 22:15:43.972 - <02><02><02><02><00><00><00><07>sWI<00>W<00><00>:
LD_sBUF(7) = Array(2, 2, 2, 2, 0, 0, 0, 7, Asc("s"), Asc("W"), Asc("I"), 0, Asc("W"), 0, 0, Asc(":"))
''    RECEIVE 2012-04-06 22:15:43.979 - <02><02><02><02><00><00><00><05>sWA<00>W2
''
''    SEND    2012-04-06 22:15:43.980 - <02><02><02><02><00><00><00><06>sWI<00>X<0ff><0ca>
LD_sBUF(8) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("X"), &HFF, &HCA)
''    RECEIVE 2012-04-06 22:15:43.987 - <02><02><02><02><00><00><00><05>sWA<00>X=
''
''    SEND    2012-04-06 22:15:43.988 - <02><02><02><02><00><00><00><06>sWI<00>Y<00>4
LD_sBUF(9) = Array(2, 2, 2, 2, 0, 0, 0, 6, Asc("s"), Asc("W"), Asc("I"), 0, Asc("Y"), 0, Asc("4"))
''    RECEIVE 2012-04-06 22:15:43.995 - <02><02><02><02><00><00><00><05>sWA<00>Y<
''
''    SEND    2012-04-06 22:15:43.995 - <02><02><02><02><00><00><00><07>sWI<00>\<00><00>1
''    SEND    2012-07-18 17:10:35.343 - <02><02><02><02><00><00><00><07>sWI<00>\<081><0ba><0a>  ''(New)
LD_sBUF(10) = Array(2, 2, 2, 2, 0, 0, 0, 7, Asc("s"), Asc("W"), Asc("I"), 0, Asc("\"), &H81, &HBA, &HA)
''    RECEIVE 2012-04-06 22:15:44.001 - <02><02><02><02><00><00><00><05>sWA<00>\9
''
''    SEND    2012-04-06 22:15:44.001 - <02><02><02><02><00><00><00><07>sWI<00>]<00><00>0
LD_sBUF(11) = Array(2, 2, 2, 2, 0, 0, 0, 7, Asc("s"), Asc("W"), Asc("I"), 0, Asc("]"), 0, 0, Asc("0"))
''    RECEIVE 2012-04-06 22:15:44.008 - <02><02><02><02><00><00><00><05>sWA<00>]8
''
''    SEND    2012-04-06 22:15:44.008 - <02><02><02><02><00><00><00><05>sMI<00><02>u
LD_sBUF(12) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, 2, Asc("u"))
''    RECEIVE 2012-04-06 22:15:44.014 - <02><02><02><02><00><00><00><06>sAI<00><02><01>x

''    SEND    2012-07-18 17:10:35.406 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-07-18 17:10:35.421 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>

''    SEND    2012-04-06 22:15:44.015 - <02><02><02><02><00><00><00><0a>sMI<00><00><03><0f4>rGD<0f1>
LD_sBUF(13) = Array(2, 2, 2, 2, 0, 0, 0, &HA, Asc("s"), Asc("M"), Asc("I"), 0, 0, 3, &HF4, Asc("r"), Asc("G"), Asc("D"), &HF1)
''    RECEIVE 2012-04-06 22:15:44.021 - <02><02><02><02><00><00><00><06>sAI<00><00><01>z

''
''    SEND    2012-04-06 22:15:44.021 - <02><02><02><02><00><00><00><0d>sWI<00>N<080><02><00><10><00><03><00><08><0ba>
LD_sBUF(14) = Array(2, 2, 2, 2, 0, 0, 0, &HD, Asc("s"), Asc("W"), Asc("I"), 0, Asc("N"), &H80, 2, 0, &H10, 0, 3, 0, 8, &HBA)
''    RECEIVE 2012-04-06 22:15:44.028 - <02><02><02><02><00><00><00><05>sWA<00>N+
''
''    SEND    2012-04-06 22:15:44.028 - <02><02><02><02><00><00><00><12>sWI<00>O<00><00><05><00><00><00><00><00><00><00><00><00><00>'
LD_sBUF(15) = Array(2, 2, 2, 2, 0, 0, 0, &H12, Asc("s"), Asc("W"), Asc("I"), 0, Asc("O"), 0, 0, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Asc("'"))
''    RECEIVE 2012-04-06 22:15:44.035 - <02><02><02><02><00><00><00><05>sWA<00>O*
''
''    SEND    2012-04-06 22:15:44.035 - <02><02><02><02><00><00><00><18>sWI<00>P<06><02><0a8><00><03><07><00><0ff><0ff><00><00><00><00><00><00><00><00><00><00><095>
LD_sBUF(16) = Array(2, 2, 2, 2, 0, 0, 0, &H18, Asc("s"), Asc("W"), Asc("I"), 0, Asc("P"), 6, 2, &HA8, 0, 3, 7, 0, &HFF, &HFF, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, &H95)
''    RECEIVE 2012-04-06 22:15:44.043 - <02><02><02><02><00><00><00><05>sWA<00>P5
''
''    SEND    2012-04-06 22:15:44.043 - <02><02><02><02><00><00><00><16>sWI<00>Q<06><00><00><00><01><00><08><00><00><00><00><00><00><00><00><00><00>3
LD_sBUF(17) = Array(2, 2, 2, 2, 0, 0, 0, &H16, Asc("s"), Asc("W"), Asc("I"), 0, Asc("Q"), 6, 0, 0, 0, 1, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Asc("3"))
''    RECEIVE 2012-04-06 22:15:44.051 - <02><02><02><02><00><00><00><05>sWA<00>Q4
''
''    SEND    2012-04-06 22:15:44.051 - <02><02><02><02><00><00><00><15>sWI<00>R<0c0><0a8><01><15><0ff><0ff><0ff><00><0c0><0a8><01><01><00><10><00><00><0c4>
LD_sBUF(18) = Array(2, 2, 2, 2, 0, 0, 0, &H15, Asc("s"), Asc("W"), Asc("I"), 0, Asc("R"), &HC0, &HA8, 1, &H15, &HFF, &HFF, &HFF, 0, &HC0, &HA8, 1, 1, 0, &H10, 0, 0, &HC4)
''    RECEIVE 2012-04-06 22:15:44.058 - <02><02><02><02><00><00><00><05>sWA<00>R7
''
''    SEND    2012-04-06 22:15:44.058 - <02><02><02><02><00><00><00><15>sWI<00>S<0e><0f8><07><080><00><00><00><00><00><00><00><00><00><00><00><00>O
LD_sBUF(19) = Array(2, 2, 2, 2, 0, 0, 0, &H15, Asc("s"), Asc("W"), Asc("I"), 0, Asc("S"), &HE, &HF8, 7, &H80, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Asc("O"))
''    RECEIVE 2012-04-06 22:15:44.065 - <02><02><02><02><00><00><00><05>sWA<00>S6
''
''    SEND    2012-04-06 22:15:44.065 - <02><02><02><02><00><00><00><0e>sWI<00>T<03><01><00><00><00><00><00><00><00>;
LD_sBUF(20) = Array(2, 2, 2, 2, 0, 0, 0, &HE, Asc("s"), Asc("W"), Asc("I"), 0, Asc("T"), 3, 1, 0, 0, 0, 0, 0, 0, 0, Asc(";"))
''    RECEIVE 2012-04-06 22:15:44.072 - <02><02><02><02><00><00><00><05>sWA<00>T1
''
''    SEND    2012-04-06 22:15:44.072 - <02><02><02><02><00><00><00><05>sMI<00><02>u
LD_sBUF(21) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, 2, Asc("u"))
''    RECEIVE 2012-04-06 22:15:44.078 - <02><02><02><02><00><00><00><06>sAI<00><02><01>x


''    <<<<RUN>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>0
''
''    SEND    2012-04-06 22:16:01.089 - <02><02><02><02><00><00><00><09>sMI<00><0b><00><01><05><0ba><0c2>
LD_sBUF(22) = Array(2, 2, 2, 2, 0, 0, 0, 9, Asc("s"), Asc("M"), Asc("I"), 0, &HB, 0, 1, 5, &HBA, &HC2)
''    RECEIVE 2012-04-06 22:16:01.096 - <02><02><02><02><00><00><00><05>sAI<00><0b>p
''
''
''    SEND    2012-04-06 22:16:01.163 - <02><02><02><02><00><00><00><05>sRI<00>Z2
LD_sBUF(23) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("Z"), Asc("2"))
''    RECEIVE 2012-04-06 22:16:01.169 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-04-06 22:16:01.174 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
LD_sBUF(24) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("c"), &HB)
''    RECEIVE 2012-04-06 22:16:01.180 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''    SEND    2012-04-06 22:16:01.268 - <02><02><02><02><00><00><00><05>sRI<00>Z2
LD_sBUF(25) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("Z"), Asc("2"))   '';LD_sBUF(22)
''    RECEIVE 2012-04-06 22:16:01.274 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-04-06 22:16:01.278 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
LD_sBUF(26) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("c"), &HB)        '';LD_sBUF(23)
''    RECEIVE 2012-04-06 22:16:01.284 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>

''    SEND    2012-04-06 22:16:04.214 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-04-06 22:16:04.220 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-04-06 22:16:04.223 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-04-06 22:16:04.229 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''
''    <<<<RUN>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>1
''
''    SEND    2012-04-06 22:16:04.276 - <02><02><02><02><00><00><00><05>sMI<00><09>~
LD_sBUF(27) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, 9, Asc("~"))
''    RECEIVE 2012-04-06 22:16:08.232 - <02><02><02><02><00><00><00><06>sAI<00><09><00>r
''
''    SEND    2012-04-06 22:16:08.232 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-04-06 22:16:08.238 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><03>9


''    SEND    2012-04-06 22:16:08.238 - <02><02><02><02><00><00><00><03>sPEf
''    RECEIVE 2012-04-06 22:16:08.243 - <02><02><02><02><00><00><00><03>sPX{
''
''    SEND    2012-04-06 22:16:08.243 - <02><02><02><02><00><00><00><05>sMI<00><0c>{
LD_sBUF(28) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("M"), Asc("I"), 0, &HC, Asc("{"))
''    RECEIVE 2012-04-06 22:16:08.249 - <02><02><02><02><00><00><00><09>sAI<00><0c><00><00><00><03>t
''
''    SEND    2012-04-06 22:16:08.368 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-04-06 22:16:08.373 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''
''    SEND    2012-04-06 22:16:08.376 - <02><02><02><02><00><00><00><09>sMI<00><0b><00><01><01><0ba><0c6>
LD_sBUF(29) = Array(2, 2, 2, 2, 0, 0, 0, 9, Asc("s"), Asc("M"), Asc("I"), 0, &HB, 0, 1, 1, &HBA, &HC6)
''    RECEIVE 2012-04-06 22:16:08.382 - <02><02><02><02><00><00><00><05>sAI<00><0b>p
''
''
''    SEND    2012-04-06 22:16:08.383 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-04-06 22:16:08.388 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><03>9


''    <<<<REQ!!>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>1
''
''    SEND    2012-04-06 22:16:08.727 - <02><02><02><02><00><00><00><05>sRI<00>b<0a>
LD_sBUF(39) = Array(2, 2, 2, 2, 0, 0, 0, 5, Asc("s"), Asc("R"), Asc("I"), 0, Asc("b"), &HA)
''    RECEIVE 2012-04-06 22:16:08.732 - <02><02><02><02><00><00><00><07>sRA<00>b<00><00><02>
''
''SSget(1) = startString + Chr(0) + Chr(5) + "s" + "R" + "I" + Chr(0) + "b" + Chr(&HA)


    LD_sBUF(41) = Array(&HD, &HD, &HD, &HD, &HD, &HA, &HA, &HA, &HA, &HA, &HD, &HA, &HD, &HA, &HD, &HA, &HD, &HA, &HD, &HA)
    
    LD_sBUF(43) = Array(Asc("s"), &HD, &HA)   '''43<--49 for DPS2590--12590


''''//// 53 43 41 4E 00 00 00 04 00 00 00 10 F0 85 33 D2 ;; SCAN........ð?3O
''''//// 47 53 43 4E 00 00 00 04 00 00 00 00 48 2F E1 C3 ;; GSCN........H/aA
''''//
''''char DPS_SCAN [16] = { 0x53, 0x43, 0x41, 0x4E, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x10, 0xF0, 0x85, 0x33, 0xD2 };
    ''''
    LD_sBUF(45) = Array(&H53, &H43, &H41, &H4E, &H0, &H0, &H0, &H4, &H0, &H0, &H0, &H10, &HF0, &H85, &H33, &HD2)

''''char DPS_GSCN [16] = { 0x47, 0x53, 0x43, 0x4E, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x00, 0x48, 0x2F, 0xE1, 0xC3 };
    ''''
    LD_sBUF(47) = Array(&H47, &H53, &H43, &H4E, &H0, &H0, &H0, &H4, &H0, &H0, &H0, &H0, &H48, &H2F, &HE1, &HC3)

    ''''ERR?????????????????????????????????
    LD_sBUF(49) = Array(&H45, &H52, &H52, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)


''    // Set Parameter SDC(Scan Data Content) to 8(=distances+PW)
''    // 53 50 52 4d 00 00 00 08 00 00 00 1f 00 00 00 08 e8 c4 07 6d;; SPRM............eA.m
''    char DPS_SPRM_SDC8[20] = { 0x53, 0x50, 0x52, 0x4d, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x1f, 0x00, 0x00, 0x00, 0x08, 0xe8, 0xc4, 0x07, 0x6d };
''
''    // Set Parameter SDC(Scan Data Content) to 4(=distances olny)
''    // 53 50 52 4d 00 00 00 08 00 00 00 1f 00 00 00 04 e1 72 4b 46;; SPRM............arKF   --5350524d000000080000001f00000004e1724b46--
''    char DPS_SPRM_SDC4[20] = { 0x53, 0x50, 0x52, 0x4d, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x1f, 0x00, 0x00, 0x00, 0x04, 0xe1, 0x72, 0x4b, 0x46 };
    
    LD_sBUF(51) = Array(&H53, &H50, &H52, &H4D, &H0, &H0, &H0, &H8, &H0, &H0, &H0, &H1F, &H0, &H0, &H0, &H4, &HE1, &H72, &H4B, &H46)

''    // Set Parameter Red Laser Marker at startup to 1(=off)
''    // 53 50 52 4d 00 00 00 08 00 00 00 07 00 00 00 01 c1 88 63 8a;; SPRM............A.c.
''    // Set Parameter Red Laser Marker at startup to 0(=off)
''    // 53 50 52 4d 00 00 00 08 00 00 00 07 00 00 00 00 b6 8f 53 1c;; SPRM............?S.
''    // Set Parameter Red Laser Marker status temporary to 1(=on)
''    // 53 50 52 4d 00 00 00 08 00 00 00 08 00 00 00 01 43 d8 f4 5b;; SPRM............CØo[

    LD_sBUF(53) = Array(&H53, &H50, &H52, &H4D, &H0, &H0, &H0, &H8, &H0, &H0, &H0, &H8, &H0, &H0, &H0, &H1, &H43, &HD8, &HF4, &H5B)

''    // Set Parameter Red Laser Marker status temporary to 0(=off)
''    // 53 50 52 4d 00 00 00 08 00 00 00 08 00 00 00 00 34 df c4 cd;; SPRM............4ßAI

    LD_sBUF(55) = Array(&H53, &H50, &H52, &H4D, &H0, &H0, &H0, &H8, &H0, &H0, &H0, &H8, &H0, &H0, &H0, &H0, &H34, &HDF, &HC4, &HCD)

End Sub





'''<<< Near Field Suppression >>>'''

''    SEND    2012-07-18 17:10:35.031 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-07-18 17:10:35.046 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''    SEND    2012-07-18 17:10:35.109 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-07-18 17:10:35.125 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-07-18 17:10:35.140 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-07-18 17:10:35.156 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>

''    SEND    2012-07-18 17:10:35.156 - <02><02><02><02><00><00><00><05>sRI<00><00>h
''    RECEIVE 2012-07-18 17:10:35.171 - <02><02><02><02><00><00><00>$sRA<00><00><00><07>LD_XXXX<00><14>V01.00.00-07.05.2008V
''    SEND    2012-07-18 17:10:35.218 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-07-18 17:10:35.218 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-07-18 17:10:35.218 - <02><02><02><02><00><00><00><0a>sMI<00><00><03><0f4>rGD<0f1>
''    RECEIVE 2012-07-18 17:10:35.250 - <02><02><02><02><00><00><00><06>sAI<00><00><01>z
''    SEND    2012-07-18 17:10:35.250 - <02><02><02><02><00><00><00><0d>sWI<00><02><00><06>NoNameo
''    RECEIVE 2012-07-18 17:10:35.265 - <02><02><02><02><00><00><00><05>sWA<00><02>g
''    SEND    2012-07-18 17:10:35.265 - <02><02><02><02><00><00><00><06>sWI<00>U<01>9
''    RECEIVE 2012-07-18 17:10:35.281 - <02><02><02><02><00><00><00><05>sWA<00>U0
''    SEND    2012-07-18 17:10:35.281 - <02><02><02><02><00><00><00><06>sWI<00>V<00>;
''    RECEIVE 2012-07-18 17:10:35.312 - <02><02><02><02><00><00><00><05>sWA<00>V3
''    SEND    2012-07-18 17:10:35.312 - <02><02><02><02><00><00><00><07>sWI<00>W<00><00>:
''    RECEIVE 2012-07-18 17:10:35.328 - <02><02><02><02><00><00><00><05>sWA<00>W2
''    SEND    2012-07-18 17:10:35.328 - <02><02><02><02><00><00><00><06>sWI<00>X<0ff><0ca>
''    RECEIVE 2012-07-18 17:10:35.328 - <02><02><02><02><00><00><00><05>sWA<00>X=
''    SEND    2012-07-18 17:10:35.328 - <02><02><02><02><00><00><00><06>sWI<00>Y<00>4
''    RECEIVE 2012-07-18 17:10:35.343 - <02><02><02><02><00><00><00><05>sWA<00>Y<
''    SEND    2012-07-18 17:10:35.343 - <02><02><02><02><00><00><00><07>sWI<00>\<081><0ba><0a>
''    RECEIVE 2012-07-18 17:10:35.359 - <02><02><02><02><00><00><00><05>sWA<00>\9
''    SEND    2012-07-18 17:10:35.359 - <02><02><02><02><00><00><00><07>sWI<00>]<00><00>0
''    RECEIVE 2012-07-18 17:10:35.390 - <02><02><02><02><00><00><00><05>sWA<00>]8
''    SEND    2012-07-18 17:10:35.390 - <02><02><02><02><00><00><00><05>sMI<00><02>u
''    RECEIVE 2012-07-18 17:10:35.406 - <02><02><02><02><00><00><00><06>sAI<00><02><01>x
''    SEND    2012-07-18 17:10:35.406 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-07-18 17:10:35.421 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''    SEND    2012-07-18 17:10:35.421 - <02><02><02><02><00><00><00><0a>sMI<00><00><03><0f4>rGD<0f1>
''    RECEIVE 2012-07-18 17:10:35.421 - <02><02><02><02><00><00><00><06>sAI<00><00><01>z
''    SEND    2012-07-18 17:10:35.437 - <02><02><02><02><00><00><00><0d>sWI<00>N<080><02><00><10><00><03><00><08><0ba>
''    RECEIVE 2012-07-18 17:10:35.437 - <02><02><02><02><00><00><00><05>sWA<00>N+
''    SEND    2012-07-18 17:10:35.437 - <02><02><02><02><00><00><00><12>sWI<00>O<00><00><05><00><00><00><00><00><00><00><00><00><00>'
''    RECEIVE 2012-07-18 17:10:35.453 - <02><02><02><02><00><00><00><05>sWA<00>O*
''    SEND    2012-07-18 17:10:35.453 - <02><02><02><02><00><00><00><18>sWI<00>P<06><02><0a8><00><03><07><00><0ff><0ff><00><00><00><00><00><00><00><00><00><00><095>
''    RECEIVE 2012-07-18 17:10:35.468 - <02><02><02><02><00><00><00><05>sWA<00>P5
''    SEND    2012-07-18 17:10:35.468 - <02><02><02><02><00><00><00><16>sWI<00>Q<06><00><00><00><01><00><08><00><00><00><00><00><00><00><00><00><00>3
''    RECEIVE 2012-07-18 17:10:35.500 - <02><02><02><02><00><00><00><05>sWA<00>Q4
''    SEND    2012-07-18 17:10:35.500 - <02><02><02><02><00><00><00><15>sWI<00>R<0c0><0a8><01><15><0ff><0ff><0ff><00><0c0><0a8><01><01><00><10><00><00><0c4>
''    RECEIVE 2012-07-18 17:10:35.515 - <02><02><02><02><00><00><00><05>sWA<00>R7
''    SEND    2012-07-18 17:10:35.515 - <02><02><02><02><00><00><00><15>sWI<00>S<0e><0f8><07><080><00><00><00><00><00><00><00><00><00><00><00><00>O
''    RECEIVE 2012-07-18 17:10:35.531 - <02><02><02><02><00><00><00><05>sWA<00>S6
''    SEND    2012-07-18 17:10:35.531 - <02><02><02><02><00><00><00><0e>sWI<00>T<03><01><00><00><00><00><00><00><00>;
''    RECEIVE 2012-07-18 17:10:35.546 - <02><02><02><02><00><00><00><05>sWA<00>T1
''    SEND    2012-07-18 17:10:35.546 - <02><02><02><02><00><00><00><05>sMI<00><02>u
''    RECEIVE 2012-07-18 17:10:35.562 - <02><02><02><02><00><00><00><06>sAI<00><02><01>x

''    SEND    2012-07-18 17:10:35.562 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-07-18 17:10:35.578 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-07-18 17:10:35.578 - <02><02><02><02><00><00><00><03>sPEf
''    RECEIVE 2012-07-18 17:10:35.593 - <02><02><02><02><00><00><00><03>sPX{
''    SEND    2012-07-18 17:10:35.593 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-07-18 17:10:35.609 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''    SEND    2012-07-18 17:10:35.609 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-07-18 17:10:35.625 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;
''    SEND    2012-07-18 17:10:35.687 - <02><02><02><02><00><00><00><05>sRI<00>c<0b>
''    RECEIVE 2012-07-18 17:10:35.687 - <02><02><02><02><00><00><00><07>sRA<00>c<00><00><03>
''    SEND    2012-07-18 17:10:35.718 - <02><02><02><02><00><00><00><05>sRI<00>Z2
''    RECEIVE 2012-07-18 17:10:35.734 - <02><02><02><02><00><00><00><09>sRA<00>Z<00><00><00><01>;



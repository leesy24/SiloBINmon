Attribute VB_Name = "modSurFit"
'=============================================================
' Descrizione.....: Form di prova per le routines di "Surface
'                   Fitting".
' Nome dei Files..: frmSurFit.frm, frmSurFit.frx
'                   frmSettings.frm, frmSettings.frx
'                   frm3D.frm, frm3D.frx
'                   frmInstr.frm, frmInstr.frx
'                   InfoCr.frm, InfoCr.frx
'                   modKTB2D.bas, modMASUB.bas, modQSHEP2D
'                   modGradiente2D.bas, Layers.bas
'                   modUtility.bas
' Data............: 21/9/2001
' Versione........: 1.0 a 32 bits.
' Sistema.........: VB6 sotto Windows NT.
' Scritto da......: F. Languasco
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.virgilio.it/flanguasco/
'                   http://www.flanguasco.org
'=============================================================
'
'   Nota:   Tutti i vettori e le matrici di queste routines
'           iniziano dall' indice 1 (ZCol() escluso).
'
Option Explicit
'
Const XDMin# = -25# ' Min of X grid
Const XDMax# = 25#  ' Max of X
Const YDMin# = -25# ' Min of Y
Const YDMax# = 25#  ' Max of Y
Const ZDMin# = 0#   ' Min of Z
Const ZDMax# = 50#  ' Max of Z
'
Dim ND&             ' Number of data in the vectors.
Dim PhiD#()         ' Angle Phi data.
Dim ThetaD#()       ' Angle Theta data.
Dim XD#()           ' Vector data values
Dim YD#()           '  of the surface
Dim ZD#()           '  to be interpolated.
Dim OXD#, OYD#      ' X/Y offset values.
Dim RD#             ' Radius of data.
'
Dim Xs#(), Ys#()    ' Coordinates of the data point grid.
'
Dim NXI&, NYI&      ' Number of columns and rows in
                    ' the interpolated points grid.
Dim PhiI#(), ThetaI#()    ' Coordinates of the interpolated points grid.
Dim XI#(), YI#()    ' Coordinates of the interpolated points grid.
Dim ZI#()           ' Interpolated surface.
Dim ZI_default#     ' Default value of Interpolated surface.
Dim ZC#()           ' Calculated surface.
'
' Settings for MASUB:
Dim TP#
'
' Impostazioni per QSHEP2D:
Dim NQ&, NW&, NR&
'
Dim ZCol&()         ' Table of colors.
Const NTCol& = 1280 ' Number of colors available in ZCol ().
Dim NLiv&           ' Number of levels to trace.
'
Dim FolderN$        ' Folder dei files dati.
'
Dim lZxy&           ' Indice della funzione di prova.
'
Dim bFilterEnabled As Boolean   ' Enable filter on data file proces.
Dim bDrawUpVal As Boolean       ' Draw up the values of the level lines.
Dim bDrawGD As Boolean          ' Draw the darts of the gradient.
Dim bDrawZC As Boolean          ' Draw the calculated surface.
Dim Title$                      ' Title of the picOrg picture.
'

Public Sub SurFit_Open(FileNameFull$)
'
'
    ZCol() = ColorTable(NTCol)
'
    NXI = 50
    NYI = 50
    NLiv = 20   ' Number of levels to trace.
'
    bFilterEnabled = True
'
'    optZxy(1).Value = True
'
    If BreakDown(FileNameFull$, FolderN$, Title$) Then
        ProcessDataFile FileNameFull$
    End If
'
'
'
End Sub

Private Sub Test_MASUB()
'
'   Interpolation of a surface with data points in
'   the XD (), YD (), ZD () vectors:
'
    Dim A#, B#, C#, D#, Px3!, Py3!
    Dim IC&, IEX&
'
    ' Prepare the PhiI(), ThetaI(), XI () and YI () vectors with
    '  the coordinates of the interpolation grid:
    GridForInterpolation A, B, C, D, 0.1
'
    ' Parameter setting for MASUB:
    IC = 1      ' First and only call.
    'IEX = 1     ' Extrapolation is required.
    ZI_default = 0
'
    If Not MASUB(IC, IEX, ND, XD(), YD(), ZD(), TP _
               , NXI, NYI, XI(), YI(), ZI(), ZI_default) Then
        MsgBox "Error in MASUB", vbCritical
        Exit Sub
    End If
    If IEX = 1 Then
        'lblNAdd = UBound(XD) - ND   ' Points added for extrapolation.
    End If
'
    DrawLevels A, B, C, D, Px3, Py3
'
    'picOrg.AutoRedraw = False
    'picSurFit.AutoRedraw = False
'
'
'
End Sub

Public Sub GridPointsData(XD#(), YD#(), XGD#(), YGD#())
'
'   Prepare the vector vectors of the data points, eliminating
'   the double values and ordering them in increasing direction:
'
    XGD() = XD()
    YGD() = YD()
'
    QuickSort XGD(), 1, UBound(XGD), 1
    QuickSort YGD(), 1, UBound(YGD), 1
'
    XGD() = Decima(XGD())
    YGD() = Decima(YGD())
'
'
'
End Sub

Private Sub DrawLevels(ByVal A#, ByVal B#, ByVal C#, ByVal D#, _
    ByRef Px3!, ByRef Py3!)
'
    ' Draw the surface in 3D:
    'frm3D.Surface XI#(), YI#(), ZI#(), Title$
    frm3D.Surface PhiD#(), ThetaD#(), XI#(), YI#(), ZI#(), Title$, False, XDMin, XDMax, YDMin, YDMax, ZDMin, ZDMax
'
'
'
End Sub

Public Sub SurFit_Close(Cancel As Integer)
'
'
    If IsLoaded(frm3D) Then Unload frm3D
'
'
'
End Sub

Private Sub GridForInterpolation(ByRef A#, ByRef B#, ByRef C#, ByRef D#, _
    Optional ByVal est# = 0)
'
'   Prepare the vectors containing abscissa and order the interpolation grid.
'   It also calculates the extremes of the interpolation coordinates,
'   possibly extending them to the Est factor:
'   to be used, mainly, for MASUB which is easily mistaken when the
'   interpolation extremes coincide with the ends of the data points.
'
    Dim I&, J&, HX#, HY#
'
    ' To delete points added by a previous call to MASUB:
    ReDim Preserve XD(1 To ND)
    ReDim Preserve YD(1 To ND)
    'lblND = ND
    'lblNAdd = "--"
'
    ' Find the Max. and Min. coordinates of the data points:
    A = DMINVAL(XD())   ' Minimum X.
    B = DMAXVAL(XD())   ' Maximum X.
    C = DMINVAL(YD())   ' Minimum Y.
    D = DMAXVAL(YD())   ' Maximum Y.
    'lblXMin = Format$(A, "#0.000")
    'lblXMax = Format$(B, "#0.000")
    'lblXMid = Format$((B + A) / 2#, "#0.000")
    'lblYMin = Format$(C, "#0.000")
    'lblYMax = Format$(D, "#0.000")
    'lblYMid = Format$((D + C) / 2#, "#0.000")
'
    ' And widens the Est factor:
    HX = (B - A)
    A = A - est * HX
    B = B + est * HX
    HY = (D - C)
    C = C - est * HY
    D = D + est * HY
'
    ReDim PhiI(1 To NXI), ThetaI#(1 To NYI) ' Angle phi and theta of the interpolated points grid.
    ReDim XI(1 To NXI), YI#(1 To NYI)   ' Coordinates of the interpolated points grid.
    ReDim ZI(1 To NXI, 1 To NYI)        ' Interpolated surface.
    ReDim ZC(1 To NXI, 1 To NYI)        ' Calculated surface.
'
    ' Abscissas of the grid of the interpolated points:
    HX = (B - A) / CDbl(NXI - 1)
    For I = 1 To NXI
        XI(I) = A + (I - 1) * HX
    Next I
'
    ' Ordinates of the grid of the interpolated points:
    HY = (D - C) / CDbl(NYI - 1)
    For J = 1 To NYI
        YI(J) = C + (J - 1) * HY
    Next J
'
'
'
End Sub

Private Sub DefaultParameters()
'
'   Attribuisce i valori di default ai parametri delle
'   routines di interpolazione.  Questa routine viene
'   richiamata ogni volta che si generano nuovi dati
'   casuali o si leggono i dati di un file.
'
    ' Impostazione parametri per MASUB:
    'IEX = 1     ' E' richiesta l' estrapolazione.
    TP = 10#    ' Tensione della superficie (TP >= 0).
'
    ' Impostazione dei parametri per QSHEP2:
    NQ = 13  ' 5 <= NQ <= MIN(40,ND-1)
    NW = 19  ' 1 <= NW <= MIN(40,ND-1)
    NR = MAX0(1, Sqr(ND / 3))       ' 1 <= NR
'
'
'
End Sub

Private Sub ProcessDataFile(ByVal FileN$)
'
'
    Dim FF%
    Dim lND&        ' Number of data in the vectors.
    Dim lPhiD#()    ' Angle Phi data.
    Dim lThetaD#()  ' Angle Theta data.
    Dim lXD#()      ' Vector data values
    Dim lYD#()      ' of the surface
    Dim lZD#()      ' to be interpolated.
    Dim lZDAvg#     ' Average of ZD().
    Dim lZDMin#     ' Min of ZD().
    Dim lDSkip() As Boolean   ' Flag data will skip.
    Dim I%, J%
'
    On Error GoTo ProcessDataFile_ERR
'
    Screen.MousePointer = vbHourglass
    DoEvents
'
    FF = FreeFile
    Open FileN$ For Input As #FF
'
    ' Read the offset X/Y data from the file:
    If (Not EOF(FF)) Then
        Input #FF, OXD, OYD, RD
        'lblOffsetX = OXD
        'lblOffsetY = OYD
        'lblRadius = RD
    End If
'
    If (bFilterEnabled = False) Then
        ' Read the data points from the file:
        ND = 0
        Do While Not EOF(FF)
            ND = ND + 1
            ReDim Preserve PhiD(1 To ND), ThetaD(1 To ND), XD(1 To ND), YD(1 To ND), ZD(1 To ND)
            Input #FF, PhiD(ND), ThetaD(ND), XD(ND), YD(ND), ZD(ND)
            'XD(ND) = XD(ND) - OXD
            'YD(ND) = YD(ND) - OYD
            'If (Sqr(XD(ND) ^ 2 + YD(ND) ^ 2) > 19#) Then
            '    ND = ND - 1
            'End If
        Loop
        'lblNAdd = 0
    Else ' Else of If (bFilterEnabled = False) Then
        ' Read the data points from the file:
        lND = 0
        Do While Not EOF(FF)
            lND = lND + 1
            ReDim Preserve lPhiD(1 To lND), lThetaD(1 To lND), _
                            lXD(1 To lND), lYD(1 To lND), lZD(1 To lND)
            Input #FF, lPhiD(lND), lThetaD(lND), lXD(lND), lYD(lND), lZD(lND)
            lXD(lND) = lXD(lND) - OXD
            lYD(lND) = lYD(lND) - OYD
            If (RD <> 0) And (Sqr(lXD(lND) ^ 2 + lYD(lND) ^ 2) > RD) Then
                lND = lND - 1
            End If
        Loop
'
        ' Sort the vectors so that the points with major Z remain behind:
        QuickSort5V lZD(), lXD(), lYD(), lPhiD(), lThetaD(), 1, lND, 1
'
        ' Skip point over Z axis.
        Dim lZOk As Boolean
        Dim lZOverCnt%
        ReDim Preserve lDSkip(1 To lND)
'
        ND = 0
        lZDAvg = 0
        lZDMin = ZDMax
        For I = 1 To lND
            lZOk = False
            lZOverCnt = 0
            For J = 1 To I - 1
                If (lDSkip(J) = False) Then
                    If (Abs(lXD(I) - lXD(J)) < 2#) And (Abs(lYD(I) - lYD(J)) < 2#) Then
                        lZOverCnt = lZOverCnt + 1
                        If (lZD(I) < lZD(J) + 3#) Then
                            lZOk = True
                            Exit For
                        End If
                    End If
                End If
            Next J
            If lZOverCnt = 0 Or lZOk = True Then
                ND = ND + 1
                ReDim Preserve PhiD(1 To ND), ThetaD(1 To ND), _
                                XD(1 To ND), YD(1 To ND), ZD(1 To ND)
                PhiD(ND) = lPhiD(I)
                ThetaD(ND) = lThetaD(I)
                XD(ND) = lXD(I)
                YD(ND) = lYD(I)
                ZD(ND) = lZD(I)
                lZDAvg = lZDAvg + lZD(I)
                If (lZD(I) < lZDMin) Then lZDMin = lZD(I)
            Else
                lDSkip(I) = True
            End If
        Next I
'
        lZDAvg = lZDAvg / ND
'
        ' Fill data for BIN shape.
        Dim BinX#, BinY#, Distance#, DistanceMin#, DistanceMinZ#
        Dim AddBin As Boolean
        Const DistanceRangeMin# = 2#
        Dim DistanceRangeMax#
        DistanceRangeMax = RD
'
        lND = ND
        'lblNAdd = 0
'
        For I = 0 To 360 - 1 Step 10
            BinX = RD * Sin(I * PI / 180#)
            BinY = RD * Cos(I * PI / 180#)
            AddBin = True
            DistanceMin = DMAX1(XDMax, YDMax)
            DistanceMinZ = lZDAvg
            For J = 1 To lND
                If (Abs(BinX - XD(J)) < DistanceRangeMin) _
                    And (Abs(BinY - YD(J)) < DistanceRangeMin) Then
                    AddBin = False
                    Exit For
                End If
                Distance = Sqr((BinX - XD(J)) ^ 2 + (BinY - YD(J)) ^ 2)
                'If (Distance < (lZDAvg - lZDMin)) Then
                If (Distance < DistanceRangeMin) Then
                    AddBin = False
                    Exit For
                End If
                If (Distance < DistanceMin) Then
                    DistanceMin = Distance
                    DistanceMinZ = ZD(J)
                End If
            Next J
            If (AddBin = True) Then
                ND = ND + 1
                ReDim Preserve PhiD(1 To ND), ThetaD(1 To ND), _
                                XD(1 To ND), YD(1 To ND), ZD(1 To ND)
                PhiD(ND) = I
                ThetaD(ND) = DistanceMin
                XD(ND) = BinX
                YD(ND) = BinY
                If (DistanceMin > DistanceRangeMax#) Then
                    ZD(ND) = lZDMin
                Else
                    ZD(ND) = _
                        (lZDMin - DistanceMinZ) _
                        / (DistanceRangeMax - 0) _
                        * DistanceMin _
                        + DistanceMinZ
                End If
            End If
        Next I
'
        'lblNAdd = ND - lND
    End If ' End of If (bFilterEnabled = False) Then
'
    Call DefaultParameters
'
    ' Prepare a grid corresponding to data points:
    GridPointsData XD(), YD(), Xs(), Ys()
'
    ' Call the interpolation routine:
    bDrawZC = False
    Test_MASUB
'
'
ProcessDataFile_ERR:
    Close FF
    Screen.MousePointer = vbDefault
'
    If (Err <> 0) Then
        MsgBox Err.Description, vbCritical, Err.Source
    End If
'
'
'
End Sub




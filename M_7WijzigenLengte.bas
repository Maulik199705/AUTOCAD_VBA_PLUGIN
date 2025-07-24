Attribute VB_Name = "M_7WijzigenLengte"
    Option Explicit
    
    Sub WijzigenLengte()
    ' --------------------------------------------------------------
    ' ACHTERAF HANDMATIG WIJZIGEN VAN LEIDINGLENGTE (2 LIJNEN + ARC)
    ' EBR: 29 APRIL 2004
    ' --------------------------------------------------------------
    
    On Error Resume Next
    
    Dim dReserveLengte As Double
    dReserveLengte = F_Main.TextBox1.Text
    
    ' --------------------------------------------------------------
    ' SELECTEREN LIJN
    ' --------------------------------------------------------------
    
    F_Main.Hide

    Dim LijnObj1 As AcadLine
    Dim bEindpuntInkortenLijn1 As Boolean
    Dim bEindpuntInkortenLijn2 As Boolean
    
    Dim P0 As Variant
    Dim p1 As Variant
    Dim p2 As Variant
    Dim P3 As Variant
    Dim P6 As Variant

    On Error Resume Next

    ThisDrawing.Utility.GetEntity LijnObj1, p1, "Selecteer leiding:"
    If Err <> 0 Then
        Err.Clear
        'MsgBox "Verkeerd geselecteerd"
        Exit Sub
    End If

    If LijnObj1.EntityName <> "AcDbLine" Then
        MsgBox "Geen lijn geselecteerd.", vbCritical, "Let op"
        Exit Sub
    End If
        

    ' --------------------------------------------------------------
    ' BEPALEN OF BEGIN OF EINDPUNT GESELECTEERD IS
    ' --------------------------------------------------------------

    If Lengte(LijnObj1.EndPoint, p1) > Lengte(LijnObj1.StartPoint, p1) Then
        'MsgBox "lijn bij beginpunt geselecteerd"
        p2 = LijnObj1.StartPoint
        P0 = LijnObj1.EndPoint
        bEindpuntInkortenLijn1 = False
    Else
        'MsgBox "lijn bij eindpunt geselecteerd"
        p2 = LijnObj1.EndPoint
        P0 = LijnObj1.StartPoint
        bEindpuntInkortenLijn1 = True
    End If
    
    
    ' --------------------------------------------------------------
    ' ZOEK ARC OP EINDPUNT VAN LIJN
    ' ZOEK VOLGENDE LIJN AAN ANDERE ZIJDE VAN GEVONDEN ARC
    ' --------------------------------------------------------------
    
    
    Dim sZoekHandle As String
    Dim ArcObj As AcadObject
    
    sZoekHandle = ZoekElementHandleRondPunt(p2, 3, LijnObj1.handle)
    Set ArcObj = ThisDrawing.HandleToObject(sZoekHandle)
    
    If ArcObj.EntityName <> "AcDbArc" Then
        MsgBox "Geen bocht (arc) gevonden.", vbInformation, "Inkorten niet mogelijk"
        Exit Sub
    End If
    
    
    'BEPAAL OF BEGIN- OF EINDPUNT VAN ARC GEVONDEN IS:
    
    Dim bBeginpuntGevondenArc As Boolean
    bBeginpuntGevondenArc = BepaalBeginpunt(p2, ArcObj)
    
    If bBeginpuntGevondenArc = True Then
        P3 = ArcObj.EndPoint
    Else
        P3 = ArcObj.StartPoint
    End If
    
    'Punt (P3)
    
    
    'ZOEK LIJN AAN ANDERE ZIJDE VAN ARC:
    
    Dim LijnObj2 As AcadLine
    sZoekHandle = ZoekElementHandleRondPunt(P3, 3, ArcObj.handle)
    
    Set LijnObj2 = ThisDrawing.HandleToObject(sZoekHandle)
    
    
    'BEPAAL OF BEGIN- OF EINDPUNT VAN LINE GEVONDEN IS:
    
    Dim sBeginpuntGevondenLijn2 As Boolean
    sBeginpuntGevondenLijn2 = BepaalBeginpunt(P3, LijnObj2)
    
    If sBeginpuntGevondenLijn2 = True Then
        bEindpuntInkortenLijn2 = False
    Else
        bEindpuntInkortenLijn2 = True
    End If
        
    
    
    ' --------------------------------------------------------------
    ' OPVRAGEN MODEMACRO
    ' --------------------------------------------------------------
         
''''    Dim sModemacro As String
''''    Dim pos As Integer
''''    Dim sTekst As String
''''
''''    Dim sGemetenLengte As String
''''    Dim sRollengte As String
''''
''''    ' sModemacro = "Gemeten lengte=5 m."
''''    sModemacro = ThisDrawing.GetVariable("modemacro")
''''
''''    pos = InStr(1, sModemacro, "Gemeten lengte=")
''''    If pos <> 0 Then
''''        sTekst = Right(sModemacro, Len(sModemacro) - 15)
''''        pos = InStr(1, sTekst, "=")
''''        sTekst = Right(sTekst, Len(sTekst) - pos)
''''        pos = InStr(1, sTekst, " m.")
''''        sTekst = Left(sTekst, pos)
''''        sTekst = Trim(sTekst)
''''        sGemetenLengte = sTekst
''''        ' MsgBox "Gemeten lengte=" & sTekst
''''    End If
''''
''''    pos = InStr(1, sModemacro, "Rollengte=")
''''    If pos <> 0 Then
''''        sTekst = Right(sModemacro, Len(sModemacro) - 10)
''''        pos = InStr(1, sTekst, "m")
''''        sTekst = Left(sTekst, pos - 1)
''''        sTekst = Trim(sTekst)
''''        sRollengte = sTekst
''''        ' MsgBox " Rollengte=" & sTekst
''''    End If
''''
''''    Dim dRollengte As Double
''''    Dim dGemetenLengte As Double
''''
''''    'dGemetenLengte = CDbl(sGemetenLengte)
''''    'dRollengte = CDbl(sRollengte)
''''
''''    sGemetenLengte = Replace(sGemetenLengte, ",", ".")
''''
''''    dGemetenLengte = Val(sGemetenLengte)
''''    If Err Then MsgBox Err.Description
''''    dRollengte = Val(sRollengte)
''''    If Err Then MsgBox Err.Description

    
    ' *************************************************************
    ' BOVENSTAANDE 14 JUNI 2004 VERVANGEN
    ' GEGEVENS NIET LEZEN UIT MODEMACRO, MAAR UIT REGISTRY
    ' *************************************************************
    
    Dim dRollengte As Double
    Dim dGemetenLengte As Double
    
    dGemetenLengte = F_LengteMonitor.TextBox1.Text
    
    dRollengte = GetSetting("Leidinglegprogramma", "Startup", "InvoerRolLengte", "")
    If Err Then
        Err.Clear
        dRollengte = 0
    End If
    
    ' --------------------------------------------------------------
    ' OPGEVEN LENGTEVERSCHIL
    ' --------------------------------------------------------------
   
         
    Dim dLengteVerschil As Double
    dLengteVerschil = InputBox("Geef lengteverschil in cm op (+ of -)" _
        & Chr(10) & Chr(13) _
        & Chr(10) & Chr(13) & "*  gemeten lengte=" & dGemetenLengte & " m" _
        & Chr(10) & Chr(13) & "-  Laatst opgegeven rollengte=" & dRollengte & " m" _
        & Chr(10) & Chr(13) & "-  reserve lengte=" & dReserveLengte & " m" _
        & Chr(10) & Chr(13) _
        & Chr(10) & Chr(13) & "-  lengteverschil=" & Round((dRollengte - dGemetenLengte - dReserveLengte), 1) & " m" _
        & Chr(10) & Chr(13) _
        & Chr(10) & Chr(13) _
        & "                           [Enter] = bevestigen" _
        , "Wijzigen leidinglengte", Round(100 * (dRollengte - dGemetenLengte - dReserveLengte), 1))
        
    dLengteVerschil = dLengteVerschil / 2
    
    If Err Then
        Err.Clear
        dLengteVerschil = 0
    End If
    
    If dLengteVerschil = 0 Then
        
        P6 = ThisDrawing.Utility.GetPoint(ArcObj.Center, "Geef eindpunt aan")
        
        dLengteVerschil = Lengte(ArcObj.Center, P6)
    
        If Lengte(P0, p1) < Lengte(P0, P6) Then
            'MsgBox "verlengen"
        Else
            'MsgBox "inkorten"
            dLengteVerschil = -dLengteVerschil
        End If
        
        
        
        If Err Then
            Err.Clear
            'Exit Sub
        End If
        
    End If
    
    
    
    ' --------------------------------------------------------------
    ' INKORTEN LIJNEN
    ' --------------------------------------------------------------
      
    LijnlengteWijzigen LijnObj1, LijnObj1.Length + dLengteVerschil, bEindpuntInkortenLijn1
    
    LijnlengteWijzigen LijnObj2, LijnObj2.Length + dLengteVerschil, bEindpuntInkortenLijn2
      
    
    ' --------------------------------------------------------------
    '  VERPLAATSEN ARC
    ' --------------------------------------------------------------
    
    'BEPALEN OF BEGIN- OF EINDPUNT VAN LIJN GEVONDEN IS
    
    Dim P5 As Variant
    If bEindpuntInkortenLijn1 = True Then
        P5 = LijnObj1.EndPoint
    Else
        P5 = LijnObj1.StartPoint
    End If
    
    
    If bBeginpuntGevondenArc = True Then
        ArcObj.Move ArcObj.StartPoint, P5
    Else
        ArcObj.Move ArcObj.EndPoint, P5
    End If
    
    ArcObj.Update

      

    
End Sub




    
    
''''    Dim sHoekInGraden As String         'LET OP STRING
''''    sHoekInGraden = ThisDrawing.Utility.AngleToString(ReturnObj1.Angle, acDegrees, 3)
''''
''''    'ThisDrawing.SetVariable "SNAPANG", sHoekInGraden           'WERKT GOED: ALS STRING MAAR NIET IN GROTER PROGRAMMADEEL ?
''''    ThisDrawing.ActiveViewport.SnapRotationAngle = ReturnObj1.Angle
''''
''''
''''
''''    'ThisDrawing.SetVariable "ORTHO", "ON"                      'GEEFT ERROR
''''    ThisDrawing.ActiveViewport.OrthoOn = True                  'LUKT OOK NIET
''''    'ThisDrawing.SendCommand ("ORTHO" & vbCr & "ON" & vbCr)      'GAAT WEL



   
     
''''   ThisDrawing.Regen acActiveViewport
''''
''''    Dim currViewport As AcadViewport
''''    Set currViewport = ThisDrawing.ActiveViewport
''''    ThisDrawing.ActiveViewport = currViewport
''''    Application.Update
''''
''''
''''    'ThisDrawing.ActiveViewport = ThisDrawing.ActiveViewport
''''    'ThisDrawing.ActiveSpace = acModelSpace
''''    'ThisDrawing.Regen True
''''
''''    P3 = ThisDrawing.Utility.GetPoint(P2, "Geef de richting aan.")
'''''
'''''    ThisDrawing.ActiveViewport.OrthoOn = False
'''''    ThisDrawing.ActiveViewport.SnapRotationAngle = 0
''''

Function BepaalBeginpunt(Punt, Obj) As Boolean
    'EBR 29 APRIL 2004
    'DEZE FUNCTIE BEPAALT OF BEGIN- OF EINDPUNT VAN LINE OF ARC GEVONDEN IS.

    Select Case Obj.EntityName
    Case "AcDbLine", "AcDbArc"
    Case Else
        MsgBox "Geen line of Arc aangegeven", vbExclamation, "Let op"
        End
    End Select
    
    
    If Lengte(Punt, Obj.EndPoint) > Lengte(Punt, Obj.StartPoint) Then
        'MsgBox "beginpunt gevonden"
        BepaalBeginpunt = True
    Else
        'MsgBox "eindpunt gevonden"
        BepaalBeginpunt = False
    End If
End Function












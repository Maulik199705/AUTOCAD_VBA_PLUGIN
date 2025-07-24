Attribute VB_Name = "M_4BerekenKopStaartNW"


'***********************************************************
' NIEUW AANGEPAST DOOR EBR: 13 FEB 2003
'***********************************************************

Public Sub BerekenKopStaart1(Rollengte As Double)

    ' --------------------------------------------------------------
    ' VOORGAANDE LOGFILE VERWIJDEREN
    ' --------------------------------------------------------------

    Call VerwijderenLogFile

    SchrijfLogFile ("-------------------------------------------------------------------")
    SchrijfLogFile (Date & "      " & Time)
    SchrijfLogFile ("-------------------------------------------------------------------")
    SchrijfLogFile ("")
    SchrijfLogFile ("")
    
    ' --------------------------------------------------------------
    ' PLAATSEN UNDO-MARK
    ' --------------------------------------------------------------
    '7 APTIL 2003 NIEUW
    
    'ThisDrawing.StartUndoMark
    'DIT WERK NIET
    
    ThisDrawing.SetVariable "modemacro", "Rollengte=" & Round(Rollengte / 100, 1) & " m"
   
    
    ' --------------------------------------------------------------
    ' UITLEZEN RESERVE LENGTE (SNIT) (controle zie F_main)
    ' --------------------------------------------------------------
    
    On Error Resume Next
    Dim dSnit As Double
    dSnit = F_Main.TextBox1.Text
    dSnit = dSnit * 100
    SchrijfLogFile ("Reservelengte (snit) = " & dSnit & " cm")
    
    ' --------------------------------------------------------------
    ' NIEUWE ROLLENGTE (INCLUSIEF SNIT)
    ' --------------------------------------------------------------

    Rollengte = Rollengte - dSnit
    ' --------------------------------------------------------------
    ' FORMULIER VERBERGEN
    ' --------------------------------------------------------------
    
    
    F_Main.Hide

    ' --------------------------------------------------------------
    ' INVOEREN BEGINLENGTE (24 FEB 2003)
    ' --------------------------------------------------------------

    Dim dInvoerLengte As Double
    dInvoerLengte = Rollengte
      
    ' dInvoerLengte = InputBox("Vul de rollengte in:", "Invoer rollengte", "1000")
    
    SchrijfLogFile ("Ingevoerde lengte=" & dInvoerLengte)
    
    ' --------------------------------------------------------------
    ' NIEUWE LAAG AANMAKEN EN AKTIEF MAKEN
    ' --------------------------------------------------------------
    
    Call LaagAktief
    
    
    ' --------------------------------------------------------------
    ' SELECTEREN 1E GEDEELTE VAN DE AANVOERPOLYLINE (1 FEB 2004)
    ' aanvoer-polyline kan in twee stukken worden gesplits om rekenfouten
    ' bij complexe aanvoer-polyline te voorkomen.
    ' --------------------------------------------------------------
    
    If F_Main.CheckBox11.Value = True Then
        'MsgBox "Selecteer 1e gedeelte van de (gesplitste) aanvoerpolyline.", vbInformation, "Selecteer"
        
        Dim ReturnObj1a As Object
        On Error Resume Next
        Dim iFoutSelectieTeller As Integer
    
Opnieuw1a:
    
        If iFoutSelectieTeller > 1 Then End
        
        
        ThisDrawing.Utility.GetEntity ReturnObj1a, basePnt1, "Selecteer 1e gedeelte van de (gesplitste) aanvoerpolyline:"
        If Err <> 0 Then
            Err.Clear
            If iFoutSelectieTeller < 1 Then MsgBox "Verkeerde selectie", vbCritical, "Let op"
            iFoutSelectieTeller = iFoutSelectieTeller + 1
            GoTo Opnieuw1
            'Exit Sub
        End If
                
        If ReturnObj1a.EntityName <> "AcDbPolyline" Then
            MsgBox "Geen aanvoer-polyline geselecteerd: " & Right$(ReturnObj1a.EntityName, Len(ReturnObj1a.EntityName) - 4), vbCritical, "Let op"
            FoutSelectieTeller = iFoutSelectieTeller + 1
            GoTo Opnieuw1a
        End If
        
        iFoutSelectieTeller = 0
        
        ReturnObj1a.Color = acByLayer
        ReturnObj1a.Highlight (True)
        ReturnObj1a.Update
        
        
        'BEREKENEN LENGTE VAN GEBROKEN (BREAK) AANVOER-POLYLINE (EERSTE GEDEELTE)
        Dim dLengteAanvoerPolylineVanEersteGedeelte As Double
        dLengteAanvoerPolylineVanEersteGedeelte = LengtePolyline2(ReturnObj1a)
        
        ReturnObj1a.Layer = ThisDrawing.ActiveLayer.Name
        
        iFoutSelectieTeller = 0
        
        
        
         'MsgBox dLengteAanvoerPolylineVanEersteGedeelte
    
        dInvoerLengte = dInvoerLengte - dLengteAanvoerPolylineVanEersteGedeelte
    
        'MsgBox "dInvoerLengte= " & dInvoerLengte
        
        
        
    End If
    
    
    
    
   
    ' --------------------------------------------------------------
    ' SELECTEREN DE AANVOERPOLYLINE (1 FEB 2004)
    ' --------------------------------------------------------------
   
    
    
    Dim ReturnObj1 As Object
    On Error Resume Next
    'Dim iFoutSelectieTeller As Integer
    
Opnieuw1:
    
    If iFoutSelectieTeller > 1 Then End
    
    
    ThisDrawing.Utility.GetEntity ReturnObj1, basePnt1, "Selecteer de aanvoer-polyline."
    If Err <> 0 Then
        Err.Clear
        If iFoutSelectieTeller < 1 Then MsgBox "Verkeerde selectie", vbCritical, "Let op"
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw1
        'Exit Sub
    End If
            
    If ReturnObj1.EntityName <> "AcDbPolyline" Then
        MsgBox "Geen aanvoer-polyline geselecteerd: " & Right$(ReturnObj1.EntityName, Len(ReturnObj1.EntityName) - 4), vbCritical, "Let op"
        FoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw1
    End If
    
    iFoutSelectieTeller = 0
    
    ReturnObj1.Color = acByLayer
    
    Dim plineObj As AcadLWPolyline
    Set plineObj = ReturnObj1
    
    
    '1 feb 200: controle of polyline al is aangewezen (eerste gedeelte van gebroken aanvoer-poly):
    If F_Main.CheckBox11.Value = True Then
        If ReturnObj1a.handle = plineObj.handle Then
            MsgBox "Polyline is niet gesplitst: dezelfde aanvoer-polyline geselecteerd.", vbInformation, "Let op"
            End
        End If
    End If
    
    plineObj.Highlight (True)
    
    
    
    
''''''' >>> 13 MEI 2004 DEZE IETS LAGER GEPLAATST (NET NA TRIMMEN AANVOER-POLYLINE)
''''''    ' --------------------------------------------------------------
''''''    ' FILLET DE AANVOER POLYLINE (24 FEB 2003)
''''''    ' --------------------------------------------------------------
''''''
''''''    Set copyCircleObj = ReturnObj1.Copy()
''''''    ReturnObj1.Erase
''''''    ThisDrawing.SendCommand "_Fillet" & vbCr & "R" & vbCr & "5" & vbCr & "P" & vbCr & "L" & vbCr
''''''    ThisDrawing.Regen acAllViewports
''''''    Set plineObj = copyCircleObj
''''''    plineObj.Layer = ThisDrawing.ActiveLayer.Name
        
    
    
    ' --------------------------------------------------------------
    ' SELECTEREN RETOUR POLYLINE (24 MAART 2003)
    ' --------------------------------------------------------------
    
    
    Dim RetourPoly As Object
    Dim basePnt2 As Variant
    'On Error Resume Next
    
Opnieuw2:
    
    If iFoutSelectieTeller > 1 Then End
    
    ThisDrawing.Utility.GetEntity RetourPoly, basePnt2, "Selecteer de retour-polyline."
    
    If Err <> 0 Then
        Err.Clear
        If iFoutSelectieTeller < 1 Then MsgBox "Verkeerd geselecteerd", vbCritical, "Let op"
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw2
        'Exit Sub
    End If
            
    If RetourPoly.EntityName <> "AcDbPolyline" Then
        If iFoutSelectieTeller < 1 Then MsgBox "Geen retour-polyline geselecteerd: " & Right$(RetourPoly.EntityName, Len(RetourPoly.EntityName) - 4), vbCritical, "Let op"
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw2
        'Exit Sub
    End If
    
    iFoutSelectieTeller = 0
    
    RetourPoly.Color = acByLayer
    RetourPoly.Highlight (True)
     
    
   
   
    
    ' --------------------------------------------------------------
    ' SELECTEREN EERSTE LIJN (BEGIN VAN DE LEIDING)
    ' --------------------------------------------------------------
    
    Dim returnObj As Object
    'On Error Resume Next
    
Opnieuw3:
    
    If iFoutSelectieTeller > 1 Then End
    
    ThisDrawing.Utility.GetEntity returnObj, basePnt, "Selecteer het beginpunt van de eerste leiding (line)."
    
    If Err <> 0 Then
        Err.Clear
        If iFoutSelectieTeller < 1 Then MsgBox "Verkeerd geselecteerd", vbCritical, "Let op"
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw3
    End If
            
    If returnObj.EntityName <> "AcDbLine" Then
        MsgBox "Geen lijn geselecteerd: " & Right$(returnObj.EntityName, Len(returnObj.EntityName) - 4)
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw3
        'Exit Sub
    End If
    
    iFoutSelectieTeller = 0
        
    returnObj.Color = acByLayer
    returnObj.Update
    returnObj.Highlight True
    
     
    Dim dTotaleLengte As Double
    dTotaleLengte = returnObj.Length
    
    
    
    
    ' --------------------------------------------------------------
    ' BEPALEN SNIJPUNT VAN AANVOER-POLYLINE MET EERSTE LEIDING
    ' --------------------------------------------------------------
    
    Dim vSnijpunt As Variant
    vSnijpunt = BepalenVanSnijpunt(returnObj, plineObj, 3)
    'MsgBox vSnijpunt(0) & Chr(10) & Chr(13) & vSnijpunt(1), vbExclamation
    'Punt (vSnijpunt)
    
     
    
        
    
'''    6 AUG 2003 ONDERSTAANDE DEEL IS UITGEZET OMDAT BOVENSTAANDE FILLET-COMMANDO HIER VOOR IN DE
'''    PLAATS IS GEKOMEN.

    ' --------------------------------------------------------------
    ' TRIMMEN AANVOERPOLYLINE TOT SNIJPUNT MET EERSTE LEIDING (25 FEB 2003)
    ' --------------------------------------------------------------

    'INDIEN ER WEL EEN SNIJPUNT IS DAN:

    If vSnijpunt(0) = 0 And vSnijpunt(1) = 0 Then
        SchrijfLogFile ("aanvoer polyline heeft geen snijpunt met de eerste leiding (en zal niet worden getrimd).")

    Else
    
            'MsgBox "6 mei 2004, onderstaande uitgezet en nieuwe functie"
            
            Call TrimPolyline(plineObj, vSnijpunt)
''''        ONDERSTAANDE IS 6 MEI 2004 UITGEZET EN VERVANGEN DOOR NIEUWE FUNCTIE


''''        'OPVRAGEN BEGINPUNT VAN POLYLINE EN HIEROP PIJL PLAATSEN
''''        Dim coord As Variant    'vertex-array
''''        Dim varVert As Variant  'vertex
''''        Dim intVCnt As Integer  'vertex-counter
''''
''''        coord = plineObj.Coordinates
''''        'MsgBox coord(0) & ", " & coord(1)
''''
''''        For Each varVert In coord
''''            intVCnt = intVCnt + 1
''''        Next
''''
''''        'OPVRAGEN EINDPUNT LWT-POLYLINE
''''        'MsgBox "Laatste x =" & coord(intVCnt - 2)
''''        'MsgBox "Laatste y =" & coord(intVCnt - 1)
''''
''''        'nieuw array, want lwt-poly bestaat heeft geen z-coordinaat en snijpunt wel
''''        Dim NieuwEindpuntPoly(0 To 1) As Double
''''        NieuwEindpuntPoly(0) = vSnijpunt(0)
''''        NieuwEindpuntPoly(1) = vSnijpunt(1)
''''
''''        'LET OP: COORDINATE BESTAAT UIT TWEE VERTEXEN (X EN Y), COORDINATE IS EEN KNOOPPUNT.
''''        'DAAROM DELEN DOOR 2
''''        'EERSTE VERTEX BEGINT BIJ 0, DAAROM LAATSTE -1
''''
''''        plineObj.Coordinate(intVCnt / 2 - 1) = NieuwEindpuntPoly
''''        plineObj.Update

    End If
    
    ' --------------------------------------------------------------
    ' FILLET DE AANVOER POLYLINE (24 FEB 2003)
    ' --------------------------------------------------------------
    
    
    Set copyCircleObj = plineObj.Copy()
    ReturnObj1.Erase
    ThisDrawing.SendCommand "_Fillet" & vbCr & "R" & vbCr & "5" & vbCr & "P" & vbCr & "L" & vbCr
    ThisDrawing.Regen acAllViewports
    Set plineObj = copyCircleObj
    plineObj.Layer = ThisDrawing.ActiveLayer.Name
    
    
    ' --------------------------------------------------------------
    ' FILLET DE RETOUR POLYLINE (24 MAART 2003)
    ' --------------------------------------------------------------
    
    Set copyRetourPoly = RetourPoly.Copy()
    RetourPoly.Erase
    ThisDrawing.SendCommand "_Fillet" & vbCr & "R" & vbCr & "5" & vbCr & "P" & vbCr & "L" & vbCr
    ThisDrawing.Regen acAllViewports
    Set RetourPoly = copyRetourPoly
    
    RetourPoly.Layer = ThisDrawing.ActiveLayer.Name
    
           
    ' --------------------------------------------------------------
    ' OPVRAGEN LENGTE VAN AANVOER-POLYLINE (VAN GETRIMDE POLYLINE)
    ' --------------------------------------------------------------
    
    Dim dLengteAanvoerPolyline As Double
    dLengteAanvoerPolyline = LengtePolyline2(plineObj)
    
    If dLengteAanvoerPolyline > dInvoerLengte Then
        MsgBox "De lengte van de aanvoer-polyline tot het snijpunt (" & Round(dLengteAanvoerPolyline, 1) / 100 & " m) is groter dan de ingevoerde rollengte.", vbExclamation, "Let op"
        SchrijfLogFile ("De lengte van de aanvoer-polyline tot het snijpunt (" & Round(dLengteAanvoerPolyline, 1) / 100 & " m) is groter dan de ingevoerde rollengte.")
        End
    End If
    
    If 2 * dLengteAanvoerPolyline > dInvoerLengte Then
        MsgBox "Lengte van aanvoer- en retour-polyline > ingevoerde lengte (er kunnen geen slingers worden gemaakt.)", vbInformation
        SchrijfLogFile ("Lengte van aanvoer- + retour-polyline > ingevoerde lengte (er kunnen geen slingers worden gemaakt.")
        End
    End If
    
    'MsgBox "Totale lengte van aanvoer-polyline=" & dLengteAanvoerPolyline, vbInformation
    SchrijfLogFile ("lengte van aanvoer-polyline (tot snijp eerste leiding)=" & dLengteAanvoerPolyline)
        
    ' -----------------------------------------------------------------------
    ' GESELECTEERDE AANVOER-POLYLINE EN EERSTE LEIDING IN JUISTE LAYER PLAATSEN
    ' -----------------------------------------------------------------------
    
    plineObj.Layer = ThisDrawing.ActiveLayer.Name
    returnObj.Layer = ThisDrawing.ActiveLayer.Name
    
         
    ' -----------------------------------------------------------------------
    ' EXPLODEREN AANVOER-POLYLINE
    ' -----------------------------------------------------------------------
    
    ' *** NIEUW OP 19 MEI 2003
    'plineObj.Explode
    'plineObj.Delete
    
    ' -----------------------------------------------------------------------
    ' EXPLODEREN RETOUR-POLYLINE VAN GETRIMDE POLYLINE (NW 17 feb 2003)
    ' HIER WORDEN LOSSE LIJNEN OVERHEEN GETEKEND (IVM LENGTE BEREKENING)
    ' -----------------------------------------------------------------------
    
    'volgende 4 regels als test:
    'Dim PlineObj As AcadLWPolyline
    'Dim lijnobj As AcadLine
    'Set PlineObj = ThisDrawing.ModelSpace.Item(0)
    'Set lijnobj = ThisDrawing.ModelSpace.Item(1)

    'MsgBox "*** NIEUW 24 MAART 03"
    Call M_4BepalenLengteNEW.TekenLijnenOverPolyline(RetourPoly)

    '**** 26 FEB 2003 GETEST EN WERKT
    'Call M_BepalenLengteNEW.VerwijderenLosseLines
    
   
   
    ' --------------------------------------------------------------
    ' BEPALEN OF BEGIN OF EINDPUNT VAN DE LEIDING GESELECTEERD IS
    ' --------------------------------------------------------------
    
    Dim bStartpointGevonden As Boolean

    If Lengte(basePnt, returnObj.EndPoint) > Lengte(basePnt, returnObj.StartPoint) Then
        'MsgBox "Lijn aan beginzijde geselecteerd"
        bStartpointGevonden = False

        '180 DRAAIEN EERSTE (GESELECTEERDE) LEIDING (nieuw 20 mrt 2003, dit om berekening lijnlengtes e.d. te vereenvoudigen)
        Dim PnwBegin As Variant
        Dim PnwEind As Variant
        PnwBegin = returnObj.EndPoint
        PnwEind = returnObj.StartPoint
        returnObj.StartPoint = PnwBegin
        returnObj.EndPoint = PnwEind

        'MsgBox "Eerste leiding 180 graden geroteerd"
        bStartpointGevonden = True
        
    Else
        'MsgBox "Lijn aan eindzijde geselecteerd"
        bStartpointGevonden = True
    End If
    
    
    ' --------------------------------------------------------------------
    ' NIEUW 5 AUG 2003 FILLET AANVOERPOLYLINE MET 1E GESELECTEERDE LEIDING
    ' --------------------------------------------------------------------
    'MsgBox "5 AUG FILLET, TOT HIER"
    
    Dim coord1 As Variant    'vertex-array
    Dim varVert1 As Variant  'vertex
    Dim intVCnt1 As Integer  'vertex-counter

    coord1 = plineObj.Coordinates
    'MsgBox coord(0) & ", " & coord(1)

    For Each varVert1 In coord1
        intVCnt1 = intVCnt1 + 1
    Next

    'OPVRAGEN EEN NA LAATSTE EINDPUNT LWT-POLYLINE
    'MsgBox "Laatste x =" & coord(intVCnt - 4)
    'MsgBox "Laatste y =" & coord(intVCnt - 3)
    
    Dim FilletP1(0 To 2) As Double
    FilletP1(0) = coord1(intVCnt1 - 4)    'EEN NA LAATSTE X-COORD
    FilletP1(1) = coord1(intVCnt1 - 3)    'EEN NA LAATSTE Y-COORD
    
    'Punt (FilletP1)
    
    
   
    
    Dim Psnij1 As Variant
    'let op, gebruik acextendthisentity, anders bij moeilijke aanvoerpolylines fout!
    Psnij1 = returnObj.IntersectWith(plineObj, acExtendThisEntity)
    'MsgBox Psnij1(0) & "  " & Psnij1(1)
    
    If IsEmpty(Psnij1) Then
        MsgBox "Geen snijpunt gevonden. Afronding bij eerste leiding met aanvoerpolyline kan niet worden gemaakt", vbInformation, " Let op"
        
    Else
        
        Call M_6Fillet.Fillet5aug(FilletP1, returnObj.StartPoint, Psnij1, 5)
       
        
        Dim FilletArcObj As AcadEntity
        Set FilletArcObj = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)
        
        
        If FilletArcObj.EntityName <> "AcDbArc" Then
            MsgBox "Afronding bij aanvoerleiding is niet gelukt.", vbInformation, "Let op"
        Else
            'onthouden oorspronkelijke hoek van geselecteerde leiding, ter controle
            Dim dHoekvanreturnObj As Double
            dHoekvanreturnObj = returnObj.Angle
            
        
            Dim dEindpuntFilletArc As Variant
            'wijzigen eindpunt (lees beginpunt) van geselecteerde lijn op eindpunt van fillet-arc.
            
            'onthouden originele eindpunt van geselecteerde leiding
            Dim dOrgReturnObjStartPoint As Variant
            dOrgReturnObjStartPoint = returnObj.StartPoint
            
            returnObj.EndPoint = FilletArcObj.EndPoint
            dEindpuntFilletArc = FilletArcObj.StartPoint
            
            
            'controleren of juiste punt van arc gekozen is, anders punt omkeren
            If Round(dHoekvanreturnObj, 4) <> Round(returnObj.Angle, 4) Then

                'lijnen terugzetten
                returnObj.EndPoint = dOrgReturnObjStartPoint

                'nu tegen andere arc-uiteinde plaatsen
                returnObj.EndPoint = FilletArcObj.StartPoint
                dEindpuntFilletArc = FilletArcObj.EndPoint
            End If

            'laatste punt van aanvoer-polyline wijzigen
            Dim NieuwEindpuntPoly1(0 To 1) As Double
            NieuwEindpuntPoly1(0) = dEindpuntFilletArc(0)
            NieuwEindpuntPoly1(1) = dEindpuntFilletArc(1)
            plineObj.Coordinate(intVCnt1 / 2 - 1) = NieuwEindpuntPoly1
            plineObj.Update

            
        
        End If
    End If
    
    
    ' -----------------------------------------------------------------------
    ' EXPLODEREN AANVOER-POLYLINE
    ' -----------------------------------------------------------------------
    
    ' *** NIEUW OP 19 MEI 2003, VERPLAATST OP 6 AUG 2003 (KOMT VAN BOVEN)
    
    plineObj.Explode
    plineObj.Delete
    
    ' --------------------------------------------------------------------------
    ' GESELECTEERDE LINE VERLENGEN TOT HET SNIJPUNT MET DE AANVOER-POLYLINE
    ' INDIEN ER EEN SNIJPUNT IS.
    ' --------------------------------------------------------------------------
    'GEHEEL NIEUW 26 MAART 2003
    '***UITGEZET OP 6 AUG 2003
    
'    If vSnijpunt(0) <> 0 And vSnijpunt(1) <> 0 Then
'        'LIJN IS 180 GEDRAAID !
'        returnObj.EndPoint = vSnijpunt
'    End If
    
    
    ' --------------------------------------------------------------
    ' ALLE VOORGAANDE SELECTIESETS VERWIJDEREN
    ' --------------------------------------------------------------
    
    Call SelectiesetsVerwijderen
   
    
    ' --------------------------------------------------------------
    ' BEREKENEN WELKE LIJNEN MET ELKAAR VERBONDEN ZIJN
    ' --------------------------------------------------------------
   
    
    Dim ReturnEindpunt As Variant
    
    Dim element As Object
    Dim Beginpunt As Variant
    Dim Eindpunt As Variant
    Dim LijnGevonden As Boolean
    Dim StartpuntGevonden As Boolean
    Dim testobj As Object
    Dim bMeerdereLijnenGevonden As Boolean
    
    
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    'Dim ssetObj As AcadSelectionSet
    Dim LijnTeller As Integer
    Dim gpCode(0) As Integer
    Dim dataValue(0) As Variant
    Dim groupCode As Variant, dataCode As Variant
    Dim mode As Integer
    
    Dim RollengteVerschil As Double
    
    Dim VoorgaandeElement As Object
    Dim VoorVoorgaandeElement As Object
    Dim bVoorVoorgaandeElement_Aanwezig As Boolean
    Dim PSnijpVoorgaande As Variant
    
    
    Dim VorigeHandle As String
    VorigeHandle = returnObj.handle
    Dim t As Integer
    
    Dim VoorgaandeArcBeginpuntGevonden As Boolean
    
    ZoomAll
    
    
    
    
    ' -------------------------------------------------------------------------------------------------
    ' *** 27 MEI 2003: INSTELLEN LIMITS OFF.
    ' LET OP, INDIEN LIMITS ZEER GROOT INGESTELD STAAN DAN GAAT BEREKENING
    ' KOP-STAAT NIET GOED: MAAK LIMITS KLEINER (GELIJK AAN EXTENTS OF MAX 5X GROTER BIJVOORBEELD)
    ' -------------------------------------------------------------------------------------------------
     
    Dim dLimMin As Variant
    Dim dLimMax As Variant
    dLimMin = ThisDrawing.GetVariable("LIMMIN")
    dLimMax = ThisDrawing.GetVariable("LIMMAX")
    
    Dim dExtMin As Variant
    Dim dExtMax As Variant
    dExtMin = ThisDrawing.GetVariable("EXTMIN")
    dExtMax = ThisDrawing.GetVariable("EXTMAX")
    
    Dim bControleerLimits As Boolean
    If dLimMin(0) / dExtMin(0) > 5 Then bControleerLimits = True
    If dLimMin(1) / dExtMin(1) > 5 Then bControleerLimits = True
    
    If dLimMax(0) / dExtMax(0) > 5 Then bControleerLimits = True
    If dLimMax(1) / dExtMax(1) > 5 Then bControleerLimits = True
    
    If bControleerLimits = True Then MsgBox "Limits staan zeer groot (5x Extents) ingesteld t.o.v. tekening-extents." _
    & Chr(10) & Chr(13) & "Stel limits goed in om rekenfouten te voorkomen.", vbCritical, "Limits niet juist"
    'deze op 31 jan 2004 uitgezet (dus toch doorgaan met berekenen)
    'End
    
'    Dim currLimits As Variant
'    currLimits = ThisDrawing.Limits
'
'
    ThisDrawing.SetVariable "LIMCHECK", 0      'limits off
    
    
'    ThisDrawing.Limits = currLimits
'    ThisDrawing.Regen (acActiveViewport)
    
    
   
    
    
    
    '***************************************************************************************************
    '********************   START HOOFSROUTINE  ********************************************************
    '***************************************************************************************************


    
opnieuw:

    t = t + 1
    
    'On Error Resume Next
    ' VOORGAAND OBJECT AANMAKEN, BEHALVE DE EERSTE KEER
    If t > 1 Then
        Set VoorVoorgaandeElement = VoorgaandeElement
        'If Err Then Err.Clear
        bVoorVoorgaandeElement_Aanwezig = True
        PSnijpVoorgaande = PsnijpuntPublic
    Else
        bVoorVoorgaandeElement_Aanwezig = False
        PSnijpVoorgaande = vSnijpunt
    End If
    
    'MsgBox vSnijpunt(0) & Chr(10) & Chr(13) & vSnijpunt(1) & Chr(10) & Chr(13) & vSnijpunt(2)
    
    
   
    
    Set VoorgaandeElement = returnObj
    'If Err Then Err.Clear
    
    

    ' MsgBox "VoorgaandeElement=" & VoorgaandeElement.handle & Chr(10) & Chr(13) & "VoorVoorgaandeElement=" & VoorVoorgaandeElement.handle
   

    


    If bStartpointGevonden = True Then
        ReturnEindpunt = returnObj.StartPoint
    Else
        ReturnEindpunt = returnObj.EndPoint
    End If
    

    LijnGevonden = False
    LijnTeller = 0
    
    '---------------------------------------------------
    'Aanmaken selectieset
    '---------------------------------------------------
    'voor 31 maart, was 0.1 (programma stopte)
   
    p1(0) = ReturnEindpunt(0) - 0.3
    p1(1) = ReturnEindpunt(1) - 0.3
    p1(2) = 0
    p2(0) = ReturnEindpunt(0) + 0.3
    p2(1) = ReturnEindpunt(1) + 0.3
    p2(2) = 0
        
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")
    If Err Then ssetObj.Clear
    mode = acSelectionSetCrossing
    ssetObj.Select mode, p2, p1
    
    If Err Then MsgBox Err.Description, vbInformation, "FOUT 0"
    
    For Each element In ssetObj
        'MsgBox element.handle
        If element.handle <> VorigeHandle Then
            If element.EntityName = "AcDbLine" Or element.EntityName = "AcDbArc" Then
                '*** Nieuw 25 maart 2003
                'MsgBox Element.Layer & "   " & ThisDrawing.ActiveLayer.Name
                
                'ORIGINEEL VOOR 7 APRIL 2003
                'If element.Layer <> ThisDrawing.ActiveLayer.Name Then
                
                'NIEUW 7 APRIL 2003
                If UCase(element.Layer) = "LEGPLAN" Then
                
                    element.Highlight True
                    LijnTeller = LijnTeller + 1
                    Set returnObj = element
                    
                    '*** MsgBox returnObj.handle, , LijnTeller
                
                    returnObj.Layer = ThisDrawing.ActiveLayer.Name
                    returnObj.Color = acByLayer
                    
                End If
            End If
        End If
    Next element
    
    'MsgBox returnObj.handle
    
    If Err Then MsgBox Err.Description, vbInformation, "FOUT 1"
    
    ssetObj.Clear
    ssetObj.Delete
    
    If Err Then MsgBox Err.Description, vbInformation, "FOUT 2"
    
   
     
   
    
    
    If LijnTeller = 0 Then
        ZoomPrevious
        MsgBox "Totale berekende lengte = " & Round(dTotaleLengte / 100, 1) & " m" _
        & Chr(10) & Chr(13) & "(dus korter dan ingevoerde rollengte: retour-polyline handmatig afmaken (explode).", vbInformation, "Einde"
        Exit Sub
        
        
''''        UITBREIDING 21 MEI 2004: SELECTIE VAN ARC OF LINE ALS ER EEN OPENING IS EN KOP
''''        STAART NIET GEVONDEN KAN WORDEN. TOCH WEER UITGEZET: COMPLEX !

'''''        Dim SelectieObj As AcadEntity
'''''        Dim Pp As Variant
'''''        ThisDrawing.Utility.GetEntity SelectieObj, Pp
'''''        Set returnObj = SelectieObj
'''''
'''''        If Err Then MsgBox Err.Description
'''''
'''''        element.Highlight True
'''''        LijnTeller = LijnTeller + 1
'''''
'''''        returnObj.Layer = ThisDrawing.ActiveLayer.Name
'''''        returnObj.Color = acByLayer
        
        
       
        
        'Exit Sub
    End If
    
    
    'origineel:
    'If LijnTeller > 1 Then ZoomPrevious: MsgBox "Meerdere lijnen gevonden   ': Einde": End
    
    'aangepast op 25 mrt 2003
    'MsgBox "tot hier 25 mrt 2003"
    '31 maart 2003 aangepast
    If LijnTeller > 1 Then
        ZoomPrevious
        SchrijfLogFile ("Meerdere lijnen gevonden.")
        'MsgBox "Meerdere lijnen gevonden   ': Einde"
        'End        'doorgaan
    End If
    
     
    '---------------------------------------------------
    ' BEPALEN TOTALE LENGTE
    '---------------------------------------------------
    
    If returnObj.EntityName = "AcDbLine" Then
        dTotaleLengte = dTotaleLengte + returnObj.Length
        Debug.Print returnObj.Length & "            " & dTotaleLengte
    Else
        dTotaleLengte = dTotaleLengte + returnObj.ArcLength
        Debug.Print returnObj.ArcLength & "            " & dTotaleLengte
    End If
    
    
    '---------------------------------------------------
    ' BEPALEN BEGIN OF EINPUNT
    '---------------------------------------------------
    
    If Lengte(ReturnEindpunt, returnObj.EndPoint) > Lengte(ReturnEindpunt, returnObj.StartPoint) Then
       ' MsgBox "Lijn aan beginzijde geselecteerd"
        bStartpointGevonden = False
    Else
       ' MsgBox "Lijn aan eindzijde geselecteerd"
        bStartpointGevonden = True
    End If
    
    VorigeHandle = returnObj.handle
    
        
    '------------------------------------------------------------
    ' BEPALEN OF VAN ARC HET BEGIN- OF EINDPUNT GEVONDEN IS
    ' 30 MAART 2003 NIEUW
    '------------------------------------------------------------
    
    If returnObj.EntityName = "AcDbArc" Then
        VoorgaandeArcBeginpuntGevonden = bStartpointGevonden
    End If
    
    '*******************************************************
    ' BEREKENEN LENGTE TOT SNIJPUNT POLYLINE NIEUW 17 FEB 2003
    '*******************************************************
    Dim Debugmode As Boolean
    Debugmode = False
    
    If returnObj.EntityName <> "AcDbArc" Then
    
        If bStartpointGevonden = False Then
            'MsgBox "Laatste lijn roteren met 180 graden, (dan is het eindpunt altijd juist)."
            Dim PStartOrgineel As Variant
            Dim PEindOrgineel As Variant
            PStartOrgineel = returnObj.StartPoint
            PEindOrgineel = returnObj.EndPoint
            returnObj.StartPoint = PEindOrgineel
            returnObj.EndPoint = PStartOrgineel
            
            bStartpointGevonden = True
        End If
        
        lengtevanpoly = M_4BepalenLengteNEW.BepalenSnijpunt(RetourPoly, returnObj)
        
       
        'MsgBox "lengtevanpoly (retour-poly tot snijpunt)= " & lengtevanpoly & Chr(10) & Chr(13) _
        & "lengte van aanvoer poly= " & dLengteAanvoerPolyline & Chr(10) & Chr(13) _
        & "dTotaleLengte (losse lines en arcs)= " & dTotaleLengte & Chr(10) & Chr(13) _
        & Chr(10) & Chr(13) _
        & "totaal = " & lengtevanpoly + dLengteAanvoerPolyline + dTotaleLengte
         
        SchrijfLogFile ("dLengteAanvoerPolyline=" & Round(dLengteAanvoerPolyline, 1) & "        lengtevanpoly=" & Round(lengtevanpoly, 1) & "         dTotaleLengte=" & Round(dTotaleLengte, 1) & "      >>> " & Round(dLengteAanvoerPolyline + lengtevanpoly + dTotaleLengte, 1))
    
         
        'LET OP: SNIT IS AL IN INGEVOERDE ROLLENGTE VERWERKT
        If (dLengteAanvoerPolyline + lengtevanpoly + dTotaleLengte) > (dInvoerLengte) Then
        
        
        
        
        
       
        
        
        
        
        
' ***************************************************************************************************
' *** EINDE LENGTE-BEREKENING
' ***************************************************************************************************


            ' MsgBox "bVoorVoorgaandeElement_Aanwezig=" & bVoorVoorgaandeElement_Aanwezig


            'VERWIJDEREN RETOUR-POLYLINE
            RetourPoly.Delete
                         
            RollengteVerschil = dInvoerLengte - (dLengteAanvoerPolyline + lengtevanpoly + dTotaleLengte)
            'MsgBox "rollengte behaald ! Verschil = " & RollengteVerschil
            SchrijfLogFile ("rollengte behaald ! Verschil = " & RollengteVerschil)
            
            
            Application.Update
            
            
            'WEERGEVEN LENGTE-VERSCHIL
            If F_Main.CheckBox1.Value = True Then
                MsgBox "Rollengte: " & (dInvoerLengte) / 100 & " m" & Chr(10) & Chr(13) _
                    & "Reserve lengte: " & dSnit / 100 & " m" & Chr(10) & Chr(13) _
                    & Chr(10) & Chr(13) _
                    & "Lengte verschil: " & Abs(Round(RollengteVerschil, 0)) / 100 & " m", vbInformation, "Lengte verschil"
            End If
                    
            
            
            'lossen lines en arcs die over polyline getekend waren weer verwijderen (26 FEB 2003)
            Call M_4BepalenLengteNEW.VerwijderenLosseLines
            
 

            Dim Situatie As Integer
            Dim Peind As Variant
            Dim Pstart As Variant
            
            'Punt (PsnijpuntPublic)
            'Punt (returnObj.EndPoint)
            'PsnijpuntPublic = snijpunt leiding (returnobj) met aanvoer polyline
            'Psnijpunt = snijpunt van de eerste (geselecteerde) leiding met (geselecteerde) aanvoerpolyline
            
            SchrijfLogFile ("")
            
            If Lengte(PsnijpuntPublic, returnObj.StartPoint) > Lengte(PsnijpuntPublic, returnObj.EndPoint) Then
                SchrijfLogFile ("*SITUATIE 2    (volgende lijn doet ook mee in lengte berekening)")
                Situatie = 2
                'DEZE SITUATIE NOG BETER BEKIJKEN (LIMCHECK ?? EBR 21 MEI)
            Else
                SchrijfLogFile ("*SITUATIE 1    (voorgaande lijn doet mee in lengte berekening)")
                Situatie = 1
            End If




'*****************************************************************************************************************
''''            returnObj.Color = acGreen               'het laatste element
''''            VoorgaandeElement.Color = acMagenta     'de arc
''''            VoorVoorgaandeElement.Color = acRed     'de voorlaatste lijn
'*****************************************************************************************************************

            
            'MsgBox "returnObj.Length=" & returnObj.Length
            'MsgBox "RollengteVerschil=" & RollengteVerschil
                    
                   
            Dim dDraaipunt As Variant
            Dim dDraaiHoek As Variant
            
            Peind = returnObj.EndPoint
            Pstart = returnObj.StartPoint
                
            If Situatie = 1 Then
                    
                    If Peind(0) < PsnijpuntPublic(0) Then
                        If Debugmode Then MsgBox "SITUATIE 1 [A]:Snijp ligt (links) NEE RECHTS ! 28 MRT 2003, vbExclamation"
                        SchrijfLogFile ("SITUATIE 1 [A]")
                        F_Main.Label25.Caption = "SITUATIE 1 [A]"
                        
                        
                        'laatste lijn draaien en verkorten:
                        Call LijnlengteWijzigen(returnObj, returnObj.Length - (Abs(RollengteVerschil) / 2), True)
                        'InsertBlock "pijl", returnObj.EndPoint, "0"
                                               
                        'de voorlaatste lijn draaien inkorten
                        Call LijnlengteWijzigen(VoorVoorgaandeElement, VoorVoorgaandeElement.Length - (Abs(RollengteVerschil)) / 2, False)
                        
                        'de arc verplaatsen
                        VoorgaandeElement.Move VoorgaandeElement.StartPoint, returnObj.EndPoint      'dus startpoint
                    Else
                        If Debugmode Then MsgBox "SITUATIE 1 [B] Snijp ligt (rechts) NEE LINKS ! 28 MRT 2003, vbExclamation"
                        SchrijfLogFile ("SITUATIE 1 [B]")
                        F_Main.Label25.Caption = "SITUATIE 1 [B]"
                        
                        'GESPIEGELD T.O.V. BOVENSTAANDE SITUATIE
                        
                        'laatste lijn draaien en verkorten
                        Call LijnlengteWijzigen(returnObj, returnObj.Length - (Abs(RollengteVerschil) / 2), True)
                        
                        'de voorlaatste lijn draaien inkorten
                        Call LijnlengteWijzigen(VoorVoorgaandeElement, VoorVoorgaandeElement.Length - (Abs(RollengteVerschil)) / 2, False)
                        
                        'MsgBox "TOT HIER", vbExclamation
                        
                        'de arc verplaatsen
                        VoorgaandeElement.Move VoorgaandeElement.EndPoint, returnObj.EndPoint      'dus startpoint
                       
                    End If
                    
                    'EINDE PROGRAMMA
                    ThisDrawing.SetVariable "LIMCHECK", 0      'limits off, 27 MEI 2003
                    ZoomPrevious
                    Exit Sub
            End If
                                
' ***************************************************************************************************
                    
            
            If Situatie = 2 Then
            
                    Dim ArcCopyObj As AcadArc
                    'nieuw 30 maart 2003
                    Dim PbeginLaatsteLijn As Variant
                    PbeginLaatsteLijn = returnObj.StartPoint
                    
                    If Round(Peind(0), 0) < Round(PsnijpuntPublic(0), 0) Then
                            SchrijfLogFile ("SITUATIE 2 [A]")
                            F_Main.Label25.Caption = "SITUATIE 2 [A]"
                            'MsgBox "SITUATIE 2 [C], NOG BEKIJKEN, Snijp ligt RECHTS NEE LINKS !!!, hier gebelevden op 28 maart om 17:26 !", vbExclamation
                            
                                                                           
                            
                            'laatste lijn draaien en verkorten
                            Call LijnlengteWijzigen(returnObj, (returnObj.Length - Abs(RollengteVerschil)) / 2, False)
                            
                            
                            'de arc kopieren en roteren
                            Set ArcCopyObj = VoorgaandeElement.Copy()    'dus startpoint
                            ArcCopyObj.Rotate ArcCopyObj.StartPoint, 3.141592
                            
                            'ORIGINEEL VOOR 30 MAART
                            'ArcCopyObj.Move ArcCopyObj.EndPoint, returnObj.StartPoint
                             
                            'NIEUW 30 MA
                            'de arc veplaatsen
                            If VoorgaandeArcBeginpuntGevonden = bStartpointGevonden Then
                                ArcCopyObj.Move ArcCopyObj.StartPoint, returnObj.StartPoint
                            Else
                                ArcCopyObj.Move ArcCopyObj.EndPoint, returnObj.StartPoint
                            End If
                    
                    ElseIf Round(Peind(0), 0) > Round(PsnijpuntPublic(0), 0) Then
                            'MsgBox "SITUATIE 2 [B]:Snijp ligt LINKS TOT."
                            SchrijfLogFile ("SITUATIE 2 [B]")
                            F_Main.Label25.Caption = "SITUATIE 2 [B]"
                            
                            
                            
                            'laatste lijn draaien en verkorten
                            Call LijnlengteWijzigen(returnObj, (returnObj.Length - Abs(RollengteVerschil)) / 2, False)
                                                   
                            'de arc kopieren en roteren
                            Set ArcCopyObj = VoorgaandeElement.Copy()    'dus startpoint
                            ArcCopyObj.Rotate ArcCopyObj.EndPoint, 3.141592
                             
                            'ORIGINEEL VOOR 30 MAART
                            'ArcCopyObj.Move ArcCopyObj.EndPoint, returnObj.StartPoint
                             
                            'NIEUW 30 MAART
                            'de arc veplaatsen
                            If VoorgaandeArcBeginpuntGevonden = bStartpointGevonden Then
                                ArcCopyObj.Move ArcCopyObj.StartPoint, returnObj.StartPoint
                            Else
                                ArcCopyObj.Move ArcCopyObj.EndPoint, returnObj.StartPoint
                            End If
                            
                            'NIEUWE LIJN TEKENEN
                            'UITGEZET 7 AUG 2003: Call Punt1(PeindLaatsteLijn, acMagenta)
                            'UITGEZET 7 AUG 2003: Call Punt1(returnObj.StartPoint, acGreen)
                            
                            
                            '*** TOT HIER GEBLEVEN OP 30 MAART 2003 !!
                            'UITGEZET 7 AUG 2003:  Call TekenLijn(returnObj.StartPoint, PbeginLaatsteLijn, "0")
                            'Call LijnlengteWijzigen(returnObj, (returnObj.Length - Abs(RollengteVerschil)) / 2, False)
                    Else
                        SchrijfLogFile ("SITUATIE 3 !!")
                        F_Main.Label25.Caption = "SITUATIE 3 !!"
                        MsgBox "SITUATIE 3 !! (21 mei 2004, controle Ebr)", vbInformation
                        
                        'MsgBox "SITUATIE 3: Snijpunt is eindpunt lijn ! Einde", vbCritical
                        ''UITGEZET 7 AUG 2003: MsgBox "SITUATIE 2 [C], NOG BEKIJKEN, Snijp ligt RECHTS NEE LINKS !!!, hier gebelevden op 28 maart om 17:26 !", vbExclamation
                                                                   
                        
                        'laatste lijn draaien en verkorten
                        Call LijnlengteWijzigen(returnObj, (returnObj.Length - Abs(RollengteVerschil)) / 2, False)
                                                
                        'de arc kopieren en roteren
                        Set ArcCopyObj = VoorgaandeElement.Copy()    'dus startpoint
                        ArcCopyObj.Rotate ArcCopyObj.StartPoint, 3.141592
                        
                        'ORIGINEEL VOOR 30 MAART
                        'ArcCopyObj.Move ArcCopyObj.EndPoint, returnObj.StartPoint
                         
                        'NIEUW 30 MA
                        'de arc veplaatsen
                        If VoorgaandeArcBeginpuntGevonden = bStartpointGevonden Then
                            ArcCopyObj.Move ArcCopyObj.StartPoint, returnObj.StartPoint
                        Else
                            ArcCopyObj.Move ArcCopyObj.EndPoint, returnObj.StartPoint
                        End If
        
                    End If
                    
                    

                    'EINDE PROGRAMMA
                    '21 MEI 2004, ONDERSTAANDE REGEL ER BIJGEZET (KOPIE UIT SITUATIE 1)
                    ThisDrawing.SetVariable "LIMCHECK", 0      'limits off, 27 MEI 2003
                    
                    ZoomPrevious
                    Exit Sub
                    
                
            End If
            
            
' ***************************************************************************************************
' ***************************************************************************************************
' ***************************************************************************************************

             
        End If
        
       
    End If
    
    
    GoTo opnieuw
    

End Sub






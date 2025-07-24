Attribute VB_Name = "M_BerekenRollengte"
'Public lengte As Double
Public InvoerLengte As Double
Public Stoppen As Boolean
Public lengtevanpoly As Double
Public vSnijpunt As Variant

Public Sub BerekenRollengte()
    MsgBox "DEZE ROUTINE IS BESCHADIGD VRIJDAG 21 MAART 2003"
    MsgBox "dus verwijderen !!! (ebr 27 mei 2004)"
   
    ' ------------------------------------------------------------------------------
    ' BEPALEN LAATSTE GROEP-NUMMER (UIT LAYER)EN AANMAKEN NIEUWE LAYER
    ' ------------------------------------------------------------------------------
    
    SchrijfLogFile ("[1]   BEPALEN LAATSTE GROEP-NUMMER (UIT LAYER)EN AANMAKEN NIEUWE LAYER")

    Dim LaagObj As AcadLayer
    Dim sLaagNaam As String
    Dim sLaagNummer As String
    Dim iLaagNummer As String
    Dim iHoogsteNummer As Integer
    Dim sNwLaagnaam As String

    For Each LaagObj In ThisDrawing.Layers
        sLaagNaam = UCase(LaagObj.Name)
        If Left$(sLaagNaam, 6) = "GROEP_" Then

            sLaagNummer = Right(sLaagNaam, Len(sLaagNaam) - 6)
            If Val(sLaagNummer) <> sLaagNummer Then
                MsgBox "Laag " & sLaagNummer & " bevat een verkeerd nummer.", vbCritical, "Let op"
                End
            Else
                iLaagNummer = sLaagNummer
                If iLaagNummer > iHoogsteNummer Then iHoogsteNummer = iLaagNummer
            End If

        End If
    Next LaagObj

    'GROEPSNUMMER OPHOGEN MET 1:
    If iLaagNummer = "" Then iLaagNummer = 0
    
    'NIEUWE LAAGNUMMER
    iLaagNummer = iHoogsteNummer + 1
    sLaagNummer = str(iLaagNummer)
    
    Dim AantalNullen As Integer
    If Len(sLaagNummer) > 3 Then MsgBox "Meer dan 99 groepsnummers in de laagnaam Groep_QQ. >> Purge de lagen.", vbCritical, "Let op": End
    
    
    ' ------------------------------------------------------------------------------
    ' LAAGNAAM SAMENSTELLEN
    ' ------------------------------------------------------------------------------
    
    sNwLaagnaam = "groep_" & String(3 - Len(sLaagNummer), "0") & iLaagNummer
    'MsgBox "Laag wordt: " & sNwLaagnaam
    
    ' ------------------------------------------------------------------------------
    ' LAAG KLEUR BEPALEN
    ' ------------------------------------------------------------------------------
    
    

    'LAAG AANMAKEN:

    Dim NwLaagObj As Object
    Set NwLaagObj = ThisDrawing.Layers.Add(sNwLaagnaam)
    
    'Kleur bepalen
    
    'oorspronkelijk:
    'nwLaagObj.Color = acYellow
    'nieuw:
    
    Dim KleurNo As Integer
    On Error Resume Next
    
    'MsgBox "Vorige laag-groep: " & "groep_" & iHoogsteNummer
    
    KleurNo = ThisDrawing.Layers.Item("groep_" & iHoogsteNummer).Color
    If Err Then Err.Clear
    
    If KleurNo = 2 Then
        NwLaagObj.Color = acGreen
    Else
        NwLaagObj.Color = acYellow
    End If
    
    
   

    

    ' ---------------------------------------------------------------------
    ' LAAG AANMAKEN EN AKTIEF MAKEN
    ' ---------------------------------------------------------------------
    SchrijfLogFile ("[2]   LAAG AKTIEF MAKEN")

    ThisDrawing.ActiveLayer = NwLaagObj

    ' AanmakenLaag(laagnaam,kleur,True)
    
    ' ---------------------------------------------------------------------
    ' INVOEREN TOTALE LEIDINGLENGTE)
    ' ---------------------------------------------------------------------
    SchrijfLogFile ("[3]   INVOEREN TOTALE LEIDINGLENGTE)")

    'Dim Invoerlengte As Double 'public maken
    Dim Lengte As Double       'public maken
    Dim sInvoerlengte As String

    On Error Resume Next
    'listboxen verwijderd 24 feb 2003
    'If F_Main.ListBox6.Text <> "" Then
    '   sInvoerlengte = ListBox6.Text
    'Else
        'If F_Main.ListBox5.Text <> "" Then
        '    sInvoerlengte = F_Main.ListBox5.Text
        'Else
            If F_Main.CheckBox1.Value = True Then sInvoerlengte = InputBox("Voer rollengte in (meters).", "Invoer lengte")
        'End If
    'End If

    'MsgBox sInvoerlengte
    Dim pos As Integer
    pos = InStr(1, sInvoerlengte, ".")
    If pos <> 0 Then MsgBox "Gebruik een komma, maar geen punt.", vbExclamation, "Let op": End
    
    InvoerLengte = sInvoerlengte

    SchrijfLogFile ("   Invoerlengte (cm)=" & InvoerLengte)
    InvoerLengte = InvoerLengte * 100

    If Err Then MsgBox "Geen juiste lengte opgegeven.", vbCritical, "Leidingleg-programma": End
    SchrijfLogFile ("   Invoerlengte (m)=" & InvoerLengte)

    

    ' ----------------------------------------------------------------------------
    ' SELECTEREN POLYLINE:
    '
    ' SELECTEREN VAN DE AANVOER-POLYLINE T.B.V. HET BEREKENEN
    ' VAN DE LENGTE VAN HET LEGPATROON
    '((((INCLUSIEF DE LENGTE
    ' VAN DE RETOURLEIDING (NA BEPALEN SNIJPUNT)))??
    ' ----------------------------------------------------------------------------
    SchrijfLogFile ("[4]   SELECTEREN POLYLINE (aanvoer) en plaatsen pijl-block")

    'ThisDrawing.SetVariable "PDMODE", 35    'Laat snijpunt zien.

    If F_Main.CheckBox2.Value = True Then

        Dim RetourObj As Object
        On Error Resume Next
        ThisDrawing.Utility.GetEntity RetourObj, basePnt, "Selecteer de aanvoer-leiding."
        If RetourObj.EntityName <> "AcDbPolyline" Then MsgBox "Geen polyline geselecteerd.", vbCritical, "Let op": Exit Sub
        If Err <> 0 Then
            Err.Clear
            MsgBox "Verkeerd geselecteerd"
            End
        End If

        RetourObj.Layer = sNwLaagnaam
        RetourObj.Color = acByLayer
        RetourObj.Highlight True


        'BEREKENEN LENGTE VAN LWT-POLYLINE

        Lengte = LengtePolyline2(RetourObj)
        SchrijfLogFile ("     lengte van geselecteerde polyline=" & Lengte)



        'PLAATSEN BLOCK-PIJL
        'OPVRAGEN BEGINPUNT VAN LWTPOLYLINE EN AANGEVEN VAN PLAATS VERDELER

        'VOORGAANDE PIJLEN (BLOCKS) VERWIJDEREN
        Dim element As Object
        For Each element In ThisDrawing.ModelSpace
            If element.EntityName = "AcDbBlockReference" Then
                If UCase(element.Name) = "PIJL" Then
                    element.Erase
                End If
            End If
        Next element
        ThisDrawing.Regen (True)

        'OPVRAGEN BEGINPUNT VAN POLYLINE EN HIEROP PIJL PLAATSEN
        Dim coord As Variant    'vertex-array
        Dim intVCnt As Integer  'vertex-counter
        Dim varVert As Variant  'vertex

        coord = RetourObj.Coordinates
        'MsgBox coord(0) & ", " & coord(1)

        Dim BlockObj As AcadBlockReference
        Dim Pins(0 To 2) As Double
        'OPVRAGEN BEGINPUNT LWT-POLYLINE
        Pins(0) = coord(0)
        Pins(1) = coord(1)
        Pins(2) = 0
        Set BlockObj = ThisDrawing.ModelSpace.InsertBlock(Pins, "pijl", 1, 1, 1, 0)
        BlockObj.Layer = "0"
        BlockObj.Update
        
        
        ' ---------------------------------------------------------------------------------
        ' BEPALEN WELKE LIJN AAN GESELECTEERDE AANVOER-POLYLINE IS VERBONDEN.
        ' ---------------------------------------------------------------------------------
        
        
        SchrijfLogFile ("[5]   BEPALEN WELKE LIJN AAN GESELECTEERDE AANVOER-POLYLINE IS VERBONDEN.")


        Dim strHandle As String
        strHandle = BepaalLijnAanEindePolyline(RetourObj)
        'MsgBox strHandle

        'berekenen lengte v.a. eerste gewone lijn:
        Dim tempObj As AcadObject
        Set tempObj = ThisDrawing.HandleToObject(strHandle)

        'tempObj = IS DE EERSTE LIJN VERBONDEN AAN AANVOER-POLY
        
        ' ---------------------------------------------------------------------------------
        ' AUTOMATISCH BEPALEN WAT DE BIJBEHORENDE RETOURLEIDING (POLYLINE)IS.
        ' ---------------------------------------------------------------------------------

        'GESELECTEERDE AANVOER-POLYLINE:    RetourObj
        'BEGINPUNT AANVOER-POLYLINE:        coord(0), coord(1), 0

         SchrijfLogFile ("[6]   AUTOMATISCH BEPALEN WAT DE BIJBEHORENDE RETOURLEIDING (POLYLINE)IS.")



        ' MAKEN SELECTIESET:

        Dim ssetObj As AcadSelectionSet
        Dim mode As Integer

        Dim Px As Double
        Dim Py As Double

        Px = coord(0)
        Py = coord(1)

        On Error Resume Next
        Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
        If Err Then
            ThisDrawing.SelectionSets.Item("TEST_SSET").Clear
            ThisDrawing.SelectionSets.Item("TEST_SSET").Delete
        End If

        mode = acSelectionSetCrossing

        Dim Tel As Integer
        'Dim sHandle As String
        Dim corner1(0 To 2) As Double
        Dim corner2(0 To 2) As Double
        Dim ssElement As Object


        Dim gpCode(0) As Integer
        Dim dataValue(0) As Variant
        gpCode(0) = 0
        dataValue(0) = "LWPOLYLINE"

        Dim groupCode As Variant, dataCode As Variant
        groupCode = gpCode
        dataCode = dataValue

        'SELECTIE WINDOW (CROSSING !) STEEDS GROTER MAKEN, TOT DE AANVOER RETOUR POLYLINE GEVONDEN IS.

''        Dim PuntObj1 As Object
''        Dim PuntObj2 As Object

        'zoomen anders geen juiste selectieset !!
        ZoomExtents
        
        Dim JuisteRetourPolyGevonden As Boolean
                

Opnieuw1:

        Tel = Tel + 1

        corner1(0) = Px - Tel: corner1(1) = Py - Tel: corner1(2) = 0
        corner2(0) = Px + Tel: corner2(1) = Py + Tel: corner2(2) = 0
        ssetObj.Select mode, corner1, corner2, groupCode, dataCode

''        Set PuntObj1 = ThisDrawing.ModelSpace.AddPoint(corner1)
''        PuntObj1.Color = tel
''        Set PuntObj2 = ThisDrawing.ModelSpace.AddPoint(corner2)
''        PuntObj2.Color = tel

        'MsgBox "teller = " & tel & "     " & "count = " & ssetObj.Count


        ' LEZEN SELECTIESET:

        If Tel > 20 Then
            ssetObj.Clear
            ssetObj.Delete
            ZoomPrevious
            MsgBox "Geen retourleiding (polyline) gevonden in laag 'Legplan' (of HOH te groot).", vbExclamation, "Let op"
            End
        End If

        If ssetObj.Count = 0 Or ssetObj.Count = 1 Then GoTo Opnieuw1

        If ssetObj.Count > 1 Then
            For Each ssElement In ssetObj
                If RetourObj.handle <> ssElement.handle Then
                        'MsgBox "gevonden retourPolyline-handle = " & ssElement.handle
                        'sHandle = element.handle
                        
''                        If ssElement.Layer <> "Legplan" Then
''                            GoTo Opnieuw1
''                        Else
''                            ssElement.Layer = sNwLaagnaam
''                            Exit For
''                        End If


                        If ssElement.Layer = "Legplan" Then
                            ssElement.Layer = sNwLaagnaam
                            ssElement.Color = acByLayer
                            JuisteRetourPolyGevonden = True
                            Exit For
                        End If
                        

                End If
            Next ssElement
            
            If JuisteRetourPolyGevonden = False Then GoTo Opnieuw1
        End If


        ssetObj.Clear
        ssetObj.Delete

        ZoomPrevious

                
        '-----------------------------------------------------------------------------------
        'BEPALEN OF DAT HET EINDPUNT OF HET BEGINPUNT VAN DE EERSTE
        'LIJN VERBONDEN IS MET HET EINDPUNT VAN DE AANVOER-POLYLINE
        'want beide richtingen moeten mogelijk zijn.
        '-----------------------------------------------------------------------------------

        For Each varVert In coord
            intVCnt = intVCnt + 1
        Next

        'OPVRAGEN EINDPUNT LWT-POLYLINE
        'MsgBox "Laatste x =" & coord(intVCnt - 2)
        'MsgBox "Laatste y =" & coord(intVCnt - 1)

        Dim PeindPoly(0 To 2) As Double
        PeindPoly(0) = coord(intVCnt - 2)
        PeindPoly(1) = coord(intVCnt - 1)
        PeindPoly(2) = 0

        Dim Afstand1 As Double
        Dim Afstand2 As Double
        Afstand1 = M_Afstand.Lengte(PeindPoly, tempObj.EndPoint)
        Afstand2 = M_Afstand.Lengte(PeindPoly, tempObj.StartPoint)

        SchrijfLogFile ("   Lengte eerste lijn verbonden aan aanvoer-polyline = " & tempObj.Length)
        Lengte = Lengte + tempObj.Length

    End If

    ' ------------------------------------------------------------------------------------
    ' BEREKENEN WELKE LIJNEN MET ELKAAR VERBONDEN ZIJN
    ' -------------------------------------------------------------------------------------

    Dim ReturnEindpunt As Variant
    Dim Beginpunt As Variant
    Dim Eindpunt As Variant
    Dim LijnGevonden As Boolean
    Dim StartpuntGevonden As Boolean
    Dim testobj As Object


    If Afstand1 < Afstand2 Then
            ReturnEindpunt = tempObj.StartPoint
            'MsgBox "Eerste lijn met eindpunt (endpoint) verbonden aan de aanvoer-polyline-eindpunt"
    Else
            ReturnEindpunt = tempObj.EndPoint
            'MsgBox "Eerste lijn met beginpunt (startpoint) verbonden aan de aanvoer-polyline-eindpunt"
    End If



    '----------------------------------------------------------------------------------------------------
    ' *** START BEREKENING LENGTE-LOOP ***
    '----------------------------------------------------------------------------------------------------
  
    'ReturnEindpunt = tempObj.EndPoint

    Set returnObj = tempObj
    returnObj.Layer = sNwLaagnaam
    returnObj.Color = acByLayer
    
    Dim bOptellen As Boolean
    Dim EindpuntLijn As Variant

opnieuw:

    LijnGevonden = False
    
   
    For Each element In ThisDrawing.ModelSpace
    If element.EntityName = "AcDbLine" Or element.EntityName = "AcDbArc" Then

        'MsgBox element.Handle
        If returnObj.handle <> element.handle Then

                Beginpunt = element.StartPoint
                Eindpunt = element.EndPoint
                'MsgBox "Zoek naar: " & ReturnEindpunt(0) & "    " & ReturnEindpunt(1)
    
                bOptellen = False

                If AFR(ReturnEindpunt(0)) = AFR(Beginpunt(0)) Then
                    If AFR(ReturnEindpunt(1)) = AFR(Beginpunt(1)) Then
                        StartpuntGevonden = True
                        bOptellen = True
                    End If
                End If
                
                If AFR(ReturnEindpunt(0)) = AFR(Eindpunt(0)) Then
                    If AFR(ReturnEindpunt(1)) = AFR(Eindpunt(1)) Then
                        StartpuntGevonden = False
                        bOptellen = True
                    End If
                End If
                
                If bOptellen = True Then
                    element.Layer = sNwLaagnaam
                    element.Color = acByLayer
                    element.Update
                    LijnGevonden = True
                    'Set ReturnObj = element
                    Set testobj = element

                    If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                    If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
                    
'''''                    If element.EntityName = "AcDbLine" Then
'''''                        'omkeren beginpunt en eindpunt van de lijn tbv aansluiten op retour.
'''''                        If StartpuntGevonden = True Then
'''''                            EindpuntLijn = element.EndPoint
'''''                            element.EndPoint = element.StartPoint
'''''                            element.StartPoint = EindpuntLijn
'''''                            StartpuntGevonden = False
'''''                            'btempStartpuntGevonden = True
'''''                        End If
'''''                    End If
                            

                    ' BEREKENEN SNIJPUNT MET RETOUR POLYLINE
                    'ORGINEEL VOOR 28 NOV 2002
                    'If F_Main.CheckBox2.Value Then Call BepalenSnijpunt1(element, RetourObj, Lengte, sNwLaagnaam)
                    
                    If F_Main.CheckBox2.Value Then Call BepalenSnijpunt1(element, ssElement, Lengte, sNwLaagnaam)
   
                End If
        End If
        
        

        'MsgBox "DE LENGTE IS : " & Lengte, vbInformation
        
'----------------------------------------------------------------------------------------------------
' ***  MELDING ALS ROLLENGTE BEHAALD IS + BEPALEN JUISTE PUNT WAARBIJ DE LENGTE OVEREENKOMT MET DE INGEVOERDE (GEWENSTE) LENGTE ***
'----------------------------------------------------------------------------------------------------

'        If F_Main.CheckBox1.Value = True Then

        If Stoppen = True Then

                'laatste gehele lijn van het legpatroon
                element.Color = acMagenta

                'Totalelengte = lengtevanpoly + Lengte
                SchrijfLogFile ("[6]   stoppen = True")
                SchrijfLogFile ("")
                SchrijfLogFile ("  ! Invoerlengte=" & InvoerLengte)
                SchrijfLogFile ("  Lengte (tot laats berekende lijn)=" & Lengte)      'Lengte van aanvoer polyline + lijns + arcs
                SchrijfLogFile ("  LengtePolylineTotSnijpunt = " & lengtevanpoly)     'public var !
                SchrijfLogFile ("  Lengte van laatste lijnstuk=" & element.Length)
                SchrijfLogFile ("")
                
                SchrijfLogFile ("  SnijpuntX=") & vSnijpunt(0)                         'public, uit snijpuntberekening anderstaande subroutine
                SchrijfLogFile ("  SnijpuntY=") & vSnijpunt(1)
                SchrijfLogFile ("  StartpuntGevonden= " & StartpuntGevonden)
                
                '-----------------------------------------------------------------
                'BEPALEN TEKENRICHTING
                '-----------------------------------------------------------------
                
                Dim EindpuntNwLijn(0 To 2) As Double
                Dim LijnObj As Object
                Dim SituatieA As Boolean
                
                Dim SnijLengte As Double    'lengte tussen beginpunt of eindpunt van de laatste lijn (element) tot het snijpunt
                                            'tot de polyline.
                                            
                If StartpuntGevonden = True Then
                    SnijLengte = M_Afstand.Lengte(element.StartPoint, vSnijpunt)
                Else
                    SnijLengte = M_Afstand.Lengte(element.EndPoint, vSnijpunt)
                End If
                
                'MsgBox "begin of eindpunt van de laatste lijn tot snijpunt aanvoerpoly = " & SnijLengte
                 SchrijfLogFile ("begin of eindpunt van de laatste lijn tot snijpunt aanvoerpoly = " & SnijLengte)
                 
                'als retourpolyline iets verschoven is dan snijlengte <> 0, anders precies 0 !
                
                '*** uitgezet op 5 dec 2002
                '*** If SnijLengte > 0 And SnijLengte < 10 Then MsgBox "Retourpolyline staat niet precies op 1x HOH naast de aanvoerpolyline."
                
                If SnijLengte < 5 Then
                    SituatieA = True
                Else
                    SituatieA = False
                End If
                
                'MsgBox "SituatieA=" & SituatieA
                SchrijfLogFile ("SituatieA=" & SituatieA)
                
                 
                '-----------------------------------------------------------------
                'AFHANKELIJK VAN DE TEKENRICHTING LAATSTE LIJN(EN) PLAATSEN
                '-----------------------------------------------------------------
                
                Dim vEindpunt As Variant
                
                'If StartpuntGevonden = True Or StartpuntGevonden = False Then
                If SituatieA = True Then
                    'SITUATIE 2
                    SchrijfLogFile ("*** SITUATIE 1 ***")

                    Dim LengteVerschil As Double
                    LengteVerschil = (Lengte + lengtevanpoly) - InvoerLengte
                    SchrijfLogFile ("Lengteverschil van de laatste lijn moet zijn:" & LengteVerschil)

                    Dim LengteVanLaatsteLijnstuk As Double
                    LengteVanLaatsteLijnstuk = element.Length - LengteVerschil
                    SchrijfLogFile ("** Lengte laatste lijnstuk moet zijn (zonder berek arcs ed): " & LengteVanLaatsteLijnstuk)
                    LengteVanLaatsteLijnstuk = LengteVanLaatsteLijnstuk - 15.7
                    LengteVanLaatsteLijnstuk = LengteVanLaatsteLijnstuk / 2
                    SchrijfLogFile ("** Lengte laatste lijnstuk juist berekend: " & LengteVanLaatsteLijnstuk)

                    If LengteVanLaatsteLijnstuk < 0 Then MsgBox "Negatieve lengte, verleng de laatste lijn tot de retourleiding !", vbInformation: End


                    '----------------------------------------------------------------------
                    '*** NIEUW
                    '----------------------------------------------------------------------
                    If StartpuntGevonden = False Then
                        'OMKEREN LAATSTE LIJN (RICHTING VERKEERD)
                        Beginpunt(0) = Eindpunt(0)
                        Beginpunt(1) = Eindpunt(1)
                        'Dim vEindpunt As Variant
                        vEindpunt = element.EndPoint
                        element.EndPoint = element.StartPoint
                        element.StartPoint = vEindpunt
                    End If


                    'PLAATSEN BEREKENDE LIJNSTUK
                    EindpuntNwLijn(0) = Beginpunt(0) + LengteVanLaatsteLijnstuk
                    EindpuntNwLijn(1) = Beginpunt(1)
                    EindpuntNwLijn(2) = 0

                    Set LijnObj = ThisDrawing.ModelSpace.AddLine(Beginpunt, EindpuntNwLijn)
                    
                    
                    LijnObj.Rotate element.StartPoint, element.Angle
                    LijnObj.Layer = sNwLaagnaam
                    LijnObj.Color = acByLayer
                    element.StartPoint = LijnObj.EndPoint
                    'Call M_InsertBlock.InsertBlock(LijnObj.EndPoint)
                   
                    
                    'OVERIGE LIJNEN PLAATSEN:
                    Dim PuntObj As AcadPoint
                    Set PuntObj = ThisDrawing.ModelSpace.AddPoint(LijnObj.EndPoint)
                    PuntObj.Color = acRed
                    
                    Set PuntObj = ThisDrawing.ModelSpace.AddPoint(LijnObj.StartPoint)
                    PuntObj.Color = acCyan
                    
                    'PLAATSEN VAN AFTAKLIJNEN:
                    If F_Main.CheckBox4 Then Call M_TekenenAftaklijnen.TekenenAftaklijnen(F_Main.ComboBox2, LijnObj.EndPoint, element.Angle)
                                       
                    
                    
                    

                Else
                    MsgBox "IN SITUATIE II NOG GEEN AFTAKLIJNEN", vbExclamation
                    SchrijfLogFile ("SituatieA=" & SituatieA)
                            
                    Dim Opening As Double
                    Opening = InvoerLengte - Lengte - lengtevanpoly
                            SchrijfLogFile ("Opening = Invoerlengte - Lengte - lengtevanpoly = " & Opening)
                            SchrijfLogFile ("Opening / 2 = " & Opening / 2)
                    LengteVanLaatsteLijnstuk = element.Length - Abs(Opening / 2)
                            SchrijfLogFile ("LengteVanLaatsteLijnstuk = " & LengteVanLaatsteLijnstuk)
                            
                    'MsgBox "lengte van laatste lijn moet worden: " & LengteVanLaatsteLijnstuk
                    SchrijfLogFile ("")
                    SchrijfLogFile ("   >>> lengte van laatste lijn moet worden: " & LengteVanLaatsteLijnstuk)
                    SchrijfLogFile ("")
                    
                    
                    'NIEUW AANGEVULD
                    'Eindpunt(0) = Beginpunt(0)
                    'Eindpunt(1) = Beginpunt(1)
                    If StartpuntGevonden = True Then
                        'OMKEREN LAATSTE LIJN (RICHTING VERKEERD)
                        Beginpunt(0) = Eindpunt(0)
                        Beginpunt(1) = Eindpunt(1)
                        'Dim vEindpunt As Variant
                        vEindpunt = element.EndPoint
                        element.EndPoint = element.StartPoint
                        element.StartPoint = vEindpunt
                    End If
                    

                    EindpuntNwLijn(0) = Eindpunt(0) - LengteVanLaatsteLijnstuk
                    EindpuntNwLijn(1) = Eindpunt(1)
                    EindpuntNwLijn(2) = 0

                    Set LijnObj = ThisDrawing.ModelSpace.AddLine(Eindpunt, EindpuntNwLijn)
                    LijnObj.Layer = sNwLaagnaam
                    LijnObj.Color = acByLayer
                    LijnObj.Rotate element.StartPoint, element.Angle

                    'OPSCHUIVEN VAN LAATSTE LIJN
                    LijnObj.Move LijnObj.EndPoint, element.StartPoint       'let op bij bovenstaande startpoint
                    'Call M_InsertBlock.InsertBlock(LijnObj.EndPoint)
                    


                    'AANPASSEN LENGTE VAN LAATSTE (OORSPRONKELIJKE) LIJN
                    element.StartPoint = LijnObj.StartPoint
                    element.Layer = "Legplan"
                    element.Color = acByLayer

                    'Call M_InsertBlock.InsertBlock(EindpuntNwLijn)



                End If

                   

                End
                'F_Main.Show
                'Exit Sub

        End If


    End If
    Next element



'----------------------------------------------------------------------------------------------
' *** EINDE BEREKENING ***
'----------------------------------------------------------------------------------------------


    If LijnGevonden = True Then
        'MsgBox "Wacht"
        Set returnObj = testobj
        If StartpuntGevonden = True Then ReturnEindpunt = returnObj.EndPoint
        If StartpuntGevonden = False Then ReturnEindpunt = returnObj.StartPoint
        LijnGevonden = False
        GoTo opnieuw
    Else

        'INDIEN DE LENGTE VAN HET GESELECTEERDE ELEMENT GELIJK IS AAN DE TOTALE LENGTE, DAN OPNIEUW BEREKENEN
        'MAAR NU VANUIT HET STARTPUNT IPV EINDPUNT VAN DE GESELECTEERDE LINE
        If Lengte = returnObj.Length Then
            MsgBox "Geen andere leidingen gevonden die aan de geselecteerde leiding verbonden zijn.", vbInformation
            ReturnEindpunt = returnObj.StartPoint
            GoTo opnieuw
        End If

        'BEREKENING IS CORRECT UITGEVOERD
        MsgBox "De totale leidinglengte is:  " & Format(Lengte / 100, "0.0") & " m" _
            & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "(korter dan " & InvoerLengte / 100 & "m)", vbInformation, "Totale lengte"
        ThisDrawing.SetVariable "modemacro", "Leidinglengte = " & Format(Lengte / 100, "0.0" & m)
        ThisDrawing.Utility.Prompt "Leidinglengte = " & Format(Lengte, "0.0")
        F_Main.show
    End If


End Sub

Function AFR(getal)
    'Afronden getal
    AFR = Format(getal, "0.0")
End Function
Sub BepalenSnijpunt1(element As Object, RetourObj As Object, Lengte, sNwLaagnaam)

' -----------------------------------------------------------------------
'   BEREKENEN INTERSECTIE (SNIJPUNT) MET RETOUR-LEIDING
' -----------------------------------------------------------------------
'   RetourObj   DIT IS DE GESELECTEERDE POLYLINE (AANVOER)
'   element     DIT IS DE BETREFFENDE LIJN (LEIDING) VAN HET LEGPATROON
'   lengte      DIT IS DE BEREKENDE LENGTE VAN HET LEGPATROON (LINES EN ARCS)
'               TOT HET BETREFFENDE ELEMENT
'   DEZE ROUTINE BEPAALD HET SNIJPUNT TUSSEN DEZE TWEE ELEMENENTEN
' -----------------------------------------------------------------------


    If RetourObj.handle = element.handle Then Exit Sub
    element.Layer = sNwLaagnaam

    Dim intPoints As Variant
    intPoints = element.IntersectWith(RetourObj, acExtendThisEntity)

    Dim I As Integer, j As Integer, k As Integer
    Dim str As String
    If VarType(intPoints) <> vbEmpty Then

        For I = LBound(intPoints) To UBound(intPoints)
            'str = "Intersection Point[" & k & "] is: " & intPoints(j) & "," & intPoints(j + 1) & "," & intPoints(j + 2)
            'MsgBox str: str = ""
            I = I + 2
            j = j + 3
            k = k + 1   'geeft aantal snijpunten aan
        Next
    End If

' -----------------------------------------------------------------------
'   CONTROLE OP HET AANTAL SNIJPUNTEN
' -----------------------------------------------------------------------

    If k = 0 Then
        'MsgBox "Leiding heeft geen snijpunt met de aanvoer/ retourleiding", vbCritical
        Exit Sub
    End If
    
    If k = 1 Then
        vSnijpunt = intPoints
    End If
        

'    If k > 1 Then
'        MsgBox "Leiding heeft meerdere snijpunten met de aanvoer/ retourleiding", vbCritical
'        Exit Sub
'    End If

' --------------------------------------------------------------------------------------------
'   BEREKENEN LENGTE VAN DE POYLINE TOT HET HIERBOVEN BEREKENENDE SNIJPUNT OP DEZE POLYLINE
' --------------------------------------------------------------------------------------------

    'Dim c As Double
    Dim LengtePolylineTotSnijpunt As Double
    
    'c = LengtePolyline(RetourObj, intPoints, False)

    LengtePolylineTotSnijpunt = LengtePolyline(RetourObj, intPoints, True)
    'MsgBox "LENGTE MET SNIJPUNT:" & LengtePolylineTotSnijpunt
    'LENGTE POLYLINE TOT HET BEREKENDE SNIJPUNT + LENGTE VAN LEGPATROON


    If F_Main.CheckBox1.Value = True Then
        If InvoerLengte < (LengtePolylineTotSnijpunt + Lengte) Then
            'POINT AFBEELDEN OP HET SNIJPUNT
            'Dim PointObj As Object
            'Set PointObj = ThisDrawing.ModelSpace.AddPoint(intPoints)
            'PointObj.Update

            'Dim TekstObj As Object
            'Set TekstObj = ThisDrawing.ModelSpace.AddText("LengtePolylineTotSnijpunt=" & LengtePolylineTotSnijpunt & "    Lengte=" & Lengte, intPoints, 6)
            'MsgBox "Stoppen"
            Stoppen = True
            lengtevanpoly = LengtePolylineTotSnijpunt       'public voor de hoofdroutine!
            'End


        End If
    End If

End Sub













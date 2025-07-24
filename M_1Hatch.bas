Attribute VB_Name = "M_1Hatch"
Public Sub PlaatsenHatch()


' -----------------------------------------------------------------------------------
' CONTROLE OPGEGEVEN HOH
' -----------------------------------------------------------------------------------

    Dim iHOH As Integer
    On Error Resume Next
    iHOH = F_Main.ComboBox2.Text
    
    If Err Then
        Err.Clear
        MsgBox "Geen juiste HOH-afstand opgegeven.", vbCritical, "Let op"
        End
    End If
    
'    Dim antw As String
'    If iHOH > 35 Then
'        antw = MsgBox("HOH-afstand is groter dan 35 cm. Toch doorgaan ?", vbYesNo + vbQuestion, "leidingleg programma")
'        If antw = vbNo Then End
'    End If

' -----------------------------------------------------------------------------------
' HIDE FORM
' -----------------------------------------------------------------------------------
    
    F_Main.Hide
    
' -----------------------------------------------------------------------------------
' JUISTE LAAG AKTIEF MAKEN
' -----------------------------------------------------------------------------------
    ' voor 27 mei 2003 was kleur 5 (blue)
    AanmakenLaag "Legplan", 2, True
    
    
' -----------------------------------------------------------------------------------
' LAAG BOUWKUNDIG UITZETTEN nieuw 7 april 2003
' -----------------------------------------------------------------------------------

    Dim LaagBouwkundig As AcadLayer
    Dim Laag As AcadLayer
    Dim LaagBouwkundigBestaat As Boolean
    For Each Laag In ThisDrawing.Layers
        If UCase(Laag.Name) = "BOUWKUNDIG" Then Set LaagBouwkundig = Laag: LaagBouwkundigBestaat = True
    Next Laag
    
    Dim sAntw As String
    If LaagBouwkundigBestaat = False Then
        sAntw = MsgBox("Er is geen laag 'Bouwkundig' aanwezig." _
        & Chr(10) & Chr(13) & "Selecteer de bouwkundige laag.", vbYesNo + vbExclamation, "Geen laag 'bouwkundig'")
        If sAntw = vbYes Then Call HernoemLaag
        
        ' INDIEN LAAG IS AANGEMAAKT DAN BEVRIEZEN
        For Each Laag In ThisDrawing.Layers
            If UCase(Laag.Name) = "BOUWKUNDIG" Then
                Set LaagBouwkundig = Laag
                LaagBouwkundig.Freeze = True
            End If
        Next Laag
        
    Else
        'If LaagBouwkundig.Lock = True Then LaagBouwkundig.Lock = False
        LaagBouwkundig.Freeze = True
    End If
    

' ___________________________________________________________________________________
' Tellen hoeveel buitengrens-lijnen en hoeveel obstakels er zijn
' ___________________________________________________________________________________

    Dim OuterlObj As Object   ' hier AcadEntity niet gebruiken
    Dim RetObj As Object
    Dim iFoutSelectieTeller1 As Integer
    
    
    
    ' 27 MEI 2003 SELECTIE AANGEPAST
Opnieuw1:
    
    ThisDrawing.Utility.GetEntity RetObj, p, "Selecteer een ruimte-polyline"
    
    If iFoutSelectieTeller1 > 0 Then End
    
    If Err <> 0 Then
        If iFoutSelectieTeller1 < 1 Then MsgBox "Verkeerde selectie", vbCritical
        iFoutSelectieTeller1 = iFoutSelectieTeller1 + 1
        Err.Clear
        GoTo Opnieuw1
    Else
        RetObj.Highlight (True)
    End If
            
    If RetObj.EntityName = "AcDbPolyline" Then
        Set OuterlObj = RetObj
    Else
        If iFoutSelectieTeller1 < 1 Then MsgBox "Geen polyline geselecteerd", vbCritical
        iFoutSelectieTeller1 = iFoutSelectieTeller1 + 1
        GoTo Opnieuw1
        'End
    End If

    ''''plineObj.Closed = True
    
' ___________________________________________________________________________________
' Geselecteerde Omtrek-polyline in laag "LegplanOmtrek"
' ___________________________________________________________________________________
    
    Call AanmakenLaag("Legplanomtrek", 6, False)
    RetObj.Layer = "Legplanomtrek"
            

' -----------------------------------------------------------------------------------
' Aangeven hoek van hatch
' -----------------------------------------------------------------------------------


    F_Main.Hide
    
    'P1: DIT IS OOK HET BEGINPUNT VOOR HET UCS.
    Dim p1 As Variant
    On Error Resume Next
    p1 = ThisDrawing.Utility.GetPoint(, "Geef de zijde aan van de eerste leiding")
    If Err Then
        Err.Clear
        End
    End If
    
    Dim p2 As Variant
    On Error Resume Next
    p2 = ThisDrawing.Utility.GetPoint(p1, "Geef de leidingrichting aan.")
    If Err Then
        Err.Clear
        End
    End If
    
    
    
    Dim retAngle As Double
    Dim retAngleOrgineel As Double
    
    retAngle = ThisDrawing.Utility.AngleFromXAxis(p1, p2)
    retAngleOrgineel = retAngle
    retAngle = retAngle - (3.14159265358979 / 4)
    
    'testdoeleinden: Dit is een hoek van 0 graden (voor de 45 graden hatch):
    'retAngle = 0 - (3.14159265358979 / 4)
    
    
' -----------------------------------------------------------------------------------
' ucs verplaatsen (iets lager of hoger)
' -----------------------------------------------------------------------------------
    
    
    'BEREKENEN POLAIRPOINT
    

    Dim Ppolair As Variant
    ' Ppolair = ThisDrawing.Utility.PolarPoint(P1, 1.5707963, 6.35)     'voor hoeken van 0 graden
    Ppolair = ThisDrawing.Utility.PolarPoint(p1, retAngleOrgineel + 1.5707963, F_Main.ComboBox2 / 2)     'halve hoh
    'Dim PUNTOBJ As AcadPoint
    'Set PUNTOBJ = ThisDrawing.ModelSpace.AddPoint(Ppolair)
   
   
   
   
    'VERPLAATSEN UCS
   
    Dim viewportObj As AcadViewport
    Set viewportObj = ThisDrawing.ActiveViewport
    
    ' Create a new UCS with origin 200, 200, 0
    Dim UcsObj As AcadUCS
    Dim origin(0 To 2) As Double            'positie
    Dim xAxisPoint(0 To 2) As Double        'x-richting
    Dim yAxisPoint(0 To 2) As Double        'y-richting
    
    origin(0) = Ppolair(0): origin(1) = Ppolair(1): origin(2) = 0
    xAxisPoint(0) = Ppolair(0) + 100: xAxisPoint(1) = Ppolair(1): xAxisPoint(2) = 0
    yAxisPoint(0) = Ppolair(0): yAxisPoint(1) = Ppolair(1) + 100: yAxisPoint(2) = 0
    
    Set UcsObj = ThisDrawing.UserCoordinateSystems.Add(origin, xAxisPoint, yAxisPoint, "UCS1")
    ThisDrawing.ActiveUCS = UcsObj
    viewportObj.UCSIconOn = True
    viewportObj.UCSIconAtOrigin = True
    ThisDrawing.ActiveViewport = viewportObj
    
    'Sysvar: UCSORG geeft aan waar het UCS zich bevindt (read only)
    'MsgBox "The origin of the UCS is: " & UcsObj.origin(0) & ", " & UcsObj.origin(1) & ", " & UcsObj.origin(2), , "Origin Example"
    
   
'----------------------------------------------------------------------------------
'INZOOMEN IN RUIMTEPOLYLINE
'----------------------------------------------------------------------------------
    
    Dim minExt As Variant
    Dim maxExt As Variant

    RetObj.GetBoundingBox minExt, maxExt
    ZoomWindow minExt, maxExt
    
    
' -----------------------------------------------------------------------------------
' Hatch-object aanmaken, maar nog niet vullen (incl schaal en hoek)
' -----------------------------------------------------------------------------------


    Dim HoHafstand As Double
    Dim HatchSchaal As Double
    
    HoHafstand = F_Main.ComboBox2.Text                ' INVOER H-op-H afstand
    'HOHafstand = 10
    HatchSchaal = HoHafstand * 0.3149606299
            
    
    Dim HatchObj As Object  'niet gebruiken: AcadHatch
    Dim patternName As String
    Dim PatternType As Long
    Dim bAssociativity As Boolean
        
    ' Define the hatch
    patternName = "ANSI31"
    PatternType = 0

    bAssociativity = True
    Set HatchObj = ThisDrawing.ModelSpace.AddHatch(PatternType, patternName, bAssociativity)
    
    
    HatchObj.PatternScale = HatchSchaal            ' declareren as double
    
    HatchObj.HatchStyle = acHatchStyleOuter
    

    
    
'''    Dim p1 As Variant
'''    Dim ArceerHoek As Double
'''
'''
'''    ''HoekInGraden = ComboBox3.Text
'''    HoekInGraden = 0
'''
'''    ' Hoek standaard al -45 graden verdraaien om deze horizontaal te zetten
'''    HoekInRad = ThisDrawing.Utility.AngleToReal(HoekInGraden - 45, acDegrees)
'''    HatchObj.PatternAngle = HoekInRad

    HatchObj.PatternAngle = retAngle


    
   
    
   
' ___________________________________________________________________________________
' OFFSETTEN POLYLINE
' ___________________________________________________________________________________
'    Dim offsetObj As Variant
'    offsetObj = retObj.Offset(HoHafstand)
'
'    offsetObj(0).Color = 53
'
'    'CONTROLEREN OF OFFSET RICHTING JUIST IS MIDDELS BOUNDINGBOX EN BEREKENEN OPPERVL.
'    'IS DIT NIET HET GEVAL DAN OFFSET-OBJ VERWIJDEREN EN OPNIEUW OFFSETTEN
'    'IN DE TEGENGESTELDE RICHTING.
'
'    Dim minExt As Variant
'    Dim maxExt As Variant
'    Dim Lengte As Double
'    Dim hoogte As Double
'    Dim Oppervlakte1 As Double
'    Dim Oppervlakte2 As Double
'
'    retObj.GetBoundingBox minExt, maxExt
'    Lengte = maxExt(0) - minExt(0)
'    hoogte = maxExt(1) - minExt(1)
'    Oppervlakte1 = Lengte * hoogte
'
'    offsetObj(0).GetBoundingBox minExt, maxExt
'    Lengte = maxExt(0) - minExt(0)
'    hoogte = maxExt(1) - minExt(1)
'    Oppervlakte2 = Lengte * hoogte
'
'    'INDIEN DE OFFSET RICHTING VERKEERD IS (NAAR BUITEN I.P.V. NAAR BINNEN).
'    If Oppervlakte2 > Oppervlakte1 Then
'        offsetObj(0).Erase
'        offsetObj = retObj.Offset(-HoHafstand / 2)
'        offsetObj(0).Color = 53
'    End If
'
'    Set OuterlObj = offsetObj(0)

'---------------------------------------------
'nieuw, vervangen door:

    Set OuterlObj = RetObj
    
' ___________________________________________________________________________________
' Outerloop aanmaken met een (1) gefilterde buitengrens-lijn
' ___________________________________________________________________________________

    'Dim outerloop(0 To 0) As Object
    Dim outerloop(0 To 0) As Object ' geen AcadEntity gebruiken
    Set outerloop(0) = OuterlObj
        
    On Error Resume Next
    HatchObj.AppendOuterLoop (outerloop)
    If Err Then
        Err.Clear
        MsgBox "De hulpoffset-lijn (magenta outerloop) is niet aangegeven. ", vbCritical, "(errcode 3)"
        'OuterlObj.Erase        27 mei 2004 uitgezet !
        OuterlObj.Highlight (True)
        
        'NIEUW 27 MEI 2004: UCS WEER TERUGPLAATSEN !!!
         origin(0) = 0: origin(1) = 0: origin(2) = 0
        xAxisPoint(0) = 100:   xAxisPoint(1) = 0:       xAxisPoint(2) = 0
        yAxisPoint(0) = 0:     yAxisPoint(1) = 100:     yAxisPoint(2) = 0

        Set UcsObj = ThisDrawing.UserCoordinateSystems.Add(origin, xAxisPoint, yAxisPoint, "UCS1")

        ThisDrawing.ActiveUCS = UcsObj
        viewportObj.UCSIconOn = True
        viewportObj.UCSIconAtOrigin = True
        ThisDrawing.ActiveViewport = viewportObj

        Set UcsObj = Nothing
        
        'Exit Sub
        End
    End If
    
' ___________________________________________________________________________________
' Innerloop aanmaken met een (5) obstakel
' ___________________________________________________________________________________

'Dim ElementTeller As Integer
'Dim ObstakelTeller As Integer
'
'For ElementTeller = 0 To ThisDrawing.ModelSpace.Count - 1
'    If ThisDrawing.ModelSpace.Item(ElementTeller).Color = acCyan Then ' **acYellow
'        ObstakelTeller = ObstakelTeller + 1
'
'        Select Case ObstakelTeller
'        Case 1
'            Dim innerLoop1(0) As Object                     ' of (0 to 0) of (0) 0f ()
'            Set innerLoop1(0) = ThisDrawing.ModelSpace.Item(ElementTeller)
'            HatchObj.AppendInnerLoop (innerLoop1)
'        Case 2
'            Dim innerLoop2(0) As Object
'            Set innerLoop2(0) = ThisDrawing.ModelSpace.Item(ElementTeller)
'            HatchObj.AppendInnerLoop (innerLoop2)
'        Case 3
'            Dim innerLoop3(0) As Object
'            Set innerLoop3(0) = ThisDrawing.ModelSpace.Item(ElementTeller)
'            HatchObj.AppendInnerLoop (innerLoop3)
'        Case 4
'            Dim innerLoop4(0) As Object
'            Set innerLoop4(0) = ThisDrawing.ModelSpace.Item(ElementTeller)
'            HatchObj.AppendInnerLoop (innerLoop4)
'        Case 5
'            Dim innerLoop5(0) As Object
'            Set innerLoop5(0) = ThisDrawing.ModelSpace.Item(ElementTeller)
'            HatchObj.AppendInnerLoop (innerLoop5)
'        End Select
'    End If
'Next ElementTeller




''??If ObstakelTeller = 0 Then MsgBox "Er zijn geen obstakels gevonden", vbInformation, "Aantal obstakels"
''??If ObstakelTeller = 1 Then MsgBox "Er is 1 obstakel gevonden", vbInformation, "Aantal obstakels"
''??If ObstakelTeller > 1 And ObstakelTeller < 6 Then MsgBox "Er zijn " & ObstakelTeller & " obstakels gevonden", vbInformation, "Aantal obstakels"
''??If ObstakelTeller > 5 Then MsgBox "Let op er zijn " & ObstakelTeller & " obstakels gevonden", vbCritical, "Meer dan 5 obstakels"

' --------------------------------------------------------------------------------
' NIEUW: INNERLOOP AANGEVEN DOOR DE GEBRUIKER (5 OBSTAKELS SELECTEREN)
' --------------------------------------------------------------------------------


    Dim returnString As String
    returnString = ThisDrawing.Utility.GetString(False, "Zijn er obstakels ?: Ja <Nee>: ")
    'MsgBox returnString, vbExclamation
    If UCase(Left$(returnString, 1)) = "N" Or returnString = "" Then GoTo PlaatsenHatch
    
    
    'UILEZEN SELECTIESET:
    

    Dim ssetObj As AcadSelectionSet
    Dim bStoppen As Boolean
    
opnieuw:
    bStoppen = False
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET1")
    
    ssetObj.SelectOnScreen
    
    If Err Then
        'enter is doorgaan
        Err.Clear
    End If
    
    Dim element As Object
    Dim ObstakelTeller As Integer
    
    
    'CONTROLE OF ER POLYLINES GESELECTEERD ZIJN:
    For Each element In ssetObj
        If element.EntityName <> "AcDbPolyline" Then
            element.Highlight = True
            bStoppen = True
        End If
    Next element
    
    If bStoppen = True Then
        MsgBox "Geselecteerde element is geen polyline.", vbCritical, "Let op"
        ssetObj.Clear
        ssetObj.Delete
        GoTo opnieuw
    End If
        
    
    
    'UITLEZEN SLECTIESET EN VAN DE POLYLINES HATCH-ISLANDS MAKEN:
    For Each element In ssetObj
    
        If element.EntityName = "AcDbPolyline" Then
        
            element.Layer = "Legplanomtrek"
    
            ObstakelTeller = ObstakelTeller + 1
        
            Select Case ObstakelTeller
            Case 1
                Dim innerLoop1(0) As Object                     ' of (0 to 0) of (0) 0f ()
                Set innerLoop1(0) = element
                HatchObj.AppendInnerLoop (innerLoop1)
            Case 2
                Dim innerLoop2(0) As Object
                Set innerLoop2(0) = element
                HatchObj.AppendInnerLoop (innerLoop2)
            Case 3
                Dim innerLoop3(0) As Object
                Set innerLoop3(0) = element
                HatchObj.AppendInnerLoop (innerLoop3)
            Case 4
                Dim innerLoop4(0) As Object
                Set innerLoop4(0) = element
                HatchObj.AppendInnerLoop (innerLoop4)
            Case 5
                Dim innerLoop5(0) As Object
                Set innerLoop5(0) = element
                HatchObj.AppendInnerLoop (innerLoop5)
            Case Else
                MsgBox "Er mogen maximaal 5 obstakels worden geselecteerd", vbCritical, "Let op"
                bStoppen = True
            End Select
            
        End If
    Next element
           
            
    
    
    ssetObj.Clear
    ssetObj.Delete
    
    If bStoppen = True Then End
        
        

    
    
    


''''''' ___________________________________________________________________________________
''''''' een innerloops in een innerloop toevoegen (niet nodig !!)
''''''' ___________________________________________________________________________________
''''''
''''''    Dim loopobj(0 To 0) As Object
''''''    Set loopobj(0) = ThisDrawing.ModelSpace.Item(2)
''''''
''''''
''''''    Dim ObstTeller As Integer
''''''    For ObstTeller = 1 To 4
''''''        Set loopobj(0) = ThisDrawing.ModelSpace.Item(ObstTeller)
''''''        HatchObj.InsertLoopAt HatchObj.NumberOfLoops, acHatchLoopTypeDefault, loopobj
''''''    Next ObstTeller
''''''
''''''    'hatchObj.InsertLoopAt hatchObj.NumberOfLoops, acHatchLoopTypeDefault, loopobj

' ___________________________________________________________________________________
' Arcering afbeelden
' ___________________________________________________________________________________



PlaatsenHatch:

    'MsgBox "Berekenen voorlopig legpatroon.", vbExclamation

    'HatchObj.AssociativeHatch = False
    HatchObj.Evaluate
    HatchObj.Color = acByLayer
    HatchObj.Layer = "Legplan"
    HatchObj.Update
    'ThisDrawing.Activate
    
   
    '-----------------------------------------

'    'hulppolyline verwijderen
'    Dim HulpPolyLine
'    For Each HulpPolyLine In ThisDrawing.ModelSpace
'        If HulpPolyLine.Color = 53 Then element.HulpPolyLine
'    Next HulpPolyLine
'

''    If Err Then
''        Err.Clear
''        MsgBox "Controleer het legplatroon (object-type)", vbInformation, "Let op"
''    End If

    UcsObj.Delete
    
    'HatchObj.Explode   'DIT IS VOOR ARCERINGEN NIET MOGELIJK (R2002 EN EERDERE VERSIES) !
    On Error Resume Next
    
    'ThisDrawing.SendCommand ("erase" & vbCr & "l" & vbCr & vbCr) 'weggehaald, polyline wordt niet meer getekend.
    ThisDrawing.SendCommand ("explode" & vbCr & "l" & vbCr & vbCr)
    'ThisDrawing.SendCommand ("USC" & vbCr & "W" & vbCr)
    ThisDrawing.Activate

'    ThisDrawing.Regen True

    'MsgBox "VOOR EXPLODE"
    'ThisDrawing.SendCommand ("(COMMAND ""_-VBARUN"" ""PART2"")" & vbCr)
    'ThisDrawing.SendCommand ("_-VBARUN " & "PART2 ")
      
      
    '----------------------------------------------------------------------------------
    'BEPALEN OF ER OBSTAKELS ZIJN, DIT EVEN UITGEZET.
    '----------------------------------------------------------------------------------
    '** UITGEZET OP 10 DECEMBER 2002
    Application.RunMacro "part2"
    
    'ThisDrawing.Regen (True)
    'MsgBox "NA SENDCOMMAND AANROEP PART2"
    
    
    '----------------------------------------------------------------------------------
    'UCS WEER TERUG PLAATSEN IN WORLD
    '----------------------------------------------------------------------------------
    
    origin(0) = 0: origin(1) = 0: origin(2) = 0
    xAxisPoint(0) = 100:   xAxisPoint(1) = 0:       xAxisPoint(2) = 0
    yAxisPoint(0) = 0:     yAxisPoint(1) = 100:     yAxisPoint(2) = 0

    Set UcsObj = ThisDrawing.UserCoordinateSystems.Add(origin, xAxisPoint, yAxisPoint, "UCS1")

    ThisDrawing.ActiveUCS = UcsObj
    viewportObj.UCSIconOn = True
    viewportObj.UCSIconAtOrigin = True
    ThisDrawing.ActiveViewport = viewportObj

    Set UcsObj = Nothing



    '----------------------------------------------------------------------------------
    'OPNIEUW INZOOMEN IN RUIMTEPOLYLINE OMDAT UCS WEER TERUGGEZET IS ?
    '----------------------------------------------------------------------------------
        
'    Dim minExt As Variant
'    Dim maxExt As Variant

    RetObj.GetBoundingBox minExt, maxExt
    ZoomWindow minExt, maxExt
    
    
    
    ' -----------------------------------------------------------------------------------
    ' LAAG BOUWKUNDIG WEER AANZETTEN nieuw 7 april 2003
    ' -----------------------------------------------------------------------------------

    If LaagBouwkundigBestaat = True Then LaagBouwkundig.Freeze = False


    
    
   
    

End Sub


Sub PART2()
    Call BepaalObstakels(False)
End Sub



Sub HernoemLaag()

    ' aanvulling 27 jan 2004
    Dim ReturnObj1 As AcadObject
    Dim p1 As Variant

    ThisDrawing.Utility.GetEntity ReturnObj1, p1, "(Laag hernoemen > Selecteer een bouwkundig element:"
    
    If Err Then
        Err.Clear
        Exit Sub
    End If
    
    Dim LaagObj As AcadLayer
    Set LaagObj = ThisDrawing.Layers.Item(ReturnObj1.Layer)
    LaagObj.Name = "BOUWKUNDIG"
    LaagObj.Color = 8
    
    If Err Then
        Err.Clear
        MsgBox "Laag " & ReturnObj1.Layer & " kan niet worden hernoemd."
        Exit Sub
    End If
    


End Sub

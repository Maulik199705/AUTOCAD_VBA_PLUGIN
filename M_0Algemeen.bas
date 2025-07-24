Attribute VB_Name = "M_0Algemeen"

'--------------------------------------------------------------------------------------------
' DE VOLGENDE SUBS EN FUNCTIONS ZIJN AANWEZIG:
'--------------------------------------------------------------------------------------------
' Public Function Lengte(P1 As Variant, p2 As Variant) As Double
' Public Sub LijnlengteWijzigen(LijnObj As AcadLine, dNieuweLengte As Double, bEindpuntInkorten As Boolean)
' Public Sub InsertBlock(Blocknaam, PInsert, sLaagnaam)
' Public Sub AanmakenLaag(Laagnaam As String, KleurNr As Integer, MaakAktief As Boolean)
' Public Sub SelectiesetsVerwijderen()
' Public Function BepalenVanSnijpunt(element1 As Object, element2 As Object, ExtendOption As Integer) As Variant
' Public Sub Punt(P)
' Public Sub Punt1(P, kleur)
' Public Sub VerwijderenLogFile()
' Public Sub SchrijfLogFile(Tekst As String)
' Public Function ZoekElementHandleRondPunt(P As Variant, dCrossingSelectieGrootte As Double) As String
' Public Sub ZoomInOpElement(element As AcadEntity, bHighlight As Boolean)
' Function GetalIsDouble(sGetal As String, Optional sMelding As String) As Boolean



Public Function Lengte(p1 As Variant, p2 As Variant) As Double

    'FUNCTIE VOOR HET BEREKENEN VAN DE AFSTAND TUSSEN TWEE PUNTEN
    
    Dim DX As Double
    Dim DY As Double
    
    DX = Abs(p2(0) - p1(0))
    DY = Abs(p2(1) - p1(1))

    Lengte = Sqr((DX ^ 2) + (DY ^ 2))
End Function

Public Sub LijnlengteWijzigen(LijnObj As AcadLine, dNieuweLengte As Double, bEindpuntInkorten As Boolean)

'EBR 28 MAART 2003

'--------------------------------------------------------------------------------------------
' DE VOLGENDE SUBROUTINE WIJZIGD DE LENGTE VAN EEN LIJN T.O.V. BEGIN- OF EINDPUNT
'--------------------------------------------------------------------------------------------

    'VOORBEELD AANROEP: Call LijnlengteWijzigen(ThisDrawing.ModelSpace.Item(0), 70, True)
    
    Dim dDraaipunt As Variant
    Dim dDraaiHoek As Double
    Dim p As Variant
    
    dDraaipunt = LijnObj.StartPoint
    dDraaiHoek = LijnObj.Angle
    LijnObj.Rotate dDraaipunt, -dDraaiHoek

    If bEindpuntInkorten = True Then
        p = LijnObj.StartPoint
        p(0) = p(0) + dNieuweLengte
        LijnObj.EndPoint = p
    Else
        p = LijnObj.EndPoint
        p(0) = p(0) - dNieuweLengte
        LijnObj.StartPoint = p
    End If
    
    LijnObj.Rotate dDraaipunt, dDraaiHoek
        

End Sub


Public Sub InsertBlock(Blocknaam, PInsert, sLaagNaam)
        
'EBR 28 MAART 2003

'--------------------------------------------------------------------------------------------
' DE VOLGENDE SUBROUTINE PLAATST BLOCK IN DE TEKENING
' DE BLOCKNAAM MAG ZOWEL MET HOOFD- ALS KLEINE LETTERS INGEGEVEN WORDEN
' INDIEN HET BLOCK NIET IN DE TEKEING BESTAAT DAN ZAL DEZE UIT DE SUPPORT-DIRECTORY
' OF DE ACAD-APPLICATIONS DIRECTORY WORDEN GELEZEN.

' DEZE SUBROUTINE MAAKT GEBRUIK VAN: CALL AANMAKENLAGEN EN FINDZOEKPAD
'--------------------------------------------------------------------------------------------


        'AANROEP: InsertBlock "pijl", Pins, "0"
        'Ingevoerde blocknaam zal altijd worden geconverteerd naar hoofdletters.

        'CONTROLE OF DE BLOCKNAME (DEFENITIE) IN DE TEKENING VOORKOMT
        Dim sBlockNaam As String
        Dim bBlockInTekening As Boolean
        Dim Block As AcadBlock
        
        For Each BlockObj In ThisDrawing.Blocks
            If UCase(BlockObj.Name) = UCase(Blocknaam) Then
                bBlockInTekening = True
                sBlockNaam = BlockObj.Name
                'bBlockInTekening = True
            End If
        Next
        
        'SYMBOOL INSERTEN
        Dim BlockRefObj As AcadBlockReference
        
        If bBlockInTekening = True Then
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(PInsert, sBlockNaam, 1, 1, 1, 0)
        Else
            'WBLOCK UIT DE ACAD-ZOEKPADEN
            sBlockNaam = Blocknaam & ".dwg"
            sBlockNaam = FindZoekpad(sBlockNaam)
            'MsgBox sBlockNaam
            
            If sBlockNaam = "ERROR" Then
                MsgBox "Het symbool (block) " & Blocknaam & " staat niet in de zoekpaden van AutoCAD.", vbCritical, "Block kan niet worden geplaatst."
                Exit Sub
            Else
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(PInsert, sBlockNaam, 1, 1, 1, 0)
                'BlockObj.Layer = "0"
                'BlockObj.Update
            End If
        End If
        
        Dim sNieuwLaagnaam As String
        sNieuwLaagnaam = sLaagNaam
        Call AanmakenLaag(sNieuwLaagnaam, 7, False)
        
        BlockRefObj.Layer = sLaagNaam
        BlockRefObj.Update
        
         'oude routine:
'        Dim sAppPath As String
'        sAppPath = Application.Path
'
'        If Dir(sAppPath & "\pijl.dwg") = "" Then
'            MsgBox "Geen symbool PIJL.DWG in het zoekpad " & sAppPath & " aanwezig.", vbExclamation, "leidinglegprogramma"
'        Else
'            Dim BlockObj As AcadBlockReference
'            Set BlockObj = ThisDrawing.ModelSpace.InsertBlock(PInsert, sAppPath & "\pijl.dwg", 1, 1, 1, 0)
'            BlockObj.Layer = "0"
'            BlockObj.Update
'        End If
        

        
        
        
        
End Sub

Public Sub AanmakenLaag(LaagNaam As String, KleurNr As Integer, MaakAktief As Boolean)

    Dim LaagObj As AcadLayer
    Dim LaagGevonden As Boolean
    Dim NwLaagObj As AcadLayer
    
    For Each LaagObj In ThisDrawing.Layers
        If UCase(LaagObj.Name) = LaagNaam Then LaagGevonden = True
    Next LaagObj
    
    If LaagGevonden = False Then
        Set NwLaagObj = ThisDrawing.Layers.Add(LaagNaam)
        NwLaagObj.Color = KleurNr
    End If
    
    On Error Resume Next
    If MaakAktief = True Then ThisDrawing.ActiveLayer = ThisDrawing.Layers.Item(LaagNaam)
    If Err Then
        Err.Clear
        MsgBox "Laag ' " & LaagNaam & " ' kan niet aktief worden gemaakt.", vbExclamation, "Let op"
    End If

End Sub


Public Sub SelectiesetsVerwijderen()

    Dim ssetObj As AcadSelectionSet
    
    For Each ssetObj In ThisDrawing.SelectionSets
        ssetObj.Clear
        ssetObj.Delete
    Next ssetObj
    
End Sub


Public Function BepalenVanSnijpunt(element1 As Object, element2 As Object, ExtendOption As Integer) 'As Variant

'26 feb 2003 EBR
'21 mei 2003: het eerste snijpunt met een polyline is het snijpunt wat het dichtst bij het
'             beginpunt van die polyline bevindt (dus rekend vanaf beginpunt) !!

'   0 = acExtendNone
'   1 = acExtendThisEntity
'   2 = acExtendOtherEntity
'   3 = acExtendBoth

'   Voorbeeld gebruik:  Snijpunt = BepalenVanSnijpunt(element1, element2, 3)
'                       Dim PuntObj As AcadPoint
'                       Set PuntObj = ThisDrawing.ModelSpace.AddPoint(vSnijpunt)

' -----------------------------------------------------------------------
'   BEREKENEN INTERSECTIE (SNIJPUNT) VAN TWEE LIJNEN
' -----------------------------------------------------------------------

    Dim intPoints As Variant
    Dim vSnijpunt(0 To 2) As Double
    
    Select Case ExtendOption
    Case 0
        'Does not extend either object.
        intPoints = element1.IntersectWith(element2, acExtendNone)
    Case 1
        'Extends the base object.
        intPoints = element1.IntersectWith(element2, acExtendThisEntity)
    Case 2
        'Extends the object passed as an argument.
        intPoints = element1.IntersectWith(element2, acExtendOtherEntity)
    Case 3
        'Extends both objects.
        intPoints = element1.IntersectWith(element2, acExtendBoth)
    End Select
     

    Dim I As Integer, j As Integer, k As Integer
    Dim str As String
    If VarType(intPoints) <> vbEmpty Then

        For I = LBound(intPoints) To UBound(intPoints)
            'str = "Intersection Point[" & k & "] is: x=" & intPoints(j) & "   y=" & intPoints(j + 1) & "   z=" & intPoints(j + 2)
            'MsgBox str: str = ""
            
            vSnijpunt(0) = intPoints(j)
            vSnijpunt(1) = intPoints(j + 1)
            vSnijpunt(2) = 0
            'MsgBox vSnijpunt(0) & "      " & vSnijpunt(1), vbCritical
            
            I = I + 2
            j = j + 3
            k = k + 1   'geeft aantal snijpunten aan
            
            '** NIEUW 21 MEI 2003: PAKT ALTIJD HET LAATSTE SNIJPUNT (IS VERSTE VANAF BEGINPUNT VAN POLYLINE !!!)
            'vSnijpunt = intPoints
            
            
            
        Next
    End If

' -----------------------------------------------------------------------
'   CONTROLE OP HET AANTAL SNIJPUNTEN
' -----------------------------------------------------------------------
        
    If k = 0 Then
        '*** deze melding uitgezet op 27 jan 2004
        MsgBox "Leiding heeft geen snijpunt met de aanvoer/ retourleiding", vbCritical, "Let op"
        Exit Function
    End If
    
    If k = 1 Then
        'vSnijpunt = intPoints         '* uitgezet op 21 mei 2003, anders wordt laatste snijpunt gepakt ipv eerste !
    End If

'    If k > 1 Then
'        MsgBox "Leiding heeft meerdere snijpunten met de aanvoer/ retourleiding", vbCritical
'        '21 mei 2003 uitgezet (was origineel)
'        'Exit Function
'    End If

' --------------------------------------------------------------------------------------------
'   SNIJPUNT WEERGEVEN
' --------------------------------------------------------------------------------------------
    
    'MsgBox vSnijpunt(0) & "   " & vSnijpunt(1), vbExclamation
    'Call Punt1(vSnijpunt, 1)
    
'    Dim PuntObj As AcadPoint
'    Dim Ppunt(0 To 2) As Double
'    Ppunt(0) = vSnijpunt(0)
'    Ppunt(1) = vSnijpunt(1)
'    Ppunt(2) = 0
'    Set PuntObj = ThisDrawing.ModelSpace.AddPoint(Ppunt)
    
    
    
    'MsgBox "Het laatste snijpunt is: " & vSnijpunt(0) & "   " & vSnijpunt(1), vbInformation
    
    BepalenVanSnijpunt = vSnijpunt

End Function


Public Sub Punt(p)
    '--------------------------------------------------------------------------------------------
    ' PUNT TEKENEN EN PD-MODE INSTELLEN
    '--------------------------------------------------------------------------------------------
    Dim PuntObj As AcadPoint
    Dim iKleur As Integer
    
    Set PuntObj = ThisDrawing.ModelSpace.AddPoint(p)
    PuntObj.Update
    
    ThisDrawing.SetVariable "PDMODE", 32
    ThisDrawing.Regen acAllViewports
End Sub
Public Sub Punt1(p, kleur)
    '--------------------------------------------------------------------------------------------
    ' PUNT TEKENEN MET KLEUR EN PD-MODE INSTELLEN
    '--------------------------------------------------------------------------------------------
    Dim PuntObj As AcadPoint
    Dim iKleur As Integer
    
    Set PuntObj = ThisDrawing.ModelSpace.AddPoint(p)
    
    iKleur = kleur
    
    PuntObj.Color = iKleur
    PuntObj.Update
    
    ThisDrawing.SetVariable "PDMODE", 32
    'ThisDrawing.Regen acAllViewports
End Sub

Public Sub TekenLijn(Pbegin As Variant, Peind As Variant, LaagNaam As String)
    MsgBox Pbegin(0) & "     " & Pbegin(1)
    MsgBox Peind(0) & "     " & Peind(1)
    
    Dim LijnObj As AcadLine
    Set LijnObj = ThisDrawing.ModelSpace.AddLine(Pbegin, Peind)
    LijnObj.Layer = LaagNaam
    LijnObj.Update
End Sub



Public Sub SchrijfLogFile(Tekst As String)
    '--------------------------------------------------------------------------------------------
    ' LOG-FILE SCHRIJVEN NAAR HARDDISK (TXT-FILE)
    '--------------------------------------------------------------------------------------------

    Dim filenum As Integer
    Dim textline As String

    'Kill ("C:\Temp\WTH-logfile.txt")
    
    filenum = FreeFile
    Open "C:\Temp\WTH-logfile.txt" For Append As filenum
        Print #filenum, Tekst
    Close filenum
    
End Sub

Public Sub VerwijderenLogFile()
    '--------------------------------------------------------------------------------------------
    ' VERWIJDEREN VOORGAANDE LOGFILE
    '--------------------------------------------------------------------------------------------
    Dim sLogFile As String
    sLogFile = "C:\Temp\WTH-logfile.txt"
    If Dir(sLogFile) <> "" Then Kill (sLogFile)
End Sub

  
'LET OP: DIT IS AANGEPAST MET EXTRA UITBREIDING ObjectNiet 29 APRIL 2004 !!
Public Function ZoekElementHandleRondPunt(p As Variant, dCrossingSelectieGrootte As Double, Optional ObjectHandleNiet As String) As String
    On Error Resume Next
    
    'DEZE FUNCTIE BEPAALD EEN ELEMENT ROND EEN OPGEGEVEN PUNT.
    'DE GROOTTE VAN DE SELECTIECROSSING KAN WORDEN OPGEGEVEN, WAARBIJ
    'DE OPGEGEVEN GROOTTE ZOWEL DE LENGTE ALS BREEDTE VORMT VAN HET
    'ZOEK-GEBIED.
    
    '1 EN 5 AUG 2003 DOOR EBR
    
    'LET OP: INDIEN GEEN ELEMENT GEVONDEN WORDT, DAN GEEFT DE FUNCTIE
    'EEN LEGE STRING TERUG !
    
    ' AANROEPEN BIJVOORBEELD ALSVOLGT:
    
    'Me.Hide
    'Dim P As Variant
    'Dim sHandle as string
    'P = ThisDrawing.Utility.GetPoint(, "Selecteer punt")
    'sHandle =  ZoekElementHandleRondPunt(P, 0)
    'IF sHandle) = "" THEN
    '   msgbox "geen element gevonden"
    'else
    '   MsgBox "gevonden: " & sHandle
    'end if
    

    '--------------------------------------------------------------
    ' INSTELLEN SELECTIE-GROOTTE
    ' --------------------------------------------------------------
   
    If dCrossingSelectieGrootte = 0 Then dCrossingSelectieGrootte = 0.1
    dCrossingSelectieGrootte = dCrossingSelectieGrootte / 2

    '--------------------------------------------------------------
    ' AANMAKEN SELECTIESET OM EINDPUNT ARC (ZOEK LIJN)
    ' --------------------------------------------------------------
   
    
    'Aanmaken selectieset
    
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double

    p1(0) = p(0) - dCrossingSelectieGrootte
    p1(1) = p(1) - dCrossingSelectieGrootte
    p1(2) = 0
    p2(0) = p(0) + dCrossingSelectieGrootte
    p2(1) = p(1) + dCrossingSelectieGrootte
    p2(2) = 0
    
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")
    If Err Then
        ssetObj.Clear
        ssetObj.Delete
        End
    End If
    
    ssetObj.Select acSelectionSetCrossing, p1, p2
    
    ' --------------------------------------------------------------
    ' UITLEZEN SELECTIESET EN ELEMENT OPVRAGEN
    ' --------------------------------------------------------------
    
    Dim element As Object
    Dim t As Integer
    Dim Object As AcadEntity
    
    For Each element In ssetObj
        'element.Color = acGreen
        'element.Highlight True
        t = t + 1
        If element.handle <> ObjectHandleNiet Then Set Object = element
    Next element


    ssetObj.Clear
    ssetObj.Delete
    
    ' --------------------------------------------------------------
    ' CONTROLE
    ' --------------------------------------------------------------
    
    If t = 0 Then
        'MsgBox "Geen lijn gevonden.", vbExclamation, "let op"
    Else
        'MsgBox Object.handle
        ZoekElementHandleRondPunt = Object.handle
    End If
    
End Function


Public Sub ZoomInOpElement(element As AcadEntity, bHighlight As Boolean)
    
    Dim minExt As Variant
    Dim maxExt As Variant

    element.GetBoundingBox minExt, maxExt
    ZoomWindow minExt, maxExt
    ZoomScaled 0.9, acZoomScaledRelative
    
    If bHighlight = True Then element.Highlight (True)

End Sub

Function GetalIsDouble(sGetal As String, Optional sMelding As String) As Boolean
        'Ebr 14 juni 2004

        Dim dGetal As Double
        
        On Error Resume Next
        
        dGetal = sGetal
        If Err Or (dGetal <> Val(sGetal)) Then
            Err.Clear
            GetalIsDouble = False
            If sMelding <> "" Then MsgBox sMelding, vbExclamation, "Let op"
        Else
            GetalIsDouble = True
        End If

End Function









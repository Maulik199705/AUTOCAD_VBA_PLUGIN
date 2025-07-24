Attribute VB_Name = "M_6TrimmenLijnen"


'---------------------------------------------------------------------------
' EBR OKTOBER 2002
' DIT PROGRAMMA TRIMT ALLE LIJNEN (AAN DE KORTSTE ZIJDE !), LANGS
' EEN OBJECT AF (LIJN, POLYLINE OF ARC)
' DE ARCS DIE AAN HET EIND OF BEGINPUNT VAN DE BETREFFENDE LIJN VERBONDEN
' (SELECTIESET) IS
' ZAL WORDEN VERPLAATST NAAR HET SNIJPUNT.
'----------------------------------------------------------------------------

Sub TrimmenLijnen()

    
   
    '-----------------------------------------------------------
    ' SELECTEER DE SNIJLIJN (LIJN OF POLYLINE)
    '-----------------------------------------------------------
    
    Dim returnObj As AcadObject
    Dim basePnt As Variant
    On Error Resume Next
    
    F_Main.Hide
        
 Dim iFoutSelectieTeller As Integer
    
              
Opnieuw1:

    If iFoutSelectieTeller > 1 Then Exit Sub 'End ***1 fewb 2004 was end
    
    ThisDrawing.Utility.GetEntity returnObj, basePnt, "Selecteer de (TRIM-) aanvoer- of retourleiding"
    
    If Err <> 0 Then
        Err.Clear
        If iFoutSelectieTeller < 1 Then MsgBox "Verkeerd geselecteerd", vbCritical, "Let op"
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw1
    Else
        'Selectieset wordt niet goed gevuld als deze buiten beeld valt, daarom zoomextents
        
        ZoomExtents
        Call BepalenSnijpunt(returnObj)
        ZoomPrevious
    End If
    
    iFoutSelectieTeller = 0
    
    ThisDrawing.EndUndoMark
    '=einde
            
End Sub

Sub BepalenSnijpunt(LijnObj As Object)

'---------------------------------------------------------------------------------
' BEPAAL SNIJPUNTEN VAN ALLE LIJNEN MET DE GESELECTEERDE SNIJLIJN
' EN BEPAAL MET SUBROUTINE 'WegtrimmenBij' WAT DE KORTSTE ZIJDE IS
' HET EINDPUNT AAN DE KORTSTE ZIJDE ZAL WORDEN GEWIJZIGD
' EINDPUNT OF BEGINPUNT ZAL GELIJK WORDEN GEMAAKT MET HET BEREKENDE SNIJPUNT.
'---------------------------------------------------------------------------------

'lijnobj: is de geselecteerde snijpolyline (aanvoer, retour of gewoon een hulppolyline)

'plaats eerst undo-mark, omdat de lijnen anders per stuk moeten worden ge-undo't
ThisDrawing.StartUndoMark


Dim element As Object
Dim intPoints As Variant
Dim iAantalSnijpunten As Integer

Dim I As Integer, j As Integer, k As Integer
Dim str As String

Dim Xsnijp As Double
Dim Ysnijp As Double
Dim TrimZijde As String

Dim LijnenGevonden As Boolean
Dim SnijPuntengevonden As Boolean
Dim TrimZijdeBeginpunt As Boolean

Dim LaagNaam As String

LaagNaam = LijnObj.Layer


For Each element In ThisDrawing.ModelSpace
If element.EntityName = "AcDbLine" And element.Layer = LijnObj.Layer Then
If element.handle <> LijnObj.handle Then
    
        TrimZijdeBeginpunt = False
        intPoints = LijnObj.IntersectWith(element, acExtendNone)
        
        
        
        If VarType(intPoints) <> vbEmpty Then
            SnijPuntengevonden = True
            
            'INDIEN MEER DAN 2 SNIJPUNTEN DAN TEKENING COTNROLEREN
            
            iAantalSnijpunten = (UBound(intPoints) + 1) / 3
            'MsgBox iAantalSnijpunten
            
            If iAantalSnijpunten > 1 Then
                MsgBox "element " & element.handle & " heeft " & iAantalSnijpunten & " snijpunten met de opschuif-polyline. " _
                & Chr(10) & Chr(13) & "Opschuiven onmogelijk: controleer de tekening", vbCritical, "Controleer tekening."
                Call ZoomInOpElement(element, True)
                End
            End If
            
            For I = LBound(intPoints) To UBound(intPoints)
                'ListBox1.AddItem ELEMENT.Handle & "  " & intPoints(j) & "," & intPoints(j + 1) & "," & intPoints(j + 2)
                              
                Xsnijp = intPoints(j)
                Ysnijp = intPoints(j + 1)
                
                
                    LijnenGevonden = True
                    TrimZijde = BepalenTrimZijde(element, Xsnijp, Ysnijp)
                    
                    If TrimZijde = "beginpunt" Then
                        TrimZijdeBeginpunt = True
                        Call VerplaatsenArcs(element.StartPoint, intPoints, TrimZijdeBeginpunt, LaagNaam)
                        element.StartPoint = intPoints
                        
                    Else
                        TrimZijdeBeginpunt = True = False
                        Call VerplaatsenArcs(element.EndPoint, intPoints, TrimZijdeBeginpunt, LaagNaam)
                        element.EndPoint = intPoints
                    End If
                    
                    'VERPLAATSEN ARCS
                    
               
                
                Exit For
                'I = I + 2
                'j = j + 3
                'k = k + 1
            Next
        End If
      
End If
End If
Next element

If LijnenGevonden = False Then MsgBox "Geen leidingen (lines) in dezelfde laag als de geselecteerde (aanvoer) polyline gevonden.", vbExclamation, "Leidingleg-programma"
If SnijPuntengevonden = False Then MsgBox "Geen leidingen worden gesneden door de geselecteerde (aanvoer) polyline.", vbExclamation, "Leidingleg-programma"
       
       
 
End Sub


Function BepalenTrimZijde(LijnObj As Object, Xsnijp As Double, Ysnijp) As String

    '-------------------------------------------------------
    'BEPALEN LENGTE VANAF HET SNIJPUNT NAAR ZOWEL HET BGIN-
    'ALS HET EINDPUNT VAN DE LIJN
    '-------------------------------------------------------

    Dim Pbegin As Variant
    Dim Peind As Variant
    
    Pbegin = LijnObj.StartPoint
    Peind = LijnObj.EndPoint
    
    'BEREKENEN LENGTE VAN SNIJPUNT TOT BEGINPUNT LIJN
    Dim Lengte1 As Double
    Lengte1 = Sqr((Pbegin(0) - Xsnijp) ^ 2 + (Pbegin(1) - Ysnijp) ^ 2)
    
    'BEREKENEN LENGTE VAN SNIJPUNT TOT EINDPUNT LIJN
    Dim Lengte2 As Double
    Lengte2 = Sqr((Peind(0) - Xsnijp) ^ 2 + (Peind(1) - Ysnijp) ^ 2)
    
    
    If Abs(Lengte2) > Abs(Lengte1) Then
        'MsgBox "WEGTRIMMEN BIJ BEGINPUNT" & LijnObj.Handle
        BepalenTrimZijde = "beginpunt"
    Else
        'MsgBox "WEGTRIMMEN BIJ EINDPUNT " & LijnObj.Handle
        BepalenTrimZijde = "eindpunt"
    End If
    
    'Update
    
End Function

Sub VerplaatsenArcs(OorspronkelijkeEindpuntLijn, NieuweEindpuntLijn, TrimZijdeBeginpunt, LaagNaam)

    'M.B.V. EEN SELECTIE-WINDOW OM HET GEVONDEN EINDPUNT (OF BEGINPUNT) VAN DE
    'LIJN DE BIJBEHORENDE ARC SELECTEREN. DEZE VERVOLGENS VERPLAATSEN.

    'If TrimZijdeBeginpunt = True Then MsgBox "1e lijn aan beginpunt-zijde inkorten"
    'If TrimZijdeBeginpunt = False Then MsgBox "1e lijn aan eindpunt-zijde inkorten"
     
    '----------------------------------------------------------------------
    'VOORGAANDE SELECTIESETS VERWIJDEREN
    '----------------------------------------------------------------------
    
    Dim ssetObj As AcadSelectionSet
    For Each ssetObj In ThisDrawing.SelectionSets
        ssetObj.Clear
        ssetObj.Delete
    Next ssetObj
    
    
    '----------------------------------------------------------------------
    'GROOTTE VAN SELECTIE-WINDOW INSTELLEN:
    '----------------------------------------------------------------------
    
    Dim P1sset(0 To 2) As Double
    Dim P2sset(0 To 2) As Double
    
    P1sset(0) = OorspronkelijkeEindpuntLijn(0) - 1
    P1sset(1) = OorspronkelijkeEindpuntLijn(1) - 1
    
    P2sset(0) = OorspronkelijkeEindpuntLijn(0) + 1
    P2sset(1) = OorspronkelijkeEindpuntLijn(1) + 1
    
    '----------------------------------------------------------------------
    'AANMAKEN SELECTIESET OM EINDPUNT/ BEGINPUNT VAN LIJN:
    '----------------------------------------------------------------------
    
    ReDim gpCode(0) As Integer
    gpCode(0) = 0
    ReDim dataValue(0) As Variant
    dataValue(0) = "Arc"
    
    Dim groupCode As Variant, dataCode As Variant
    groupCode = gpCode
    dataCode = dataValue
    
     
    
    Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
    'ssetObj.SelectOnScreen
    ssetObj.Select acSelectionSetCrossing, P1sset, P2sset, groupCode, dataCode
    
    'UITFILTEREN ARC:
    
    Dim element As Object
    Dim ArcObj As AcadArc
    Dim ArcGevonden As Boolean
    
    For Each element In ssetObj
        If element.EntityName = "AcDbArc" Then
            Set ArcObj = element
            ArcGevonden = True
            Exit For
        End If
    Next element
    
    ssetObj.Clear
    ssetObj.Delete
    
    '----------------------------------------------------------------------
    'AANMAKEN SELECTIESET OM EINDPUNT/ BEGINPUNT VAN LIJN:
    '----------------------------------------------------------------------
   
    
    Dim PuntObj As AcadPoint
    Dim Pss As Variant      'punt selectieset
    Dim StartpointArcGevonden As Boolean
    
    If ArcGevonden = True Then

            'Bepalen of start o endpoint van de arc gevonden is.
            If Lengte(ArcObj.StartPoint, OorspronkelijkeEindpuntLijn) > Lengte(ArcObj.EndPoint, OorspronkelijkeEindpuntLijn) Then
                'MsgBox "Endpunt van Arc gevonden"
                StartpointArcGevonden = True
                'Set Puntobj = ThisDrawing.ModelSpace.AddPoint(ArcObj.StartPoint)
                Pss = ArcObj.StartPoint
            Else
                'MsgBox "Startpoint van Arc gevonden"
                StartpointArcGevonden = False
                'Set Puntobj = ThisDrawing.ModelSpace.AddPoint(ArcObj.EndPoint)
                Pss = ArcObj.EndPoint
            End If
            
            
            'SELECTIE MAKEN OM START- OF ENDPOINT VAN ARC
            '*** 26 MAART 2003 WAS 4, GEWIJZIGD IN 1:
            P1sset(0) = Pss(0) - 1
            P1sset(1) = Pss(1) - 1
            
            P2sset(0) = Pss(0) + 1
            P2sset(1) = Pss(1) + 1
            
            '----------------------------------------------------------------------
            'AANMAKEN SELECTIESET OM EINDPUNT/ BEGINPUNT VAN GEVONDEN ARC:
            '----------------------------------------------------------------------
            
            ReDim gpCode(0) As Integer
            gpCode(0) = 0
            ReDim dataValue(0) As Variant
            dataValue(0) = "Line"
            
            'Dim groupCode As Variant, dataCode As Variant
            groupCode = gpCode
            dataCode = dataValue
            
            Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
            'ssetObj.SelectOnScreen
            ssetObj.Select acSelectionSetCrossing, P1sset, P2sset, groupCode, dataCode
            
            Dim LijnObj2 As AcadLine
            Dim LijnObj2Gevonden As Boolean
            
            For Each element In ssetObj
                If element.EntityName = "AcDbLine" Then
                    '*** nieuw 26 maart 2003: if conditie toegevoegd omdat verkeerde elementen werden gevonden
                    'daarom ook laagnaam als argument in de subroutine toegevoegd.
                    
                    If element.Layer = LaagNaam Then
                        Set LijnObj2 = element
                        LijnObj2Gevonden = True
                        Exit For
                    End If
                End If
            Next element
            
            ssetObj.Clear
            ssetObj.Delete
            
            ArcObj.Move OorspronkelijkeEindpuntLijn, NieuweEindpuntLijn
            
            If LijnObj2Gevonden = True Then
                'MsgBox LijnObj2.handle
                
                If Lengte(LijnObj2.StartPoint, Pss) > Lengte(LijnObj2.EndPoint, Pss) Then
'                    MsgBox "Eindpunt van lijn 2 gevonden"
'
                    If StartpointArcGevonden = True Then
                        LijnObj2.EndPoint = ArcObj.StartPoint
                    Else
                        LijnObj2.EndPoint = ArcObj.EndPoint
                    End If

                    'LijnObj2.Color = acMagenta
                
                Else
                    'MsgBox "Startpunt van lijn 2 gevonden"

                    If StartpointArcGevonden = True Then
                        LijnObj2.StartPoint = ArcObj.StartPoint
                    Else
                        LijnObj2.StartPoint = ArcObj.EndPoint
                    End If

                    'LijnObj2.Color = acMagenta
                
                End If
                
                
                
            End If
                       
                      
            
            'UPDATEN IS ZEER BELANGRIJK, ANDERS FOUTEN !
            ArcObj.Update
    End If
    'End 'test eenmalige lijnverschuiving
    
End Sub


Sub Undomark()
    'UNDO HET TRIMMEN EN VERPLAATEN VAN DE ARCS TOT DE UNDO-MARK
    ThisDrawing.Regen True
    Application.RunMacro "Undomark1"
End Sub

Sub Undomark1()
    ThisDrawing.SendCommand "UNDO" & vbCr & "B" & vbCr
    Application.Update
End Sub




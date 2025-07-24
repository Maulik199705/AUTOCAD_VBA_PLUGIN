Attribute VB_Name = "M_4BepalenLengteNEW"


'------------------------------------------------------------------------------------
' !!!!  NIEUW FEBRUARI 2003: BEPALEN LENGTE POLYLINE T.O.V. SNIJPUNT MET LEIDING !!!!
'------------------------------------------------------------------------------------


Dim ArrHandle() As String
Dim ArrLengte() As Double
Dim sLaatseHandle As String

' Public PsnijpuntPublic as Variant   ' *** aangepast 21 mei 2003, nu double
Public PsnijpuntPublic(0 To 2) As Double


Private Sub CommandButton1_Click()
    
'    ReDim ArrHandle(0)
'    ReDim ArrLengte(0)

    Dim plineObj As AcadLWPolyline
    Dim LijnObj As AcadLine
    
    Set plineObj = ThisDrawing.ModelSpace.Item(0)
    Set LijnObj = ThisDrawing.ModelSpace.Item(1)

    Call M_BepalenLengteNEW.TekenLijnenOverPolyline(plineObj)
    Call M_BepalenLengteNEWBepalenSnijpunt(plineObj, LijnObj)

End Sub

Sub TekenLijnenOverPolyline(plineObj As AcadLWPolyline)

    ReDim ArrHandle(0)
    ReDim ArrLengte(0)

    '-----------------------------------------------------------------------
    'EXPLODEREN POLYLINE
    '----------------------------------------------------------------------
    
    Dim PolyObj As AcadLWPolyline
    'Set PolyObj = ThisDrawing.ModelSpace.Item(0)
    Set PolyObj = plineObj
    
    Dim ExplodedObj As Variant
    Dim I As Integer
    
    ExplodedObj = PolyObj.Explode
    'PolyObj.Explode 'dit mag ook indien geen return-objecten nodig
    
    ' DOORLOOP DE GEEXPLODEERDE ELEMENTEN: (ELEMENT 0 IS HET EERST GETEKENDE ELEMENT)
    
    For I = 0 To UBound(ExplodedObj)
        If I = 0 Then
            'ExplodedObj(I).Color = acYellow
            ExplodedObj(I).Layer = ThisDrawing.ActiveLayer.Name
        Else
            'ExplodedObj(I).Color = acYellow
            ExplodedObj(I).Layer = ThisDrawing.ActiveLayer.Name
        End If
        
        ExplodedObj(I).Update
        'MsgBox "Exploded Object " & I & ": " & ExplodedObj(I).ObjectName
    
        '---------------------------------------------------------------------
        'OPSLAAN HANDLES VAN DE LOSSEN ELEMENT (OM LATER LENGTE TE BEREKENEN)
        '---------------------------------------------------------------------
        
        'OPSLAAN HANDLE:
        ReDim Preserve ArrHandle(UBound(ArrHandle) + 1)
        ArrHandle(UBound(ArrHandle)) = ExplodedObj(I).handle
        
        Dim VorigeLengte As Double
        VorigeLengte = ArrLengte(UBound(ArrLengte))
        
        'OPSLAAN LINE-LENGTE OF ARC-LENGTE
        ReDim Preserve ArrLengte(UBound(ArrLengte) + 1)
        If ExplodedObj(I).EntityName = "AcDbLine" Then
            ArrLengte(UBound(ArrLengte)) = VorigeLengte + ExplodedObj(I).Length
        Else
            ArrLengte(UBound(ArrLengte)) = VorigeLengte + ExplodedObj(I).ArcLength
        End If
        
     Next
    

End Sub


Function BepalenSnijpunt(PolyObj As Object, LijnObj As Object) As Double

    '-------------------------------------------------------------
    'BEPALEN SNIJPUNT VAN LEIDING (LIJN) MET AANVOER-POLYLINE
    '-------------------------------------------------------------
    
    
     Dim intPoints As Variant
     Dim x As Double
     Dim Y As Double
     
     intPoints = LijnObj.IntersectWith(PolyObj, acExtendThisEntity)
    
     Dim I As Integer, j As Integer, k As Integer
     Dim str As String
     If VarType(intPoints) <> vbEmpty Then
         For I = LBound(intPoints) To UBound(intPoints)
             x = intPoints(j)
             Y = intPoints(j + 1)
             
             I = I + 2
             j = j + 3
             k = k + 1
         Next
    End If
    
    'origineel voor 25 mrt 2003
    'If k = 0 Then MsgBox "Geen snijpunt (leiding met retour-polyline) gevonden", vbCritical: End
    'nieuw
   
    If k = 0 Then
        'MsgBox "Geen snijpunt (leiding met retour-polyline) gevonden", vbCritical
        
        '5 mei 2004 dit uitgezet want lengte nu in modemacro geplaatst !
        'ThisDrawing.SetVariable "modemacro", "* geen snijp. gevonden. *"
        Exit Function
    End If
    
    '*** 21 mei 2003 deze regel uitgevinkt
    'If k > 1 Then MsgBox "Leiding heeft meerdere snijpunten met de aanvoer-leiding", vbCritical ': End
    
    '*** 21 mei 2003 deze regel nieuw
    If k > 1 Then SchrijfLogFile ("Leiding heeft meerdere snijpunten met de aanvoer-leiding")
    'Punt (intPoints)
    
      
    '----------------------------------------------------------------------------------
    'MAAK SELECTIESET (WINDOW) OP HET SNIJPUNT
    'OP DIT SNIJPUNT LIGT DE POLYLINE EN HIEROVER EEN LOSSE LINE OF ARC
    'VAN DIT LOSSE ELEMENT WORDT DE HANDLE BEPAALD OM VERVOLGENS DE LENGTE TE BEREKENEN
    '----------------------------------------------------------------------------------
    Dim point1(0 To 2) As Double
    Dim Point2(0 To 2) As Double
    
    point1(0) = x - 0.5
    point1(1) = Y - 0.5
    Point2(0) = x + 0.5
    Point2(1) = Y + 0.5
    
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET1")
    ssetObj.Select acSelectionSetCrossing, point1, Point2
    
    'MsgBox "Aantal elementen in selectieset"& ssetObj.Count

    Dim element As Object
    Dim LijnTeller As Integer
    Dim LaatsteObj As AcadEntity
    
    For Each element In ssetObj
        If element.EntityName = "AcDbLine" Or element.EntityName = "AcDbArc" Then
            'de volgende if is nieuw op 25 maart 2003
            'MsgBox Element.Layer
            'MsgBox Element.handle
            If element.Layer = ThisDrawing.ActiveLayer.Name Then
            If element.handle <> LijnObj.handle Then
                    LijnTeller = LijnTeller + 1
                    'MsgBox Element.handle
                    'Element.Color = acMagenta
                    Set LaatsteObj = element
            End If
            End If
        End If
    Next element
    
    
    'MsgBox "Laatste lijn (in het snijpunt) is: " & LaatsteObj.handle
    
    
    
    ssetObj.Clear
    ssetObj.Delete
    
    If LijnTeller = 0 Then MsgBox "Geen hulpsnijlijn gevonden": End
    'If LijnTeller > 1 Then MsgBox "Meer dan 1 hulpsnijlijn gevonden"    ': End
    'If LijnTeller > 1 Then MsgBox "Meer dan 1 hulpsnijlijn gevonden"
    If LijnTeller > 1 Then SchrijfLogFile ("* Meer dan 1 hulpsnijlijn gevonden bij het snijpunt met de geexplodeerde aanvoer-polyline.")
    
    
    sLaatseHandle = LaatsteObj.handle
    '-------------------------------------------------------------
    'ENDPOINT VAN GEVONDEN LIJN INKORTEN TOT SNIJPUNT
    '-------------------------------------------------------------
    Dim Peind(0 To 2) As Double
    Peind(0) = x
    Peind(1) = Y
    
    '*** 26 maart 2003 UITGEZET
    'If LaatsteObj.EntityName = "AcDbLine" Then LaatsteObj.EndPoint = Peind
    
    
    '-------------------------------------------------------------
    'BEPALEN LENGTE VAN DE VOORGAANDE LIJNSTUKKEN
    '-------------------------------------------------------------
    '*** 26 maart 2003 ORIGINEEL:
'    Dim LaatsteObjectLengte As Double
'    If LaatsteObj.EntityName = "AcDbLine" Then
'        LaatsteObjectLengte = LaatsteObj.Length
'    Else
'        LaatsteObjectLengte = LaatsteObj.ArcLength
'    End If
    
    
    Dim LaatsteObjectLengte As Double
    If LaatsteObj.EntityName = "AcDbLine" Then
        LaatsteObjectLengte = Lengte(LaatsteObj.StartPoint, Peind)
        'LaatsteObjectLengte = LaatsteObj.Length
    Else
        LaatsteObjectLengte = LaatsteObj.ArcLength
    End If
        
    
    
    
    Dim t As Integer                                'tellen lijnen (aanvoer over de polyline)
    Dim LengteTotSnijp As Double
    
    For t = 1 To UBound(ArrHandle)
        
        'MsgBox "handle=" & ArrHandle(t) & Chr(10) & Chr(13) & "totale lengte =" & ArrLengte(t)
        
        If ArrHandle(t) = LaatsteObj.handle Then
            LengteTotSnijp = ArrLengte(t - 1) + LaatsteObjectLengte
        End If
        
    Next t
    
    'MsgBox "Lengte polyline tot snijpunt=" & LengteTotSnijp
    
    BepalenSnijpunt = LengteTotSnijp
    
    ' *** uitgezet: dit gaat niet meer goed 21 mei 2003
    ' PsnijpuntPublic = intPoints         'opslaan om lengte van eerste lijn te kunnen bepalen
    PsnijpuntPublic(0) = x
    PsnijpuntPublic(1) = Y
    PsnijpuntPublic(2) = 0

End Function

Sub VerwijderenLosseLines()
    
    
    '-----------------------------------------------------------------------------
    'LOSSE LINES EN ARC VAN DE GEEXPLODEERDE POLYLINE VERWIJDEREN VANAF
    'HET LAATSTE ELEMENT BIJ HET SNIJPUNT
    'HET ARRAY MET HANDLES WORDT ACHTERSTEVOREN DOORLOPEN
    '------------------------------------------------------------------------------

    'MsgBox sLaatseHandle

    Dim t As Integer
    Dim LengteTotSnijp As Double
    Dim element As AcadEntity
    
    'For t = 1 To UBound(ArrHandle)
    For t = UBound(ArrHandle) To 1 Step -1
        
        'MsgBox ArrHandle(t)
        Set element = ThisDrawing.HandleToObject(ArrHandle(t))
        
        'EERSTE LIJN (BEGINLIJN VAN GEEXPOLDEERDE RETOUR-POLYLINE (=EERSTE IN ARRAY) TRIMMEN
        'TOT LAATSTE SNIJPUNT, REST VERWIJDEREN
        
        'VOOR 31 MAART 2003:
        If ArrHandle(t) = sLaatseHandle Then
        'MsgBox sLaatseHandle, vbInformation
        
        'If ArrHandle(t) = ArrHandle(1) Then
            '28 maar aangevuld alleen met volgende regel
            If element.EntityName = "AcDbLine" Then element.EndPoint = PsnijpuntPublic
            ' ** 21 MEI 2003, OMDAT VARIABELE NIET GOED WAS (VARIANT IPV DOUBLE-ARRAY)
            ' SPRONG HET PROGRAMMA HIER UIT DEZE SUB ? IS NU AANGEPAST.
            Exit For
        Else
            'MsgBox "13 mei 2004 erase teruggezet !"
            element.Color = acRed
            element.Erase
        End If
      
        
    Next t

End Sub




 '-------------------------------------------------------------
    'AFBEELDEN PUNT OP SNIJPUNT
    '-------------------------------------------------------------
    
'    Dim retCoord As Variant
'    retCoord = PlineObj.Coordinates
'
'    Dim X1, Y1, X2, Y2 As Double
'
'    Dim LijnstukTeller As Integer
'    For t = LBound(retCoord) To UBound(retCoord) - 2 Step 2
'
'        LijnstukTeller = LijnstukTeller + 1
'
'        X1 = retCoord(t)
'        Y1 = retCoord(t + 1)
'        X2 = retCoord(t + 2)
'        Y2 = retCoord(t + 3)
'
'        'Call TekenLijntje(X1, Y1, X2, Y2)
'
'        'MsgBox "X1=" & X1 & "     Y1=" & Y1 & Chr(10) & Chr(13) _
'            & "X2=" & X2 & "     Y2=" & Y2, , "lijnstuk: " & LijnstukTeller
'
'    Next t
    


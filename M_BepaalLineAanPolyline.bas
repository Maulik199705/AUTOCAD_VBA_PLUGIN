Attribute VB_Name = "M_BepaalLineAanPolyline"
Public Function BepaalLijnAanEindePolyline(ObjPolyline As Object) As String

    ZoomExtents         'LET OP, ZOOM EXTENTS GEHEEL BOVEN SELECTIESET PLAATSEN, ANDERS LEGE SELECTIESET !!
    
    '2 okt 2002
    '-----------------------------------------------------------------------------------
    'DEZE FUNCTIE BEPAALT WELKE LIJN ZICH BEVINDT AAN HET AANPUNT VAN EEN LWT-POLYLINE
    'INPUT: OBJECT LWT-POLYLINE
    'OUTPUT: HANDLE VAN DE LINE
    '-----------------------------------------------------------------------------------
    
    'MsgBox ObjPolyline.handle
    'Dim objPline As Object
    'Set objPline = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)
    
    Dim intVCnt As Integer
    Dim varCords As Variant
    Dim varVert As Variant
    
    coord = ObjPolyline.Coordinates
     
    'MsgBox "Eerste x =" & coord(0)
    'MsgBox "Eerste Y =" & coord(1)
    
    'bepalen of tellen van het aantal vertexten
    For Each varVert In coord
      intVCnt = intVCnt + 1
    Next
     
    'MsgBox "Laatste x =" & coord(intVCnt - 2)
    'MsgBox "Laatste y =" & coord(intVCnt - 1)
    
    '----------------------------------------------------------
    ' Create the selection set
    '----------------------------------------------------------
    
    Dim ssetObj As AcadSelectionSet
    For Each ssetObj In ThisDrawing.SelectionSets
        'ssetObj.Clear
        ssetObj.Delete
    Next ssetObj
    
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")

    Dim mode As Integer
    Dim corner1(0 To 2) As Double
    Dim corner2(0 To 2) As Double

    mode = acSelectionSetCrossing
    corner1(0) = coord(intVCnt - 2) - 2
    corner1(1) = coord(intVCnt - 1) - 2
    corner1(2) = 0

    corner2(0) = coord(intVCnt - 2) + 2
    corner2(1) = coord(intVCnt - 1) + 2
    corner2(2) = 0

    ssetObj.Select mode, corner1, corner2
 
    
    
    Dim element As Object
    Dim tl As Integer
    Dim LijnObj As Object
    
    '*** MsgBox "HIER FATEL ERROR ?", vbExclamation
    
    
    
    For Each element In ssetObj
        If element.handle <> ObjPolyline.handle Then
            If element.EntityName = "AcDbLine" Then
                tl = tl + 1
                Set LijnObj = element
            End If
        End If
    Next element
        
    ssetObj.Clear
    ssetObj.Delete
    
    If tl = 0 Then SchrijfLogFile ("Geen leiding (lijn) verbonden met het eindpunt van de aanvoer leiding (polyline) (of UCS-niet in World ?) OF eerst zoom-extents")
    If tl > 1 Then SchrijfLogFile ("Meerdere leidingen (lijnen) verbonden met het eindpunt van de aanvoer leiding (polyline)")
    
    If tl = 0 Then MsgBox "Geen leiding (lijn) verbonden met het eindpunt van de aanvoer leiding (polyline) (UCS niet in World ?) OF eerst zoom-extents", vbCritical: End
    If tl > 1 Then MsgBox "Meerdere leidingen (lijnen) verbonden met het eindpunt van de aanvoer leiding (polyline)", vbCritical: End
   

    SchrijfLogFile ("Gevonden lijn aan einde polyline: handle=" & LijnObj.handle)
    'MsgBox "gevonden lijn: " & LijnObj.handle
    BepaalLijnAanEindePolyline = LijnObj.handle
    
    ZoomPrevious
    
End Function


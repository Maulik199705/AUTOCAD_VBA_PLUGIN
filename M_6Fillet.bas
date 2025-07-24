Attribute VB_Name = "M_6Fillet"
    
' totale fillet command d.m.v. sendcommand:  5 aug 2003  1,5 uur
' 6 aug 2003: 8:15 t/m
    
Sub Fillet5aug(FilletP1, FilletP2, Psnijpunt1, dFilletStraal)
'
'     Punt1 FilletP1, 1
'     Punt1 FilletP2, 2
'     Punt1 Psnijpunt1, 3

    ' ----------------------------------------------------------------------------
    ' Fillet d.m.v. sendcommand
    ' Door EBR 5 aug 2003
    ' ----------------------------------------------------------------------------
        

    Dim plineObj As AcadLWPolyline
    Dim Points(0 To 5) As Double
    
    Points(0) = FilletP1(0):      Points(1) = FilletP1(1)
    Points(2) = Psnijpunt1(0):     Points(3) = Psnijpunt1(1)
    Points(4) = FilletP2(0):      Points(5) = FilletP2(1)
    Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(Points)
    plineObj.Update
    
    
    ThisDrawing.SendCommand ("FILLET" & vbCr & "R" & vbCr & dFilletStraal & vbCr & "P" & vbCr & "L" & vbCr)
    
    
    ' EXPLODEREN LAATST GETEKENDE POLYLINE EN LOSSE LINES VERWIJDEREN
    
    Dim explodedObjects As Variant
    Dim I As Integer
    Dim FilletArcObj As AcadEntity
    
    explodedObjects = plineObj.Explode
    
    For I = 0 To UBound(explodedObjects)
        'MsgBox explodedObjects(I).EntityName
        If explodedObjects(I).EntityName = "AcDbArc" Then Set FilletArcObj = explodedObjects(I)
        If explodedObjects(I).EntityName = "AcDbLine" Then explodedObjects(I).Delete
    Next
    
    plineObj.Delete
    
    If IsEmpty(FilletArcObj) Then
        MsgBox "Plaatsen bocht (fillet) bij aanvoerleiding met de eerste leiding is niet gelukt", vbInformation, "Let op"
        Exit Sub
    End If
        
    
    
    
End Sub

' ORIGINEEL PRINCIPE DOOR EBR:



'''Sub Fillet
'''
'''
'''    ' ----------------------------------------------------------------------------
'''    ' Fillet d.m.v. sendcommand
'''    ' Door EBR 5 aug 2003
'''    ' ----------------------------------------------------------------------------
'''
'''
'''    Dim LijnObj1 As AcadLine
'''    Dim LijnObj2 As AcadLine
'''    Dim Psnij As Variant
'''
'''    Set LijnObj1 = ThisDrawing.ModelSpace.Item(0)
'''    Set LijnObj2 = ThisDrawing.ModelSpace.Item(1)
'''
'''    Psnij = LijnObj1.IntersectWith(LijnObj2, acExtendBoth)
'''
'''    If IsEmpty(Psnij) Then MsgBox "Geen snijpunt gevonden", vbCritical: End
'''
'''    Dim Peind1 As Variant
'''    Dim Peind2 As Variant
'''
'''    Peind1 = LijnObj1.StartPoint
'''    Peind2 = LijnObj2.StartPoint
'''
'''
'''    Dim plineObj As AcadLWPolyline
'''    Dim points(0 To 5) As Double
'''    points(0) = Peind1(0):      points(1) = Peind1(1)
'''    points(2) = Psnij(0):       points(3) = Psnij(1)
'''    points(4) = Peind2(0):      points(5) = Peind2(1)
'''    Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
'''
'''
'''    Dim dStraal As Double
'''    dStraal = 5
'''    ThisDrawing.SendCommand ("FILLET" & vbCr & "R" & vbCr & dStraal & vbCr & "P" & vbCr & "L" & vbCr)
'''
'''End Sub
    
    
    
    

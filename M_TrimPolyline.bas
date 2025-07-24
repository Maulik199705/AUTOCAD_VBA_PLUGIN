Attribute VB_Name = "M_TrimPolyline"

'''TEST VOOR ONDERSTAANDE ROUTINE
'''Private Sub CommandButton38_Click()
'''    Dim PolyObj As AcadLWPolyline
'''    Dim p As Variant
'''
'''    Me.Hide
'''opnieuw:
'''
'''    ThisDrawing.Utility.GetEntity PolyObj, p
'''    Call M_TrimPolyline.TrimPolyline(PolyObj, p)
'''
'''    GoTo opnieuw
'''End Sub







Public Sub TrimPolyline(PolyObj As AcadLWPolyline, Psnij As Variant)
    
    'Punt (Psnij)
    'MsgBox "trimpolyline", vbInformation
    
    '-----------------------------------------------------------------------------
    'DOOR EBR 6 MEI 2004:
    'DEZE FUNCTIE TRIMT EEN POLYLINE. HET EINDE VAN POLYLINE ZAL HIERBIJ WORDEN
    'VERWIJDERD (DUS NIET VANAF BEGINPUNT)
    'HET SNIJPUNT MOET WEL "EXACT" OP DE POLYLINE LIGGEN.
    'FUNCTIE GAAT NIET GOED ALS EEN PUNT OP ARC OP DE POLYLINE GESELECTEERD WORDT
    'DUS POYLINE MAG VOOR GOEDE WERKING GEEN BULGES BEVATTEN.
    '(ECHTER, EEN POLYLINE KAN AAN HET UITEINDE NOOIT GEEN BULGE HEBBEN)
    '-----------------------------------------------------------------------------
        
    Dim varCords
    Dim t As Integer
    Dim bStoppen As Boolean
    
    Dim dHoekRad As Double
    Dim dHoekRadnw As Double
    
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    
    
opnieuw:

    varCords = PolyObj.Coordinates
    'MsgBox UBound(varCords)
    
    'INDIEN POLYLINE UIT SLECHTS 1 LIJN BESTAAT, DAN EINDPUNT GELIJKMAKEN AAN SNIJPUNT
    If UBound(varCords) < 4 Then
        'MsgBox "te trimmen polyline bestaat uit 1 lijnsegment (0 t/m 3 is 4 vertexen)"
        varCords(UBound(varCords) - 1) = Psnij(0)
        varCords(UBound(varCords) - 0) = Psnij(1)
        
        PolyObj.Coordinates = varCords
        'PolyObj.Color = acRed
        Exit Sub
    End If
    
    
    'HOEK BEPALEN VAN LAATSTE VAN POLYLINE
    'EINDPUNT VAN POLYLINE GELIJK MAKEN AAN SNIJPUNT EN
    'OPNIEUW HOEK OPVRAGEN
    
    p1(0) = varCords(UBound(varCords) - 3)
    p1(1) = varCords(UBound(varCords) - 2)
    p2(0) = varCords(UBound(varCords) - 1)
    p2(1) = varCords(UBound(varCords) - 0)
    
    dHoekRad = ThisDrawing.Utility.AngleFromXAxis(p1, p2)
    
    varCords(UBound(varCords) - 1) = Psnij(0)
    varCords(UBound(varCords) - 0) = Psnij(1)
    
    p1(0) = varCords(UBound(varCords) - 3)
    p1(1) = varCords(UBound(varCords) - 2)
    p2(0) = varCords(UBound(varCords) - 1)
    p2(1) = varCords(UBound(varCords) - 0)
    
    dHoekRadnw = ThisDrawing.Utility.AngleFromXAxis(p1, p2)


    'INDIEN HOEKEN NIET GELIJK, DAN COORDINATEN NIET AANPASSEN EN STOPPEN
    'ANDERS LAATSTE X EN Y-WAARDE VAN POLILYNE VERWIJDEREN EN OPNIEUW
    
    'LET OP: ROUND, ANDERS TE NAUWKEURIG !
    
    If Round(dHoekRad, 0) = Round(dHoekRadnw, 0) Then
        bStoppen = True
    Else
        'laatste x en y waarde verwijderen
        ReDim Preserve varCords(UBound(varCords) - 2)

        p1(0) = varCords(UBound(varCords) - 3)
        p1(1) = varCords(UBound(varCords) - 2)
        p2(0) = varCords(UBound(varCords) - 1)
        p2(1) = varCords(UBound(varCords) - 0)
    End If
    

    PolyObj.Coordinates = varCords
    'PolyObj.Color = acGreen
    
    If bStoppen = True Then Exit Sub
    GoTo opnieuw
    

End Sub

'manier 1:
'Dim Points
'Points = varCords
'ReDim Preserve Points(UBound(Points) - 2)

                        
'    Dim PolyObj2 As AcadLWPolyline
'    Set PolyObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(varCords)
'    PolyObj2.Color = acYellow
'    PolyObj2.Update



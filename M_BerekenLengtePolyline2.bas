Attribute VB_Name = "M_BerekenLengtePolyline2"
Public Function LengtePolyline2(objPline As Object) As Double
 
    
    'MsgBox objPline.handle
    '-------------------------------------------------------------------------------------
    'DIT PROGRAMMA
    '   - BEREKEND DE LENGTE VAN EEN LWT-POLYLINE
    '-------------------------------------------------------------------------------------
            
    
    If objPline.EntityName <> "AcDbPolyline" Then
        MsgBox "Het geslecteerde element is geen LWT-Polyline", vbCritical, "Let op"
        Exit Function
    End If
       
    
    '------------------------------------------------------
    'BEREKENEN LENGTE
    '------------------------------------------------------
    
    Dim intVCnt As Integer  'interval counter
    Dim varCords As Variant
    Dim varVert As Variant  'vertexen
    Dim varCord As Variant
    Dim varNext As Variant
    Dim intCrdCnt As Integer
    Dim dblTemp As Double
    Dim dblArc As Double
    Dim dblAng As Double
    Dim dblChord As Double
    Dim dblInclAng As Double
    Dim dblRad As Double
    
      
    varCords = objPline.Coordinates
    
    For Each varVert In varCords
      intVCnt = intVCnt + 1
    Next
    
    For intCrdCnt = 0 To intVCnt / 2 - 1              'For LWPoly 2 - 1 else 3 -1
           
      
      If intCrdCnt < intVCnt / 2 - 1 Then
        If objPline.GetBulge(intCrdCnt) = 0 Then
          varCord = objPline.Coordinate(intCrdCnt)
          varNext = objPline.Coordinate(intCrdCnt + 1)
          
          'computes a simple Pythagorean length
          dblTemp = dblTemp + Sqr(((varCord(0) - varNext(0)) ^ 2) + ((varCord(1) - varNext(1)) ^ 2))
          
                    
        Else
          'If there is a bulge we need to get an arc length
          varCord = objPline.Coordinate(intCrdCnt)
          varNext = objPline.Coordinate(intCrdCnt + 1)
          dblChord = Sqr(((varCord(0) - varNext(0)) ^ 2) + ((varCord(1) - varNext(1)) ^ 2))
          
          'Bulge is the tangent of 1/4 of the included angle between
          'vertices. So we reverse the process to get the included angle
          dblInclAng = Atn(Abs(objPline.GetBulge(intCrdCnt))) * 4
          dblAng = (dblInclAng / 2) - ((Atn(1) * 4) / 2)
          dblRad = (dblChord / 2) / (Cos(dblAng))
          dblArc = dblInclAng * dblRad
          dblTemp = dblTemp + dblArc
        End If
      End If
      
    Next intCrdCnt
      
    'WEERGEVEN LENGTE IN FUNCTIEAANROEP
    LengtePolyline2 = dblTemp
    
End Function





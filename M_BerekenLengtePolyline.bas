Attribute VB_Name = "M_BerekenLengtePolyline"
Public BerekendeAfstTotVertex As Double

'VOORBEELD AANROEP VANUIT FORM:
'Private Sub CommandButton8_Click()
'    Dim objPline As Object
'    Dim LengtePolyline As Double'
'    Set objPline = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)
'
'    Dim PSnij(0 To 2) As Double
'    PSnij(0) = 50
'    PSnij(1) = 50
'    PSnij(2) = 0
'
'    LengtePolyline = M_BerekenLengtePolyline.LengtePolyline(objPline, PSnij, False)
'    MsgBox "LengtePolyline=" & LengtePolyline'
'End Sub



Public Function LengtePolyline(objPline As Object, Psnij As Variant, bZoekLengte As Boolean) As Double
   
    'Dim DebugMode As Boolean
    'DebugMode = True
    
    Dim AfstSnijpTotVertex As Double
    Dim KortsteLengte As Double
    KortsteLengte = 100000
    
    
    '-------------------------------------------------------------------------------------
    'DIT PROGRAMMA
    '   - BEREKEND DE LENGTE VAN EEN LWT-POLYLINE
    '   - BEPAALD DE KORTSTE AFSTAND VAN DE POLYLINEVERTEX MET EEN OPGEGEVEN SNIJPUNT
    '-------------------------------------------------------------------------------------
    
    F_Main.ListBox3.Clear   'opslaan coordinaten van lwt-polyline
    
    
    If objPline.EntityName <> "AcDbPolyline" Then
        MsgBox "Het geslecteerde element is geen LWT-Polyline", vbCritical, "Let op"
        Exit Function
    End If
       
    Dim VoorgaandeLengte As Double
    '------------------------------------------------------
    'BEREKENEN LENGTE
    '------------------------------------------------------
    
    Dim intVCnt As Integer
    Dim varCords As Variant
    Dim varVert As Variant
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
          
          VoorgaandeLengte = dblTemp
          'computes a simple Pythagorean length
          dblTemp = dblTemp + Sqr(((varCord(0) - varNext(0)) ^ 2) + ((varCord(1) - varNext(1)) ^ 2))
                            
            'AFBEELDEN COORDINATEN VAN LWT-POLYLINE
            F_Main.ListBox3.AddItem AFR(varCord(0)) & " EN " & AFR(varCord(1))
                       
                    
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
      
        ' -------------------------------------------------------------------------
        ' BEREKENEN KORTSTE AFSTAND TOT SNIJPUNT
        ' -------------------------------------------------------------------------
        AfstSnijpTotVertex = Sqr(((varCord(0) - Psnij(0)) ^ 2) + ((varCord(1) - Psnij(1)) ^ 2))
        If AfstSnijpTotVertex < KortsteLengte Then KortsteLengte = AfstSnijpTotVertex
        
        If intCrdCnt Mod 2 = 0 Then
            'ALLEEN ALS HET BEGINPUNT IS
            If Debugmode = True Then MsgBox "Afstand van het opgegeven snijpunt tot de 'aktieve' vertex =" & AfstSnijpTotVertex
        End If
        
        If bZoekLengte = True Then
            If AfstSnijpTotVertex = BerekendeAfstTotVertex Then
            
                If Debugmode = True Then MsgBox "LENGTE POLYLINE VAN BEGINPUNT (POLYLINE) TOT HET OPGEGEVEN SNIJPUNT = " & AfstSnijpTotVertex + VoorgaandeLengte, vbInformation
                        
                'LengtePolyline = dblTemp                                       'LENGTE VAN SNIJPUNT TOT DICHTSTBIJZIJNDE VERTEX
                  
                        
                ' BEPALEN OF DE VERTEX EEN 'BEGIN- OF EINDPUNT' IS
                If intCrdCnt Mod 2 = 0 Then
                    'MsgBox "Snijpunt bij Beginpunt"
                    'SNIJPUNT BOVEN HET MIDDEN (AAN KANT VAN BEGINPUNT)
                    LengtePolyline = VoorgaandeLengte + AfstSnijpTotVertex
                Else
                    'MsgBox "Snijpunt bij eindpunt"
                    'SNIJPUNT ONDER HET MIDDEN (AAN KANT VAN BEGINPUNT)
                    LengtePolyline = VoorgaandeLengte - AfstSnijpTotVertex
                End If
                        
                 
                Exit Function
            End If
        End If
    
    Next
    
        'AFBEELDEN LAATSTE COORDINAATVAN LWT-POLYLINE
        F_Main.ListBox3.AddItem AFR(varNext(0)) & " EN " & AFR(varNext(1))
        
        If Debugmode = True Then MsgBox "KortsteLengte=" & KortsteLengte
        BerekendeAfstTotVertex = KortsteLengte  'DEZE OPSLAAN ALS PUBLIC, OM HER TE GEBRUIKEN !! NA EERSTE BEREKENING VOOR DE TWEEDE KAAR HIER NAARTOE REKENEN.
        
    VoorgaandeLengte = dblTemp
    LengtePolyline = dblTemp
    
  
End Function




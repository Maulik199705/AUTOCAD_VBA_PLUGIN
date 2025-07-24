VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_TrimPolyline_TEST 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8568.001
   OleObjectBlob   =   "F_TrimPolyline_TEST.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_TrimPolyline_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Call TrimPolyline2(ThisDrawing.ModelSpace.Item(0), 80, 0, 0)
End Sub

Public Sub TrimPolyline2(objPline As Object, InvoerLengte, LaatsteX, LaatsteY)
     
    'WERKING:
    'ALS X = 0 EN Y = 0 DAN ZAL HET LAATSTE LIJNSTUK ZIJN OORSPRONKELIJKE LENGTE BEHOUDEN
    'ALS X <> 0 OF Y <> 0 DAN ZAL HET LAATSTE LIJNSTUK TOT HET OPGEGEVEN COORDINAAT WORDEN VERKORT
    
    'MsgBox objPline.handle
    
    '-------------------------------------------------------------------------------------
    'DIT PROGRAMMA BEREKEND DE LENGTE VAN EEN LWT-POLYLINE
    '-------------------------------------------------------------------------------------
            
    
    If objPline.EntityName <> "AcDbPolyline" Then
        MsgBox "Het geslecteerde element is geen LWT-Polyline", vbCritical, "Let op"
        Exit Sub
    End If
       
    
    Dim intVCnt As Integer  'interval counter
    Dim varCords As Variant
    Dim varVert As Variant  'vertexen
    Dim varCord As Variant
    Dim varNext As Variant
    Dim intCrdCnt As Integer
    Dim dLengte As Double
    Dim dblArc As Double
    Dim dblAng As Double
    Dim dblChord As Double
    Dim dblInclAng As Double
    Dim dblRad As Double
    
    Dim LaatsteVertex As Integer
    
      
    varCords = objPline.Coordinates
    
    'AANTAL VERTEXEN (COORDINATEN VAN POLYLINE)
    For Each varVert In varCords
      intVCnt = intVCnt + 1
    Next
     
    Dim punten() As Double      'LET OP, MOET ALS DOUBLE ANDERS ERROR BIJ AddLightWeightPolyline(punten) !
    ReDim punten(0)
       
            
    For intCrdCnt = 0 To intVCnt / 2 - 1                'For LWPoly 2 - 1 else 3 -1
      If intCrdCnt < intVCnt / 2 - 1 Then               'noodzakelijk anders error bij een enkele polyline
      
                '-------------------------------------------------------------------------
                'BEPAAL LIJN-LENGTE (INDIEN GETBULGE = 0, DAN GEEN ARC)
                '-------------------------------------------------------------------------
            
                If objPline.GetBulge(intCrdCnt) = 0 Then
                  varCord = objPline.Coordinate(intCrdCnt)
                  varNext = objPline.Coordinate(intCrdCnt + 1)
                  
                  dLengte = dLengte + Sqr(((varCord(0) - varNext(0)) ^ 2) + ((varCord(1) - varNext(1)) ^ 2))
                                          
                Else
                
                '-------------------------------------------------------------------------
                'BEPAAL ARC-LENGTE (INDIEN GETBULGE <> 0, DAN ARC)
                '-------------------------------------------------------------------------
                    varCord = objPline.Coordinate(intCrdCnt)
                    varNext = objPline.Coordinate(intCrdCnt + 1)
                    dblChord = Sqr(((varCord(0) - varNext(0)) ^ 2) + ((varCord(1) - varNext(1)) ^ 2))
                    
                    'Bulge is the tangent of 1/4 of the included angle between
                    'vertices. So we reverse the process to get the included angle
                    dblInclAng = Atn(Abs(objPline.GetBulge(intCrdCnt))) * 4
                    dblAng = (dblInclAng / 2) - ((Atn(1) * 4) / 2)
                    dblRad = (dblChord / 2) / (Cos(dblAng))
                    dblArc = dblInclAng * dblRad
                    dLengte = dLengte + dblArc
                End If
                
                
                
                
           
                ' ***********************************************************
                ' *** UITBREIDING 26 FEB 2003 DOOR EBR: OPSLAAN VERTEXEN: ***
                ' ***********************************************************
                
                'MsgBox "x=" & varCord(0) & "   y=" & varCord(1)
                
                'NIEUW VERTEXEN ARRAY
                If intCrdCnt = 0 Then
                    punten(UBound(punten)) = varCord(0)
                    ReDim Preserve punten(UBound(punten) + 1)
                    punten(UBound(punten)) = varCord(1)
                Else
                    ReDim Preserve punten(UBound(punten) + 1)
                    punten(UBound(punten)) = varCord(0)
                    ReDim Preserve punten(UBound(punten) + 1)
                    punten(UBound(punten)) = varCord(1)
                End If
                                 
                LaatsteVertex = intCrdCnt
                If dLengte > InvoerLengte Then Exit For
            
            
        End If
    Next intCrdCnt
      
    'WEERGEVEN LENGTE IN FUNCTIEAANROEP:
    'LengtePolyline2 = dLengte
    
    MsgBox "LaatsteVertex=" & LaatsteVertex
    MsgBox "Lengte verschil=" & dLengte - InvoerLengte
    
    'OPSLAAN LAATSTE TWEE VERTEXEN
    ReDim Preserve punten(UBound(punten) + 1)
    punten(UBound(punten)) = varNext(0)
    ReDim Preserve punten(UBound(punten) + 1)
    punten(UBound(punten)) = varNext(1)
    
    '-------------------------------------------------------------------------
    'AFBEELDEN VERTEXEN POLYLINE (DEBUG)
    '------------------------------------------------------------------------
    
'    For t = LBound(punten) To UBound(punten)
'        MsgBox punten(t), vbExclamation
'    Next t
        
    '-------------------------------------------------------------------------
    'TEKEN NIEUWE POLYLINE (MET DE JUISTE LENGTE)  25 FEB 2003
    '-------------------------------------------------------------------------
           
    
    'TEKENEN NIEUWE GETRIMDE POLYLINE:
    Dim NieuwPoly As AcadLWPolyline
    Set NieuwPoly = ThisDrawing.ModelSpace.AddLightWeightPolyline(punten)
    NieuwPoly.Color = acRed
    NieuwPoly.Update
    
    'NADERHAND ARCS (BULGES) IN POLYLINE PLAATSEN:
    varCords = NieuwPoly.Coordinates
    For Each varVert In varCords
      intVCnt = intVCnt + 1
    Next
    
    'If LaatsteVertex = 0 Then LaatsteVertex = intCrdCnt + 2

    For intCrdCnt = 0 To LaatsteVertex - 1
         NieuwPoly.SetBulge intCrdCnt, objPline.GetBulge(intCrdCnt)
    Next intCrdCnt
    
  
    
    
End Sub







Private Sub CommandButton3_Click()
'WERKT NIET: GEEFT ERROR
Dim Obj As AcadEntity

SelecterenObject ("Selecteer leiding")
Obj.Color = acYellow

End Sub
Sub SelecterenObject(sBeschrijving As String) 'As AcadEntity

    'Dim ReturnObj1
    Me.Hide
    
    ThisDrawing.Utility.GetEntity ReturnObj1, basePnt1, sBeschrijving
    If Err <> 0 Then
        Err.Clear
        MsgBox "Verkeerd geselecteerd"
        Exit Sub
    End If
    
    If UCase(ReturnObj1.EntityName) <> ObjecType Then
        MsgBox "Verkeerde element geselecteer", vbCritical, "Let op"
        Exit Sub
    End If
    
    Set SelecterenObject = ReturnObj1
    
End Sub































Attribute VB_Name = "M_6PlaatsenSlingers"
'gemaakt op 30 + 31 juli 2003

'30 juli 2003 C.A. 4 UUR
'31 JULI 2003 C.A. 4 UUR
'1+5 AUG2003 C.A. 2 UUR (ZIE OOK M_0Algemeen > function ZoekElementHandleRondPunt(P1, 0.1) !!

' ---------------------------------------------------------------------
' INSERT BLOCK MET SLINGER (ARC'S EN LINES)
' ---------------------------------------------------------------------


Sub OpvragenHOH()

    Dim dHOH As Double
    On Error Resume Next
    dHOH = F_Main.ComboBox2.Text
    If Err Then
        Err.Clear
        MsgBox "Opgegeven HOH-afstand is niet juist", vbCritical, "Let op"
        Exit Sub
    End If
    
    If dHOH = 0 Then MsgBox "Opgegeven HOH-afstand is niet juist", vbCritical, "Let op": Exit Sub
  
    Call PlaatsenSlingers(dHOH)
End Sub

Sub PlaatsenSlingers(dHOH)
    
    
    '--------------------------------------------------------------------
    'OPVRAGEN OSNAP-INSTELLING EN OSNAP INSTELLEN
    '--------------------------------------------------------------------
    
    Dim iOsnap As Integer
    iOsnap = ThisDrawing.GetVariable("OSMODE")
    
    'AANZETTEN OSNAP OP NEAREST + ENDPOINT
    ThisDrawing.ObjectSnapMode = True
    ThisDrawing.SetVariable "OSMODE", 513      ' NEAREST + ENDPOINT
    'ThisDrawing.SetVariable "OSMODE", 512        ' NEAREST toch fout 18 juni


    'TERUGSETTEN OSNAP
    'ThisDrawing.SetVariable "OSMODE", iOsnap
    '--------------------------------------------------------------------
    
    Dim p1 As Variant
    ' Dim GetentObj As AcadEntity
    Dim BlockRefObj As AcadBlockReference
    Dim sLaagNaam As String
    'Dim dHOH As Double 'aangeroepen vanuit subroutine
    Dim dRotatie As Double
    Dim sInspoint As String
    
    ' Me.Hide
    ' F_Main.Hide
    
    '--------------------------------------------------------------------
    'OPVRAGEN ELEMENT (IVM LAAGNAAM) ROND OPGEGEVEN PUNT
    '--------------------------------------------------------------------
    
    Dim sHandle As String
    Dim GevondenObj As AcadEntity
    
    ' ThisDrawing.Utility.GetEntity GetentObj, P1, "Selecteer beginpunt"        ' werkt niet mooi want geeft selectie-pickbox
    p1 = ThisDrawing.Utility.GetPoint(, "Selecteer beginpunt")
    sHandle = ZoekElementHandleRondPunt(p1, 0.1)
    
    If sHandle = "" Then
        MsgBox "Juiste laag kan niet worden bepaald.", vbInformation, "Geen element rond snappunt gevonden"
        sLaagNaam = ThisDrawing.ActiveLayer.Name
    Else
        Set GevondenObj = ThisDrawing.HandleToObject(sHandle)
        sLaagNaam = GevondenObj.Layer
    End If
            
     
    
            
    ' ------------------------------------
    ' sLaagNaam = GetentObj.Layer
    ' ThisDrawing.ActiveLayer.Name
    ' ------------------------------------
    
    
    'BEPALEN VAN DE KWADRANTEN
    Dim bXrichting As Boolean
    Dim bPositief As Boolean
    Dim bSpiegelen As Boolean
    
'''''    If Abs(P3(0) - P1(0)) > Abs(P3(1) - P1(1)) Then
'''''        bXrichting = True
'''''
'''''        If P3(0) - P1(0) > 0 Then
'''''            bPositief = True
'''''        Else
'''''            bPositief = False
'''''        End If
'''''    Else
'''''        bXrichting = False
'''''
'''''        If P3(1) - P1(1) > 0 Then
'''''            bPositief = True
'''''        Else
'''''            bPositief = False
'''''        End If
'''''    End If
    
'''''   MsgBox "bXrichting = " & bXrichting _
        & Chr(10) & Chr(13) & "bPositief = " & bPositief _
        & Chr(10) & Chr(13) & "dRotatie = " & dRotatie
   

    
    Dim p2 As Variant
    p2 = ThisDrawing.Utility.GetPoint(p1, "Selecteer eindpunt")
    
    'dHOH = 20
    dRotatie = ThisDrawing.Utility.AngleFromXAxis(p1, p2)
    
    '--------------------------------------------------------------------------------------
    'TERUGSETTEN OSNAP
    ThisDrawing.SetVariable "OSMODE", iOsnap
    '--------------------------------------------------------------------------------------
     
    'eerst inserten om te controleren of deze bestaat (wblock), IS NIET ECHT NODIG
    Dim Pins(0 To 2) As Double
    Pins(0) = 0
    Pins(1) = 0
    Pins(2) = 0
    InsertBlock "slinger-block", Pins, "0"
    Dim TestBlock As Object
    Set TestBlock = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)
    If TestBlock.EntityName = "AcDbBlockReference" Then TestBlock.Delete Else MsgBox "Geen slinger-block gevonden": End
    
    'LET OP, OMDAT X, Y EN Z-SCALE NIET GELIJKS ZIJN ZULLEN DE ARC'S NA HET EXPLODEREN
    'VAN HET BLOCK VANZELF ELLIPSEN WORDEN.
    'DUS Z Y EN Z SCALE MOETEN GELIJK ZIJN. DOOR Z-SCALE GELIJK AAN DHOH TE MAKEN.
    'EVENTUEEL KAN OOK VERSCHALEN MET BlockRefObj.ScaleEntity P1, dHOH
    
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(p1, "slinger-block", dHOH, dHOH, dHOH, dRotatie)
    
    BlockRefObj.Layer = sLaagNaam
    BlockRefObj.Update
    
    
        
    
    
     
    
    '--------------------------------------------------------------------------------------
    'MIRROR BLOCKREFERENCE (uitgezet)
    '--------------------------------------------------------------------------------------
    
    ZoomWindow p1, p2
    ' ZoomScaled 0.3, acZoomScaledRelative      voor 14 juni 2004 aanepast
    ZoomScaled 0.6, acZoomScaledRelative
    
'''    ' OUD VOOR 14 JUNI 2004
'''    Dim antw As String
'''    antw = MsgBox("leiding-slinger spiegelen ?", vbQuestion + vbYesNo, "Spiegelen")
'''    If antw = vbYes Then bSpiegelen = True
    
    
    ' IN COMMANDREGEL OPGEVEN JA OF NEE:
    Dim returnString As String
    returnString = ThisDrawing.Utility.GetString(True, "Spiegelen ?: Ja <Nee>")
    returnString = UCase(returnString)
    
    If returnString = "" Or UCase(Left(returnString, 1)) = "N" Then
        bSpiegelen = False
    Else
        bSpiegelen = True
    End If

    
    
    
    
    
    
    
'    If bSpiegelen = True Then
'        Dim MirrorObj As AcadBlockReference
'        Set MirrorObj = BlockRefObj.Mirror(P1, P2)
'        BlockRefObj.Delete
'        Set BlockRefObj = MirrorObj
'    End If
'
    ' VOLGENDE MANIER GEEFT EEN PREVIEW VAN BLOCKINSERTION:
    
    ' sInspoint = P1(0) & "," & P1(1) & "," & P1(2)
    ' ThisDrawing.SendCommand ("_-insert" & vbCr & "arcs" & vbCr & "Scale" & vbCr & dHoh & vbCr & sInspoint & vbCr)
    
'
'    If ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1).EntityName <> "AcDbBlockReference" Then
'        MsgBox "Het laatst geplaatste element is geen blockreference", vbCritical
'        Exit Sub
'    Else
'        Set BlockRefObj = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)
'    End If

    ' ---------------------------------------------------------------------
    ' EXPLODEREN BLOCK
    ' ---------------------------------------------------------------------


    Dim explodedObjects As Variant
    explodedObjects = BlockRefObj.Explode
    BlockRefObj.Delete
    
    Dim I As Integer
    For I = 0 To UBound(explodedObjects)
        'explodedObjects(I).Color = acRed
        explodedObjects(I).Layer = sLaagNaam
        explodedObjects(I).Update
        ' MsgBox explodedObjects(I).EntityName
    Next
    
    ' OBJECTEN IN DE BLOCKREFERENCE ZIJN GENUMMER VANAF HET GESELECTEERDE BEGINPUNT:
    
    ' BIJ INSERTEN VAN ARC EN EXPLODEREN WORDT EEN ARC EEN ELLIPSE (SCALE)
    Dim Obj1Arc As AcadArc
    Dim Obj2VertLijn As AcadLine
    Dim Obj3Arc As AcadArc
    Dim Obj4HorLijn As AcadLine
    Dim Obj5Arc As AcadArc
    Dim Obj6HorLijn As AcadLine
    Dim Obj7Arc As AcadArc
    Dim Obj8HorLijn As AcadLine
    Dim Obj9HorBeginLijn As AcadLine
    
        
    On Error Resume Next
        
    Set Obj1Arc = explodedObjects(0)
    Set Obj2VertLijn = explodedObjects(1)
    Set Obj3Arc = explodedObjects(2)
    Set Obj4HorLijn = explodedObjects(3)
    Set Obj5Arc = explodedObjects(4)
    Set Obj6HorLijn = explodedObjects(5)
    Set Obj7Arc = explodedObjects(6)
    Set Obj8HorLijn = explodedObjects(7)
    Set Obj9HorBeginLijn = explodedObjects(8)
    
    If Err Then
        Err.Clear
        MsgBox "Volgorde van entiteiten in block ARCS is niet juist.", vbCritical, "Let op"
        Exit Sub
    End If
        
    
    
    ' ---------------------------------------------------------------------
    ' STRETCHEN VAN DE LINES
    ' ---------------------------------------------------------------------
   
    Obj4HorLijn.Delete
    Obj6HorLijn.Delete
    
    'Obj8HorLijn.Color = acGreen
    Dim PeindObj8HorLijn As Variant
    PeindObj8HorLijn = Obj8HorLijn.EndPoint
    PeindObj8HorLijn(0) = p2(0)
    PeindObj8HorLijn(1) = p2(1)
    Obj8HorLijn.EndPoint = PeindObj8HorLijn
    
    
    
    Set Obj6HorLijn = Obj8HorLijn.Copy
    'Obj6HorLijn.Color = acBlue
    Obj6HorLijn.Move Obj6HorLijn.StartPoint, Obj7Arc.EndPoint
    
    
    
    Set Obj4HorLijn = Obj8HorLijn.Copy
    'Obj4HorLijn.Color = acYellow
    Obj4HorLijn.Move Obj4HorLijn.StartPoint, Obj3Arc.EndPoint
    '*** 6 aug: nu test op startpoint
    'Obj4HorLijn.Move Obj4HorLijn.StartPoint, Obj3Arc.StartPoint
    
    
    ' ---------------------------------------------------------------------
    ' VERPLAATSEN ARC
    ' ---------------------------------------------------------------------
    
    
    Obj5Arc.Move Obj5Arc.EndPoint, Obj6HorLijn.EndPoint
    '*** 6 aug: nu test op startpoint
    'Obj5Arc.Move Obj5Arc.EndPoint, Obj4HorLijn.EndPoint
    'Obj5Arc.Color = acRed
        
    
    
    'Obj4HorLijn.Color = acMagenta
    Obj4HorLijn.EndPoint = Obj5Arc.StartPoint
   
    
    'ThisDrawing.ActiveViewport.OrthoOn
    'ThisDrawing.ObjectSnapMode
    'ThisDrawing.ActiveViewport.OrthoOn
    
    '-------------------------------------------------------------------------
    'MIRROR DE AFZONDERLIJKE ELEMENTEN 6 AUG 2003
    '-------------------------------------------------------------------------
    
    If bSpiegelen = True Then
        Dim MirrorObj As AcadEntity
        Set MirrorObj = Obj1Arc.Mirror(p1, p2)
        Obj1Arc.Delete
        Set MirrorObj = Obj2VertLijn.Mirror(p1, p2)
        Obj2VertLijn.Delete
        Set MirrorObj = Obj3Arc.Mirror(p1, p2)
        Obj3Arc.Delete
        Set MirrorObj = Obj4HorLijn.Mirror(p1, p2)
        Obj4HorLijn.Delete
        Set MirrorObj = Obj5Arc.Mirror(p1, p2)
        Obj5Arc.Delete
        Set MirrorObj = Obj6HorLijn.Mirror(p1, p2)
        Obj6HorLijn.Delete
        Set MirrorObj = Obj7Arc.Mirror(p1, p2)
        Obj7Arc.Delete
        Set MirrorObj = Obj8HorLijn.Mirror(p1, p2)
        Obj8HorLijn.Delete
    End If
            
    
    '--------------------------------------------------------------------
    'EINDPUNT VAN GEVONDEN LIJN TRIMMEN (6 aug 2003)
    '--------------------------------------------------------------------
        
    'BEPALEN AAN WELKE KANT DE GESELECTEERDE LIJN GETRIMED MOET GAAN WORDEN
    
    
    If sHandle <> "" And GevondenObj.EntityName = "AcDbLine" Then
        If Lengte(p2, GevondenObj.StartPoint) > Lengte(p2, GevondenObj.EndPoint) Then
            GevondenObj.EndPoint = Obj9HorBeginLijn.EndPoint
        Else
            GevondenObj.StartPoint = Obj9HorBeginLijn.EndPoint
        End If
        
        Obj9HorBeginLijn.Erase
    Else
        MsgBox "Het geselecteerde element is geen lijn: element kan niet worden afgekort.", vbInformation, "Letop"
    End If
            
    
    
End Sub

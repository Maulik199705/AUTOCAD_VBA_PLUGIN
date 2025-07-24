Attribute VB_Name = "M_Meten_Flexfix"
Sub Meten_Flexfix(RetObj, Pbase, eg)
    'Dim retobj As Object
    'Dim Pbase As Variant
    Dim element As Object
    Dim Lengte As Double
    Dim Lengte3 As Double
    Dim totallengte As String       'EBR, DEC 2002
    'Dim eg As Double         'EBR, DEC 2002: declareren van variabelen
    Dim optel As Integer            'EBR, DEC 2002: declareren van variabelen
    Dim unittel As Integer          'EBR, DEC 2002: declareren van variabelen
    Dim groepsnummer As String      'EBR, DEC 2002: declareren van variabelen
    check_layernaam = Left(RetObj.Layer, 4)
    'MsgBox check_layernaam
      
    frmGroeptekst.Hide
    'ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "GT" & vbCr & vbCr
    Update
    
    'On Error Resume Next           'EBR, DEC 2002 uitgezet !
    
   '----------------------------------------------------------------
    ' BEPALEN MIDDELPUNT OF CENTERPUNT BEPALEN
    '----------------------------------------------------------------
    
    If RetObj.EntityName = "AcDbLine" Then
        Dim PBEGIN As Variant
        Dim Peind As Variant
        Dim Pmid(0 To 2) As Double
        'Dim Pmid As Variant
        
        PBEGIN = RetObj.startPoint
        Peind = RetObj.endPoint
        
        Pmid(0) = (PBEGIN(0) + Peind(0)) / 2
        Pmid(1) = (PBEGIN(1) + Peind(1)) / 2
        Pmid(2) = 0
        'Dim PUNT As AcadPoint
        'Set PUNT = ThisDrawing.ModelSpace.AddPoint(Pmid)
        'PUNT.Color = acGreen
    Else
        
        Pmid(0) = Pbase(0)
        Pmid(1) = Pbase(1)
        Pmid(2) = 0
        
    End If
    
    '----------------------------------------------------------------
    ' MAKEN SELECTIESET
    '----------------------------------------------------------------
    
        
    Dim ssetObj As AcadSelectionSet
    Dim mode As Integer
    Dim pointsArray(0 To 14) As Double
    
    'NIEUW
    'On Error Resume Next
    Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
    If Err Then ThisDrawing.SelectionSets.Item("TEST_SSET").Delete
    mode = acSelectionSetCrossingPolygon
    
    Dim tel As Integer
    Dim sHandle As String
    Dim LijnGevonden As Boolean
    Dim ArcGevonden As Boolean
    
    
Opnieuw:
    
    tel = tel + 1
    
    pointsArray(0) = Pmid(0) - tel: pointsArray(1) = Pmid(1) - tel: pointsArray(2) = 0
    pointsArray(3) = Pmid(0) + tel: pointsArray(4) = Pmid(1) - tel: pointsArray(5) = 0
    pointsArray(6) = Pmid(0) + tel: pointsArray(7) = Pmid(1) + tel: pointsArray(8) = 0
    pointsArray(9) = Pmid(0) - tel: pointsArray(10) = Pmid(1) + tel: pointsArray(11) = 0
    pointsArray(12) = Pmid(0) - tel: pointsArray(13) = Pmid(1) - tel: pointsArray(14) = 0
    
    ssetObj.SelectByPolygon mode, pointsArray
   
    '----------------------------------------------------------------
    ' LEZEN SELECTIESET
    '----------------------------------------------------------------
    
    If tel > 120 Then
        ssetObj.Clear
        ssetObj.Delete
        MsgBox "Geen juiste HOH gevonden.!!" & (Chr(13) & Chr(10)) & "Ik zoom nu uit," & Chr(13) & "wijs daarna de groep weer aan.", vbInformation
        tel = 0
        ZoomPrevious
        GoTo Opnieuw
        'End
    End If

     If ssetObj.Count > 1 Then
        
            For Each element In ssetObj
            If RetObj.Handle <> element.Handle Then
                If element.Layer = RetObj.Layer Then
                        If element.EntityName = "AcDbLine" And RetObj.EntityName = "AcDbLine" Then
                            sHandle = element.Handle
                            LijnGevonden = True
                            Exit For
                        End If
                        If element.EntityName = "AcDbArc" And RetObj.EntityName = "AcDbArc" Then
                            sHandle = element.Handle
                            ArcGevonden = True
                            Exit For
                        End If
                End If
            End If
            Next element
            
            If LijnGevonden = False And ArcGevonden = False Then GoTo Opnieuw
    Else
            GoTo Opnieuw
    End If
    
     lang = element.Length 't.b.v. flexfix

    
    ssetObj.Clear
    ssetObj.Delete
    
    '----------------------------------------------------------------
    ' SNIJPUNT MET LIJN BEREKENEN EN BEPALEN HOH
    '----------------------------------------------------------------
    Dim HOH As Double
       
    If RetObj.EntityName = "AcDbLine" Then
    
            Dim LijnObj As AcadLine
            Dim Peind1(0 To 2) As Double
            
            Peind1(0) = Pmid(0) + 35
            Peind1(1) = Pmid(1)
            Peind1(2) = 0
            
        
            Set LijnObj = ThisDrawing.ModelSpace.AddLine(Pmid, Peind1)
            LijnObj.Rotate Pmid, RetObj.Angle + 1.57075
                          
            Dim Ps As Variant   'snijpunt
            Dim IntersectObject As Object
            Set IntersectObject = ThisDrawing.HandleToObject(sHandle)
            
            Ps = LijnObj.IntersectWith(IntersectObject, acExtendBoth)
            
            '-------------------------------------
            'NIEUW
            '-------------------------------------
            ' weergeven van alle snijpunten
            Dim I As Integer, j As Integer, k As Integer
            Dim str As String
            If VarType(Ps) <> vbEmpty Then
                For I = LBound(Ps) To UBound(Ps)
            'uitgezet GCH  MsgBox "x=" & Round(Ps(j), 0) & Chr(10) & Chr(13) & "y=" & Round(Ps(j + 1), 0) & Chr(10) & Chr(13) & "z=" & Round(Ps(j + 2), 0), , "Snijpunt"
                    I = I + 2
                    j = j + 3
                    'k = k + 1
                Next
            End If
            
            '-------------------------------------
            LijnObj.Erase
             Dim puntobj As Object
            'uitgezet GCH  Set puntobj = ThisDrawing.ModelSpace.AddPoint(Ps)
            'uitgezet GCH  Set puntobj = ThisDrawing.ModelSpace.AddPoint(Pmid)
            ' BEPALEN HOH
            HOH = 0
            HOH = Lengte2(Ps, Pmid)
            
            
    Else
            
            'Dim puntobj2 As Object
            'puntobj2 = acMagenta
     'Set puntobj2 = ThisDrawing.ModelSpace.AddPoint(Peind1)
            
    '----------------------------------------------------------------
    ' STRAAL BEPALEN EN BEPALEN HOH
    '----------------------------------------------------------------
        
        Dim GevondenArcObj As AcadArc
        Dim ArcRadius As Double
        
        Set GevondenArcObj = ThisDrawing.HandleToObject(sHandle)
        ArcRadius = GevondenArcObj.radius
        HOH = Abs(RetObj.radius - ArcRadius)
    End If

    '----------------------------------------------------------------
    ' AFBEELDEN HOH
    '----------------------------------------------------------------
   
    HOH = Format(HOH, "0.0")
    'MsgBox HOH
    'frmGroeptekst.TextBox8 = "HOH " & HOH
    
    'EBR, DEC 2002: TOT HIER WERKT HET NU GOED !!!
    'MsgBox "EBR: v.a. hier ging het fout."
    'Exit Sub
   lang2 = lang + HOH  'FLEXFIX BREEDTE
   lang2 = Round(lang2, 0) 'FLEXFIX BREEDTE

    '-------------------
    'METEN VAN DE GROEP
    '-------------------
    If Err <> 0 Then
        Err.Clear
        MsgBox "Geen element geselecteerd", vbCritical, "Let op."
    Else
        For Each element In ThisDrawing.ModelSpace
             If element.Layer = RetObj.Layer Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                     If element.EntityName = "AcDbArc" Then
                     Lengte = Lengte + element.ArcLength
                     arctel = arctel + 1
                     End If
              End If
             Next element
             
          faanvoer = RetObj.Layer & "_Flexfix_aanvoer"
          fretour = RetObj.Layer & "_Flexfix_retour"
          
          'als de layers hernoemd zijn en er tijdelijk een h achter de layer staat
          myaanvoer = Right(RetObj.Layer, 1)
          'myaanvoerlengte = Len(retobj.Layer)
          If myaanvoer = "h" Then
            myaanvoerleft = Split(RetObj.Layer, "_")
            faanvoer = myaanvoerleft(0) & "_Flexfix_aanvoerh"
            fretour = myaanvoerleft(0) & "_Flexfix_retourh"
            
            'myaanvoerlengte = myaanvoerlengte - 1
            'myaanvoerleft = Left(retobj.Layer, myaanvoerlengte)
            'faanvoer = myaanvoerleft & "_aanvoer"
            'fretour = myaanvoerleft & "_retour"
            End If
           'als de layers hernoemd zijn en er tijdelijk een h achter de layer staat
           
         'faanvoerh = retobj.Layer & "_Flexfix_aanvoerh"
   
         For Each elementaanvoer In ThisDrawing.ModelSpace
             If elementaanvoer.Layer = faanvoer Then
                'BEREKENEN TOTALE Lengteaanvoer
                If elementaanvoer.EntityName = "AcDbLine" Then lengteaanvoer = lengteaanvoer + elementaanvoer.Length
                     If elementaanvoer.EntityName = "AcDbArc" Then
                     lengteaanvoer = lengteaanvoer + elementaanvoer.ArcLength
                     End If
              End If
             Next elementaanvoer
             
           
           'fretourh = retobj.Layer & "_Flexfix_retourh"
          For Each elementretour In ThisDrawing.ModelSpace
             If elementretour.Layer = fretour Then
                'BEREKENEN TOTALE Lengteretour
                If elementretour.EntityName = "AcDbLine" Then lengteretour = lengteretour + elementretour.Length
                     If elementretour.EntityName = "AcDbArc" Then
                     lengteretour = lengteretour + elementretour.ArcLength
                     End If
              End If
             Next elementretour
        End If
    
    'LENGTE IN METERS
    reserve = 3
    reserve = reserve / 2
    
    If frmGroeptekst.CheckBox7 = True Then
      reserve = (Val(frmGroeptekst.TextBox24))
      reserve = reserve / 2
    End If
    'MsgBox reserve
    
    lengteaanvoer = Round((lengteaanvoer / 100), 1) + reserve
    lengteaanvoer = lengteaanvoer + (Val(frmGroeptekst.TextBox28) / 2)
    lengteretour = Round((lengteretour / 100), 1) + reserve
    lengteretour = lengteretour + (Val(frmGroeptekst.TextBox28) / 2)
    
    Lengte = Lengte / 100
    Lengte = Lengte + lengteaanvoer + lengteretour
    Lengte = Round(Lengte, 1)
    'Lengte = Lengte
    
    lang3 = (lang2 / 100)
    lang4 = lang3 * 2
    
    slingeraanvoer1 = lengteaanvoer / lang4 'lang2 = breedte flexfix
    slingeraanvoer2 = Fix(slingeraanvoer1 + 1)
    'MsgBox slingeraanvoer & "-" & lengteaanvoer & "-" & lang2
    slingerretour1 = lengteretour / lang4 'lang2 = breedte flexfix
    slingerretour2 = Fix(slingerretour1 + 1)
    
    '-------------------
    'EINDE METEN VAN DE GROEP
    '-------------------
    
    'MsgBox "Lengte=" & Lengte, , "Ebr: tot hier werkt het goed !"
    'EBR, DEC 2002: TOT HIER WERKT HET NU GOED !!!
    ' Exit Sub
      '---------------------------------
    '[meten flexfix aanvoer en retour]
    '---------------------------------
       
    zz = frmGroeptekst.Label11.Caption & "_Flexfix"
    
    If zz = RetObj.Layer Then
    optel = frmGroeptekst.TextBox10        'textbox10 uitlezen
    totallengte = " "
    ffoptie = 1
    Else
    'MsgBox "eg= " & eg
    If eg = 0 Then optel = frmGroeptekst.TextBox10 'uitlezen 1e keer
    If eg <> 0 Then optel = frmGroeptekst.TextBox10 + 1 '1 erbij optellen
    totallengte = Lengte & " meter"
    End If
    If optel > 0 And optel < 10 Then groeponder10 = "0"
    
    unittel = frmGroeptekst.TextBox9
       
    
       
       
    If frmGroeptekst.CheckBox3.Value = False Then
    If unittel > 0 And unittel < 10 Then unitonder10 = "0"
    groepsnummer = "groep " & unitonder10 & frmGroeptekst.TextBox9 & "." & groeponder10 & optel    'tekst samenvoegen
    End If
    If frmGroeptekst.CheckBox3.Value = True Then
    groepsnummer = "groep " & frmGroeptekst.TextBox9 & "." & groeponder10 & optel    'tekst samenvoegen
    End If
   
       
    zz = groepsnummer
    'frmGroeptekst.Label11.Caption = Clear
    frmGroeptekst.Label11.Caption = zz
    
    eg = 1 'EERSTE GROEP
    frmGroeptekst.TextBox10 = optel
    'MsgBox optel
    
    'groepslengte op de commandregel
    zzz = groepsnummer & " = " & Lengte & " meter." & " [" & Val(frmGroeptekst.TextBox28) & "]"
    ThisDrawing.Utility.Prompt zzz
    ThisDrawing.SendCommand Chr(27)
    
    
     If frmGroeptekst.CheckBox6.Value = True Then '--- flexfix matnummer
        unittel = frmGroeptekst.TextBox10
         If unittel > 0 And unittel < 10 Then
         unitonder11 = "0"
         mm = unitonder11 & frmGroeptekst.TextBox10 'tekst samenvoegen
         Else
         mm = frmGroeptekst.TextBox10 'tekst samenvoegen
         End If
     End If
       
    '-------------------------------------------------
    '  groeptekst plaatsen in de tekening
    '-------------------------------------------------
    Dim newLayer As AcadLayer
    
    
    Set newLayer = ThisDrawing.Layers.Add("GT")
    ThisDrawing.ActiveLayer = newLayer
    Update
        
    Dim hohafstand As String
    hohafstand = "H.O.H. " & HOH & " cm."
    If check_layernaam = "wand" Then hohafstand = "Wandverwarming"
    On Error Resume Next
    
    Dim Zoekpad As String
    Dim blockRefObj As Object       'EBR: var declareren
    
    ' wth
    Zoekpad = "C:\ACAD2002\DWG\groeptekstbloknew.dwg"
    
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(Pbase, Zoekpad, 1, 1, 1, 0)
    
    'DOOR EBR DEC 2002:
    If Err Then
        Err.Clear
        MsgBox "symbool " & Zoekpad & " niet gevonden.", vbCritical, "Let op":
        End
    End If
    
'DOOR EBR DEC 2002: enkele aanpassingen

    Dim symbool As Object       'EBR altijd juist declareren
    Dim attributen As Variant   'EBR altijd juist declareren
    Dim attribuut As Object     'EBR altijd juist declareren
    If frmGroeptekst.OptionButton1 = True Then LW = "WTH-ZD 20*3,4 mm"
    If frmGroeptekst.OptionButton9 = True Then LW = "WTH-ZD 16*2,7 mm"
    If frmGroeptekst.OptionButton2 = True Then LW = frmGroeptekst.ComboBox4
    If frmGroeptekst.OptionButton6 = True Then LW = frmGroeptekst.ComboBox4
   If frmGroeptekst.OptionButton7 = True Then
      RL = "RM"
      Else
      RL = " "
    End If
    If frmGroeptekst.OptionButton8 = True Then
              groepsnummer = " "
              RL = "RZ"
    End If
    
    arctel2 = (arctel / 2)
    arctel3 = Round(arctel2, 1)
    arctel3 = arctel3 / 2  '7-6

    lengtearc = Len(arctel3)
    If lengtearc > 2 Then arctel3 = arctel3 + 0.5
    arctel3 = Fix(arctel3)
        'MsgBox arctel3
   
    'If CheckBox6.Value = True Then FF = "ja"
'    If ffoptie = 1 Then
'        LANG2 = "-" 'breedte flexfix
'        arctel3 = "-" 'FLEXFIX AANTAL BOCHTEN
'        lengteaanvoer = "-"
'        lengteretour = "-"
'    End If
     
    
    
    
    For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOKNEW" Then
                Set symbool = element
                If symbool.HasAttributes Then
                    attributen = symbool.GetAttributes
                    For I = LBound(attributen) To UBound(attributen)
                    Set attribuut = attributen(I)
                         If attribuut.TagString = "RINGLEIDING" And attribuut.textstring = "" Then attribuut.textstring = RL
                         If attribuut.TagString = "LEIDINGSOORT" And attribuut.textstring = "" Then attribuut.textstring = LW
                         If attribuut.TagString = "UNITNUMMER" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 & frmGroeptekst.TextBox9
                         If attribuut.TagString = "GROEPTEKST" And attribuut.textstring = "" Then attribuut.textstring = groepsnummer
                         If attribuut.TagString = "HOHAFSTAND" And attribuut.textstring = "" Then attribuut.textstring = hohafstand
                         If attribuut.TagString = "ROLLENGTE" And attribuut.textstring = "" Then attribuut.textstring = " "
                         If attribuut.TagString = "TLVL" And attribuut.textstring = "" Then attribuut.textstring = totallengte
                         
                            
                           If frmGroeptekst.CheckBox6.Value = True And ffoptie <> 1 Then
                            For z = LBound(attributen) To UBound(attributen)
                            Set attribuut = attributen(z)
                             If attribuut.TagString = "GROEPTEKST" Then gp = attribuut.textstring
                              If gp = groepsnummer Then
                              If attribuut.TagString = "FLEXFIX" Then attribuut.textstring = "ja" 'FLEXFIX
                              If attribuut.TagString = "AANTAL_SLINGERS_FLEXFIX" Then attribuut.textstring = arctel3   'FLEXFIX AANTAL BOCHTEN
                              If attribuut.TagString = "FLEXFIX_AANVOER" Then attribuut.textstring = lengteaanvoer  'aanvoer flexfix groep
                              If attribuut.TagString = "FLEXFIX_RETOUR" Then attribuut.textstring = lengteretour    'retour flexfix groep
                              If attribuut.TagString = "MATNUMMER_FLEXFIX" Then attribuut.textstring = mm
                              If attribuut.TagString = "BREEDTE_FLEXFIX" Then attribuut.textstring = lang2   'FLEXFIX BREEDTE
                              If attribuut.TagString = "SL_AANVOER" Then attribuut.textstring = slingeraanvoer2 'aantal slingers aanvoer
                              If attribuut.TagString = "SL_RETOUR" Then attribuut.textstring = slingerretour2  'aantal slingers retour
                              End If
                             Next z
                            End If
                    Next I
                End If
            End If
        End If
    Next element

'   Call ThisDrawing.mflex(groepsnummer)

  '  Dim layerObj As AcadLayer
  '  If frmGroeptekst.CheckBox6 = False Then
  '  For Each layerObj In ThisDrawing.Layers
  '       If layerObj.Name = retobj.Layer Then layerObj.Name = groepsnummer
  '  Next
  '  End If
    
    If frmGroeptekst.CheckBox6 = True Then
    groepsnummer1 = groepsnummer & "_Flexfix"
    groepsnummer2 = groepsnummer & "_Flexfix_aanvoer"
    groepsnummer3 = groepsnummer & "_Flexfix_retour"
    'MsgBox groepsnummer1 & " - " & groepsnummer2 & " - " & groepsnummer3
    
    check20 = RetObj.Layer & "_Flexfix_aanvoer"
    check21 = RetObj.Layer & "_Flexfix_retour"
    check22 = myaanvoerleft(0) & "_Flexfix_aanvoerh"
    check23 = myaanvoerleft(0) & "_Flexfix_retourh"
   ' MsgBox check22 & "   -    " & check23
    
    For Each layerObj In ThisDrawing.Layers
         If layerObj.Name = RetObj.Layer Then layerObj.Name = groepsnummer1
         If layerObj.Name = check20 Then layerObj.Name = groepsnummer2
         If layerObj.Name = check21 Then layerObj.Name = groepsnummer3
         If layerObj.Name = check22 Then layerObj.Name = groepsnummer2
         If layerObj.Name = check23 Then layerObj.Name = groepsnummer3
    Next
    End If
   ' MsgBox "hier"
  
   
    Update
    End Sub
Public Function Lengte2(p1 As Variant, p2 As Variant) As Double

    'FUNCTIE VOOR HET BEREKENEN VAN DE AFSTAND TUSSEN TWEE PUNTEN
    
    Dim DX As Double
    Dim DY As Double
    
    DX = Abs(p2(0) - p1(0))
    DY = Abs(p2(1) - p1(1))

    Lengte2 = Sqr((DX ^ 2) + (DY ^ 2))
End Function



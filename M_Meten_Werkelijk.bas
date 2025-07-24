Attribute VB_Name = "M_Meten_Werkelijk"
Sub Meten_Werkelijk(RetObj, Pbase, eg)
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
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
              End If
        Next element
    End If
             z = 0
             Dim cirkel As Object
             For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.Layer = RetObj.Layer Then
                If cirkel.EntityName = "AcDbCircle" Then z = z + 1
                End If
                Next cirkel
                
       If frmGroeptekst.ToggleButton1.Value = False Then wandhoogte = 2.5
       If frmGroeptekst.ToggleButton1.Value = True Then wandhoogte = 2
       If frmGroeptekst.OptionButton5 = True Then wandhoogte = frmGroeptekst.TextBox13
                
                If z <> 0 Then zlengte = (z * (100 * wandhoogte)) + 100
    
    Lengte = Lengte + zlengte
  
   'LENGTE IN METERS
    Lengte = Lengte / 100
    Lengte = Round(Lengte, 1)
    Lengte = Lengte + Val(frmGroeptekst.TextBox28) ' + 3 ' 07-04-06    <--------------------------------------------------------------------------
     
    'MsgBox "Lengte=" & Lengte, , "Ebr: tot hier werkt het goed !"
    'EBR, DEC 2002: TOT HIER WERKT HET NU GOED !!!
    ' Exit Sub
    
     
    
    zz = frmGroeptekst.Label11.Caption
    'MsgBox zz & " | " & retobj.Layer
    If zz = RetObj.Layer Then
    optel = frmGroeptekst.TextBox10        'textbox10 uitlezen
    totallengte = " "
    leegwerklengte = "ja"
    Else
    'MsgBox "eg= " & eg
    If eg = 0 Then optel = frmGroeptekst.TextBox10 'uitlezen 1e keer
    If eg <> 0 Then optel = frmGroeptekst.TextBox10 + 1 '1 erbij optellen
    totallengte = Lengte & " meter"
    leegwerklengte = "ja"
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
    'groepsnummer = "groep " & frmGroeptekst.TextBox9 & "." & groeponder10 & optel  'tekst samenvoegen
    
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
    
    ' ebr
    ' Zoekpad = "C:\temp\groeptekstbloknew.dwg"
    
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
                         If attribuut.TagString = "WERKLENGTE" And attribuut.textstring = "" Then attribuut.textstring = leegwerklengte
                         If attribuut.TagString = "WANDHOOGTE" And check_layernaam = "groe" And attribuut.textstring = "" Then attribuut.textstring = " "
                         If attribuut.TagString = "WANDHOOGTE" And check_layernaam = "wand" And attribuut.textstring = "" Then attribuut.textstring = wandhoogte & " meter hoog"
                    
                    Next I
                End If
            End If
        End If
    Next element
    
    '---layer hernoemen
    'If optel > 0 And optel < 10 Then groeponder10 = "0"
    'groepsnummer = "groep " & frmGroeptekst.TextBox9 & "." & groeponder10 & optel  'tekst samenvoegen
           
    Dim layerObj As AcadLayer
    For Each layerObj In ThisDrawing.Layers
         If layerObj.Name = RetObj.Layer Then layerObj.Name = groepsnummer
    Next
    
    
    End Sub
Public Function Lengte2(p1 As Variant, p2 As Variant) As Double

    'FUNCTIE VOOR HET BEREKENEN VAN DE AFSTAND TUSSEN TWEE PUNTEN
    
    Dim DX As Double
    Dim DY As Double
    
    DX = Abs(p2(0) - p1(0))
    DY = Abs(p2(1) - p1(1))

    Lengte2 = Sqr((DX ^ 2) + (DY ^ 2))
End Function


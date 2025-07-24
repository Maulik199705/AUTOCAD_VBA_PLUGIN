Attribute VB_Name = "Update_module"
Sub update_unitlogo()
On Error Resume Next
If frmGroeptekst.ComboBox2 = "" Then
    MsgBox "Je bent 'Type unit' vergeten in te vullen..!!"
    frmGroeptekst.ComboBox2.SetFocus
    Exit Sub
    End If
If frmGroeptekst.ComboBox1 = "" Then
    MsgBox "Je bent 'Bevestigingsmateriaal' vergeten in te vullen..!!"
    frmGroeptekst.ComboBox1.SetFocus
    Exit Sub
    End If
If frmGroeptekst.ComboBox4 = "" And frmGroeptekst.OptionButton6 = True Then
    MsgBox "Je bent 'Type buis' vergeten in te vullen..!!"
    frmGroeptekst.ComboBox4.SetFocus
    Exit Sub
    End If
If frmGroeptekst.ComboBox4 = "" And frmGroeptekst.OptionButton2 = True Then
    MsgBox "Je bent 'Type buis' vergeten in te vullen..!!"
    frmGroeptekst.ComboBox4.SetFocus
    Exit Sub
    End If
frmGroeptekst.Hide
frmGroeptekst.TextBox9.Locked = False
    Set newLayer = ThisDrawing.Layers.Add("BLOKLOGO")
    ThisDrawing.ActiveLayer = newLayer
    Update
    ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "*" & vbCr & vbCr
    Update
Call frmGroeptekst.Schaal(scaal)
If frmGroeptekst.OptionButton1.Value = True And (frmGroeptekst.TextBox11 = "0" And frmGroeptekst.TextBox25 <> "0") Then
bestand = "c:\acad2002\dwg\Mat_spe_FLEX.dwg"
'MsgBox bestand
Else
bestand = "c:\acad2002\dwg\Mat_spe_ZD.dwg"
End If


unittel = frmGroeptekst.TextBox9

If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If


For Each element10 In ThisDrawing.ModelSpace
      If element10.ObjectName = "AcDbBlockReference" Then
       If UCase(element10.Name) = "HERZ" Or UCase(element10.Name) = "KMV" Or UCase(element10.Name) = "RUH-N" _
       Or UCase(element10.Name) = "RUH-R" Or UCase(element10.Name) = "RUH-RT" Or UCase(element10.Name) = "RUB-R" _
       Or UCase(element10.Name) = "RUB-RT" Or UCase(element10.Name) = "RUBK-R" Or UCase(element10.Name) = "RUBK-RT" _
       Or UCase(element10.Name) = "LT" Or UCase(element10.Name) = "LT-N" Or UCase(element10.Name) = "LTS" _
       Or UCase(element10.Name) = "LT-VK" Or UCase(element10.Name) = "RINGLEIDING" Or UCase(element10.Name) = "RU-WW" _
       Or UCase(element10.Name) = "RUV" Or UCase(element10.Name) = "RUH-S" Or UCase(element10.Name) = "RUB-S" _
       Or UCase(element10.Name) = "VSKO" Or UCase(element10.Name) = "LTS-N" Or UCase(element10.Name) = "RU-WWN" _
       Or UCase(element10.Name) = "RU-WWS" Or UCase(element10.Name) = "RU-WK" Or UCase(element10.Name) = "RU-WKN" Or UCase(element10.Name) = "RU-WKS" Then
       Set symbool = element10
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
         If attribuut.TagString = "UNITNUMMER" And attribuut.textstring <> "" Then aa = attribuut.textstring 'UNITNUMMER
        
          If aa = unitonder10 Then
           insp = element10.InsertionPoint
           'element10.Highlight (True)
           element10.Erase
           aa = ""
          End If
        Next I
       End If
     End If
   End If
  Next element10


 
  

If frmGroeptekst.OptionButton2.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_PE.dwg"
If frmGroeptekst.OptionButton6.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_ALU.dwg"
If frmGroeptekst.OptionButton1.Value = True And (frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True) Then bestand = "c:\acad2002\dwg\Mat_spe_ZDringleiding.dwg"
If frmGroeptekst.OptionButton2.Value = True And (frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True) Then bestand = "c:\acad2002\dwg\Mat_spe_PEringleiding.dwg"
If frmGroeptekst.OptionButton6.Value = True And (frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True) Then bestand = "c:\acad2002\dwg\Mat_spe_ALUringleiding.dwg"
bestand20 = "c:\acad2002\dwg\Mat_spe_FLEX_Aankoppel.dwg"
Update

If Err Then
    frmGroeptekst.Show
    Exit Sub
    End If
Dim element2 As Object
If frmGroeptekst.OptionButton1.Value = True Then
aantal_groepen = (Val(frmGroeptekst.TextBox1)) + (Val(frmGroeptekst.TextBox2)) + (Val(frmGroeptekst.TextBox3)) + (Val(frmGroeptekst.TextBox4)) + (Val(frmGroeptekst.TextBox5)) + _
(Val(frmGroeptekst.TextBox6)) + (Val(frmGroeptekst.TextBox7)) + (Val(frmGroeptekst.TextBox15)) + (Val(frmGroeptekst.TextBox16)) + (Val(frmGroeptekst.TextBox17))
If (Val(frmGroeptekst.TextBox1)) = 0 Then frmGroeptekst.TextBox1 = "-"
If (Val(frmGroeptekst.TextBox2)) = 0 Then frmGroeptekst.TextBox2 = "-"
If (Val(frmGroeptekst.TextBox3)) = 0 Then frmGroeptekst.TextBox3 = "-"
If (Val(frmGroeptekst.TextBox4)) = 0 Then frmGroeptekst.TextBox4 = "-"
If (Val(frmGroeptekst.TextBox5)) = 0 Then frmGroeptekst.TextBox5 = "-"
If (Val(frmGroeptekst.TextBox6)) = 0 Then frmGroeptekst.TextBox6 = "-"
If (Val(frmGroeptekst.TextBox7)) = 0 Then frmGroeptekst.TextBox7 = "-"
If (Val(frmGroeptekst.TextBox15)) = 0 Then frmGroeptekst.TextBox15 = "-"
If (Val(frmGroeptekst.TextBox16)) = 0 Then frmGroeptekst.TextBox16 = "-"
aantal_groepen = aantal_groepen + (Val(frmGroeptekst.TextBox26))
'MsgBox aantal_groepen
If frmGroeptekst.ComboBox3.Enabled = False Then
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen 'ZONDER REGELING
Else
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen & "/" & frmGroeptekst.ComboBox3 'met regeling
End If
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Then REGELTYPE = frmGroeptekst.ComboBox2

If frmGroeptekst.ComboBox2 = "RUW-Groot" Or frmGroeptekst.ComboBox2 = "RUW-Klein" Then
REGELTYPE = "RUW" & " " & aantal_groepen
End If
'bloklogo invullen WTH-ZD
  
unittel = frmGroeptekst.TextBox9

If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "MAT_SPE_ZD" Or UCase(element2.Name) = "MAT_SPE_ZDRINGLEIDING" Or UCase(element2.Name) = "MAT_SPE_ZD_1627" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "RNU" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
            If bb = unitonder10 Then
              For k = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(k)
              'element2.Highlight (True)
              If attribuut.TagString = "WTHZD" And attribuut.textstring <> "" Then attribuut.textstring = "WTH-ZD 20*3,4 mm" 'REGELUNITNUMMER
              If attribuut.TagString = "WTH250" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox1  '250 METER
              If attribuut.TagString = "WTH165" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox2  '165 METER
              If attribuut.TagString = "WTH125" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox3  '125 METER
              If attribuut.TagString = "WTH105" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox4  '105 METER
              If attribuut.TagString = "WTH90" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox5  '90 METER
              If attribuut.TagString = "WTH75" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox6  '75 METER
              If attribuut.TagString = "WTH63" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox7  '63 METER
              If attribuut.TagString = "WTH50" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox15  '40 METER
              If attribuut.TagString = "WTH40" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox16  '50 METER
              If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring <> "" Then attribuut.textstring = REGELTYPE 'TYPE REGELUNIT
              If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.ComboBox1   'BEVESTIGING
              'If attribuut.TagString = "ROLGROTER250" And CheckBox1.Value = True Then attribuut.TextString = frmGroeptekst.TextBox14 & " meter :" 'ROL GROTER DAN 250 METER
              Next k
             bb = ""
             End If
       Next I
       
      End If
      End If
      End If
  Next element2
 End If

Dim ELEMENT7
If frmGroeptekst.OptionButton9.Value = True Then
aantal_groepen = (Val(frmGroeptekst.TextBox4)) + (Val(frmGroeptekst.TextBox5)) + _
(Val(frmGroeptekst.TextBox6)) + (Val(frmGroeptekst.TextBox7)) + (Val(frmGroeptekst.TextBox17))
If (Val(frmGroeptekst.TextBox4)) = 0 Then frmGroeptekst.TextBox4 = "-"
If (Val(frmGroeptekst.TextBox5)) = 0 Then frmGroeptekst.TextBox5 = "-"
If (Val(frmGroeptekst.TextBox6)) = 0 Then frmGroeptekst.TextBox6 = "-"
If (Val(frmGroeptekst.TextBox7)) = 0 Then frmGroeptekst.TextBox7 = "-"
aantal_groepen = aantal_groepen + (Val(frmGroeptekst.TextBox26))
'MsgBox aantal_groepen
If frmGroeptekst.ComboBox3.Enabled = False Then
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen 'ZONDER REGELING
Else
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen & "/" & frmGroeptekst.ComboBox3 'met regeling
End If
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Then REGELTYPE = frmGroeptekst.ComboBox2

If frmGroeptekst.ComboBox2 = "RUW-Groot" Or frmGroeptekst.ComboBox2 = "RUW-Klein" Then
REGELTYPE = "RUW" & " " & aantal_groepen
End If
'bloklogo invullen WTH-ZD 16*2,7
  
unittel = frmGroeptekst.TextBox9

If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If



For Each ELEMENT7 In ThisDrawing.ModelSpace
      If ELEMENT7.ObjectName = "AcDbBlockReference" Then
      If UCase(ELEMENT7.Name) = "MAT_SPE_ZD_1627" Then
      Set symbool = ELEMENT7
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "RNU" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
            If bb = unitonder10 Then
              For k = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(k)
              'ELEMENT7.Highlight (True)
              If attribuut.TagString = "WTHZD" And attribuut.textstring <> "" Then attribuut.textstring = "WTH-ZD 16 * 2,7 mm" 'REGELUNITNUMMER
              If attribuut.TagString = "WTH105" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox4  '105 METER
              If attribuut.TagString = "WTH90" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox5  '90 METER
              If attribuut.TagString = "WTH75" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox6  '75 METER
              If attribuut.TagString = "WTH63" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox7  '63 METER
              If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring <> "" Then attribuut.textstring = REGELTYPE 'TYPE REGELUNIT
              If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.ComboBox1   'BEVESTIGING
              'If attribuut.TagString = "ROLGROTER250" And CheckBox1.Value = True Then attribuut.TextString = frmGroeptekst.TextBox14 & " meter :" 'ROL GROTER DAN 250 METER
              Next k
             bb = ""
             End If
       Next I
       
      End If
      End If
      End If
  Next ELEMENT7
  
 End If

 For Each element6 In ThisDrawing.ModelSpace
      If element6.ObjectName = "AcDbBlockReference" Then
      If UCase(element6.Name) = "MAT_SPE_FLEX" Or element6.Name = "Mat_spe_FLEX_Aankoppel" Then
      Set symbool = element6
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        
        If attribuut.TagString = "RNU" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
            If bb = unitonder10 Then
              For k = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(k)
                'element6.Highlight (True)
                'If attribuut.TagString = "RNU" And attribuut.TextString <> "" Then attribuut.TextString = unitonder10 'REGELUNITNUMMER
                If attribuut.TagString = "FLEX_BUIS" And attribuut.textstring <> "" Then attribuut.textstring = "WTH-ZD 16 * 2,7 mm" 'REGELUNITNUMMER
                If attribuut.TagString = "FLEX_METERS" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox25  'AANTAL METER
                If attribuut.TagString = "FLEX_MATTEN" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox26  'AANTAL METER
                If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring <> "" Then attribuut.textstring = REGELTYPE  'TYPE REGELUNIT
                If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.ComboBox1  'BEVESTIGING
             Next k
             bb = ""
           End If
       
       
       Next I
       
      End If
      End If
      End If
  Next element6

 Update
 
 Dim element3 As Object
If frmGroeptekst.OptionButton2.Value = True Then
aantal_groepen = (Val(frmGroeptekst.TextBox1)) + (Val(frmGroeptekst.TextBox2)) + (Val(frmGroeptekst.TextBox3)) + (Val(frmGroeptekst.TextBox17))
If (Val(frmGroeptekst.TextBox1)) = 0 Then frmGroeptekst.TextBox1 = "-"
If (Val(frmGroeptekst.TextBox2)) = 0 Then frmGroeptekst.TextBox2 = "-"
If (Val(frmGroeptekst.TextBox3)) = 0 Then frmGroeptekst.TextBox3 = "-"
If frmGroeptekst.ComboBox3.Enabled = False Then
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen
Else
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen & "/" & frmGroeptekst.ComboBox3
End If
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Then REGELTYPE = frmGroeptekst.ComboBox2

'PE-RT
unittel = frmGroeptekst.TextBox9
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If

For Each element3 In ThisDrawing.ModelSpace
      If element3.ObjectName = "AcDbBlockReference" Then
      If UCase(element3.Name) = "MAT_SPE_PE" Or element3.Name = "Mat_spe_PEringleiding" Or element3.Name = "Mat_spe_PE800" Then
      Set symbool = element3
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
          If attribuut.TagString = "RNU" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
            If bb = unitonder10 Then
              For k = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(k)
                'element3.Highlight (True)
                'If attribuut.TagString = "RNU" And attribuut.TextString <> "" Then attribuut.TextString = unitonder10 'REGELUNITNUMMER
                If attribuut.TagString = "PE" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.ComboBox4 'TYPE BUIS
                If attribuut.TagString = "PE120" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox1 '120 METER
                If attribuut.TagString = "PE90" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox2  '90 METER
                If attribuut.TagString = "PE60" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox3 '60 METER
                If attribuut.TagString = "LMETER" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox27 'werkelijke meters totaal
                If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring <> "" Then attribuut.textstring = REGELTYPE 'TYPE REGELUNIT
                If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.ComboBox1  'BEVESTIGING
             Next k
             bb = ""
           End If
       Next j
       
        End If
      End If
      End If
  Next element3
  
 End If
 Update

If frmGroeptekst.OptionButton6 = True Then
aantal_groepen = (Val(frmGroeptekst.TextBox1)) + (Val(frmGroeptekst.TextBox2)) + (Val(frmGroeptekst.TextBox3)) + (Val(frmGroeptekst.TextBox17))
If (Val(frmGroeptekst.TextBox1)) = 0 Then frmGroeptekst.TextBox1 = "-"
If (Val(frmGroeptekst.TextBox2)) = 0 Then frmGroeptekst.TextBox2 = "-"
If (Val(frmGroeptekst.TextBox3)) = 0 Then frmGroeptekst.TextBox3 = "-"
If frmGroeptekst.ComboBox3.Enabled = False Then
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen
Else
REGELTYPE = frmGroeptekst.ComboBox2 & " " & aantal_groepen & "/" & frmGroeptekst.ComboBox3
End If
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Then REGELTYPE = frmGroeptekst.ComboBox2

'ALUFLEX
unittel = frmGroeptekst.TextBox9
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If

 Dim element4 As Object
For Each element4 In ThisDrawing.ModelSpace
      If element4.ObjectName = "AcDbBlockReference" Then
      If UCase(element4.Name) = "MAT_SPE_ALU" Or element4.Name = "Mat_spe_ALUringleiding" Then
      Set symbool = element4
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "RNU" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
            If bb = unitonder10 Then
              For k = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(k)
             'element4.Highlight (True)
             'If attribuut.TagString = "RNU" And attribuut.TextString <> "" Then attribuut.TextString = unitonder10 'REGELUNITNUMMER
             If attribuut.TagString = "ALU" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.ComboBox4 'TYPE BUIS
             If attribuut.TagString = "ALU200" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox1  '120 METER
             If attribuut.TagString = "ALU100" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox2  '90 METER
             If attribuut.TagString = "ALU50" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.TextBox3 '60 METER
             If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring <> "" Then attribuut.textstring = REGELTYPE 'TYPE REGELUNIT
             If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring <> "" Then attribuut.textstring = frmGroeptekst.ComboBox1  'BEVESTIGING
             If attribuut.TagString = "PEALU200" And OptionButton6 = True Then attribuut.textstring = "200 meter :" 'Aluflex 200 METER
             If attribuut.TagString = "PEALU100" And OptionButton6 = True Then attribuut.textstring = "100 meter :" 'Aluflex 100 METER
             If attribuut.TagString = "PEALU50" And OptionButton6 = True Then attribuut.textstring = " 50 meter :" 'Aluflex 100 METER
             Next k
             bb = ""
             End If
             
           Next j
       
        End If
      End If
      End If
  Next element4
 End If
  Update
  
  
'  Dim pb2(0 To 2) As Double
'  If OptionButton1.Value = True Then zakken = 210
'  If OptionButton2.Value = True Or OptionButton6.Value = True Then zakken = 179
'  If TextBox11 = "0" And TextBox25 <> "0" Then zakken = 179
'
'  pb2(0) = pb1(0) - (scaal * 460)
'  pb2(1) = pb1(1) + (scaal * zakken) '179) '177'210)
'  pb2(2) = pb1(2)
'
  'juiste regelunitblokje.dwg inserten in de tekening
  Dim bestand2 As String
  bestand2 = frmGroeptekst.ComboBox2 & ".dwg"
  If frmGroeptekst.ComboBox2 = "RUW-Groot" Then bestand2 = "RUW" & ".dwg"
  If frmGroeptekst.ComboBox2 = "RUW-Klein" Then bestand2 = "RUW" & ".dwg"
  If frmGroeptekst.ComboBox2 = "RUB-R" And aantal_groepen > 4 Then bestand2 = "RUH-R" & ".dwg"
  If frmGroeptekst.ComboBox2 = "RUB-RT" And aantal_groepen > 4 Then bestand2 = "RUH-RT" & ".dwg"
  If frmGroeptekst.ComboBox2 = "RUB-S" And aantal_groepen > 4 Then bestand2 = "RUH-S" & ".dwg"
  If frmGroeptekst.ComboBox2 = "VSKO" Then bestand2 = "VSKO-B" & ".dwg"
   Dim bestand3 As String
  bestand3 = "c:\acad2002\dwg\" & frmGroeptekst.ComboBox2 & ".txt"  'tekstbestand met afmetingen
  If frmGroeptekst.ComboBox2 = "VSKO" Then bestand3 = "c:\acad2002\dwg\" & frmGroeptekst.ComboBox2 & "-B.txt"  'tekstbestand met afmetingen
  
  Update
  'MsgBox bestand2 & " - " & bestand3

 Call unitblock(aantal_groepen, bestand2, bestand3, insp, scaal)
  
End Sub
Sub unitblock(aantal_groepen, bestand2, bestand3, insp, scaal)
'unitblok  ----------------
'unitblok  ----------------
'unitblok  ----------------
'unitblok  ----------------
 
 Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(insp, bestand2, scaal, scaal, 1, 0)
  Update
 
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Or frmGroeptekst.ComboBox2 = "RINGLEIDING" Then
 afmunitRING = "RINGLEIDING"
Else


 'open text bestand om afmetingen van de unit uit te lezen
  aantal_groepen2 = aantal_groepen
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fs, a, afmunit
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(bestand3, ForReading, False)
Do While aantal_groepen2 <> 0
    afmunit = a.ReadLine
    aantal_groepen2 = aantal_groepen2 - 1
Loop
a.Close 'sluiten van tekstbestand
End If
 

unittel = frmGroeptekst.TextBox9
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If


Dim element5 As Object
unit = frmGroeptekst.ComboBox2
If frmGroeptekst.ComboBox2 = "RUW-Groot" Then unit = "RUW"
If frmGroeptekst.ComboBox2 = "RUW-Klein" Then unit = "RUW"
If frmGroeptekst.ComboBox2 = "RUB-R" And aantal_groepen > 4 Then unit = "RUH-R"
If frmGroeptekst.ComboBox2 = "RUB-RT" And aantal_groepen > 4 Then unit = "RUH-RT"
If frmGroeptekst.ComboBox2 = "RUB-S" And aantal_groepen > 4 Then unit = "RUH-S"
If frmGroeptekst.ComboBox2 = "VSKO" Then unit = "VSKO-B"


For Each element5 In ThisDrawing.ModelSpace
      If element5.ObjectName = "AcDbBlockReference" Then
      If UCase(element5.Name) = unit Then
      Set symbool = element5
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
         If unit = "RINGLEIDING" Then
          If attribuut.TagString = "AFMETINGEN" And attribuut.textstring = "" Then attribuut.textstring = afmunitRING 'AFMETING VAN DEUNIT
          Else
          If attribuut.TagString = "AFMETINGEN" And attribuut.textstring = "" Then attribuut.textstring = afmunit 'AFMETING VAN DEUNIT
        End If
        If attribuut.TagString = "UNITNUMMER" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'UNITNUMMER
        Next I
       End If
      End If
      End If
 Next element5


  Call RESET
  'Unload frmGroeptekst
  'Unload Me
End Sub


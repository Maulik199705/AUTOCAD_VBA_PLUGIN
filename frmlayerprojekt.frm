VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmlayerprojekt 
   Caption         =   "UserForm1"
   ClientHeight    =   3060
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   6504
   OleObjectBlob   =   "frmlayerprojekt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmlayerprojekt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()
'Call Checklayer.Checklayer  'module kijkt of er lege layers zijn

ListBox1.Clear
ListBox2.Clear
Update
On Error Resume Next
  
 Dim cirkel As Object
 Dim element As Object
 Dim Lengte As Double
 Dim laagobj As Object
 
 Dim minaantal As Integer
 Dim maxaantal As Integer
 Dim i As Integer
 i = 0
 minaantal = 0
 maxaantal = ThisDrawing.Layers.Count
 For Each laagobj In ThisDrawing.Layers
 i = i + 1
 ProgressBar1.Min = minaantal
 ProgressBar1.Max = maxaantal
 ProgressBar1.Value = i
 
      mystr = Left(laagobj.Name, 5)
      If mystr = "GROEP" Then mystr = "groep"
      If Not mystr = "groep" Then
      GoTo wand
      Else
        For Each element In ThisDrawing.ModelSpace
          If element.layer = laagobj.Name Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
             End If 'elementlayer
             Next element
            
              For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.layer = laagobj.Name Then
                If cirkel.EntityName = "AcDbCircle" Then
                Z = Z + 1
                wvaanwezig = " (WV)"
                zlengte = Z * 100 '(ComboBox1.Text * 100)) + 100
                End If
                End If
                Next cirkel
                
                
                
                
wand:
        If mystr = "WAND_" Then mystr = "wand_"
        If mystr = "wand_" Then
             For Each element In ThisDrawing.ModelSpace
             If element.layer = laagobj.Name Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
                         
             End If 'elementlayer
             Next element
             
                For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.layer = laagobj.Name Then
                If cirkel.EntityName = "AcDbCircle" Then
                Z = Z + 1
                wvaanwezig = " (WV)"
                End If
                End If
                Next cirkel
                zlengte = (Z * (ComboBox1.Text * 100)) + 100
                End If '2e mystr
                
    Lengte = (Lengte * TextBox19) + zlengte
    'Lengte = Lengte + zlengte
    Lengte = (Lengte / 100) + Val(TextBox27)
    Lengte = Round(Lengte, 2)
    totalrollen = totalrollen + Lengte
    
    'nieuwe tekeningen
''''    If frmLayerlijst.CheckBox3.Value = False Then
''''   'WTH-ZD-leiding
''''    If Lengte >= 162.5 And Lengte < 250 And OptionButton1.Value = True Then
''''    q250 = q250 + 1
''''    E = "--> wordt een rol van 250 meter"
''''    End If
''''    If Lengte >= 122.5 And Lengte < 162.5 And OptionButton1.Value = True Then
''''    q165 = q165 + 1
''''    E = "--> wordt een rol van 165 meter"
''''    End If
''''    If Lengte >= 102.5 And Lengte < 122.5 And OptionButton1.Value = True Then
''''    q125 = q125 + 1
''''    E = "--> wordt een rol van 125 meter"
''''    End If
''''    If Lengte >= 87.5 And Lengte < 102.5 And OptionButton1.Value = True Then
''''    q105 = q105 + 1
''''    E = "--> wordt een rol van 105 meter"
''''    End If
''''    If Lengte >= 72.5 And Lengte < 87.5 And OptionButton1.Value = True Then
''''    q90 = q90 + 1
''''    E = "--> wordt een rol van 90 meter"
''''    End If
''''    If Lengte >= 60.5 And Lengte < 72.5 And OptionButton1.Value = True Then
''''    q75 = q75 + 1
''''    E = "--> wordt een rol van 75 meter"
''''    End If
''''    If Lengte >= 47.5 And Lengte < 60.5 And OptionButton1.Value = True Then
''''    q63 = q63 + 1
''''    E = "--> wordt een rol van 63 meter"
''''    End If
''''    If Lengte >= 37.5 And Lengte < 47.5 And OptionButton1.Value = True Then
''''    q50 = q50 + 1
''''    E = "--> wordt een rol van 50 meter"
''''    End If
''''    If Lengte >= 10 And Lengte < 37.5 And OptionButton1.Value = True Then
''''    q40 = q40 + 1
''''    E = "--> wordt een rol van 40 meter"
''''    End If
    
    
'''''    'PE-RT-leiding
'''''    If Lengte >= 87.5 And Lengte < 120 And OptionButton2.Value = True Then
'''''    qpe120 = qpe120 + 1
'''''    E = "--> wordt een rol van 120 meter"
'''''    End If
'''''    If Lengte >= 120 And OptionButton2.Value = True And CheckBox5.Value = True Then
'''''    qpe120 = qpe120 + 1
'''''    E = "--> wordt een rol van 120 meter"
'''''    End If
'''''    If Lengte >= 57.5 And Lengte < 87.5 And OptionButton2.Value = True Then
'''''    qpe90 = qpe90 + 1
'''''    E = "--> wordt een rol van 90 meter"
'''''    End If
'''''    If Lengte >= 10 And Lengte < 57.5 And OptionButton2.Value = True Then
'''''    qpe60 = qpe60 + 1
'''''    E = "--> wordt een rol van 60 meter"
'''''    End If
'''''    End If 'nieuwe tekeningen
'''''
'''''  'oude tekeningen
''''''    If frmLayerlijst.CheckBox3.Value = True Then
''''''   'WTH-ZD-leiding
''''''    If Lengte >= 163.6 And Lengte < 250 And OptionButton1.Value = True Then
''''''    q250 = q250 + 1
''''''    E = "--> wordt een rol van 250 meter"
''''''    End If
''''''    If Lengte >= 123.6 And Lengte < 163.6 And OptionButton1.Value = True Then
''''''    q165 = q165 + 1
''''''    E = "--> wordt een rol van 165 meter"
''''''    End If
''''''    If Lengte >= 103.6 And Lengte < 123.6 And OptionButton1.Value = True Then
''''''    q125 = q125 + 1
''''''    E = "--> wordt een rol van 125 meter"
''''''    End If
''''''    If Lengte >= 88.6 And Lengte < 103.6 And OptionButton1.Value = True Then
''''''    q105 = q105 + 1
''''''    E = "--> wordt een rol van 105 meter"
''''''    End If
''''''    If Lengte >= 73.6 And Lengte < 88.6 And OptionButton1.Value = True Then
''''''    q90 = q90 + 1
''''''    E = "--> wordt een rol van 90 meter"
''''''    End If
''''''    If Lengte >= 61.6 And Lengte < 73.6 And OptionButton1.Value = True Then
''''''    q75 = q75 + 1
''''''    E = "--> wordt een rol van 75 meter"
''''''    End If
''''''    If Lengte >= 48.6 And Lengte < 61.6 And OptionButton1.Value = True Then
''''''    q63 = q63 + 1
''''''    E = "--> wordt een rol van 63 meter"
''''''    End If
''''''    If Lengte >= 38.6 And Lengte < 48.6 And OptionButton1.Value = True Then
''''''    q50 = q50 + 1
''''''    E = "--> wordt een rol van 50 meter"
''''''    End If
''''''    If Lengte >= 10 And Lengte < 38.6 And OptionButton1.Value = True Then
''''''    q40 = q40 + 1
''''''    E = "--> wordt een rol van 40 meter"
''''''    End If
    
    
''''''    'PE-RT-leiding
''''''    If Lengte >= 88.6 And Lengte < 120 And OptionButton2.Value = True Then
''''''    qpe120 = qpe120 + 1
''''''    E = "--> wordt een rol van 120 meter"
''''''    End If
''''''    If Lengte >= 120 And OptionButton2.Value = True And CheckBox5.Value = True Then
''''''    qpe120 = qpe120 + 1
''''''    E = "--> wordt een rol van 120 meter"
''''''    End If
''''''    If Lengte >= 58.6 And Lengte < 88.6 And OptionButton2.Value = True Then
''''''    qpe90 = qpe90 + 1
''''''    E = "--> wordt een rol van 90 meter"
''''''    End If
''''''    If Lengte >= 10 And Lengte < 58.6 And OptionButton2.Value = True Then
''''''    qpe60 = qpe60 + 1
''''''    E = "--> wordt een rol van 60 meter"
''''''    End If
''''''    End If 'oude tekeningen
''''''
  
''''''''   'wth-zd
''''''''    If Lengte > 250 And OptionButton1.Value = True Then
''''''''    Label15 = " !!!..DE MAX. ROLLENGTE WORDT OVERSCHREDEN...!!!"
''''''''    Label15.BackColor = &HFFFF&
''''''''    'e = "--> ROL IS TE LANG.!!!"
''''''''    overschrijding1 = 1
''''''''    'frmLayerlijst.Height = 243
''''''''    End If
''''''''    'PE-RT
''''''''    If Lengte > 120 And OptionButton2.Value = True And CheckBox5.Value = False Then
''''''''    Label15 = " !!!..DE MAX. ROLLENGTE WORDT OVERSCHREDEN...!!!"
''''''''    Label15.BackColor = &HFFFF&
''''''''    'e = "--> ROL IS TE LANG.!!!"
''''''''    overschrijding1 = 1
''''''''    'frmLayerlijst.Height = 243
''''''''    End If
''''''''
''''''''
    If mystr = "groep" Then
    s = " = "
    Else
    s = "  = " ' HIERO
    End If
    
    If mystr = "wand_" Or mystr = "WAND_" Or mystr = "groep" Or mystr = "GROEP" Then
    If Lengte < 10 Or Lengte > 120 And OptionButton2.Value = True And CheckBox5.Value = False Or Lengte > 250 And OptionButton1.Value = True Then
    d = "LET OP!! " & laagobj.Name & s & Lengte & " meter."
    ListBox2.AddItem (d)
    overschrijding2 = 1
    Else
    d = laagobj.Name & s & Lengte & " meter"
    mylen = Len(d)
    ListBox1.AddItem (d)
    End If
    End If
    
'''''    If OptionButton1.Value = True Then
'''''    TextBox2 = q250
'''''    TextBox3 = q165
'''''    TextBox4 = q125
'''''    TextBox5 = q105
'''''    TextBox6 = q90
'''''    TextBox7 = q75
'''''    TextBox8 = q63
'''''    TextBox10 = q50
'''''    TextBox11 = q40
'''''
'''''
'''''    If TextBox2 = "" Then TextBox2 = "0"
'''''    If TextBox3 = "" Then TextBox3 = "0"
'''''    If TextBox4 = "" Then TextBox4 = "0"
'''''    If TextBox5 = "" Then TextBox5 = "0"
'''''    If TextBox6 = "" Then TextBox6 = "0"
'''''    If TextBox7 = "" Then TextBox7 = "0"
'''''    If TextBox8 = "" Then TextBox8 = "0"
'''''    If TextBox10 = "" Then TextBox10 = "0"
'''''    If TextBox11 = "" Then TextBox11 = "0"
'''''
'''''
'''''
'''''    totaal2 = q250 + q165 + q125 + q105 + q90 + q75 + q63 + q50 + q40
'''''    End If
    
''''''    If OptionButton2.Value = True Then
''''''    TextBox2 = qpe120
''''''    TextBox3 = qpe90
''''''    TextBox4 = qpe60
''''''    TextBox5.Visible = False
''''''    TextBox6.Visible = False
''''''    TextBox7.Visible = False
''''''    TextBox8.Visible = False
''''''    TextBox10.Visible = False
''''''    TextBox11.Visible = False
''''''
''''''
''''''    If TextBox2 = "" Then TextBox2 = "0"
''''''    If TextBox3 = "" Then TextBox3 = "0"
''''''    If TextBox4 = "" Then TextBox4 = "0"
''''''    If TextBox21 = "" Then TextBox21 = "0"
''''''    If TextBox22 = "" Then TextBox22 = "0"
''''''    If TextBox23 = "" Then TextBox23 = "0"
''''''    If TextBox24 = "" Then TextBox24 = "0"
''''''    Label1.Caption = "120 meter"
''''''    Label3.Caption = "90 meter"
''''''    Label5.Caption = "60 meter"
''''''    Label7.Caption = Clear
''''''    Label8.Caption = Clear
''''''    Label10.Caption = Clear
''''''    Label13.Caption = Clear
''''''    Label21.Caption = Clear
''''''    Label22.Caption = Clear
''''''
''''''    totaal2 = qpe120 + qpe90 + qpe60
''''''    End If
    Lengte = 0 'Lengte leeggooien voordat de volgende groep wordt gemeten
    zlengte = 0
    Z = 0
    wvaanwezig = ""
  End If  'end if  mystr
 Next laagobj
 ProgressBar1.Value = minaantal
 Update
 
''''''  Label16 = totaal2
''''''  Label17 = " Totaal: " & totaal2 & " groep(en)"
'''''' If q250 = 0 And q165 = 0 And q125 = 0 And q105 = 0 And q90 = 0 And q75 = 0 _
'''''' And q63 = 0 And q50 = 0 And q40 = 0 And qpe120 = 0 And qpe90 = 0 And qpe60 = 0 Then
'''''' Cmdprint.Enabled = False
'''''' frmLayerlijst.Height = 346
'''''' Label15.Caption = " !!!...GEEN GROEPLAYERS AANWEZIG...!!!"
'''''' Label15.BackColor = &HFFFF&
'''''' Else
'''''' frmlayerprojekt.Width = 364 ''''''''''''''''''''''''''''''''''''''''''HIER
'''''' Cmdprint.Enabled = True
'''''' End If
'''''' If overschrijding1 = 1 Then frmLayerlijst.Height = 263 '243
'''''' If overschrijding2 = 1 Then frmLayerlijst.Height = 346 '290
'''''' If CheckBox5.Value = False Then
'''''' Label24.Visible = True
'''''' TextBox12.Visible = True
'''''' Else
'''''' Label24.Visible = False
'''''' TextBox12.Visible = False
'''''' End If
'''''' If OptionButton2.Value = True Then
'''''' CheckBox5.Visible = True
'''''' Else
'''''' CheckBox5.Visible = False
'''''' End If
''''''
   'lijst rangschikken
  Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For i = 0 To ListBox1.ListCount - 1
    textstring2 = ListBox1.List(i)
    Veld(i) = textstring2
   
   Dim LB&, UB&, TEMP$, Pos&, x&
 
    LB = LBound(Veld)
    UB = UBound(Veld)
 
    While UB > LB
      Pos = LB
 
      For x = LB To UB - 1
        If Veld(x) > Veld(x + 1) Then
          TEMP = Veld(x + 1)
          Veld(x + 1) = Veld(x)
          Veld(x) = TEMP
          Pos = x
        End If
      Next x
 
      UB = Pos
    Wend
    
  Next i
  ListBox1.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then ListBox1.AddItem Veld(x)
  Next x
 TextBox12 = totalrollen
 TextBox20 = TextBox12
 TextBox26 = TextBox12
 If OptionButton2.Value = True Then CheckBox5.Enabled = False
 

End Sub

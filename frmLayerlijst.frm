VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLayerlijst 
   Caption         =   "LAYERLIJST "
   ClientHeight    =   6405
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   15324
   HelpContextID   =   5
   OleObjectBlob   =   "frmLayerlijst.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLayerlijst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'26-01-2004 Layerlijst genereren
'M.Bosch en G.C.Haak
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
        ByVal hWnd As LongPtr) As Long

    Private Declare PtrSafe Function GetMenuItemCount Lib "user32" ( _
        ByVal hMenu As LongPtr) As Long

    Private Declare PtrSafe Function GetSystemMenu Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr

    Private Declare PtrSafe Function RemoveMenu Lib "user32" ( _
        ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function DrawMenuBar Lib "user32" ( _
        ByVal hWnd As Long) As Long

    Private Declare Function GetMenuItemCount Lib "user32" ( _
        ByVal hMenu As Long) As Long

    Private Declare Function GetSystemMenu Lib "user32" ( _
        ByVal hWnd As Long, ByVal bRevert As Long) As Long

    Private Declare Function RemoveMenu Lib "user32" ( _
        ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If


Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000



Private Sub UserForm_Initialize()
#If VBA7 Then
    Dim lngHwnd As LongPtr
    Dim lngMenu As LongPtr
#Else
    Dim lngHwnd As Long
    Dim lngMenu As Long
#End If
  Dim lngCnt As Long
  lngHwnd = FindWindow(vbNullString, Me.Caption)
  lngMenu = GetSystemMenu(lngHwnd, 0)
  If lngMenu Then
    lngCnt = GetMenuItemCount(lngMenu)
    Call RemoveMenu(lngMenu, lngCnt - 1, _
    MF_REMOVE Or MF_BYPOSITION)
    Call DrawMenuBar(lngHwnd)
  End If
 'frmLayerlijst.Width = 273
 'frmLayerlijst.Height = 219.75
 On Error Resume Next
 frmLayerlijst.Height = 250
 frmLayerlijst.Width = 352
 
 Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "GERARD" Then
    TextBox27.Visible = True
    CheckBox6.Visible = True
End If
 
 
 
 Call extreemtek
 
 Kill ("c:\acad2002\layerlijst.txt")
 ComboBox1.AddItem "2.5"
 ComboBox1.AddItem "2"
 ComboBox1.Text = ComboBox1.List(0)
 Call LEIDINGSOORT
 If TextBox14 > 30 Then Call importuserr2
 If TextBox14 < 31 Then TextBox13 = ThisDrawing.GetVariable("USERR2")
 End Sub
Private Sub CheckBox5_Click()
If ListBox1.ListCount = 0 Then TextBox12 = 0
 
If CheckBox5.Value = True Then
Frame1.Visible = True
Label24.Visible = False
TextBox12.Visible = False
frmLayerlijst.Width = 364
TextBox20 = TextBox12
'TextBox25 = Val(TextBox25) * (Val(TextBox2) + Val(TextBox3) + Val(TextBox4))
TextBox25 = 0
TextBox26 = TextBox12
Else
'frmLayerlijst.Width = 352
Label24.Visible = True
TextBox12.Visible = True
Frame1.Visible = False
frmLayerlijst.Width = 352
End If
End Sub
Private Sub TextBox21_Change()
On Error Resume Next
Dim b As Double
b = TextBox21.Text
  If Err Then
   If TextBox21 = "" Then TextBox21 = 0
  Exit Sub
  End If

TextBox26 = Val(TextBox26) - (800 * Val(TextBox21))

If Label30.Caption = "0" Then Label30 = TextBox21
If TextBox21 = "0" Or TextBox21 = "" Then TextBox26 = Val(TextBox26) + (800 * Val(Label30))
Label30.Caption = TextBox21
End Sub
Private Sub TextBox22_Change()
On Error Resume Next
Dim b As Double
b = TextBox22.Text
  If Err Then
   If TextBox22 = "" Then TextBox22 = 0
  Exit Sub
  End If

TextBox26 = Val(TextBox26) - (120 * Val(TextBox22))

If Label31.Caption = "0" Then Label31 = TextBox22
If TextBox22 = "0" Or TextBox22 = "" Then TextBox26 = Val(TextBox26) + (120 * Val(Label31))
Label31.Caption = TextBox22
End Sub
Private Sub TextBox23_Change()
On Error Resume Next
Dim b As Double
b = TextBox23.Text
  If Err Then
   If TextBox23 = "" Then TextBox23 = 0
  Exit Sub
  End If

TextBox26 = Val(TextBox26) - (90 * Val(TextBox23))

If Label32.Caption = "0" Then Label32 = TextBox23
If TextBox23 = "0" Or TextBox23 = "" Then TextBox26 = Val(TextBox26) + (90 * Val(Label32))
Label32.Caption = TextBox23
End Sub
Private Sub TextBox24_Change()
On Error Resume Next
Dim b As Double
b = TextBox24.Text
  If Err Then
   If TextBox24 = "" Then TextBox24 = 0
  Exit Sub
  End If

TextBox26 = Val(TextBox26) - (60 * Val(TextBox24))

If Label33.Caption = "0" Then Label33 = TextBox24
If TextBox24 = "0" Or TextBox24 = "" Then TextBox26 = Val(TextBox26) + (60 * Val(Label33))
Label33.Caption = TextBox24
End Sub
Private Sub TextBox25_Change()
TextBox26 = (Val(TextBox20)) + (Val(TextBox25) * (Val(TextBox2) + Val(TextBox3) + Val(TextBox4)))
End Sub
Sub exportuserr2()
TT = ThisDrawing.GetVariable("USERR2")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fg, fh
'Dim s1 As AcadSelectionSet
Set fg = CreateObject("Scripting.FileSystemObject")

Set fh = fg.OpenTextFile("c:\acad2002\userr2.txt", ForWriting, -2)
    fh.write TT
    fh.Close
End Sub
Sub importuserr2()
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fl, m, TS
Set fl = CreateObject("Scripting.FileSystemObject")
Set m = fl.OpenTextFile("c:\acad2002\userr2.txt", ForReading, False)
    TS = m.ReadLine
    TextBox13 = TS
m.Close 'sluiten van tekstbestand
End Sub
Private Sub CommandButton1_Click()
    'Ensure ListBox contains list items
    If ListBox1.ListCount >= 1 Then
        'If no selection, choose last list item.
        If ListBox1.ListIndex = -1 Then
            ListBox1.ListIndex = _
                    ListBox1.ListCount - 1
        End If
        ListBox1.RemoveItem (ListBox1.ListIndex)
    End If
    Update
ListBox1.AddItem TextBox1.Text
Update

 'lijst rangschikken
  Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For I = 0 To ListBox1.ListCount - 1
    textstring2 = ListBox1.List(I)
    Veld(I) = textstring2
   
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
    
  Next I
  ListBox1.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then ListBox1.AddItem Veld(x)
  Next x

TextBox1 = Clear
CommandButton1.Locked = True
'ListBox1.RemoveItem (laagnaam)
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim naam As Object
laagnaam = ListBox1.Text
TextBox1 = laagnaam
Update

End Sub

Private Sub OptionButton1_Click()
frmLayerlijst.Height = 250
Cmdprint.Enabled = False
ListBox1.Clear
ListBox2.Clear
frmLayerlijst.Width = 352
Label2 = q250: Label4 = q165: Label6 = q125: Label9 = q105: Label11 = q90: Label12 = q75: Label12 = q63
TextBox5.Visible = True: TextBox6.Visible = True: TextBox7.Visible = True: TextBox8.Visible = True
TextBox9.Visible = True: TextBox10.Visible = True: TextBox11.Visible = True
Label1.Caption = "250 meter": Label3.Caption = "165 meter": Label5.Caption = "125 meter"
Label8.Caption = "105 meter": Label10.Caption = "90 meter": Label7.Caption = "75 meter"
Label13.Caption = "63 meter": Label21.Caption = "50 meter": Label22.Caption = "40 meter"
Label15.Caption = Clear
Label17.Caption = Clear
Label15.BackColor = &HC0C0C0
CheckBox5.Visible = False
Frame1.Visible = False
Label24.Visible = False
TextBox12.Visible = False


    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    TextBox7 = ""
    TextBox8 = ""
    TextBox10 = ""
    TextBox11 = ""
End Sub

Private Sub OptionButton2_Click()
Label24.Visible = False
TextBox12.Visible = False
frmLayerlijst.Height = 250
Cmdprint.Enabled = False
ListBox1.Clear
ListBox2.Clear
frmLayerlijst.Width = 352
Label15.Caption = Clear
Label17.Caption = Clear
Label15.BackColor = &HC0C0C0
CheckBox5.Visible = True
CheckBox5.Enabled = True

    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5.Visible = False
    TextBox6.Visible = False
    TextBox7.Visible = False
    TextBox8.Visible = False
    TextBox10.Visible = False
    TextBox11.Visible = False
    Label1.Caption = "120 meter"
    Label3.Caption = "90 meter"
    Label5.Caption = "60 meter"
    Label7.Caption = Clear
    Label8.Caption = Clear
    Label10.Caption = Clear
    Label13.Caption = Clear
    Label21.Caption = Clear
    Label22.Caption = Clear
End Sub
Private Sub TextBox1_Change()
If TextBox1 <> "" Then CommandButton1.Locked = False
End Sub
Private Sub cmdLayers_Click()
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
 Dim I As Integer
 I = 0
 minaantal = 0
 maxaantal = ThisDrawing.Layers.Count
 For Each laagobj In ThisDrawing.Layers
 I = I + 1
 ProgressBar1.Min = minaantal
 ProgressBar1.Max = maxaantal
 ProgressBar1.Value = I
 
      mystr = Left(laagobj.Name, 5)
      If mystr = "GROEP" Then mystr = "groep"
      If Not mystr = "groep" Then
      GoTo wand
      Else
        For Each element In ThisDrawing.ModelSpace
          If element.Layer = laagobj.Name Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
             End If 'elementlayer
             Next element
            
              For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.Layer = laagobj.Name Then
                If cirkel.EntityName = "AcDbCircle" Then
                Z = Z + 1
                wvaanwezig = " (WV)"
                zlengte = (Z * (ComboBox1.Text * 100)) + 100
                End If
                End If
                Next cirkel
                
                
                
                
wand:
        If mystr = "WAND_" Then mystr = "wand_"
        If mystr = "wand_" Then
             For Each element In ThisDrawing.ModelSpace
             If element.Layer = laagobj.Name Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
                         
             End If 'elementlayer
             Next element
             
                For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.Layer = laagobj.Name Then
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
    Lengte = Round(Lengte, 1)
    totalrollen = totalrollen + Lengte
    
    'nieuwe tekeningen
    If frmLayerlijst.CheckBox3.Value = False Then
   'WTH-ZD-leiding
    If Lengte >= 162.5 And Lengte < 250 And OptionButton1.Value = True Then
    q250 = q250 + 1
    E = "--> wordt een rol van 250 meter"
    End If
    If Lengte >= 122.5 And Lengte < 162.5 And OptionButton1.Value = True Then
    q165 = q165 + 1
    E = "--> wordt een rol van 165 meter"
    End If
    If Lengte >= 102.5 And Lengte < 122.5 And OptionButton1.Value = True Then
    q125 = q125 + 1
    E = "--> wordt een rol van 125 meter"
    End If
    If Lengte >= 87.5 And Lengte < 102.5 And OptionButton1.Value = True Then
    q105 = q105 + 1
    E = "--> wordt een rol van 105 meter"
    End If
    If Lengte >= 72.5 And Lengte < 87.5 And OptionButton1.Value = True Then
    q90 = q90 + 1
    E = "--> wordt een rol van 90 meter"
    End If
    If Lengte >= 60.5 And Lengte < 72.5 And OptionButton1.Value = True Then
    q75 = q75 + 1
    E = "--> wordt een rol van 75 meter"
    End If
    If Lengte >= 47.5 And Lengte < 60.5 And OptionButton1.Value = True Then
    q63 = q63 + 1
    E = "--> wordt een rol van 63 meter"
    End If
    If Lengte >= 37.5 And Lengte < 47.5 And OptionButton1.Value = True Then
    q50 = q50 + 1
    E = "--> wordt een rol van 50 meter"
    End If
    If Lengte >= 10 And Lengte < 37.5 And OptionButton1.Value = True Then
    q40 = q40 + 1
    E = "--> wordt een rol van 40 meter"
    End If
    
    
    'PE-RT-leiding
    If Lengte >= 87.5 And Lengte < 120 And OptionButton2.Value = True Then
    qpe120 = qpe120 + 1
    E = "--> wordt een rol van 120 meter"
    End If
    If Lengte >= 120 And OptionButton2.Value = True And CheckBox5.Value = True Then
    qpe120 = qpe120 + 1
    E = "--> wordt een rol van 120 meter"
    End If
    If Lengte >= 57.5 And Lengte < 87.5 And OptionButton2.Value = True Then
    qpe90 = qpe90 + 1
    E = "--> wordt een rol van 90 meter"
    End If
    If Lengte >= 10 And Lengte < 57.5 And OptionButton2.Value = True Then
    qpe60 = qpe60 + 1
    E = "--> wordt een rol van 60 meter"
    End If
    End If 'nieuwe tekeningen
  
  'oude tekeningen
    If frmLayerlijst.CheckBox3.Value = True Then
   'WTH-ZD-leiding
    If Lengte >= 163.6 And Lengte < 250 And OptionButton1.Value = True Then
    q250 = q250 + 1
    E = "--> wordt een rol van 250 meter"
    End If
    If Lengte >= 123.6 And Lengte < 163.6 And OptionButton1.Value = True Then
    q165 = q165 + 1
    E = "--> wordt een rol van 165 meter"
    End If
    If Lengte >= 103.6 And Lengte < 123.6 And OptionButton1.Value = True Then
    q125 = q125 + 1
    E = "--> wordt een rol van 125 meter"
    End If
    If Lengte >= 88.6 And Lengte < 103.6 And OptionButton1.Value = True Then
    q105 = q105 + 1
    E = "--> wordt een rol van 105 meter"
    End If
    If Lengte >= 73.6 And Lengte < 88.6 And OptionButton1.Value = True Then
    q90 = q90 + 1
    E = "--> wordt een rol van 90 meter"
    End If
    If Lengte >= 61.6 And Lengte < 73.6 And OptionButton1.Value = True Then
    q75 = q75 + 1
    E = "--> wordt een rol van 75 meter"
    End If
    If Lengte >= 48.6 And Lengte < 61.6 And OptionButton1.Value = True Then
    q63 = q63 + 1
    E = "--> wordt een rol van 63 meter"
    End If
    If Lengte >= 38.6 And Lengte < 48.6 And OptionButton1.Value = True Then
    q50 = q50 + 1
    E = "--> wordt een rol van 50 meter"
    End If
    If Lengte >= 10 And Lengte < 38.6 And OptionButton1.Value = True Then
    q40 = q40 + 1
    E = "--> wordt een rol van 40 meter"
    End If
    
    
    'PE-RT-leiding
    If Lengte >= 88.6 And Lengte < 120 And OptionButton2.Value = True Then
    qpe120 = qpe120 + 1
    E = "--> wordt een rol van 120 meter"
    End If
    If Lengte >= 120 And OptionButton2.Value = True And CheckBox5.Value = True Then
    qpe120 = qpe120 + 1
    E = "--> wordt een rol van 120 meter"
    End If
    If Lengte >= 58.6 And Lengte < 88.6 And OptionButton2.Value = True Then
    qpe90 = qpe90 + 1
    E = "--> wordt een rol van 90 meter"
    End If
    If Lengte >= 10 And Lengte < 58.6 And OptionButton2.Value = True Then
    qpe60 = qpe60 + 1
    E = "--> wordt een rol van 60 meter"
    End If
    End If 'oude tekeningen
  
  
   'wth-zd
    If Lengte > 250 And OptionButton1.Value = True Then
    Label15 = " !!!..DE MAX. ROLLENGTE WORDT OVERSCHREDEN...!!!"
    Label15.BackColor = &HFFFF&
    'e = "--> ROL IS TE LANG.!!!"
    overschrijding1 = 1
    'frmLayerlijst.Height = 243
    End If
    'PE-RT
    If Lengte > 120 And OptionButton2.Value = True And CheckBox5.Value = False Then
    Label15 = " !!!..DE MAX. ROLLENGTE WORDT OVERSCHREDEN...!!!"
    Label15.BackColor = &HFFFF&
    'e = "--> ROL IS TE LANG.!!!"
    overschrijding1 = 1
    'frmLayerlijst.Height = 243
    End If
    
    
    If mystr = "groep" Then
    S = " = "
    Else
    S = "  = " ' HIERO
    End If
    
    If mystr = "wand_" Or mystr = "WAND_" Or mystr = "groep" Or mystr = "GROEP" Then
    If Lengte < 10 Or Lengte > 120 And OptionButton2.Value = True And CheckBox5.Value = False Or Lengte > 250 And OptionButton1.Value = True Then
    d = "LET OP!! " & laagobj.Name & S & Lengte & " meter."
    ListBox2.AddItem (d)
    overschrijding2 = 1
    Else
    d = laagobj.Name & S & Lengte & " meter"
    mylen = Len(d)
    'MsgBox mylen
    If mylen = 20 Then f = "     "
    If mylen = 22 Then f = "   "
    If mylen = 23 Then f = "  "
    If mylen = 24 Then f = " "
    d = d & f & E & wvaanwezig
    ListBox1.AddItem (d)
    End If
    End If
    
    If OptionButton1.Value = True Then
    TextBox2 = q250
    TextBox3 = q165
    TextBox4 = q125
    TextBox5 = q105
    TextBox6 = q90
    TextBox7 = q75
    TextBox8 = q63
    TextBox10 = q50
    TextBox11 = q40
    
      
    If TextBox2 = "" Then TextBox2 = "0"
    If TextBox3 = "" Then TextBox3 = "0"
    If TextBox4 = "" Then TextBox4 = "0"
    If TextBox5 = "" Then TextBox5 = "0"
    If TextBox6 = "" Then TextBox6 = "0"
    If TextBox7 = "" Then TextBox7 = "0"
    If TextBox8 = "" Then TextBox8 = "0"
    If TextBox10 = "" Then TextBox10 = "0"
    If TextBox11 = "" Then TextBox11 = "0"
    
    
    
    totaal2 = q250 + q165 + q125 + q105 + q90 + q75 + q63 + q50 + q40
    End If
    
    If OptionButton2.Value = True Then
    TextBox2 = qpe120
    TextBox3 = qpe90
    TextBox4 = qpe60
    TextBox5.Visible = False
    TextBox6.Visible = False
    TextBox7.Visible = False
    TextBox8.Visible = False
    TextBox10.Visible = False
    TextBox11.Visible = False
    
         
    If TextBox2 = "" Then TextBox2 = "0"
    If TextBox3 = "" Then TextBox3 = "0"
    If TextBox4 = "" Then TextBox4 = "0"
    If TextBox21 = "" Then TextBox21 = "0"
    If TextBox22 = "" Then TextBox22 = "0"
    If TextBox23 = "" Then TextBox23 = "0"
    If TextBox24 = "" Then TextBox24 = "0"
    Label1.Caption = "120 meter"
    Label3.Caption = "90 meter"
    Label5.Caption = "60 meter"
    Label7.Caption = Clear
    Label8.Caption = Clear
    Label10.Caption = Clear
    Label13.Caption = Clear
    Label21.Caption = Clear
    Label22.Caption = Clear
    
    totaal2 = qpe120 + qpe90 + qpe60
    End If
    Lengte = 0 'Lengte leeggooien voordat de volgende groep wordt gemeten
    zlengte = 0
    Z = 0
    wvaanwezig = ""
  End If  'end if  mystr
 Next laagobj
 ProgressBar1.Value = minaantal
 Update
 
  Label16 = totaal2
  Label17 = " Totaal: " & totaal2 & " groep(en)"
 If q250 = 0 And q165 = 0 And q125 = 0 And q105 = 0 And q90 = 0 And q75 = 0 _
 And q63 = 0 And q50 = 0 And q40 = 0 And qpe120 = 0 And qpe90 = 0 And qpe60 = 0 Then
 Cmdprint.Enabled = False
 frmLayerlijst.Height = 346
 Label15.Caption = " !!!...GEEN GROEPLAYERS AANWEZIG...!!!"
 Label15.BackColor = &HFFFF&
 Else
 frmLayerlijst.Width = 364 ''''''''''''''''''''''''''''''''''''''''''HIER
 Cmdprint.Enabled = True
 End If
 If overschrijding1 = 1 Then frmLayerlijst.Height = 263 '243
 If overschrijding2 = 1 Then frmLayerlijst.Height = 346 '290
 If CheckBox5.Value = False Then
 Label24.Visible = True
 TextBox12.Visible = True
 Else
 Label24.Visible = False
 TextBox12.Visible = False
 End If
 If OptionButton2.Value = True Then
 CheckBox5.Visible = True
 Else
 CheckBox5.Visible = False
 End If
 
   'lijst rangschikken
  Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For I = 0 To ListBox1.ListCount - 1
    textstring2 = ListBox1.List(I)
    Veld(I) = textstring2
   
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
    
  Next I
  ListBox1.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then ListBox1.AddItem Veld(x)
  Next x
 TextBox12 = totalrollen
 TextBox20 = TextBox12
 TextBox26 = TextBox12
 If OptionButton2.Value = True Then CheckBox5.Enabled = False
 
Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "GERARD" Then
   frmLayerlijst.Width = 509
   Call extr_control
End If
 
End Sub
Private Sub Cmdprint_Click()
Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "ILONA" Then
   Call ILONA
  Else




Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fs, f
Dim s1 As AcadSelectionSet
Set fs = CreateObject("Scripting.FileSystemObject")
teknaam = ThisDrawing.GetVariable("dwgname")
pad = ThisDrawing.GetVariable("dwgprefix")
usernaam = ThisDrawing.GetVariable("loginname")
Dim MyDate
MyDate = DateValue(Date)    ' Return a date.


Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
    f.write "     " & "Tekenaar: " & usernaam & " |Datum: " & MyDate
    f.write Chr(13) + Chr(10)
    f.write Chr(13) + Chr(10)
    f.write "     " & pad
    f.write Chr(13) + Chr(10)
    f.write "     " & teknaam
    f.write Chr(13) + Chr(10)
    f.write Chr(13) + Chr(10)
    f.Close
    
If CheckBox5.Value = False Then
    teller = ListBox1.ListCount
    For I = 0 To teller - 1
       'Define the text object
        textstring = ListBox1.List(I)
        Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
        f.write "     " & textstring
        f.write Chr(13) + Chr(10)
        f.Close
    Next I
    

    If OptionButton1.Value = True Then
    totaal1 = TextBox2 & " * 250|" & TextBox3 & " * 165|" & TextBox4 & " * 125|" & _
    TextBox5 & " * 105|" & TextBox6 & " * 90|" & TextBox7 & " * 75|" & _
    TextBox8 & " * 63|" & TextBox10 & " * 50|" & TextBox11 & " * 40"
    
    totaal2 = "Totaal: " & Label16 & " groepen. (WTH-ZD leiding)"
    roltel = Val(TextBox2) + Val(TextBox3) + Val(TextBox4) + Val(TextBox5) + Val(TextBox6) + _
    Val(TextBox7) + Val(TextBox8) + Val(TextBox10) + Val(TextBox11)
    Else
    totaal1 = TextBox2 & " * 120 | " & TextBox3 & " * 90 | " & TextBox4 & " * 60"
    totaal2 = "Totaal: " & Label16 & " groepen. (PE-RT leiding)"
    roltel = Val(TextBox2) + Val(TextBox3) + Val(TextBox4)
    End If
    eindrol = "Totaal: " & roltel & " rol(len)"
    compleet = "Totaal: " & TextBox12 & " meters."
    uitleg = "(WV) = Groep met wandverwarming"
    If TextBox13 <> "" Then aantalmeters = "Oppervlakte t.b.v.Isolatie/Netten: " & TextBox13 & " m2"
    
    Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal1
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal2
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & eindrol
        f.write Chr(13) + Chr(10)
        f.write "     " & compleet
        f.write Chr(13) + Chr(10)
        f.write "     " & uitleg
        If TextBox13 <> 0 Then
        f.write Chr(13) + Chr(10)
        f.write "     " & aantalmeters
        End If
        f.Close
        
    teller2 = ListBox2.ListCount
    If frmLayerlijst.CheckBox4.Value = True Then
        Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
        f.write Chr(13) + Chr(10)
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & "Afwijkende rollengte's"
        f.write Chr(13) + Chr(10)
        f.Close
    For q = 0 To teller2 - 1
       'Define the text object
        textstring2 = ListBox2.List(q)
        Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
        f.write "     " & textstring2
        f.write Chr(13) + Chr(10)
        f.Close
    Next q
    End If
Else  '800 meter rollen
    teller = ListBox1.ListCount
    For I = 0 To teller - 1
       'Define the text object
        textstring = ListBox1.List(I)
        textstring2 = Split(textstring, "wordt")
        Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
        f.write "     " & textstring2(0)
        f.write Chr(13) + Chr(10)
        f.Close
    Next I
  
    totaal1 = TextBox21 & " * 800|" & TextBox22 & " * 120|" & TextBox23 & " * 90|" & TextBox24 & " * 60|"
    totaal2 = "Totaal: " & Label16 & " groepen. (PE-RT leiding)"
    roltel = Val(TextBox21) + Val(TextBox22) + Val(TextBox23) + Val(TextBox24)
    eindrol = "Totaal: " & roltel & " rol(len)"
    roltotal = (Val(TextBox20)) + (Val(TextBox25) * (Val(TextBox2) + Val(TextBox3) + Val(TextBox4)))
    compleet = "Totaal: " & roltotal & " meters.(inclusief restlengte - " & TextBox25 & " meter per groep.)"
    compleet2 = "Totaal: " & TextBox12 & " meters.(exclusief restlengte)"
    uitleg = "(WV) = Groep met wandverwarming"
    If TextBox13 <> "" Then aantalmeters = "Oppervlakte t.b.v.Isolatie/Netten: " & TextBox13 & " m2"
    
    Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal1
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal2
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & eindrol
        f.write Chr(13) + Chr(10)
        f.write "     " & compleet2
        f.write Chr(13) + Chr(10)
        f.write "     " & compleet
        f.write Chr(13) + Chr(10)
        f.write "     " & uitleg
        If TextBox13 <> 0 Then
        f.write Chr(13) + Chr(10)
        f.write "     " & aantalmeters
        End If
        f.Close
        
    teller2 = ListBox2.ListCount
    If frmLayerlijst.CheckBox4.Value = True Then
        Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
        f.write Chr(13) + Chr(10)
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & "Afwijkende rollengte's"
        f.write Chr(13) + Chr(10)
        f.Close
    For q = 0 To teller2 - 1
       'Define the text object
        textstring2 = ListBox2.List(q)
        Set f = fs.OpenTextFile("c:\acad2002\layerlijst.txt", ForAppending, -2)
''''''        f.write Chr(13) + Chr(10)
''''''        f.write Chr(13) + Chr(10)
''''''        f.write "     " & "--------------------------------------------------------------------"
''''''        f.write Chr(13) + Chr(10)
''''''        f.write "     " & "Afwijkende rollengte's"
''''''        f.write Chr(13) + Chr(10)
        f.write "     " & textstring2
        f.write Chr(13) + Chr(10)
        f.Close
    Next q
    End If

End If  '800 meter rollen




    
Dim RetVal
RetVal = Shell("C:\acad2002\vba\layerlijst.bat", 1)    ' uitprinten textfile.

If TextBox14 > 30 Then ThisDrawing.Close
 If TextBox14 > 30 Then
  bestandnaam3 = TextBox16 & TextBox17 & "-meten" & ".dwg"
  Kill (bestandnaam3)
 End If
ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "oppervlakte" & vbCr & vbCr
End If 'ILONA
Unload Me
ThisDrawing.SendCommand "setvar" & vbCr & "acadlspasdoc" & vbCr & "1" & vbCr
End Sub
Sub LEIDINGSOORT()

'TextBox13 = ThisDrawing.GetVariable("userr2")

 For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOK" Then
                Set symbool = element
                If symbool.HasAttributes Then
                    attributen = symbool.GetAttributes
                    For I = LBound(attributen) To UBound(attributen)
                         Set attribuut = attributen(I)
                         If attribuut.TagString = "LEIDINGSOORT" Then WSL = attribuut.textstring
                         If attribuut.TagString = "UNITNUMMER" Then UNITSjek = attribuut.textstring
                    Next I
                End If
            End If
        End If
    Next element
If WSL = "WTH-ZD" Then frmLayerlijst.OptionButton1 = True
If WSL = "PE-RT" Then
     CheckBox5.Visible = True
     frmLayerlijst.OptionButton2 = True
End If
UNITSJEK1 = Len(UNITSjek)
If UNITSJEK1 = 1 Then frmLayerlijst.CheckBox3.Value = True
End Sub
Private Sub cmdAfsluiten_Click()
'checkopexport = Right(TextBox20, 5)
If TextBox18 <> "" Then
 checkopexport = Split(TextBox18, ".")
 checkopexport1 = Right(checkopexport(0), 5)
 If checkopexport1 = "meten" Then ThisDrawing.Close
 bestandnaam3 = TextBox16 & TextBox17 & "-meten" & ".dwg"
 Kill (bestandnaam3)
End If
ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "oppervlakte" & vbCr & vbCr
Unload Me
ThisDrawing.SendCommand "setvar" & vbCr & "acadlspasdoc" & vbCr & "1" & vbCr
End Sub

Sub extreemtek()

'Call zoekblad(bladnummer)
'TextBox14 = bladnummer

 'If TextBox14 = "" Then
 '    Exit Sub
 'End If
 
 
TextBox14 = Clear
For Each layerObj In ThisDrawing.Layers
     If Left(layerObj.Name, 5) = "groep" Then TextBox14 = (Val(TextBox14)) + 1  'aantal groepen
Next 'layerobj
If TextBox14 = "" Then TextBox14 = 0

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "KADERLOGO" Or UCase(element2.Name) = "LOGOTGH" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For k = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(k)
        If attribuut.TagString = "SCHAAL" Then schaalfaktor = attribuut.textstring
        Next k
       
      End If
      End If
      End If
Next element2

vartab = ThisDrawing.GetVariable("EXTMAX")

' A3 T/M A0+
If vartab(0) >= 2045 And vartab(0) <= 2055 Or vartab(0) >= 2915 And vartab(0) <= 2925 _
    Or vartab(0) >= 4175 And vartab(0) <= 4185 Or vartab(0) >= 5890 And vartab(0) <= 5900 _
    Or vartab(0) >= 7840 And vartab(0) <= 7850 Then
    maxwaarde = 1
Else
    maxwaarde = 2
End If
If vartab(0) >= 16715 And vartab(0) <= 16725 Or vartab(0) >= 23575 And vartab(0) <= 23585 _
    Or vartab(0) >= 31375 And vartab(0) <= 31385 Then
    maxwaarde = 4
End If
Update

If schaalfaktor = "1:50" And maxwaarde = 1 Then scaal = 1
If schaalfaktor = "1:100" And maxwaarde = 1 Then scaal = 2
If schaalfaktor = "1:200" And maxwaarde = 1 Then scaal = 4

If schaalfaktor = "1:50" And maxwaarde = 2 Then scaal = 0.5
If schaalfaktor = "1:100" And maxwaarde = 2 Then scaal = 1
If schaalfaktor = "1:200" And maxwaarde = 2 Then scaal = 2

If schaalfaktor = "1:200" And maxwaarde = 4 Then scaal = 1
TextBox19.Value = scaal

teknaam = ThisDrawing.GetVariable("dwgname")    'teknaam = ThisDrawing.GetVariable("dwgname")
TextBox15 = teknaam
TextBox16 = "c:\acad2002\"  'pad = ThisDrawing.GetVariable("dwgprefix")
teknaam1 = Split(teknaam, ("."))                'teknaam1 = Split(teknaam, ("."))

         Dim mystr As Variant
         mystr = Len(teknaam)
         over = mystr - 4 'aantal karakters
         teknaam6 = Left(teknaam, over)

'TextBox18 = teknaam1(0)
TextBox17 = teknaam6


'aantal groepen groter dan 30 ????
If TextBox14 > 30 Then
Call exportuserr2
frmLayerlijst.Hide

 
'bestandnaam2 = pad & teknaam1(0) & "-export" & ".dwg"
bestandnaam2 = TextBox16 & TextBox17 & "-meten" & ".dwg"
TextBox18 = bestandnaam2

'Kill (bestandnaam)
'Call zoekblad(bladnummer)

ThisDrawing.SendCommand "-layer" & vbCr & "unlock" & vbCr & "*" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "set" & vbCr & "0" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "*" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "groep*" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "gt" & vbCr & vbCr

ThisDrawing.SendCommand "_copyclip" & vbCr & "all" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "*" & vbCr & vbCr

'''templatenaam = ThisDrawing.GetVariable("acadver")
'''If templatenaam = "15.06s (LMS Tech)" Then ThisDrawing.Application.Documents.Add ("acad2000.dwt")
''''If templatenaam = "16.0s (LMS Tech)" Then ThisDrawing.Application.Documents.Add ("acad2004.dwt")
'''If templatenaam >= "16.0s (LMS Tech)" Then
ThisDrawing.SendCommand "setvar" & vbCr & "acadlspasdoc" & vbCr & "0" & vbCr
ThisDrawing.Application.Documents.Add ("autocad.dwt")

ThisDrawing.SendCommand "_pasteclip" & vbCr & "0,0" & vbCr

ThisDrawing.SaveAs (bestandnaam2)
ZoomExtents
End If


End Sub
Sub extr_control()

Dim element As Object
    For Each element In ThisDrawing.ModelSpace
          If element.ObjectName = "AcDbBlockReference" Then
          If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
          Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
             If attribuut.TagString = "GROEPTEKST" Then gra = attribuut.textstring
                If attribuut.TagString = "ROLLENGTE" And attribuut.textstring <> " " Then
                vul = gra & " | " & attribuut.textstring
                ListBox3.AddItem (vul)
                End If
           Next I
    End If
    End If
    End If
    Next element
    
    
 Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For I = 0 To ListBox3.ListCount - 1
    textstring2 = ListBox3.List(I)
    Veld(I) = textstring2
   
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
    
  Next I
  ListBox3.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then ListBox3.AddItem Veld(x)
  Next x
  Label34 = ListBox3.ListCount
  
  
  
  Dim jet As String
  Dim textstring6 As String
For k = 0 To ListBox1.ListCount - 1
    textstring5 = ListBox1.List(k)
    textstring6 = ListBox3.List(k)
   
    jel = Split(textstring5, "=") 'groeptekst
    'MsgBox Len(jel(0)) 'groeptekst
    jek = Split(jel(1), "-->")
    'MsgBox jek(1)
    jem = Split(jek(1), " ")
    'MsgBox Len(jem(5)) 'rollengte
    jet = jel(0) & "| " & jem(5) & " meter"
        If jet <> textstring6 Then
          TTS = "[layers]- " & jet & " -FOUT- " & textstring6 & " -[groeptekstblok]"
          ListBox2.AddItem (TTS)
          frmLayerlijst.Height = 344.25
        End If
    'MsgBox JET & " - " & TEXTSTRING6
    R = Len(jet)
    S = Len(textstring6)
   ' MsgBox R & " - " & S
Next k
ListBox2.Width = 345
 
End Sub
Sub ILONA()
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fs, f
Dim s1 As AcadSelectionSet
Set fs = CreateObject("Scripting.FileSystemObject")
teknaam = ThisDrawing.GetVariable("dwgname")
pad = ThisDrawing.GetVariable("dwgprefix")
usernaam = ThisDrawing.GetVariable("loginname")
Dim MyDate
MyDate = DateValue(Date)    ' Return a date.


         Dim mystr As Variant
         mystr = Len(teknaam)
         over = mystr - 4 'aantal karakters
         teknaam6 = Left(teknaam, over)

txtbest = pad & "\" & teknaam6 & ".txt"


Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
    f.write "     " & "Tekenaar: " & usernaam & " |Datum: " & MyDate
    f.write Chr(13) + Chr(10)
    f.write Chr(13) + Chr(10)
    f.write "     " & pad
    f.write teknaam
    f.write Chr(13) + Chr(10)
    f.write Chr(13) + Chr(10)
    f.Close
    
If CheckBox5.Value = False Then
    teller = ListBox1.ListCount
    For I = 0 To teller - 1
       'Define the text object
        textstring = ListBox1.List(I)
        Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
        f.write "     " & textstring
        f.write Chr(13) + Chr(10)
        f.Close
    Next I
    

    If OptionButton1.Value = True Then
    totaal1 = TextBox2 & " * 250|" & TextBox3 & " * 165|" & TextBox4 & " * 125|" & _
    TextBox5 & " * 105|" & TextBox6 & " * 90|" & TextBox7 & " * 75|" & _
    TextBox8 & " * 63|" & TextBox10 & " * 50|" & TextBox11 & " * 40"
    
    totaal2 = "Totaal: " & Label16 & " groepen. (WTH-ZD leiding)"
    roltel = Val(TextBox2) + Val(TextBox3) + Val(TextBox4) + Val(TextBox5) + Val(TextBox6) + _
    Val(TextBox7) + Val(TextBox8) + Val(TextBox10) + Val(TextBox11)
    Else
    totaal1 = TextBox2 & " * 120 | " & TextBox3 & " * 90 | " & TextBox4 & " * 60"
    totaal2 = "Totaal: " & Label16 & " groepen. (PE-RT leiding)"
    roltel = Val(TextBox2) + Val(TextBox3) + Val(TextBox4)
    End If
    eindrol = "Totaal: " & roltel & " rol(len)"
    compleet = "Totaal: " & TextBox12 & " meters."
    uitleg = "(WV) = Groep met wandverwarming"
    If TextBox13 <> "" Then aantalmeters = "Oppervlakte t.b.v.Isolatie/Netten: " & TextBox13 & " m2"
    
    Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal1
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal2
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & eindrol
        f.write Chr(13) + Chr(10)
        f.write "     " & compleet
        f.write Chr(13) + Chr(10)
        f.write "     " & uitleg
        If TextBox13 <> 0 Then
        f.write Chr(13) + Chr(10)
        f.write "     " & aantalmeters
        End If
        f.Close
        
    teller2 = ListBox2.ListCount
    If frmLayerlijst.CheckBox4.Value = True Then
        Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
        f.write Chr(13) + Chr(10)
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & "Afwijkende rollengte's"
        f.write Chr(13) + Chr(10)
        f.Close
    For q = 0 To teller2 - 1
       'Define the text object
        textstring2 = ListBox2.List(q)
        Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
        f.write "     " & textstring2
        f.write Chr(13) + Chr(10)
        f.Close
    Next q
    End If
Else  '800 meter rollen
    teller = ListBox1.ListCount
    For I = 0 To teller - 1
       'Define the text object
        textstring = ListBox1.List(I)
        textstring2 = Split(textstring, "wordt")
        Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
        f.write "     " & textstring2(0)
        f.write Chr(13) + Chr(10)
        f.Close
    Next I
  
    totaal1 = TextBox21 & " * 800|" & TextBox22 & " * 120|" & TextBox23 & " * 90|" & TextBox24 & " * 60|"
    totaal2 = "Totaal: " & Label16 & " groepen. (PE-RT leiding)"
    roltel = Val(TextBox21) + Val(TextBox22) + Val(TextBox23) + Val(TextBox24)
    eindrol = "Totaal: " & roltel & " rol(len)"
    roltotal = (Val(TextBox20)) + (Val(TextBox25) * (Val(TextBox2) + Val(TextBox3) + Val(TextBox4)))
    compleet = "Totaal: " & roltotal & " meters.(inclusief restlengte - " & TextBox25 & " meter per groep.)"
    compleet2 = "Totaal: " & TextBox12 & " meters.(exclusief restlengte)"
    uitleg = "(WV) = Groep met wandverwarming"
    If TextBox13 <> "" Then aantalmeters = "Oppervlakte t.b.v.Isolatie/Netten: " & TextBox13 & " m2"
    
    Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal1
        f.write Chr(13) + Chr(10)
        f.write "     " & totaal2
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & eindrol
        f.write Chr(13) + Chr(10)
        f.write "     " & compleet2
        f.write Chr(13) + Chr(10)
        f.write "     " & compleet
        f.write Chr(13) + Chr(10)
        f.write "     " & uitleg
        If TextBox13 <> 0 Then
        f.write Chr(13) + Chr(10)
        f.write "     " & aantalmeters
        End If
        f.Close
        
    teller2 = ListBox2.ListCount
    If frmLayerlijst.CheckBox4.Value = True Then
        Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
        f.write Chr(13) + Chr(10)
        f.write Chr(13) + Chr(10)
        f.write "     " & "--------------------------------------------------------------------"
        f.write Chr(13) + Chr(10)
        f.write "     " & "Afwijkende rollengte's"
        f.write Chr(13) + Chr(10)
        f.Close
    For q = 0 To teller2 - 1
       'Define the text object
        textstring2 = ListBox2.List(q)
        Set f = fs.OpenTextFile(txtbest, ForAppending, -2)
''''''        f.write Chr(13) + Chr(10)
''''''        f.write Chr(13) + Chr(10)
''''''        f.write "     " & "--------------------------------------------------------------------"
''''''        f.write Chr(13) + Chr(10)
''''''        f.write "     " & "Afwijkende rollengte's"
''''''        f.write Chr(13) + Chr(10)
        f.write "     " & textstring2
        f.write Chr(13) + Chr(10)
        f.Close
    Next q
    End If

End If  '800 meter rollen

Dim RetVal
Dim retval2

RetVal = "> HP LaserJet 4100"
retval2 = "type " & txtbest & " " & RetVal & vbCr


ThisDrawing.SendCommand retval2


If TextBox14 > 30 Then ThisDrawing.Close
 If TextBox14 > 30 Then
  bestandnaam3 = TextBox16 & TextBox17 & "-meten" & ".dwg"
  Kill (bestandnaam3)
 End If
ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "oppervlakte" & vbCr & vbCr
Unload Me
End Sub


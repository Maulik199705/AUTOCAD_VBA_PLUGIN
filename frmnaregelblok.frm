VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmnaregelblok 
   Caption         =   "NAREGELBLOK"
   ClientHeight    =   9540.001
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11088
   OleObjectBlob   =   "frmnaregelblok.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmnaregelblok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As LongPtr) As Long
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Private Sub UserForm_Initialize()
Dim lngHwnd As LongPtr
  Dim lngMenu As LongPtr
  Dim lngCnt As Long
  lngHwnd = FindWindow(vbNullString, Me.Caption)
  lngMenu = GetSystemMenu(lngHwnd, 0)
  If lngMenu Then
    lngCnt = GetMenuItemCount(lngMenu)
    Call RemoveMenu(lngMenu, lngCnt - 1, _
    MF_REMOVE Or MF_BYPOSITION)
    Call DrawMenuBar(lngHwnd)
  End If
 ' Call zoekblok3
  
  Dim lognaam
   lognaam = ThisDrawing.GetVariable("loginname")
   lognaam = UCase(lognaam)
   If lognaam = "GERARD" Then
   TextBox2.Visible = True
   CheckBox5.Visible = True
   End If
  Call bloklezen.uitlez
frmnaregelblok.StartUpPosition = 0
End Sub

Private Sub CancelButton_Click()
Call zoekblok2
 groepinvoerBox1.Value = Clear
 groepinvoerBox2.Value = Clear
groepinvoerBox2.SetFocus
'Label2.Caption = Clear
Unload Me
'ThisDrawing.SendCommand "_vbaunload" & vbCr & "montageblok.dvb" & vbCr
End Sub

Private Sub groepinvoerBox1_Change()
Dim a As Double
On Error Resume Next
a = groepinvoerBox1.Text
If groepinvoerBox1 <> "" And groepinvoerBox2 <> "" Then PlaatsButton.Enabled = True

If Err Then
   groepinvoerBox1 = Clear
   PlaatsButton.Enabled = False
   frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    frmnaregelblok.TextBox214.Visible = True: frmnaregelblok.ComboBox214.Visible = True: frmnaregelblok.CheckBox214.Visible = True
    frmnaregelblok.TextBox215.Visible = True: frmnaregelblok.ComboBox215.Visible = True: frmnaregelblok.CheckBox215.Visible = True
    frmnaregelblok.TextBox216.Visible = True: frmnaregelblok.ComboBox216.Visible = True: frmnaregelblok.CheckBox216.Visible = True
    frmnaregelblok.TextBox217.Visible = True: frmnaregelblok.ComboBox217.Visible = True: frmnaregelblok.CheckBox217.Visible = True
    frmnaregelblok.TextBox218.Visible = True: frmnaregelblok.ComboBox218.Visible = True: frmnaregelblok.CheckBox218.Visible = True
    frmnaregelblok.TextBox219.Visible = True: frmnaregelblok.ComboBox219.Visible = True: frmnaregelblok.CheckBox219.Visible = True
    frmnaregelblok.Height = 64.5
  Exit Sub
  End If

If a > 20 Then
   MsgBox "Groter dan 20 groepen is niet toegestaan..!!!!", vbCritical
   groepinvoerBox1 = Clear
   groepinvoerBox1.SetFocus
End If
   
If a > -1 And a < 1 Then
    groepinvoerBox1 = Clear
    groepinvoerBox1.SetFocus
End If
Call bloklezen.view
frmnaregelblok.Height = 402
End Sub
Private Sub groepinvoerBox2_Change()
Dim b As Double
On Error Resume Next
groepinvoerBox2.SetFocus

b = groepinvoerBox2.Text
c = "Regelunit" & " " & b
'Label2.Caption = c

If Err Then
   groepinvoerBox2 = Clear
   'Label2.Caption = Clear
   PlaatsButton.Enabled = False
   CommandButton1.Enabled = False
  Exit Sub
  End If

If b > -1 And b < 1 Then
    groepinvoerBox2 = Clear
    groepinvoerBox2.SetFocus
End If

If groepinvoerBox2 <> "" Then
    CommandButton1.Enabled = True
Else
    CommandButton1.Enabled = False
End If

If groepinvoerBox1 <> "" And groepinvoerBox2 <> "" Then PlaatsButton.Enabled = True
End Sub
Private Sub CommandButton1_Click()
Call zoekblok1
Call bloklezen.view
End Sub
Sub zoekblok1()
For Each element20 In ThisDrawing.ModelSpace
      If element20.ObjectName = "AcDbBlockReference" Then
      'If UCase(element20.Name) = "MAT_SPE_ZD" Or UCase(element20.Name) = "MAT_SPE_PE" Then
      If element20.Name = "Mat_spe_ZD" Or element20.Name = "Mat_spe_PE" Or element20.Name = "Mat_spe_PE800" _
      Or element20.Name = "Mat_spe_ALU" Or element20.Name = "Mat_spe_ZDringleiding" Or element20.Name = "Mat_spe_PEringleiding" Or _
      element20.Name = "Mat_spe_ALUringleiding" Or element20.Name = "Mat_spe_FLEX" Then
      Set SYMBOOL = element20
        
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
          If ATTRIBUUT.TagString = "RNU" Then
               rnc = ATTRIBUUT.textstring
               rnc2 = Len(rnc)
'               If rnc2 = 1 Then CheckBox3.Value = True
           End If
             If Val(groepinvoerBox2) = rnc Then
              b0 = element20.InsertionPoint
              For L = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
              Set ATTRIBUUT = ATTRIBUTEN(L)
              element20.Highlight (True)
                          
          If ATTRIBUUT.TagString = "REGELUNITTYPE" Then
          rgl = ATTRIBUUT.textstring 'TYPE REGELUNIT
          rgl2 = Split(rgl, " ")
          rgl3 = Len(rgl2(1))
           If rgl3 > 2 Then
                'MsgBox "meer dan 2 stuks"
              rgl4 = Split(rgl2(1), "/")
              groepinvoerBox1 = rgl4(0)
           End If
           
           If rgl3 = 1 Or rgl3 = 2 Then
           'MsgBox "klopt 2 stuks"
              rgl4 = Val(rgl2(1))
              groepinvoerBox1 = rgl4
            End If
          End If 'regelunittype
              Next L
                
         End If
       Next i
       
      End If
      End If
      End If
  Next element20
 Update
 
If groepinvoerBox1 = "" Then
   a = "Het bloklogo van regelunit " & groepinvoerBox2 & " is er niet."
   MsgBox a, vbExclamation
     groepinvoerBox2 = Clear
     groepinvoerBox2.SetFocus
     Exit Sub
 End If
 
 gg = groepinvoerBox2
 If gg > 9 Then gg1 = groepinvoerBox2
 If gg > 0 And gg < 10 Then gg1 = "0" & groepinvoerBox2
 
 
frmnaregelblok.TextBox200 = "groep " & gg1 & "." & "01"
frmnaregelblok.TextBox201 = "groep " & gg1 & "." & "02"
frmnaregelblok.TextBox202 = "groep " & gg1 & "." & "03"
frmnaregelblok.TextBox203 = "groep " & gg1 & "." & "04"
frmnaregelblok.TextBox204 = "groep " & gg1 & "." & "05"
frmnaregelblok.TextBox205 = "groep " & gg1 & "." & "06"
frmnaregelblok.TextBox206 = "groep " & gg1 & "." & "07"
frmnaregelblok.TextBox207 = "groep " & gg1 & "." & "08"
frmnaregelblok.TextBox208 = "groep " & gg1 & "." & "09"
frmnaregelblok.TextBox209 = "groep " & gg1 & "." & "10"
frmnaregelblok.TextBox210 = "groep " & gg1 & "." & "11"
frmnaregelblok.TextBox211 = "groep " & gg1 & "." & "12"
frmnaregelblok.TextBox212 = "groep " & gg1 & "." & "13"
frmnaregelblok.TextBox213 = "groep " & gg1 & "." & "14"
frmnaregelblok.TextBox214 = "groep " & gg1 & "." & "15"
frmnaregelblok.TextBox215 = "groep " & gg1 & "." & "16"
frmnaregelblok.TextBox216 = "groep " & gg1 & "." & "17"
frmnaregelblok.TextBox217 = "groep " & gg1 & "." & "18"
frmnaregelblok.TextBox218 = "groep " & gg1 & "." & "19"
frmnaregelblok.TextBox219 = "groep " & gg1 & "." & "20"



 Dim b1(0 To 2) As Double
  b1(0) = b0(0) - 1500 '1000
  b1(1) = b0(1) - 400  '500
  b1(2) = 0
 
 Dim b2(0 To 2) As Double
  b2(0) = b0(0) + 500
  b2(1) = b0(1) + 1000 '750
  b2(2) = 0
  
  
  Dim lognaam
  lognaam = ThisDrawing.GetVariable("loginname")
  lognaam = UCase(lognaam)
  If lognaam = "GERARD" And CheckBox5.Value = False Then
  ZoomWindow b1, b2
  End If
  
  
'Call PlaatsButton_Click
 

  
End Sub
Sub zoekblok2()
For Each element20 In ThisDrawing.ModelSpace
      If element20.ObjectName = "AcDbBlockReference" Then
      'If UCase(element20.Name) = "MAT_SPE_ZD" Or UCase(element20.Name) = "MAT_SPE_PE" Then
       If element20.Name = "Mat_spe_ZD" Or element20.Name = "Mat_spe_PE" Or element20.Name = "Mat_spe_PE800" _
      Or element20.Name = "Mat_spe_ALU" Or element20.Name = "Mat_spe_ZDringleiding" Or element20.Name = "Mat_spe_PEringleiding" Or _
      element20.Name = "Mat_spe_ALUringleiding" Or element20.Name = "Mat_spe_FLEX" Then
      Set SYMBOOL = element20
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
          If ATTRIBUUT.TagString = "RNU" Then rnc = Val(ATTRIBUUT.textstring)
             If Val(groepinvoerBox2) = rnc Then element20.Highlight (False)
         Next i
       
      End If
      End If
      End If
  Next element20
 Update
End Sub
Private Sub PlaatsButton_Click()
Call Schaal(scaal)
On Error Resume Next

 
 Dim pb1 As Variant
 Dim pe2(0 To 2) As Double
 Dim blockRefObj As AcadBlockReference
 Dim element As AcadEntity
 Dim nieuwelement As AcadEntity
 ThisDrawing.ObjectSnapMode = False

If groepinvoerBox1 = "1" Then
     MsgBox "Een 1 groeps naregelblok vermelden we niet op de tekening..!!!!", vbCritical
     groepinvoerBox1 = Clear
     groepinvoerBox1.SetFocus
     PlaatsButton.Enabled = False
     Exit Sub
     End If
     

Dim newLayer As AcadLayer
Set newLayer = ThisDrawing.Layers.Add("NAREGELBLOK")
ThisDrawing.ActiveLayer = newLayer
Update

If frmnaregelblok.TextBox2 <> "" Then scaal = frmnaregelblok.TextBox2
 frmnaregelblok.Hide
 On Error Resume Next
 Dim bestand As String
    If frmnaregelblok.groepinvoerBox1 = 2 Then bestand = "c:\acad2002\dwg\naregelblok2.dwg"
    If frmnaregelblok.groepinvoerBox1 = 3 Then bestand = "c:\acad2002\dwg\naregelblok3.dwg"
    If frmnaregelblok.groepinvoerBox1 = 4 Then bestand = "c:\acad2002\dwg\naregelblok4.dwg"
    If frmnaregelblok.groepinvoerBox1 = 5 Then bestand = "c:\acad2002\dwg\naregelblok5.dwg"
    If frmnaregelblok.groepinvoerBox1 = 6 Then bestand = "c:\acad2002\dwg\naregelblok6.dwg"
    If frmnaregelblok.groepinvoerBox1 = 7 Then bestand = "c:\acad2002\dwg\naregelblok7.dwg"
    If frmnaregelblok.groepinvoerBox1 = 8 Then bestand = "c:\acad2002\dwg\naregelblok8.dwg"
    If frmnaregelblok.groepinvoerBox1 = 9 Then bestand = "c:\acad2002\dwg\naregelblok9.dwg"
    If frmnaregelblok.groepinvoerBox1 = 10 Then bestand = "c:\acad2002\dwg\naregelblok10.dwg"
    If frmnaregelblok.groepinvoerBox1 = 11 Then bestand = "c:\acad2002\dwg\naregelblok11.dwg"
    If frmnaregelblok.groepinvoerBox1 = 12 Then bestand = "c:\acad2002\dwg\naregelblok12.dwg"
    If frmnaregelblok.groepinvoerBox1 = 13 Then bestand = "c:\acad2002\dwg\naregelblok13.dwg"
    If frmnaregelblok.groepinvoerBox1 = 14 Then bestand = "c:\acad2002\dwg\naregelblok14.dwg"
    If frmnaregelblok.groepinvoerBox1 = 15 Then bestand = "c:\acad2002\dwg\naregelblok15.dwg"
    If frmnaregelblok.groepinvoerBox1 = 16 Then bestand = "c:\acad2002\dwg\naregelblok16.dwg"
    If frmnaregelblok.groepinvoerBox1 = 17 Then bestand = "c:\acad2002\dwg\naregelblok17.dwg"
    If frmnaregelblok.groepinvoerBox1 = 18 Then bestand = "c:\acad2002\dwg\naregelblok18.dwg"
    If frmnaregelblok.groepinvoerBox1 = 19 Then bestand = "c:\acad2002\dwg\naregelblok19.dwg"
    If frmnaregelblok.groepinvoerBox1 = 20 Then bestand = "c:\acad2002\dwg\naregelblok20.dwg"
                    
 pb1 = ThisDrawing.Utility.GetPoint(, "Plaats startpunt")
 Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb1, bestand, scaal, scaal, 1, 0)
 
 If Err Then
    ThisDrawing.ObjectSnapMode = True
    groepinvoerBox1.SetFocus
    frmnaregelblok.show
    Exit Sub
    End If

 blockRefObj.Update
 
 
 Z = Val(groepinvoerBox2.Value)
 If Z > 0 And Z < 10 Then Z = "0" & groepinvoerBox2.Value
 f = "Regelunit" & " " & Z
  
  
For Each element20 In ThisDrawing.ModelSpace
      If element20.ObjectName = "AcDbBlockReference" Then
      If UCase(element20.Name) = "NAREGELBLOK2" Or UCase(element20.Name) = "NAREGELBLOK3" Or UCase(element20.Name) = "NAREGELBLOK4" Or _
         UCase(element20.Name) = "NAREGELBLOK5" Or UCase(element20.Name) = "NAREGELBLOK6" Or UCase(element20.Name) = "NAREGELBLOK7" Or _
         UCase(element20.Name) = "NAREGELBLOK8" Or UCase(element20.Name) = "NAREGELBLOK9" Or UCase(element20.Name) = "NAREGELBLOK10" Or _
         UCase(element20.Name) = "NAREGELBLOK11" Or UCase(element20.Name) = "NAREGELBLOK12" Or UCase(element20.Name) = "NAREGELBLOK13" Or _
         UCase(element20.Name) = "NAREGELBLOK14" Or UCase(element20.Name) = "NAREGELBLOK15" Or UCase(element20.Name) = "NAREGELBLOK16" Or _
         UCase(element20.Name) = "NAREGELBLOK17" Or UCase(element20.Name) = "NAREGELBLOK18" Or UCase(element20.Name) = "NAREGELBLOK19" Or _
         UCase(element20.Name) = "NAREGELBLOK20" Then
      Set SYMBOOL = element20
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
          If ATTRIBUUT.TagString = "TYPE_RNU" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = f
          If ATTRIBUUT.TagString = "TYPE_RU" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = groepinvoerBox2.Value
          
          If ATTRIBUUT.TagString = "GROEPSNUMMER1" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox200
          If ATTRIBUUT.TagString = "RUIMTENAAM1" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox200
          If ATTRIBUUT.TagString = "NAGEREGELD1" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox200.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER2" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox201
          If ATTRIBUUT.TagString = "RUIMTENAAM2" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox201
          If ATTRIBUUT.TagString = "NAGEREGELD2" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox201.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER3" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox202
          If ATTRIBUUT.TagString = "RUIMTENAAM3" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox202
          If ATTRIBUUT.TagString = "NAGEREGELD3" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox202.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER4" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox203
          If ATTRIBUUT.TagString = "RUIMTENAAM4" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox203
          If ATTRIBUUT.TagString = "NAGEREGELD4" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox203.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER5" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox204
          If ATTRIBUUT.TagString = "RUIMTENAAM5" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox204
          If ATTRIBUUT.TagString = "NAGEREGELD5" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox204.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER6" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox205
          If ATTRIBUUT.TagString = "RUIMTENAAM6" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox205
          If ATTRIBUUT.TagString = "NAGEREGELD6" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox205.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER7" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox206
          If ATTRIBUUT.TagString = "RUIMTENAAM7" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox206
          If ATTRIBUUT.TagString = "NAGEREGELD7" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox206.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER8" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox207
          If ATTRIBUUT.TagString = "RUIMTENAAM8" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox207
          If ATTRIBUUT.TagString = "NAGEREGELD8" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox207.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER9" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox208
          If ATTRIBUUT.TagString = "RUIMTENAAM9" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox208
          If ATTRIBUUT.TagString = "NAGEREGELD9" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox208.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER10" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox209
          If ATTRIBUUT.TagString = "RUIMTENAAM10" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox209
          If ATTRIBUUT.TagString = "NAGEREGELD10" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox209.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER11" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox210
          If ATTRIBUUT.TagString = "RUIMTENAAM11" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox210
          If ATTRIBUUT.TagString = "NAGEREGELD11" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox210.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER12" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox211
          If ATTRIBUUT.TagString = "RUIMTENAAM12" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox211
          If ATTRIBUUT.TagString = "NAGEREGELD12" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox211.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER13" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox212
          If ATTRIBUUT.TagString = "RUIMTENAAM13" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox212
          If ATTRIBUUT.TagString = "NAGEREGELD13" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox212.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER14" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox213
          If ATTRIBUUT.TagString = "RUIMTENAAM14" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox213
          If ATTRIBUUT.TagString = "NAGEREGELD14" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox213.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER15" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox214
          If ATTRIBUUT.TagString = "RUIMTENAAM15" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox214
          If ATTRIBUUT.TagString = "NAGEREGELD15" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox214.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER16" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox215
          If ATTRIBUUT.TagString = "RUIMTENAAM16" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox215
          If ATTRIBUUT.TagString = "NAGEREGELD16" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox215.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER17" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox216
          If ATTRIBUUT.TagString = "RUIMTENAAM17" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox216
          If ATTRIBUUT.TagString = "NAGEREGELD17" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox216.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER18" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox217
          If ATTRIBUUT.TagString = "RUIMTENAAM18" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox217
          If ATTRIBUUT.TagString = "NAGEREGELD18" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox217.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER19" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox218
          If ATTRIBUUT.TagString = "RUIMTENAAM19" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox218
          If ATTRIBUUT.TagString = "NAGEREGELD19" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox218.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
          If ATTRIBUUT.TagString = "GROEPSNUMMER20" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = TextBox219
          If ATTRIBUUT.TagString = "RUIMTENAAM20" And ATTRIBUUT.textstring = "" Then ATTRIBUUT.textstring = ComboBox219
          If ATTRIBUUT.TagString = "NAGEREGELD20" And ATTRIBUUT.textstring = "-" Then
                         If CheckBox219.Value = True Then
                                ATTRIBUUT.textstring = "NAGEREGELD"
                                Else
                                ATTRIBUUT.textstring = " "
                         End If
          End If
         Next i
       
      End If
      End If
      End If
  Next element20
 Update
  
 Call zoekblok2
 groepinvoerBox1.Value = Clear 'aantal groepen invoerveld
 groepinvoerBox2.Value = Clear 'regelunit invoerveld
 groepinvoerBox2.SetFocus
 Label2.Caption = Clear
 ThisDrawing.ObjectSnapMode = False
 frmnaregelblok.show


End Sub

Private Sub resetButton_Click()
Call zoekblok2
 groepinvoerBox1.Value = Clear 'aantal groepen invoerveld
 groepinvoerBox2.Value = Clear 'regelunit invoerveld
 groepinvoerBox2.SetFocus
 'Label2.Caption = Clear
 TextBox1 = ""
' OptionButton1.Value = False
' OptionButton2.Value = False
 CheckBox5.Value = False
 
 
    frmnaregelblok.TextBox200.Value = Clear: frmnaregelblok.ComboBox200.Value = Clear: frmnaregelblok.CheckBox200.Value = True
    frmnaregelblok.TextBox201.Value = Clear: frmnaregelblok.ComboBox201.Value = Clear: frmnaregelblok.CheckBox201.Value = True
    frmnaregelblok.TextBox202.Value = Clear: frmnaregelblok.ComboBox202.Value = Clear: frmnaregelblok.CheckBox202.Value = True
    frmnaregelblok.TextBox203.Value = Clear: frmnaregelblok.ComboBox203.Value = Clear: frmnaregelblok.CheckBox203.Value = True
    frmnaregelblok.TextBox204.Value = Clear: frmnaregelblok.ComboBox204.Value = Clear: frmnaregelblok.CheckBox204.Value = True
    frmnaregelblok.TextBox205.Value = Clear: frmnaregelblok.ComboBox205.Value = Clear: frmnaregelblok.CheckBox205.Value = True
    frmnaregelblok.TextBox206.Value = Clear: frmnaregelblok.ComboBox206.Value = Clear: frmnaregelblok.CheckBox206.Value = True
    frmnaregelblok.TextBox207.Value = Clear: frmnaregelblok.ComboBox207.Value = Clear: frmnaregelblok.CheckBox207.Value = True
    frmnaregelblok.TextBox208.Value = Clear: frmnaregelblok.ComboBox208.Value = Clear: frmnaregelblok.CheckBox208.Value = True
    frmnaregelblok.TextBox209.Value = Clear: frmnaregelblok.ComboBox209.Value = Clear: frmnaregelblok.CheckBox209.Value = True
    frmnaregelblok.TextBox210.Value = Clear: frmnaregelblok.ComboBox210.Value = Clear: frmnaregelblok.CheckBox210.Value = True
    frmnaregelblok.TextBox211.Value = Clear: frmnaregelblok.ComboBox211.Value = Clear: frmnaregelblok.CheckBox211.Value = True
    frmnaregelblok.TextBox212.Value = Clear: frmnaregelblok.ComboBox212.Value = Clear: frmnaregelblok.CheckBox212.Value = True
    frmnaregelblok.TextBox213.Value = Clear: frmnaregelblok.ComboBox213.Value = Clear: frmnaregelblok.CheckBox213.Value = True
    frmnaregelblok.TextBox214.Value = Clear: frmnaregelblok.ComboBox214.Value = Clear: frmnaregelblok.CheckBox214.Value = True
    frmnaregelblok.TextBox215.Value = Clear: frmnaregelblok.ComboBox215.Value = Clear: frmnaregelblok.CheckBox215.Value = True
    frmnaregelblok.TextBox216.Value = Clear: frmnaregelblok.ComboBox216.Value = Clear: frmnaregelblok.CheckBox216.Value = True
    frmnaregelblok.TextBox217.Value = Clear: frmnaregelblok.ComboBox217.Value = Clear: frmnaregelblok.CheckBox217.Value = True
    frmnaregelblok.TextBox218.Value = Clear: frmnaregelblok.ComboBox218.Value = Clear: frmnaregelblok.CheckBox218.Value = True
    frmnaregelblok.TextBox219.Value = Clear: frmnaregelblok.ComboBox219.Value = Clear: frmnaregelblok.CheckBox219.Value = True
 
 
End Sub
Sub Schaal(scaal)
frmnaregelblok.Hide
On Error Resume Next
vartab = ThisDrawing.GetVariable("EXTMAX")
If vartab(0) >= 2145 And vartab(0) <= 2155 Or vartab(0) >= 4095 And vartab(0) <= 4105 _
    Or vartab(0) >= 5835 And vartab(0) <= 5845 Or vartab(0) >= 8355 And vartab(0) <= 8365 _
    Or vartab(0) >= 11785 And vartab(0) <= 11795 Or vartab(0) >= 15685 And vartab(0) <= 15695 Then
    scaal = 2
Else
    scaal = 1
End If
If vartab(0) >= 16715 And vartab(0) <= 16725 Or vartab(0) >= 23575 And vartab(0) <= 23585 _
    Or vartab(0) >= 31375 And vartab(0) <= 31385 Then
    scaal = 4
End If
Update
End Sub



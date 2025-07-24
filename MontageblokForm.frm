VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MontageblokForm 
   Caption         =   "MONTAGEBLOK"
   ClientHeight    =   5124
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11088
   OleObjectBlob   =   "MontageblokForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MontageblokForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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



Private Sub CommandButton2_Click()
    Dim retobj As Object
    Dim Pbase As Variant
 On Error Resume Next
 
MontageblokForm.Hide

opnieuw5:

ThisDrawing.Utility.GetEntity retobj, Pbase, "Selecteer een blok."

 If retobj.ObjectName = "AcDbBlockReference" Then
      If retobj.Name = "Montagetsblok" Then
      Set symbool = retobj
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        
        If Err Then
          MontageblokForm.Show
        TextBox1 = TextBox1 - 1
        Exit Sub
        End If
        
        attribuut.textstring = TextBox1 'REGELUNITNUMMER
        
        
       Next j
       
        End If
      End If
      End If

TextBox1 = TextBox1 + 1
GoTo opnieuw5
End Sub

Sub ESelectOnScreen()
  
End Sub
Private Sub TextBox1_Change()
On Error Resume Next
Dim c As Double
c = TextBox1

If Err Then
   TextBox1 = Clear
  Exit Sub
  End If
End Sub

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
  Call zoekblok3
  
  Dim lognaam
   lognaam = ThisDrawing.GetVariable("loginname")
   lognaam = UCase(lognaam)
   If lognaam = "GERARD" Then
   TextBox2.Visible = True
   CheckBox5.Visible = True
   'CheckBox6.Visible = True
   End If
  
End Sub

Private Sub CancelButton_Click()
Call zoekblok2
 groepinvoerBox1.Value = Clear
 groepinvoerBox2.Value = Clear
groepinvoerBox2.SetFocus
Label2.Caption = Clear
Unload Me
'ThisDrawing.SendCommand "_vbaunload" & vbCr & "montageblok.dvb" & vbCr
End Sub

Private Sub groepinvoerBox1_Change()
Dim a As Double
On Error Resume Next
a = groepinvoerBox1.Text
If groepinvoerBox1 <> "" And groepinvoerBox2 <> "" Then PlaatsButton.Enabled = True
If groepinvoerBox1.Value = Clear Then Call ThisDrawing.imagelijst1
If a = 1 Then Call ThisDrawing.imagelijst2
If a = 2 Then Call ThisDrawing.imagelijst3
If a = 3 Then Call ThisDrawing.imagelijst4
If a = 4 Then Call ThisDrawing.imagelijst5
If a = 5 Then Call ThisDrawing.imagelijst6
If a = 6 Then Call ThisDrawing.imagelijst7
If a = 7 Then Call ThisDrawing.imagelijst8
If a = 8 Then Call ThisDrawing.imagelijst9
If a = 9 Then Call ThisDrawing.imagelijst10
If a = 10 Then Call ThisDrawing.imagelijst11
If a = 11 Then Call ThisDrawing.imagelijst12
If a = 12 Then Call ThisDrawing.imagelijst13
If a = 13 Then Call ThisDrawing.imagelijst14
If a = 14 Then Call ThisDrawing.imagelijst15
If a = 15 Then Call ThisDrawing.imagelijst16
If a = 16 Then Call ThisDrawing.imagelijst17
If a = 17 Then Call ThisDrawing.imagelijst18
If a = 18 Then Call ThisDrawing.imagelijst19
If a = 19 Then Call ThisDrawing.imagelijst20
If a = 20 Then Call ThisDrawing.imagelijst21

If Err Then
   groepinvoerBox1 = Clear
   PlaatsButton.Enabled = False
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
   
End Sub
Private Sub groepinvoerBox2_Change()
Dim b As Double
On Error Resume Next
groepinvoerBox2.SetFocus

b = groepinvoerBox2.Text
c = "Regelunit" & " " & b
Label2.Caption = c

If Err Then
   groepinvoerBox2 = Clear
   Label2.Caption = Clear
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
End Sub
Sub zoekblok1()
For Each element20 In ThisDrawing.ModelSpace
      If element20.ObjectName = "AcDbBlockReference" Then
      'If UCase(element20.Name) = "MAT_SPE_ZD" Or UCase(element20.Name) = "MAT_SPE_PE" Then
      If element20.Name = "Mat_spe_ZD" Or element20.Name = "Mat_spe_PE" Or element20.Name = "Mat_spe_PE800" _
      Or element20.Name = "Mat_spe_ALU" Or element20.Name = "Mat_spe_ZDringleiding" Or element20.Name = "Mat_spe_PEringleiding" Or _
      element20.Name = "Mat_spe_ALUringleiding" Or element20.Name = "Mat_spe_FLEX" Or element20.Name = "Mat_spe_ZD_1627" Or _
      element20.Name = "Mat_spe_ZD_1627500" Then

      
      Set symbool = element20
        
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
          If attribuut.TagString = "RNU" Then
               rnc = attribuut.textstring
               rnc2 = Len(rnc)
               If rnc2 = 1 Then CheckBox3.Value = True
           End If
             If Val(groepinvoerBox2) = rnc Then
              b0 = element20.InsertionPoint
              For l = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(l)
              element20.Highlight (True)
                          
          If attribuut.TagString = "REGELUNITTYPE" Then
          rgl = attribuut.textstring 'TYPE REGELUNIT
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
              Next l
                
         End If
       Next I
       
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
     'Unload Me
     Exit Sub
     
 End If
 

 Dim b1(0 To 2) As Double
  b1(0) = b0(0) - 1500 '1000
  b1(1) = b0(1) - 400  '500
  b1(2) = 0
 
 Dim b2(0 To 2) As Double
  b2(0) = b0(0) + 500
  b2(1) = b0(1) + 1000 '750
  b2(2) = 0
  
  
''''  Dim lognaam
''''  lognaam = ThisDrawing.GetVariable("loginname")
''''  lognaam = UCase(lognaam)
''''  If lognaam = "GERARD" And CheckBox5.Value = False Then
''''  ZoomWindow b1, b2
''''  End If
  
  If CheckBox5.Value = False Then ZoomWindow b1, b2
   
Call PlaatsButton_Click

End Sub
Sub zoekblok2()
For Each element20 In ThisDrawing.ModelSpace
      If element20.ObjectName = "AcDbBlockReference" Then
      'If UCase(element20.Name) = "MAT_SPE_ZD" Or UCase(element20.Name) = "MAT_SPE_PE" Then
       If element20.Name = "Mat_spe_ZD" Or element20.Name = "Mat_spe_PE" Or element20.Name = "Mat_spe_PE800" _
      Or element20.Name = "Mat_spe_ALU" Or element20.Name = "Mat_spe_ZDringleiding" Or element20.Name = "Mat_spe_PEringleiding" Or _
      element20.Name = "Mat_spe_ALUringleiding" Or element20.Name = "Mat_spe_FLEX" Then
      Set symbool = element20
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
          If attribuut.TagString = "RNU" Then rnc = Val(attribuut.textstring)
             If Val(groepinvoerBox2) = rnc Then element20.Highlight (False)
         Next I
       
      End If
      End If
      End If
  Next element20
 Update
End Sub
Sub zoekblok3()
For Each element20 In ThisDrawing.ModelSpace
      If element20.ObjectName = "AcDbBlockReference" Then
      'If UCase(element20.Name) = "MAT_SPE_ZD" Or UCase(element20.Name) = "MAT_SPE_PE" Then
       If element20.Name = "Mat_spe_ZD" Or element20.Name = "Mat_spe_PE" Or element20.Name = "Mat_spe_PE800" _
      Or element20.Name = "Mat_spe_ALU" Or element20.Name = "Mat_spe_ZDringleiding" Or element20.Name = "Mat_spe_PEringleiding" Or _
      element20.Name = "Mat_spe_ALUringleiding" Or element20.Name = "Mat_spe_FLEX" Then
      Set symbool = element20
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
          If attribuut.TagString = "RNU" Then
               rnc = attribuut.textstring
               rnc2 = Len(rnc)
               If rnc2 = 1 Then CheckBox3.Value = True
           End If
             
           Next I
       
      End If
      End If
      End If
  Next element20
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
     MsgBox "Een 1 groeps montageblok vermelden we niet op de tekening..!!!!", vbCritical
     groepinvoerBox1 = Clear
     groepinvoerBox1.SetFocus
     PlaatsButton.Enabled = False
     Exit Sub
     End If
     

Dim newLayer As AcadLayer
Set newLayer = ThisDrawing.Layers.Add("MONTAGEBLOK")
ThisDrawing.ActiveLayer = newLayer
Update

If MontageblokForm.TextBox2 <> "" Then scaal = MontageblokForm.TextBox2
 MontageblokForm.Hide
 On Error Resume Next
 Dim bestand As String
 bestand = "c:\acad2002\dwg\Montagetsblok.dwg"
 pb1 = ThisDrawing.Utility.GetPoint(, "Plaats startpunt")
   
'''''   If OptionButton1.Value = True Or OptionButton2.Value = True Then
'''''    Dim ssetObj As AcadSelectionSet
'''''    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")
'''''    pbz = ThisDrawing.Utility.GetPoint(, "Plaats eindpunt")
'''''    ssetObj.Select acSelectionSetCrossing, pbz, pb1
'''''    ssetObj.Erase
'''''   End If
 
 
 Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb1, bestand, scaal, scaal, 1, 0)
 
 If OptionButton1.Value = True Then  'autonummering
    ' Get the attributes for the block reference
    Dim varAttributes As Variant
    varAttributes = blockRefObj.GetAttributes

    ' Move the attribute tags and values into a string to be displayed in a Msgbox
    Dim strAttributes As String
    Dim I As Integer
    For I = LBound(varAttributes) To UBound(varAttributes)
        strAttributes = strAttributes & "  Tag: " & varAttributes(I).TagString & _
                        "   Value: " & varAttributes(I).textstring & "    "
    Next
    'MsgBox "The attributes for blockReference " & blockrefobj.Name & " are: " & strAttributes, , ""

     ' Change the value of the attribute
    ' Note: There is no SetAttributes. Once you have the variant array, you have the objects.
    ' Changing them changes the objects in the drawing.
   'varAttributes(0).textstring = "NEW VALUE!"
  varAttributes(0).textstring = groepinvoerBox1
 End If 'autonummering
 
 If OptionButton2.Value = True Then  'autonummering
    ' Get the attributes for the block reference
    'Dim varAttributes As Variant
    varAttributes = blockRefObj.GetAttributes

    ' Move the attribute tags and values into a string to be displayed in a Msgbox
    'Dim strAttributes As String
    'Dim I As Integer
    For I = LBound(varAttributes) To UBound(varAttributes)
        strAttributes = strAttributes & "  Tag: " & varAttributes(I).TagString & _
                        "   Value: " & varAttributes(I).textstring & "    "
    Next
    'MsgBox "The attributes for blockReference " & blockrefobj.Name & " are: " & strAttributes, , ""

     ' Change the value of the attribute
    ' Note: There is no SetAttributes. Once you have the variant array, you have the objects.
    ' Changing them changes the objects in the drawing.
   'varAttributes(0).textstring = "NEW VALUE!"
  varAttributes(0).textstring = "1"
 End If 'autonummering
 
 
 
 
 
 
 If Err Then
    ThisDrawing.ObjectSnapMode = True
    groepinvoerBox1.SetFocus
    MontageblokForm.Show
    Exit Sub
    End If

 blockRefObj.Update
 
  
  a = groepinvoerBox1 'aantal groepen
  b = a - 1  'aantal groepen -1
  qw = b
  
  If OptionButton2.Value = True Then
  a = groepinvoerBox1 'aantal groepen
  b = a - 1  'aantal groepen -1
  qw = 2
  End If
  
  
  
 For I = 1 To b
      ' Change the value of the attribute
    ' Note: There is no SetAttributes. Once you have the variant array, you have the objects.
    ' Changing them changes the objects in the drawing.
    If OptionButton1.Value = True Then varAttributes(0).textstring = qw
    If OptionButton2.Value = True Then varAttributes(0).textstring = qw
    
 pe2(0) = pb1(0) - ((I * 44.6) * scaal)
 pe2(1) = pb1(1)
 pe2(2) = 0
 Set nieuwelement = blockRefObj.Copy
 nieuwelement.Move pb1, pe2
 If OptionButton1.Value = True Then qw = qw - 1
 If OptionButton2.Value = True Then qw = qw + 1
 Next I
 ' MEESTE RECHTSE BLOK
 If OptionButton1.Value = True Then varAttributes(0).textstring = groepinvoerBox1  'OPTIONBUTTON1.Value = True
 If OptionButton2.Value = True Then varAttributes(0).textstring = "1" 'CheckBox4.Value = True
 nieuwelement.Update
 
 d = (a * 44.6)
 Dim LijnObj As AcadPolyline
' Dim Lijnobj As Object
 Dim pb2(0 To 2) As Double 'punt van rechthoek
 Dim pb3(0 To 2) As Double 'punt van rechthoek
 Dim pb4(0 To 2) As Double 'punt van rechthoek
 Dim pb5(0 To 2) As Double 'punt voor regelunittekst
 
 pb1(0) = pb1(0) - 3    'montage blok op zelfde hoogte als insertionpoint
 pb1(1) = pb1(1) + 3    'montage blok op zelfde hoogte als insertionpoint
 pb1(2) = pb1(2)        'montage blok op zelfde hoogte als insertionpoint
 
 
 pb3(0) = pb1(0) - ((d + 100) * scaal)
 pb3(1) = pb1(1) + (250 * scaal)
 pb3(2) = 0
 
 pb2(0) = pb1(0)
 pb2(1) = pb3(1)
 pb2(2) = 0
 
 pb4(0) = pb3(0)
 pb4(1) = pb1(1)
 pb4(2) = 0
 
 pb5(0) = pb3(0)
 pb5(1) = pb3(1) + (13 * scaal)
 pb5(2) = 0
 


   
Dim points(0 To 14) As Double
points(0) = pb1(0): points(1) = pb1(1): points(2) = pb1(2)
points(3) = pb2(0): points(4) = pb2(1): points(5) = pb2(2)
points(6) = pb3(0): points(7) = pb3(1): points(8) = pb3(2)
points(9) = pb4(0): points(10) = pb4(1): points(11) = pb4(2)
points(12) = pb1(0): points(13) = pb1(1): points(14) = pb1(2)

 Set LijnObj = ThisDrawing.ModelSpace.AddPolyline(points)
 LijnObj.Offset (3)
 LijnObj.Update
 
 If MontageblokForm.CheckBox3.Value = False Then
 Z = Val(groepinvoerBox2.Value)
 If Z > 0 And Z < 10 Then Z = "0" & groepinvoerBox2.Value
 End If
 If MontageblokForm.CheckBox3.Value = True Then
    Z = groepinvoerBox2.Value
 End If
 
 Dim E As Double
 E = groepinvoerBox2.Text
 f = "Regelunit" & " " & Z
 
 Dim textobj As AcadText
 Set textobj = ThisDrawing.ModelSpace.AddText(f, pb5, (15 * scaal))
 textobj.Update
 
 Dim pb6(0 To 2) As Double
   
 E = (d \ 2)
 pb6(0) = pb3(0) + ((E + 48) * scaal)
 pb6(1) = pb3(1) - (30 * scaal)
 pb6(2) = 0
 
 'Dim textObj As AcadText
 Set textobj = ThisDrawing.ModelSpace.AddText("AANSLUITSCHEMA", pb6, (14 * scaal))
 textobj.Alignment = acAlignmentMiddleCenter
 textobj.TextAlignmentPoint = pb6
 textobj.Update
   
 pb6(0) = pb6(0)
 pb6(1) = pb6(1) - (25 * scaal)
 pb6(2) = 0
 
 Set textobj = ThisDrawing.ModelSpace.AddText("MONTEURS WTH", pb6, (14 * scaal))
 textobj.Alignment = acAlignmentMiddleCenter
 textobj.TextAlignmentPoint = pb6
 textobj.Update
 
 Call zoekblok2
 groepinvoerBox1.Value = Clear 'aantal groepen invoerveld
   If CheckBox6.Enabled = True Then
        groepinvoerBox2 = groepinvoerBox2 + 1
        Else
        groepinvoerBox2.Value = Clear 'regelunit invoerveld
   End If
 groepinvoerBox2.SetFocus
 Label2.Caption = Clear
 ThisDrawing.ObjectSnapMode = False
 MontageblokForm.Show
If CheckBox6.Enabled = True Then
    MontageblokForm.Hide
    Call zoekblok1
End If
End Sub

Private Sub resetButton_Click()
Call zoekblok2
 groepinvoerBox1.Value = Clear 'aantal groepen invoerveld
 groepinvoerBox2.Value = Clear 'regelunit invoerveld
 groepinvoerBox2.SetFocus
 Label2.Caption = Clear
 TextBox1 = ""
 OptionButton1.Value = False
 OptionButton2.Value = False
 CheckBox5.Value = False
End Sub
Sub Schaal(scaal)
MontageblokForm.Hide
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



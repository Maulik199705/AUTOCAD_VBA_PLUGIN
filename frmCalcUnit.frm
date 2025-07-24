VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalcUnit 
   Caption         =   "Materiaalspecificatie"
   ClientHeight    =   4080
   ClientLeft      =   828
   ClientTop       =   432
   ClientWidth     =   3804
   OleObjectBlob   =   "frmCalcUnit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalcUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'G.C.Haak
'kruisje uitschakelen
'BESTANDEN: GROEPTEKSTBLOK.DWG
' "c:\acad2002\dwg\Mat_calculatie.dwg"
' DWG - Herz_calc|Ruh-r_calc|Ruh-N_calc|ruh-rt_calc|rub-r_calc|rub-rt_calc|rubk-r_calc|rubk-rt_calc|LT_calc|lt-s_calc|lt-n_calc
' DWG - RUW_calc|ruv_calc|ruh-s_calc|rub-s_calc|kmv_calc

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

Call combolijst1 'combobox1 vullen
Call Combolijst2 'combobox2 vullen
Call Combolijst3 'combobox3 vullen
Call combolijst4 'combobox4 vullen
Call combolijst5 'combobox5 vullen
Call combolijst6 'combobox6 vullen
Call leidingtype

ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "gt" & vbCr & "ON" & vbCr & "gt" & vbCr & "T" & vbCr & "gt" & vbCr & vbCr
Label33 = 0
TextBox9.SetFocus

Dim newLayer As AcadLayer
Set newLayer = ThisDrawing.Layers.Add("GT")
ThisDrawing.ActiveLayer = newLayer

End Sub

Private Sub ComboBox1_Change()
If ComboBox1 = "IFD-Polystyreen" Then ComboBox4 = "PE-RT 14*2 mm"
End Sub
Private Sub ComboBox2_Change()
'type unit
If ComboBox2.Value = "RUH-R" Or ComboBox2.Value = "RUH-RT" Then
  ComboBox3.Enabled = True
  Else
  ComboBox3.Enabled = False
 End If
End Sub
Private Sub CommandButton2_Click()
Unload Me
frmUnitlogo.Show
End Sub


Private Sub commandbutton6_click()
Call Update_calcmodule.update_unitlogo_calc
Unload Me
End Sub
Sub leidingtype()
Dim element As Object
Dim layerObj As AcadLayer

For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If UCase(element.Name) = "MAT_CALCULATIE" Then
      
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
                
         If attribuut.TagString = "BEVESTIGINGSTYPE" Then ComboBox1 = attribuut.textstring
         If attribuut.TagString = "TB" Then ComboBox4 = attribuut.textstring
         If attribuut.TagString = "AFWERKVLOER" Then ComboBox5 = attribuut.textstring
         If attribuut.TagString = "EXTRA_MATERIALEN" Then ComboBox6 = attribuut.textstring
         
         If attribuut.TagString = "REGELUNITTYPE" Then
          RT = attribuut.textstring
          CONTROLE = InStr(1, RT, "/", vbBinaryCompare)  'staat er een komma in??
            If CONTROLE <> 0 Then
               trimstring = Split(RT, ("/"))
               'MsgBox trimstring(0) & trimstring(1)
               ComboBox2 = trimstring(0)
               ComboBox3 = trimstring(1)
            End If
            If CONTROLE = 0 Then ComboBox2 = RT
         End If
         
       Next I
       
        End If
      End If
      End If
  Next element

End Sub
Private Sub cmdAfsluiten_Click()
Unload Me
End Sub
Private Sub TextBox9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox9 = ""
End Sub
Private Sub TextBox10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox10 = ""
End Sub
Private Sub TextBox9_Change()
Dim b As Double
On Error Resume Next
TextBox9.SetFocus
b = TextBox9.Text

If Err Then
   TextBox9 = Clear
   CmdBloklogo.Enabled = False
   Exit Sub
  End If
If TextBox9.Text = "" Then CmdBloklogo.Enabled = False
If TextBox9.Text <> "" Then CmdBloklogo.Enabled = True
Label33 = 0

TextBox27 = Clear: TextBox27.Visible = False

End Sub
Private Sub TextBox11_Change()
On Error Resume Next
Dim TA As Double
TA = TextBox11

'----30-6-2004-aangepast omdat ie de rollen niet goed optelde in sommige gevallen
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + (Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16)) + (Val(TextBox17))
'----30-6-2004-aangepast omdat ie de rollen niet goed optelde in sommige gevallen
Update
If Err Then
   TextBox11 = "0"
  Exit Sub
  End If

If TA > 20 And (frmCalcUnit.OptionButton7 = False And frmCalcUnit.OptionButton8 = False) Then
   'TextBox11.BackColor = &HFFFF&
   MsgBox "PAS OP!!, MEER DAN 20 groepen....!!!!", vbCritical
   TextBox1 = "0": TextBox2 = "0": TextBox3 = "0": TextBox4 = "0"
   TextBox5 = "0": TextBox6 = "0": TextBox7 = "0": TextBox15 = "0": TextBox16 = "0": TextBox11 = "0"
   Call RESET
   'TextBox11.BackColor = &HFFFFFF
   Exit Sub
End If

If Label33 = 0 Then
  If TextBox11 > 15 And ComboBox2 = "KMV" And (frmCalcUnit.OptionButton7 = False And frmCalcUnit.OptionButton8 = False) Then
   MsgBox "PAS OP!!, KMV-verdeler gaat maar tot 15 groepen....!!!!", vbCritical
   Label33 = 1
  End If
End If

If Label33 = 0 Then
  If TextBox11 > 16 And ComboBox2 = "HERZ" And (frmCalcUnit.OptionButton7 = False And frmCalcUnit.OptionButton8 = False) Then
   MsgBox "PAS OP!!, HERZ-verdeler gaat maar tot 16 groepen....!!!!", vbCritical
   Label33 = 1
  End If
End If

If TextBox11 <> "0" Then TextBox17.Enabled = True


End Sub
Private Sub TextBox12_Change()
Dim C As Double
On Error Resume Next

C = TextBox12.Text

If Err Then
   TextBox12 = Clear
   Exit Sub
  End If
End Sub
Sub combolijst1()
ComboBox1.AddItem "Vlechtdraad"
ComboBox1.AddItem "Witmarmerbeugels"
ComboBox1.AddItem "Beugels/Nagels"
ComboBox1.AddItem "Eigen middelen"
ComboBox1.AddItem "IFD-Polystyreen"
ComboBox1.AddItem "Isoclips"
ComboBox1.AddItem "Keg"
ComboBox1.AddItem "Montagestrip"
ComboBox1.AddItem "Noppenplaat"
ComboBox1.AddItem "Schietbeugels"
ComboBox1.AddItem "Ty-raps"
ComboBox1.AddItem "Varisoclips"
ComboBox1.ListIndex = 1

End Sub
Sub Combolijst2()
ComboBox2.AddItem "HERZ"
ComboBox2.AddItem "KMV"
ComboBox2.AddItem "LT"
ComboBox2.AddItem "LT-N"
ComboBox2.AddItem "LTS" 'aangepast
ComboBox2.AddItem "LTS-N" 'aangepast
ComboBox2.AddItem "LT-VK" 'aangepast
ComboBox2.AddItem "RUB-R"
ComboBox2.AddItem "RUB-RT"
ComboBox2.AddItem "RUBK-R"
ComboBox2.AddItem "RUBK-RT"
ComboBox2.AddItem "RUB-S"
ComboBox2.AddItem "RU-EE"
ComboBox2.AddItem "RUH-N"
ComboBox2.AddItem "RUH-R"
ComboBox2.AddItem "RUH-RT"
ComboBox2.AddItem "RUH-S"
ComboBox2.AddItem "RUV"
ComboBox2.AddItem "RU-WK"
ComboBox2.AddItem "RU-WKN" 'aangepast
ComboBox2.AddItem "RU-WKS" 'aangepast
ComboBox2.AddItem "RU-WW"
ComboBox2.AddItem "RU-WWN" 'aangepast
ComboBox2.AddItem "RU-WWS" 'aangepast
'ComboBox2.AddItem "VSKO"
ComboBox2.ListIndex = 15
ComboBox2.MatchEntry = fmMatchEntryFirstLetter
End Sub
Sub Combolijst3()
ComboBox3.AddItem "KT220"
ComboBox3.AddItem "KT24"
ComboBox3.AddItem "TH7420"
ComboBox3.ListIndex = 2
ComboBox3.MatchEntry = fmMatchEntryFirstLetter
End Sub
Sub combolijst4()
 ComboBox4.AddItem "ALUFLEX 14*2 mm"
 ComboBox4.AddItem "ALUFLEX 16*2 mm"
 ComboBox4.AddItem "ALUFLEX 18*2 mm"
 ComboBox4.AddItem "ALUFLEX 20*2 mm"
 ComboBox4.AddItem "PE-RT 10*1,25 mm"
 ComboBox4.AddItem "PE-RT 14*2 mm"
 ComboBox4.AddItem "PE-RT 16*2 mm"
 ComboBox4.AddItem "WTH-ZD 16*2,7 mm"
 ComboBox4.AddItem "WTH-ZD 20*3,4 mm"
 ComboBox4.ListIndex = 8
End Sub
Sub combolijst5()
 ComboBox5.AddItem "Zand-cement"
 ComboBox5.AddItem "Gietvloer"
 ComboBox5.AddItem "Monolitisch"
 ComboBox5.AddItem "Constructievloer"
 ComboBox5.ListIndex = 0
End Sub
Sub combolijst6()
 ComboBox6.AddItem "Geen"
 ComboBox6.AddItem "Bevestigingsnetten"
 ComboBox6.AddItem "Krimpnetten"
 ComboBox6.AddItem "Variso isolatie 20mm thermisch"
 ComboBox6.AddItem "Variso isolatie 20mm akoestisch"
 ComboBox6.AddItem "Variso isolatie 30mm thermisch"
 ComboBox6.AddItem "Variso isolatie 30mm akoestisch"
 ComboBox6.AddItem "Polystyreen isolatie"
 ComboBox6.AddItem "Folie"
 ComboBox6.AddItem "Noppenplaat"
 ComboBox6.ListIndex = 0
End Sub
Private Sub CmdBloklogo_Click()
On Error Resume Next
If ComboBox2 = "" Then
    MsgBox "Je bent 'Type unit' vergeten in te vullen..!!"
    ComboBox2.SetFocus
    Exit Sub
    End If
If ComboBox1 = "" Then
    MsgBox "Je bent 'Bevestigingsmateriaal' vergeten in te vullen..!!"
    ComboBox1.SetFocus
    Exit Sub
    End If
If ComboBox4 = "" Then
    MsgBox "Je bent 'Type buis' vergeten in te vullen..!!"
    ComboBox4.SetFocus
    Exit Sub
    End If

frmCalcUnit.Hide
TextBox9.Locked = False
    Set newLayer = ThisDrawing.Layers.Add("BLOKLOGO")
    ThisDrawing.ActiveLayer = newLayer
    Update
    ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "*" & vbCr & vbCr
    Update
Call Schaal(scaal)

 bestand = "c:\acad2002\dwg\Mat_calculatie.dwg"



If frmCalcUnit.TextBox29 <> "" Then scaal = frmCalcUnit.TextBox29
Dim pb1 As Variant
pb1 = ThisDrawing.Utility.GetPoint(, "Plaats startpunt....")
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb1, bestand, scaal, scaal, 1, 0)
''''If TextBox11 <> "0" And TextBox25 <> "0" Then Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb1, bestand20, scaal, scaal, 1, 0)
Update
If Err Then
    frmCalcUnit.Show
    Exit Sub
    End If

Dim element2 As Object

If ComboBox3.Enabled = False Then
REGELTYPE = ComboBox2 'ZONDER REGELING
Else
REGELTYPE = ComboBox2 & " / " & ComboBox3  'met regeling
End If
''If frmCalcUnit.OptionButton7 = True Or frmCalcUnit.OptionButton8 = True Then REGELTYPE = ComboBox2

unittel = frmCalcUnit.TextBox9

If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmCalcUnit.TextBox9
If unittel > 9 Then unitonder10 = frmCalcUnit.TextBox9


For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "MAT_CALCULATIE" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "RNU" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'REGELUNITNUMMER
        If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring = "" Then attribuut.textstring = REGELTYPE  'TYPE REGELUNIT
        If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox1  'BEVESTIGING
        If attribuut.TagString = "TB" And attribuut.textstring = "" Then attribuut.textstring = ComboBox4 'TYPE BUIS
        If attribuut.TagString = "AFWERKVLOER" And attribuut.textstring = "" Then attribuut.textstring = ComboBox5 'AFWERKVLOER
        If attribuut.TagString = "EXTRA_MATERIALEN" And attribuut.textstring = "" Then attribuut.textstring = ComboBox6 'EXTRA MATERIALEN
        Next I
       
      End If
      End If
      End If
  Next element2
 
 Update
 
   
  Dim pb2(0 To 2) As Double
 
  pb2(0) = pb1(0) - (scaal * 460)
  pb2(1) = pb1(1) + (scaal * 225) '177'210)
  pb2(2) = pb1(2)
  
  'juiste regelunitblokje.dwg inserten in de tekening
  bestand2 = ComboBox2 & "_calc.dwg"
  
  If ComboBox2 = "RUB-R" Or ComboBox2 = "RUB-RT" Then
     Call calc_Waarschuwing.waarschuwing
     If frmCalcUnit.TextBox30 = "2" Then bestand2 = "RUH-R" & "_calc.dwg"
  End If
  
 
  Call Calc_Unitblok.Unitblok(scaal, pb2, bestand2)
  Unload Me

End Sub
Sub MESSAGE(XWAARDE)

End Sub
Sub Schaal(scaal)
frmCalcUnit.Hide
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


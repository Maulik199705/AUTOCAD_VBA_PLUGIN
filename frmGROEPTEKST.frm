VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGROEPTEKST 
   Caption         =   "GROEPTEKST PLAATSEN "
   ClientHeight    =   11235
   ClientLeft      =   828
   ClientTop       =   432
   ClientWidth     =   13380
   OleObjectBlob   =   "frmGROEPTEKST.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGroeptekst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'17-12-2002 Meten van groepen
'M.Bosch en G.C.Haak
'kruisje uitschakelen
'BESTANDEN: GROEPTEKSTBLOK.DWG
'Mat_spe_zd.dwg | Mat_spe_pe.dwg | Mat_spe_flex.dwg |
'TXT en DWG - Herz|Ruh-r|Ruh-N|ruh-rt|rub-r|rub-rt|rubk-r|rubk-rt|LT|lt-s|lt-n
'TXT en DWG - RUW|ruw-klein.txt|ruw-groot.txt|ruv|ruh-s|rub-s|kmv

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
Private Sub TextBox28_Change()
Dim b As Double
On Error Resume Next
TextBox28.SetFocus
b = TextBox28.Text
If Err Then
   TextBox28 = "0"
Exit Sub
End If
End Sub

Private Sub TextBox30_Change()
If OptionButton9.Value = True And CheckBox2.Value = True Then texbox11 = TextBox30
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
  frmGroeptekst.CheckBox8.Visible = False

Call combolijst1 'combobox1 vullen
Call Combolijst2 'combobox2 vullen
Call Combolijst3 'combobox3 vullen
Call combolijst4 'combobox4 vullen
Call leidingtype
Call LEIDINGSOORT
Call wklengte

frmGroeptekst.Height = 400
frmGroeptekst.Width = 444
TextBox23 = ThisDrawing.GetVariable("userr4")
ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "gt" & vbCr & "ON" & vbCr & "gt" & vbCr & "T" & vbCr & "gt" & vbCr & vbCr
Label33 = 0
TextBox9.SetFocus
Label11.Caption = Clear 'wordt gevuld en uitgelezen met groepsnummer

Dim newLayer As AcadLayer
Set newLayer = ThisDrawing.Layers.Add("GT")
ThisDrawing.ActiveLayer = newLayer

Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "GERARD" Then
    'CheckBox9.Visible = True
    TextBox29.Visible = True
    OptionButton9.Visible = True
End If
If lognaam = "GERARD" Then frmGroeptekst.StartUpPosition = 0

End Sub
Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
     CheckBox6.Enabled = False
     CheckBox7.Enabled = False
End If
If CheckBox2.Value = False Then
     CheckBox6.Enabled = True
     CheckBox7.Enabled = True
End If

End Sub

Private Sub CheckBox6_Click()
If CheckBox6.Value = True Then CheckBox2.Enabled = False
If CheckBox6.Value = False Then CheckBox2.Enabled = True
End Sub

Private Sub CheckBox7_Click()
If CheckBox7 = True Then TextBox24.Visible = True
If CheckBox7 = False Then TextBox24.Visible = False
End Sub

Private Sub ComboBox1_Change()
If ComboBox1 = "IFD-Polystyreen" Then ComboBox4 = "PE-RT 14*2 mm"
End Sub


Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
 If OptionButton1.Value = True Then
   Label25.Caption = "WTHZD"
   Else
   Label25.Caption = "PE"
 End If
'OptionButton1.Enabled = False
'OptionButton2.Enabled = False
TextBox14.Visible = True: Label24.Visible = True: TextBox1.Visible = False
TextBox2.Visible = False: TextBox3.Visible = False: TextBox4.Visible = False
TextBox5.Visible = False: TextBox6.Visible = False: TextBox7.Visible = False
TextBox15.Visible = False: TextBox16.Visible = False
Label2.Visible = False: Label3.Visible = False: Label4.Visible = False
Label5.Visible = False: Label6.Visible = False: Label7.Visible = False
Label27.Visible = False: Label28.Visible = False
Label1.Caption = Clear
TextBox14.SetFocus
TextBox14 = Clear
End If

If CheckBox1.Value = False Then
'OptionButton1.Enabled = True: OptionButton2.Enabled = True
TextBox14.Visible = False: Label24.Visible = False
Label25.Caption = Clear
Call RESET
End If
End Sub
Private Sub OptionButton7_Click()
If frmGroeptekst.OptionButton7 = True Then
     cmdTelrollen.Enabled = True
     ComboBox2 = "RINGLEIDING"
     End If
End Sub
Private Sub OptionButton8_Click()
If frmGroeptekst.OptionButton8 = True Then
     cmdTelrollen.Enabled = True
     ComboBox2 = "RINGLEIDING"
     End If
End Sub
Private Sub CommandButton2_Click()
Unload Me
frmUnitlogo.show
End Sub

Private Sub CommandButton3_Click()
frmGroeptekst.Hide
frmKaderlogo.show
End Sub
Private Sub commandbutton4_click()
Call Checklayer2.Checklayer2
End Sub
Private Sub commandbutton6_click()
Call Update_module.update_unitlogo
Unload Me
End Sub
Private Sub OptionButton5_Click()
If OptionButton5.Value = True Then TextBox13.Visible = True
End Sub
Private Sub TextBox14_Change()
Label1.Caption = TextBox14 & " meter"
TextBox1.Visible = True
Dim TA As Double
On Error Resume Next
TextBox14.SetFocus
TA = TextBox14.Text
If Err Then
   TextBox14 = Clear
  Exit Sub
  End If
End Sub
Private Sub ToggleButton1_Click()
If ToggleButton1.Value = False Then
   ToggleButton1.ForeColor = RGB(0, 0, 0)
   ToggleButton1.Caption = "2,5 meter"
   End If
If ToggleButton1.Value = True Then
   ToggleButton1.ForeColor = RGB(255, 0, 0)
   ToggleButton1.Caption = "2 meter"
   OptionButton5.Value = False: TextBox13.Visible = False
End If
End Sub
Sub leidingtype()
Dim element As Object
Dim layerObj As AcadLayer

For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_PE" Or element.Name = "Mat_spe_ZD" Or element.Name = "Mat_spe_ALU" Or _
      element.Name = "Mat_spe_PEringleiding" Or element.Name = "Mat_spe_ZDringleiding" Or _
      element.Name = "Mat_spe_ALUringleiding" Or element.Name = "Mat_spe_PE800" _
      Or element.Name = "Mat_spe_FLEX" Or element.Name = "Mat_spe_ZD_1627" Or element.Name = "Mat_spe_ZD_1627500" Then
      
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
       
         If attribuut.TagString = "PE" Then
         OptionButton2.Value = True
         ComboBox4.Value = attribuut.textstring
         CheckBox8.Visible = True
         End If
         If attribuut.TagString = "PE" And Left(attribuut.textstring, 7) = "ALUFLEX" Then
         OptionButton6.Value = True
         ComboBox4.Value = attribuut.textstring
         End If
         If attribuut.TagString = "WTHZD" And (attribuut.textstring = "WTH-ZD 16 * 2,7 mm" Or attribuut.textstring = "WTH-ZD 16*2,7 mm") Then OptionButton9 = True
         If attribuut.TagString = "PE" And (attribuut.textstring = "WTH-ZD 16 * 2,7 mm" Or attribuut.textstring = "WTH-ZD 16*2,7 mm") Then OptionButton9 = True
         If attribuut.TagString = "WTHZD" And OptionButton1 = True Then ComboBox4.Value = attribuut.textstring
         If attribuut.TagString = "ALU" Then ComboBox4.Value = attribuut.textstring
         If attribuut.TagString = "BEVESTIGINGSTYPE" Then ComboBox1.Value = attribuut.textstring
         If attribuut.TagString = "LMETER" Then CheckBox8.Value = True
         
         If attribuut.TagString = "REGELUNITTYPE" Then
           RT = attribuut.textstring
         
          If RT <> "RINGLEIDING" Then
          trimstring = Split(RT, (" "))
          Dim mystr As Variant
          mystr = Len(trimstring(1))
          ComboBox2 = trimstring(0)
                 
           If mystr > 2 Then
           trimstring2 = Split(trimstring(1), ("/"))
           ComboBox3 = trimstring2(1)
           End If
         
          End If
         If RT = "RINGLEIDING" Then OptionButton7.Value = True

         
       End If
         
       Next I
       
        End If
      End If
      End If
  Next element
End Sub
Sub LEIDINGSOORT()
 For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOK" Then
                Set symbool = element
                If symbool.HasAttributes Then
                    attributen = symbool.GetAttributes
                    For I = LBound(attributen) To UBound(attributen)
                         Set attribuut = attributen(I)
                         If attribuut.TagString = "RINGLEIDING" Then RML = attribuut.textstring
                         If attribuut.TagString = "LEIDINGSOORT" Then WSL = attribuut.textstring
                         If attribuut.TagString = "UNITNUMMER" Then UNITSjek = attribuut.textstring
                    Next I
                End If
            End If
        End If
    Next element
    
    
If RML = "RM" Then frmGroeptekst.OptionButton7 = True
If RML = "RZ" Then frmGroeptekst.OptionButton8 = True
If WSL = "WTH-ZD" Then frmGroeptekst.OptionButton1 = True
If WSL = "PE-RT" Then frmGroeptekst.OptionButton2 = True
If WSL = "ALUFLEX" Then frmGroeptekst.OptionButton6 = True
UNITSJEK1 = Len(UNITSjek)
If UNITSJEK1 = 1 Then frmGroeptekst.CheckBox3.Value = True
End Sub
Sub wklengte()
Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
            If attribuut.TagString = "WERKLENGTE" And attribuut.textstring = "ja" Then frmGroeptekst.CheckBox2.Value = True
            If attribuut.TagString = "FLEXFIX" And attribuut.textstring = "ja" Then frmGroeptekst.CheckBox6.Value = True
        Next I
End If
End If
End If
Next element
End Sub
Private Sub cmdAfsluiten_Click()
Unload Me
End Sub
Private Sub CmdErase_Click()
On Error Resume Next
TextBox23 = 1
ThisDrawing.SetVariable "lispinit", 0
ThisDrawing.SetVariable "userr4", 1



If TextBox12 = "" Then
  MsgBox "Eerst unitnummer hieronder invullen.!!!!", vbExclamation, "Let op"
  TextBox12.SetFocus
End If


unittel = frmGroeptekst.TextBox12
If unittel > 0 And unittel < 10 Then
   frmGroeptekst.TextBox12 = "0" & frmGroeptekst.TextBox12
End If



Call delwand

Dim element As Object
Dim layerObj As AcadLayer
'groeptekstblok verwijderen
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
               'groepen tijdelijk hernummeren
               If attribuut.TagString = "UNITNUMMER" Then unitnum = attribuut.textstring
               If attribuut.TagString = "GROEPTEKST" Then grpt = attribuut.textstring
                                           
               If unitnum = TextBox12 Or unitnum = unittel Then
               hernummer1 = grpt
               hernummer2 = hernummer1 & "h"
                 For Each layerObj In ThisDrawing.Layers
                    If layerObj.Name = hernummer1 Then layerObj.Name = hernummer2
                 Next 'layerobj
               End If
               
               If CheckBox6.Value = True Then  'flexfix
                If unitnum = TextBox12 Or unitnum = unittel Then
                
                hernummer3 = grpt & "_Flexfix_aanvoer"
                hernummer4 = hernummer3 & "h"
                hernummer5 = grpt & "_Flexfix_retour"
                hernummer6 = hernummer5 & "h"
                hernummer7 = grpt & "_Flexfix"
                hernummer8 = hernummer7 & "h"
                
                 For Each layerObj In ThisDrawing.Layers
                    If layerObj.Name = hernummer3 Then layerObj.Name = hernummer4
                    If layerObj.Name = hernummer5 Then layerObj.Name = hernummer6
                    If layerObj.Name = hernummer7 Then layerObj.Name = hernummer8
                 Next 'layerobj
                End If
               End If
          If unitnum = TextBox12 Or unitnum = unittel Then element.Erase
       Next I
       
        End If
      End If
      End If
       unitnum = 0
       grpt = 0

  Next element
  


Call layzondertekst 'uitgezet op 29-11  niet meer nodig

  
  'Mat_spe_zd en Mat_spe_pe en Mat_spe_ALU verwijderen
  For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_ZD" Or element.Name = "Mat_spe_PE" Or element.Name = "Mat_spe_PE800" _
      Or element.Name = "Mat_spe_ALU" Or element.Name = "Mat_spe_ZDringleiding" Or element.Name = "Mat_spe_PEringleiding" Or _
      element.Name = "Mat_spe_ALUringleiding" Or element.Name = "Mat_spe_FLEX" Or element.Name = "Mat_spe_ZD_1627" Or _
      element.Name = "Mat_spe_FLEX_Aankoppel" Or element.Name = "Mat_spe_ZD_1627500" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "RNU" Or attribuut.TagString = "rnu" Then
           If attribuut.textstring = TextBox12 Or attribuut.textstring = unittel Then
               element.Erase
            End If 'TEXTBOX 12
        End If 'if unitnummer
       Next I
       
        End If
      End If
      End If
  Next element
  Update
  
  'REGELUNITBLOKJE VERWIJDEREN
   For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
If element.Name = "HERZ" Or element.Name = "RUH-R" Or element.Name = "RUH-RT" _
Or element.Name = "RUB-R" Or element.Name = "RUB-RT" Or element.Name = "RUBK-R" Or element.Name = "LT-VK" _
Or element.Name = "RUBK-RT" Or element.Name = "LT" Or element.Name = "LTS" Or element.Name = "LT-N" Or element.Name = "LTS-N" _
Or element.Name = "RUW" Or element.Name = "RUV" Or element.Name = "RUH-S" Or element.Name = "VSKO-B" _
Or element.Name = "RUB-S" Or element.Name = "KMV" Or element.Name = "RUH-N" Or element.Name = "RU-WW" Or element.Name = "RINGLEIDING" _
Or element.Name = "RU-WWN" Or element.Name = "RU-WWS" Or element.Name = "RU-WKN" Or element.Name = "RU-WKS" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "UNITNUMMER" Then
           If attribuut.textstring = TextBox12 Or attribuut.textstring = unittel Then
               element.Erase
            End If 'TEXTBOX 12
        End If 'if unitnummer
       Next I
      End If
     End If
     End If
   Next element
 Update
'EINDE REGELUNITBLOKJE VERWIJDEREN
  TextBox12 = Clear

End Sub
Private Sub CommandButton5_Click()
TextBox23 = 1
ThisDrawing.SetVariable "lispinit", 0
ThisDrawing.SetVariable "userr4", 1



If frmGroeptekst.CheckBox3.Value = False Then
    If TextBox18 > 0 And TextBox18 < 10 Then
       frmGroeptekst.TextBox18 = "0" & frmGroeptekst.TextBox18
    End If
End If

TextBox21 = TextBox19 - 1
TextBox22 = TextBox20

If TextBox19 > 0 And TextBox19 < 10 Then
   frmGroeptekst.TextBox19 = "0" & frmGroeptekst.TextBox19
End If


SAMENV = "groep " & TextBox18 & "." & TextBox19
ListBox2.AddItem (SAMENV)

Do Until TextBox21 = TextBox22
 TextBox21 = TextBox21 + 1
  If TextBox21 > 0 And TextBox21 < 10 Then samenv2 = "groep " & TextBox18 & "." & "0" & TextBox21
  'MsgBox SAMENV
  If TextBox21 > 9 Then samenv2 = "groep " & TextBox18 & "." & TextBox21
  ListBox2.AddItem (samenv2)
Loop

Dim element As Object
Dim layerObj As AcadLayer
'groeptekstblok verwijderen
  For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
      Set symbool = element
        If symbool.HasAttributes Then
         attributen = symbool.GetAttributes
        
tellerj = ListBox2.ListCount
For j = 0 To tellerj - 1
  uitlees = ListBox2.List(j)
        
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
               'groepen tijdelijk hernummeren
                If attribuut.TagString = "UNITNUMMER" Then unitnum = attribuut.textstring
                If attribuut.TagString = "GROEPTEKST" Then grpt = attribuut.textstring
            
                If unitnum = TextBox18 And uitlees = grpt Then element.Erase
                'MsgBox uitlees & " - " & grpt
                uitlees = 0
        Next I
Next j
       
         End If
      End If
      End If
       unitnum = 0
       grpt = 0
  Next element

 Update


tellerz = ListBox2.ListCount
For z = 0 To tellerz - 1
   'Define the text object
     If CheckBox6.Value = False Then
     hernaam = ListBox2.List(z)
     hernaam2 = hernaam & "h"
        For Each layerObj In ThisDrawing.Layers
         If layerObj.Name = hernaam Then layerObj.Name = hernaam2
        Next 'layerobj
     End If
    
    If CheckBox6.Value = True Then
        hernaam3 = ListBox2.List(z)
        hernaam4 = hernaam3 & "_Flexfix"
        hernaam5 = hernaam3 & "_Flexfix_aanvoer"
        hernaam6 = hernaam3 & "_Flexfix_retour"
     For Each layerObj In ThisDrawing.Layers
        If layerObj.Name = hernaam Then layerObj.Name = hernaam2
        If layerObj.Name = hernaam4 Then layerObj.Name = hernaam4 & "h"
        If layerObj.Name = hernaam5 Then layerObj.Name = hernaam5 & "h"
        If layerObj.Name = hernaam6 Then layerObj.Name = hernaam6 & "h"
     Next 'layerobj
    End If
Next z

TextBox18 = Clear: TextBox19 = Clear: TextBox20 = Clear: TextBox21 = Clear: TextBox22 = Clear
ListBox2.Clear
End Sub
Sub layzondertekst()
For Each layerObj In ThisDrawing.Layers
     mystr = Left(layerObj.Name, 5)
     MYSTR2 = Right(layerObj.Name, 1)
     
     If mystr = "groep" And MYSTR2 <> "h" Then
     splijt = Split(layerObj.Name, " ")
     splijt2 = Split(splijt(1), ".")
     If splijt2(0) = TextBox12 And splijt(0) = "groep" Then layerObj.Name = layerObj.Name & "h"
     End If 'mystr
Next 'layerobj
End Sub
Sub delwand()
ListBox1.Clear

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "HOHAFSTAND" Then
           If attribuut.textstring = "Wandverwarming" Then
           For j = LBound(attributen) To UBound(attributen)
           Set attribuut = attributen(j)
            If attribuut.TagString = "GROEPTEKST" Then ListBox1.AddItem (attribuut.textstring)
           Next j
        End If 'TEXTBOX 9
      End If 'if unitnummer
    Next I
End If
End If
End If
Next element

teller = ListBox1.ListCount
For I = 0 To teller - 1
   'Define the text object
    tstring = ListBox1.List(I)
    tstring2 = Split(tstring, (" "))
    nostring = Split(tstring2(1), ("."))
     
    If nostring(0) = TextBox12 Then
        tstring3 = "wand " & tstring2(1) & "h"
        For Each layerObj In ThisDrawing.Layers
             If layerObj.Name = tstring Then layerObj.Name = tstring3
        Next 'layerobj
    End If
  Next I

End Sub
Private Sub CmdReset_Click()
Call RESET 'alle waardes op nul zetten
End Sub
Sub RESET()
cmdTelrollen.Enabled = False
ComboBox1.Clear: ComboBox2.Clear: ComboBox4.Clear
CheckBox1.Value = False: CheckBox2.Value = False
CheckBox6.Value = False: CheckBox7.Value = False
TextBox1 = Clear: TextBox2 = Clear: TextBox3 = Clear: TextBox4 = Clear
TextBox5 = Clear: TextBox6 = Clear: TextBox7 = Clear: TextBox15 = Clear: TextBox16 = Clear
TextBox14 = Clear: TextBox25 = Clear: TextBox26 = Clear
TextBox1 = "0": TextBox2 = "0": TextBox3 = "0": TextBox4 = "0"
TextBox5 = "0": TextBox6 = "0": TextBox7 = "0": TextBox15 = "0": TextBox16 = "0": TextBox11 = "0"
TextBox25 = "0": TextBox26 = "0": TextBox28 = "0"
TextBox15 = "0": TextBox16 = "0": TextBox17 = "0": TextBox9 = Clear
TextBox10 = Clear: TextBox18 = Clear: TextBox19 = Clear: TextBox20 = Clear: TextBox21 = Clear: TextBox13.Visible = False
TextBox9.Locked = False: OptionButton1.Value = True: TextBox22 = Clear
TextBox9.SetFocus
OptionButton5.Value = False
ToggleButton1.Value = False
TextBox1.Visible = True
TextBox2.Visible = True: TextBox3.Visible = True: TextBox4.Visible = True
TextBox5.Visible = True: TextBox6.Visible = True: TextBox7.Visible = True
TextBox15.Visible = True: TextBox16.Visible = True
Label1.Visible = True: Label2.Visible = True: Label3.Visible = True: Label4.Visible = True
Label5.Visible = True: Label6.Visible = True: Label7.Visible = True
Label27.Visible = True: Label28.Visible = True
Label1.Caption = "250 meter"
CmdBloklogo.Enabled = False 'knop plaatsen bloklogo
ComboBox4.Visible = False: Label26.Visible = False
ListBox1.Clear: ListBox2.Clear
Label30.Caption = "0"
TextBox17.Enabled = False
frmGroeptekst.OptionButton7.Value = False
frmGroeptekst.OptionButton8.Value = False
'CheckBox3.Value = False
Call Combolijst2
Call combolijst1
End Sub
Private Sub OptionButton1_Click()
'wth-zd
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True:
TextBox4.Visible = True: TextBox5.Visible = True: TextBox6.Visible = True:
TextBox7.Visible = True: TextBox15.Visible = True: TextBox16.Visible = True:
Label1.Caption = "250 meter": Label2.Caption = "165 meter":
Label3.Caption = "125 meter": Label4.Caption = "105 meter": Label5.Caption = "90 meter"
Label6.Caption = "75 meter": Label7.Caption = "63 meter": Label27.Caption = "40 meter": Label27.Visible = True
Label28.Caption = "50 meter"
ComboBox4.Clear
ComboBox4.Visible = False: Label26.Visible = False
CmdBloklogo.top = 72
TextBox11.top = 198
Label22.top = 198
CheckBox8.Visible = False
Frame3.Height = 103
End Sub
Private Sub OptionButton9_Click()
'wth16*2,7
'TextBox1.Visible = False: TextBox2.Visible = False: TextBox3.Visible = False:
'Label1.Caption = "105 meter": Label2.Caption = "90 meter"
'Label3.Caption = "75 meter": Label4.Caption = "63 meter"
'Label5.Caption = Clear: Label27.Caption = Clear
'Label6.Caption = Clear: Label7.Caption = Clear: Label28.Caption = Clear
'TextBox4.Visible = True
'TextBox5.Visible = False: TextBox6.Visible = False:
'TextBox7.Visible = False: TextBox15.Visible = False: TextBox16.Visible = False
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True:
TextBox4.Visible = True: TextBox5.Visible = True: TextBox6.Visible = True:
TextBox7.Visible = True: TextBox15.Visible = True: TextBox16.Visible = True:
Label1.Caption = "250 meter": Label2.Caption = "165 meter":
Label3.Caption = "125 meter": Label4.Caption = "105 meter": Label5.Caption = "90 meter"
Label6.Caption = "75 meter": Label7.Caption = "63 meter": Label27.Caption = "40 meter": Label27.Visible = True
Label28.Caption = "50 meter"
ComboBox4.Clear
ComboBox4.Visible = False: Label26.Visible = False
CmdBloklogo.top = 72
TextBox11.top = 198
Label22.top = 198
CheckBox8.Visible = False
Frame3.Height = 103
End Sub
Private Sub OptionButton2_Click()
'wth-pe-rt
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False:
TextBox7.Visible = False: TextBox15.Visible = False: TextBox16.Visible = False:
Label4.Caption = Clear: Label5.Caption = Clear: Label27.Caption = Clear: Label27.Visible = False
Label6.Caption = Clear: Label7.Caption = Clear: Label28.Caption = Clear
Label1.Caption = "120 meter": Label2.Caption = "90 meter": Label3.Caption = "60 meter"
ComboBox4.Clear
ComboBox4.Visible = True: Label26.Visible = True
CmdBloklogo.top = 95
Frame3.Height = 127
TextBox11.top = 180
Label22.top = 180
CheckBox8.Visible = True
Call combolijst4
'Call leidingtype
End Sub
Private Sub OptionButton6_Click()
'wth-aluflex
If OptionButton6 = True Then
CheckBox2 = True
Else
CheckBox2 = False
End If
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False:
TextBox7.Visible = False: TextBox15.Visible = False: TextBox16.Visible = False:
Label4.Caption = Clear: Label5.Caption = Clear: Label27.Caption = Clear: Label28.Caption = Clear
Label6.Caption = Clear: Label7.Caption = Clear: Label1.Caption = "200 meter"
Label2.Caption = "100 meter": Label3.Caption = "50 meter"
ComboBox4.Clear
ComboBox4.Visible = True: Label26.Visible = True
CmdBloklogo.top = 95
Frame3.Height = 127
TextBox11.top = 198
Label22.top = 198
CheckBox8.Visible = False
Call combolijst4
'Call leidingtype
End Sub
Private Sub TextBox9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox9 = ""
End Sub
Private Sub TextBox10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox10 = ""
End Sub
Private Sub TextBox1_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox1 = ""
End Sub
Private Sub TextBox2_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox2 = ""
End Sub
Private Sub TextBox3_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox3 = ""
End Sub
Private Sub TextBox4_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox4 = ""
End Sub
Private Sub TextBox5_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox5 = ""
End Sub
Private Sub TextBox6_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox6 = ""
End Sub
Private Sub TextBox7_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7))
End Sub
Private Sub TextBox7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox7 = ""
End Sub
Private Sub TextBox15_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox15 = ""
End Sub
Private Sub TextBox16_Change()
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16))
End Sub
Private Sub TextBox16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox16 = ""
End Sub
Private Sub TextBox17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox17 = ""
End Sub
Private Sub TextBox17_Change()
On Error Resume Next
Dim b As Double
b = TextBox17.Text
  If Err Then
   TextBox17 = Clear
   TextBox17 = 0
   Exit Sub
  End If



If Label30.Caption = "0" Then Label30 = TextBox11
If TextBox17 = "0" Or TextBox17 = "" Then TextBox11 = Label30.Caption
TextBox11 = (Val(TextBox11)) + (Val(TextBox17))
End Sub
Private Sub TextBox9_Change()
Dim b As Double
On Error Resume Next
TextBox9.SetFocus
b = TextBox9.Text
If Err Then
   TextBox9 = Clear
   Cmdmeten.Enabled = False
   cmdTelrollen.Enabled = False
   Exit Sub
  End If
If TextBox9.Text <> "" And TextBox10.Text <> "" Then Cmdmeten.Enabled = True
If TextBox9.Text <> "" Then cmdTelrollen.Enabled = True
Label33 = 0

TextBox1.Text = "0"
TextBox2.Text = "0"
TextBox3.Text = "0"
TextBox4.Text = "0"
TextBox5.Text = "0"
TextBox6.Text = "0"
TextBox7.Text = "0"
TextBox15.Text = "0"
TextBox16.Text = "0"
TextBox11.Text = "0"
TextBox27 = Clear: TextBox27.Visible = False

End Sub
Private Sub TextBox10_Change()
Dim C As Double
On Error Resume Next
TextBox10.SetFocus
C = TextBox10.Text
If Err Then
   TextBox10 = Clear
   Cmdmeten.Enabled = False
   Exit Sub
  End If
If TextBox9.Text <> "" And TextBox10.Text <> "" Then Cmdmeten.Enabled = True
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

If TA > 20 And (frmGroeptekst.OptionButton7 = False And frmGroeptekst.OptionButton8 = False) Then
   'TextBox11.BackColor = &HFFFF&
   MsgBox "PAS OP!!, MEER DAN 20 groepen....!!!!", vbCritical
   TextBox1 = "0": TextBox2 = "0": TextBox3 = "0": TextBox4 = "0"
   TextBox5 = "0": TextBox6 = "0": TextBox7 = "0": TextBox15 = "0": TextBox16 = "0": TextBox11 = "0"
   Call RESET
   'TextBox11.BackColor = &HFFFFFF
   Exit Sub
End If

If Label33 = 0 Then
  If TextBox11 > 15 And ComboBox2 = "KMV" And (frmGroeptekst.OptionButton7 = False And frmGroeptekst.OptionButton8 = False) Then
   MsgBox "PAS OP!!, KMV-verdeler gaat maar tot 15 groepen....!!!!", vbCritical
   Label33 = 1
  End If
End If

If Label33 = 0 Then
  If TextBox11 > 16 And ComboBox2 = "HERZ" And (frmGroeptekst.OptionButton7 = False And frmGroeptekst.OptionButton8 = False) Then
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
Private Sub TextBox18_Change()

Dim D As Double
On Error Resume Next

D = TextBox18.Text

If Err Then
   TextBox18 = Clear
    CommandButton5.Enabled = False
   Exit Sub
  End If
  If TextBox18.Text <> "" And TextBox19.Text <> "" And TextBox20.Text <> "" Then CommandButton5.Enabled = True
End Sub
Private Sub TextBox19_Change()
Dim e As Double
On Error Resume Next

e = TextBox19.Text

If Err Then
   TextBox19 = Clear
    CommandButton5.Enabled = False
   Exit Sub
  End If
  If TextBox18.Text <> "" And TextBox19.Text <> "" And TextBox20.Text <> "" Then CommandButton5.Enabled = True
End Sub
Private Sub TextBox20_Change()
Dim f As Double
On Error Resume Next

f = TextBox20.Text

If Err Then
   TextBox20 = Clear
   CommandButton5.Enabled = False
   Exit Sub
  End If
If TextBox18.Text <> "" And TextBox19.Text <> "" And TextBox20.Text <> "" Then CommandButton5.Enabled = True
End Sub
Private Sub Cmdmeten_Click()
ThisDrawing.SendCommand "ucs" & vbCr & "world" & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "gt" & vbCr & vbCr
SendKeys "{capslock}"

On Error Resume Next
   Call Checktekst.Checktekst(a) 'kijken of de groeptekst al in de tekening staat
       If a = 1 Then
    Exit Sub
    End If
   
If TextBox23 = 1 Then
   Call Checklayer4.Checklayer4
   ThisDrawing.SetVariable "userr4", 0
End If

   'If a = 1 Then
   '  Call CmdErase_Click
      'Exit Sub
  ' End If
    
   'Call Checklayer1.Checklayer1  'module kijkt of er lege layers zijn
    
    Dim RetObj As AcadObject
    Dim Pbase As Variant
    Dim mystr As String
    Dim mystr1 As String
    Dim zz As String
    On Error Resume Next
    eg = 0 'eerste groep
    cmdTelrollen.Enabled = True
        
    frmGroeptekst.Hide
    ThisDrawing.SendCommand "undo" & vbCr & "BE" & vbCr 'markeringspunt tussen elke unit
    ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "*" & vbCr & vbCr
  
Opnieuw1:
    Lengte = 0  'lengte groep op nul zetten
    zlengte = 0
    Err.Clear
    

        
    ThisDrawing.Utility.GetEntity RetObj, Pbase, "Selecteer een lijn."
    
    
    If Err = 0 Then
    a = RetObj.startPoint
    b = a(2)
       If b <> 0 Then
            MsgBox "De Z-waarde's staan niet op nul.!!!!" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) & _
             "Ik zal nu proberen om de" & (Chr(13) & Chr(10)) & _
            "Z-waarde's voor je op nul te zetten.", vbExclamation
             ThisDrawing.SendCommand "Flatten" & vbCr
       End If
    End If
     
    
    If Err <> 0 Then
        Err.Clear
        TextBox10 = (Val(TextBox10)) + 1
        frmGroeptekst.show
        Exit Sub
        'End
    End If
    
    
       
     
    mystr = Left(RetObj.Layer, 4)
    mystr1 = LCase(mystr)
    If mystr1 <> "groe" And mystr <> "wand" Then
        MsgBox " Let op, dit is geen vloer of wandgroep!!!!!", vbExclamation, "Let op"
        GoTo Opnieuw1
    End If
   
   If CheckBox2 = False And CheckBox6 = False Then Call M_Meten.Meten(RetObj, Pbase, eg)
   If CheckBox2 = True Then Call M_Meten_Werkelijk.Meten_Werkelijk(RetObj, Pbase, eg)
   If CheckBox6 = True Then Call M_Meten_Flexfix.Meten_Flexfix(RetObj, Pbase, eg)
   GoTo Opnieuw1
End Sub
Private Sub cmdTelrollen_Click()
check1 = 0
Call zoek_blok.zoek_blok
Call CHECKWERKLENG(check1)
'MsgBox check1

If frmGroeptekst.CheckBox2.Value = False Then Call nmeten 'standaard
If (frmGroeptekst.CheckBox2.Value = True And frmGroeptekst.CheckBox3.Value = False) And check1 = 0 Then Call nmetenw 'nieuwe tekeningen
If (frmGroeptekst.CheckBox2.Value = True And frmGroeptekst.CheckBox3.Value = True) And check1 = 0 Then Call nmetenwoud 'oude tekeningen
check1 = 0
If frmGroeptekst.TextBox2 <> "0" And OptionButton1.Value = True Then Call waarschuwing.waarschuwing
End Sub
Sub CHECKWERKLENG(check1)
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
'MsgBox unitonder10

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "UNITNUMMER" Or attribuut.TagString = "unitnummer" Then
           If attribuut.textstring = unitonder10 Then
           
           For j = LBound(attributen) To UBound(attributen)
           Set attribuut = attributen(j)
            If attribuut.TagString = "WERKLENGTE" Then wklen = attribuut.textstring
            If wklen = "ja" Then
            check1 = 0
            Else
            check1 = 1
            End If
           Next j
        End If 'TEXTBOX 9
      End If 'if unitnummer
    Next I
End If
End If
End If
Next element
'MsgBox wklen

End Sub
Sub nmeten()

unittel = frmGroeptekst.TextBox9
'MsgBox "nmeten"
If frmGroeptekst.CheckBox3.Value = False Then
   If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
   If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If

If CheckBox6.Value = True Then
    Dim element As Object
    For Each element In ThisDrawing.ModelSpace
          If element.ObjectName = "AcDbBlockReference" Then
          If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
          Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
             If attribuut.TagString = "UNITNUMMER" Or attribuut.TagString = "unitnummer" Then
               If attribuut.textstring = unitonder10 Then
               For j = LBound(attributen) To UBound(attributen)
               Set attribuut = attributen(j)
                 If attribuut.TagString = "FLEXFIX" Then
                    ffch = attribuut.textstring
                    
                    For z = LBound(attributen) To UBound(attributen)
                    Set attribuut = attributen(z)
                       If attribuut.TagString = "ROLLENGTE" And ffch = "ja" Then
                       rolleesFF = attribuut.textstring
                         If rolleesFF <> " " Then
                          mystr = Split(rolleesFF, " ")
                          TextBox25 = (Val(TextBox25)) + mystr(0)
                          TextBox26 = (Val(TextBox26)) + 1
                          End If
                       'mystr = Left(ROLLEESFF, 3)
                       'MsgBox mystr(0)
                        End If 'rollengte
                        Next z
                 End If 'if not flexfix
    
              Next j
            End If 'TEXTBOX 9
          End If 'if unitnummer
    
        Next I
    End If
    End If
    End If
    Next element
End If 'checkbox6
            
Dim element200 As Object
    For Each element200 In ThisDrawing.ModelSpace
          If element200.ObjectName = "AcDbBlockReference" Then
          If element200.Name = "GROEPTEKSTBLOK" Or element200.Name = "groeptekstblok" Then
          Set symbool = element200
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
             If attribuut.TagString = "UNITNUMMER" Or attribuut.TagString = "unitnummer" Then
               If attribuut.textstring = unitonder10 Then
               For j = LBound(attributen) To UBound(attributen)
               Set attribuut = attributen(j)
                 If attribuut.TagString = "FLEXFIX" Then
                    ffch = attribuut.textstring
                    
                    For z = LBound(attributen) To UBound(attributen)
                    Set attribuut = attributen(z)
                       If (attribuut.TagString = "ROLLENGTE" And attribuut.textstring <> "") And ffch <> "ja" Then
                       ROLLEES = attribuut.textstring
                       mystr = Left(ROLLEES, 3)
                            If frmGroeptekst.CheckBox1 = True Then
                              If mystr = TextBox14.Value Then TextBox1 = (Val(TextBox1)) + 1
                            End If
                            If mystr = 250 And OptionButton1.Value = True Then TextBox1 = (Val(TextBox1)) + 1
                            If mystr = 165 And OptionButton1.Value = True Then TextBox2 = (Val(TextBox2)) + 1
                            If mystr = 125 And OptionButton1.Value = True Then TextBox3 = (Val(TextBox3)) + 1
                            If mystr = 105 And OptionButton1.Value = True Then TextBox4 = (Val(TextBox4)) + 1
                            If mystr = 90 And OptionButton1.Value = True Then TextBox5 = (Val(TextBox5)) + 1
                            If mystr = 75 And OptionButton1.Value = True Then TextBox6 = (Val(TextBox6)) + 1
                            If mystr = 63 And OptionButton1.Value = True Then TextBox7 = (Val(TextBox7)) + 1
                            If mystr = 50 And OptionButton1.Value = True Then TextBox15 = (Val(TextBox15)) + 1
                            If mystr = 40 And OptionButton1.Value = True Then TextBox16 = (Val(TextBox16)) + 1
                            'wthzd 16*2,7
                            If mystr = 105 And OptionButton9.Value = True Then TextBox4 = (Val(TextBox4)) + 1
                            If mystr = 90 And OptionButton9.Value = True Then TextBox5 = (Val(TextBox5)) + 1
                            If mystr = 75 And OptionButton9.Value = True Then TextBox6 = (Val(TextBox6)) + 1
                            If mystr = 63 And OptionButton9.Value = True Then TextBox7 = (Val(TextBox7)) + 1
                            If OptionButton9.Value = True Then
                            TextBox1.Visible = False: TextBox2.Visible = False: TextBox3.Visible = False: TextBox15.Visible = False: TextBox16.Visible = False
                            Label1.Visible = False: Label2.Visible = False: Label3.Visible = False: Label27.Visible = False: Label28.Visible = False
                            End If
                            'PE-RT
                            If mystr = 120 And OptionButton2.Value = True Then TextBox1 = (Val(TextBox1)) + 1
                            If mystr = 90 And OptionButton2.Value = True Then TextBox2 = (Val(TextBox2)) + 1
                            If mystr = 60 And OptionButton2.Value = True Then TextBox3 = (Val(TextBox3)) + 1
                            'aluflex
                            If mystr = 200 And OptionButton6.Value = True Then TextBox1 = (Val(TextBox1)) + 1
                            If mystr = 100 And OptionButton6.Value = True Then TextBox2 = (Val(TextBox2)) + 1
                            If mystr = 50 And OptionButton6.Value = True Then TextBox3 = (Val(TextBox3)) + 1
                        End If 'rollengte
                        Next z
                 End If 'if not flexfix
    
              Next j
            End If 'TEXTBOX 9
          End If 'if unitnummer
    
        Next I
    End If
    End If
    End If
Next element200

 
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + (Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16)) + (Val(TextBox17))
CmdBloklogo.Enabled = True
cmdTelrollen.Enabled = False
TextBox1.Locked = False: TextBox2.Locked = False: TextBox3.Locked = False
TextBox4.Locked = False: TextBox5.Locked = False: TextBox6.Locked = False
TextBox7.Locked = False: TextBox15.Locked = False: TextBox16.Locked = False
TextBox9.Locked = False
TextBox9.SetFocus
Update
End Sub
Sub nmetenw()
'MsgBox "nmetenw"
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

  Dim element As Object
  
    For Each element In ThisDrawing.ModelSpace
          If element.ObjectName = "AcDbBlockReference" Then
          If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
          Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
             If attribuut.TagString = "UNITNUMMER" Or attribuut.TagString = "unitnummer" Then
               If attribuut.textstring = unitonder10 Then
               For j = LBound(attributen) To UBound(attributen)
               Set attribuut = attributen(j)
                 If attribuut.TagString = "FLEXFIX" Then
                    ffch = attribuut.textstring
                    
                    For z = LBound(attributen) To UBound(attributen)
                    Set attribuut = attributen(z)
                    
                       If (attribuut.TagString = "ROLLENGTE" And attribuut.textstring <> "") And ffch = "ja" Then
                       'If attribuut.TagString = "ROLLENGTE" And ffch = "ja" Then
                       rolleesFF = attribuut.textstring
                         If rolleesFF <> " " Then
                          mystr = Split(rolleesFF, " ")
                          TextBox25 = (Val(TextBox25)) + mystr(0)
                          TextBox26 = (Val(TextBox26)) + 1
                          End If
                       'mystr = Left(ROLLEESFF, 3)
                       'MsgBox mystr(0)
                        End If 'rollengte
                        Next z
                 End If 'if not flexfix
    
              Next j
            End If 'TEXTBOX 9
          End If 'if unitnummer
    
        Next I
    End If
    End If
    End If
    Next element

            
Dim element200 As Object

    For Each element200 In ThisDrawing.ModelSpace
          If element200.ObjectName = "AcDbBlockReference" Then
          If element200.Name = "GROEPTEKSTBLOK" Or element200.Name = "groeptekstblok" Then
          Set symbool = element200
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
             If attribuut.TagString = "UNITNUMMER" Or attribuut.TagString = "unitnummer" Then
               If attribuut.textstring = unitonder10 Then
               For j = LBound(attributen) To UBound(attributen)
               Set attribuut = attributen(j)
                 If attribuut.TagString = "FLEXFIX" Then
                    ffch = attribuut.textstring
                    
                    For z = LBound(attributen) To UBound(attributen)
                    Set attribuut = attributen(z)
                       If (attribuut.TagString = "ROLLENGTE" And attribuut.textstring <> "") And ffch <> "ja" Then
                       'MsgBox "ELEMENT200"
                      ROLLEES2 = Split(attribuut.textstring)
                     ' MsgBox attribuut.TextString
                 If ROLLEES2(0) <> "" Then
                     MYSTR2 = ROLLEES2(0)
                     Else
                     MYSTR2 = 0
                 End If
            'mystr2 = Left(ROLLEES2, 4)
            If OptionButton1.Value = True Then
            If frmGroeptekst.CheckBox1 = True Then
               If MYSTR2 = TextBox14.Value Then TextBox1 = (Val(TextBox1)) + 1
            End If
            If (MYSTR2 >= 162.5 And MYSTR2 < 247.5) And OptionButton1.Value = True Then TextBox1 = (Val(TextBox1)) + 1
            If (MYSTR2 >= 122.5 And MYSTR2 < 162.5) And OptionButton1.Value = True Then TextBox2 = (Val(TextBox2)) + 1
            If (MYSTR2 >= 102.5 And MYSTR2 < 122.5) And OptionButton1.Value = True Then TextBox3 = (Val(TextBox3)) + 1
            If (MYSTR2 >= 87.5 And MYSTR2 < 102.5) And OptionButton1.Value = True Then TextBox4 = (Val(TextBox4)) + 1
            If (MYSTR2 >= 72.5 And MYSTR2 < 87.5) And OptionButton1.Value = True Then TextBox5 = (Val(TextBox5)) + 1
            If (MYSTR2 >= 60.5 And MYSTR2 < 72.5) And OptionButton1.Value = True Then TextBox6 = (Val(TextBox6)) + 1
            If (MYSTR2 >= 47.5 And MYSTR2 < 60.5) And OptionButton1.Value = True Then TextBox7 = (Val(TextBox7)) + 1
            If (MYSTR2 >= 37.5 And MYSTR2 < 47.5) And OptionButton1.Value = True Then TextBox15 = (Val(TextBox15)) + 1
            If (MYSTR2 >= 10 And MYSTR2 < 37.5) And OptionButton1.Value = True Then TextBox16 = (Val(TextBox16)) + 1
            frmGroeptekst.TextBox27.Visible = True
            totalrol = totalrol + Round(Val(MYSTR2), 2)
            'frmGroeptekst.TextBox27 = Val(frmGroeptekst.TextBox27) + mystr2
            End If
            'PE-RT
            If OptionButton2.Value = True Then
            If (MYSTR2 >= 87.5 And MYSTR2 < 117.5) And OptionButton2.Value = True Then TextBox1 = (Val(TextBox1)) + 1
            If (MYSTR2 >= 57.5 And MYSTR2 < 87.5) And OptionButton2.Value = True Then TextBox2 = (Val(TextBox2)) + 1
            If (MYSTR2 >= 10 And MYSTR2 < 57.5) And OptionButton2.Value = True Then TextBox3 = (Val(TextBox3)) + 1
            If CheckBox8.Value = True And MYSTR2 >= 117.5 Then TextBox1 = (Val(TextBox1)) + 1
            frmGroeptekst.TextBox27.Visible = True
            'frmGroeptekst.TextBox27 = Val(frmGroeptekst.TextBox27) + mystr2
            totalrol = totalrol + Round(Val(MYSTR2), 2)
            End If
            'PE-RT 16 x 2,7
            If OptionButton9.Value = True And CheckBox2.Value = True Then
            If (MYSTR2 >= 87.5 And MYSTR2 < 125) Then TextBox4 = (Val(TextBox4)) + 1
            If (MYSTR2 >= 72.5 And MYSTR2 < 87.5) Then TextBox5 = (Val(TextBox5)) + 1
            If (MYSTR2 >= 60.5 And MYSTR2 < 72.5) Then TextBox6 = (Val(TextBox6)) + 1
            If (MYSTR2 >= 10 And MYSTR2 < 60.5) Then TextBox7 = (Val(TextBox7)) + 1
            TextBox1.Visible = False: TextBox2.Visible = False: TextBox3.Visible = False: TextBox15.Visible = False: TextBox16.Visible = False
            Label1.Visible = False: Label2.Visible = False: Label3.Visible = False: Label27.Visible = False: Label28.Visible = False
            frmGroeptekst.TextBox27.Visible = True
            'frmGroeptekst.TextBox27 = Val(frmGroeptekst.TextBox27) + mystr2
            totalrol = totalrol + Round(Val(MYSTR2), 2)
            End If
            'aluflex
            If OptionButton6.Value = True Then
            If (MYSTR2 >= 97.5 And MYSTR2 < 197.5) And OptionButton6.Value = True Then TextBox1 = (Val(TextBox1)) + 1
            If (MYSTR2 >= 47.5 And MYSTR2 < 97.5) And OptionButton6.Value = True Then TextBox2 = (Val(TextBox2)) + 1
            If (MYSTR2 >= 10 And MYSTR2 < 47.5) And OptionButton6.Value = True Then TextBox3 = (Val(TextBox3)) + 1
            frmGroeptekst.TextBox27.Visible = True
            'frmGroeptekst.TextBox27 = Val(frmGroeptekst.TextBox27) + mystr2
            totalrol = totalrol + Round(Val(MYSTR2), 2)
            End If
            End If 'ROLLENGTE
                        Next z
                 End If 'if not flexfix
    
              Next j
            End If 'TEXTBOX 9
          End If 'if unitnummer
    
        Next I
    End If
    End If
    End If
Next element200
            
            
               
'''               If Not attribuut.TagString = "FLEXFIX" Then
'''                 If attribuut.TagString = "ROLLENGTE" And attribuut.TextString <> " " Then
'''                 rollees2 = Split(attribuut.TextString)
'''                 mystr2 = rollees2(0)
'''
'''            'mystr2 = Left(ROLLEES2, 4)
'''            If OptionButton1.Value = True Then
'''            If frmGroeptekst.CheckBox1 = True Then
'''               If mystr2 = TextBox14.Value Then TextBox1 = (Val(TextBox1)) + 1
'''            End If
'''            If (mystr2 >= 162.5 And mystr2 < 247.5) And OptionButton1.Value = True Then TextBox1 = (Val(TextBox1)) + 1
'''            If (mystr2 >= 122.5 And mystr2 < 162.5) And OptionButton1.Value = True Then TextBox2 = (Val(TextBox2)) + 1
'''            If (mystr2 >= 102.5 And mystr2 < 122.5) And OptionButton1.Value = True Then TextBox3 = (Val(TextBox3)) + 1
'''            If (mystr2 >= 87.5 And mystr2 < 102.5) And OptionButton1.Value = True Then TextBox4 = (Val(TextBox4)) + 1
'''            If (mystr2 >= 72.5 And mystr2 < 87.5) And OptionButton1.Value = True Then TextBox5 = (Val(TextBox5)) + 1
'''            If (mystr2 >= 60.5 And mystr2 < 72.5) And OptionButton1.Value = True Then TextBox6 = (Val(TextBox6)) + 1
'''            If (mystr2 >= 47.5 And mystr2 < 60.5) And OptionButton1.Value = True Then TextBox7 = (Val(TextBox7)) + 1
'''            If (mystr2 >= 37.5 And mystr2 < 47.5) And OptionButton1.Value = True Then TextBox15 = (Val(TextBox15)) + 1
'''            If (mystr2 >= 10 And mystr2 < 37.5) And OptionButton1.Value = True Then TextBox16 = (Val(TextBox16)) + 1
'''            frmGroeptekst.TextBox27.Visible = True
'''            totalrol = totalrol + Round(Val(mystr2), 2)
'''            'frmGroeptekst.TextBox27 = Val(frmGroeptekst.TextBox27) + mystr2
'''            End If
'''            'PE-RT
'''            If OptionButton2.Value = True Then
'''            If (mystr2 >= 87.5 And mystr2 < 117.5) And OptionButton2.Value = True Then TextBox1 = (Val(TextBox1)) + 1
'''            If (mystr2 >= 57.5 And mystr2 < 87.5) And OptionButton2.Value = True Then TextBox2 = (Val(TextBox2)) + 1
'''            If (mystr2 >= 10 And mystr2 < 57.5) And OptionButton2.Value = True Then TextBox3 = (Val(TextBox3)) + 1
'''            If CheckBox8.Value = True And mystr2 >= 117.5 Then TextBox1 = (Val(TextBox1)) + 1
'''            frmGroeptekst.TextBox27.Visible = True
'''            'frmGroeptekst.TextBox27 = Val(frmGroeptekst.TextBox27) + mystr2
'''            totalrol = totalrol + Round(Val(mystr2), 2)
'''            End If
'''            'aluflex
'''            If OptionButton6.Value = True Then
'''            If (mystr2 >= 97.5 And mystr2 < 197.5) And OptionButton6.Value = True Then TextBox1 = (Val(TextBox1)) + 1
'''            If (mystr2 >= 47.5 And mystr2 < 97.5) And OptionButton6.Value = True Then TextBox2 = (Val(TextBox2)) + 1
'''            If (mystr2 >= 10 And mystr2 < 47.5) And OptionButton6.Value = True Then TextBox3 = (Val(TextBox3)) + 1
'''            frmGroeptekst.TextBox27.Visible = True
'''            'frmGroeptekst.TextBox27 = Val(frmGroeptekst.TextBox27) + mystr2
'''            totalrol = totalrol + Round(Val(mystr2), 2)
'''            End If
'''            End If 'ROLLENGTE
'''               End If 'if not flexfix
'''
'''          Next j
'''        End If 'TEXTBOX 9
'''
'''      End If 'if unitnummer
'''
'''    Next I
'''End If
'''End If
'''End If
'''Next element



frmGroeptekst.TextBox27 = totalrol
TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + (Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16)) + (Val(TextBox17))
CmdBloklogo.Enabled = True
cmdTelrollen.Enabled = False
TextBox1.Locked = False: TextBox2.Locked = False: TextBox3.Locked = False
TextBox4.Locked = False: TextBox5.Locked = False: TextBox6.Locked = False
TextBox7.Locked = False: TextBox15.Locked = False: TextBox16.Locked = False
TextBox9.Locked = False
TextBox9.SetFocus
Update
End Sub
Sub nmetenwoud()
unittel = frmGroeptekst.TextBox9
'MsgBox "nmetenoud"
If frmGroeptekst.CheckBox3.Value = False Then
   If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
   If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If

Dim element3 As Object
For Each element3 In ThisDrawing.ModelSpace
      If element3.ObjectName = "AcDbBlockReference" Then
      If element3.Name = "GROEPTEKSTBLOK" Or element3.Name = "groeptekstblok" Then
      Set symbool = element3
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For k = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(k)
        If attribuut.TagString = "UNITNUMMER" Or attribuut.TagString = "unitnummer" Then
           If attribuut.textstring = unitonder10 Then
           For l = LBound(attributen) To UBound(attributen)
           Set attribuut = attributen(l)
            If attribuut.TagString = "ROLLENGTE" And attribuut.textstring <> " " Then
            rollees3 = Split(attribuut.textstring)
            mystr3 = rollees3(0)
            'ListBox2.AddItem (mystr3)
            If frmGroeptekst.CheckBox1 = True Then
               If mystr3 = TextBox14.Value Then TextBox1 = (Val(TextBox1)) + 1
            End If
            If (mystr3 >= 163.6 And mystr3 < 250) And OptionButton1.Value = True Then TextBox1 = (Val(TextBox1)) + 1
            If (mystr3 >= 123.6 And mystr3 < 163.6) And OptionButton1.Value = True Then TextBox2 = (Val(TextBox2)) + 1
            If (mystr3 >= 103.6 And mystr3 < 123.6) And OptionButton1.Value = True Then TextBox3 = (Val(TextBox3)) + 1
            If (mystr3 >= 88.6 And mystr3 < 103.6) And OptionButton1.Value = True Then TextBox4 = (Val(TextBox4)) + 1
            If (mystr3 >= 73.6 And mystr3 < 88.6) And OptionButton1.Value = True Then TextBox5 = (Val(TextBox5)) + 1
            If (mystr3 >= 61.6 And mystr3 < 73.6) And OptionButton1.Value = True Then TextBox6 = (Val(TextBox6)) + 1
            If (mystr3 >= 48.6 And mystr3 < 61.6) And OptionButton1.Value = True Then TextBox7 = (Val(TextBox7)) + 1
            If (mystr3 >= 38.6 And mystr3 < 48.6) And OptionButton1.Value = True Then TextBox15 = (Val(TextBox15)) + 1
            If (mystr3 >= 10 And mystr3 < 38.6) And OptionButton1.Value = True Then TextBox16 = (Val(TextBox16)) + 1
            'PE-RT
            If (mystr3 >= 88.6 And mystr3 < 118.6) And OptionButton2.Value = True Then TextBox1 = (Val(TextBox1)) + 1
            If (mystr3 >= 58.6 And mystr3 < 88.6) And OptionButton2.Value = True Then TextBox2 = (Val(TextBox2)) + 1
            If (mystr3 >= 10 And mystr3 < 58.6) And OptionButton2.Value = True Then TextBox3 = (Val(TextBox3)) + 1
            'aluflex
            If (mystr3 >= 98.6 And mystr3 < 200) And OptionButton6.Value = True Then TextBox1 = (Val(TextBox1)) + 1
            If (mystr3 >= 48.6 And mystr3 < 98.6) And OptionButton6.Value = True Then TextBox2 = (Val(TextBox2)) + 1
            If (mystr3 >= 10 And mystr3 < 48.6) And OptionButton6.Value = True Then TextBox3 = (Val(TextBox3)) + 1
            End If 'ROLLENGTE
           Next l
        End If 'TEXTBOX 9
      End If 'if unitnummer
    Next k
End If
End If
End If
Next element3

TextBox11 = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + (Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16)) + (Val(TextBox17))
CmdBloklogo.Enabled = True
cmdTelrollen.Enabled = False
TextBox1.Locked = False: TextBox2.Locked = False: TextBox3.Locked = False
TextBox4.Locked = False: TextBox5.Locked = False: TextBox6.Locked = False
TextBox7.Locked = False: TextBox15.Locked = False: TextBox16.Locked = False
TextBox9.Locked = False
TextBox9.SetFocus
Update
End Sub
Sub combolijst1()
ComboBox1.AddItem "Vlechtdraad"
ComboBox1.AddItem "Witmarmerbeugels"
ComboBox1.AddItem "Beugels/Nagels"
ComboBox1.AddItem "Eigen middelen"
'ComboBox1.AddItem "IFD-Karton"
ComboBox1.AddItem "IFD-Polystyreen"
ComboBox1.AddItem "Isoclips"
ComboBox1.AddItem "Keg"
ComboBox1.AddItem "Montagestrip"
ComboBox1.AddItem "Noppenplaat"
ComboBox1.AddItem "Schietbeugels"
ComboBox1.AddItem "Ty-raps"
ComboBox1.AddItem "Varisoclips"
'ComboBox1.AddItem "Wedi Plaat"
'ComboBox1.listindex = 7
'ComboBox1.MatchEntry = fmMatchEntryFirstLetter
End Sub
Private Sub ComboBox2_Change()
'type unit
If ComboBox2.Value = "RUH-R" Or ComboBox2.Value = "RUH-RT" Then
  ComboBox3.Enabled = True
  Else
  ComboBox3.Enabled = False
 End If

If TextBox11 > 15 And ComboBox2 = "KMV" And (frmGroeptekst.OptionButton7.Value = False Or frmGroeptekst.OptionButton8.Value = False) Then
   MsgBox "PAS OP!!, KMV-verdeler gaat maar tot 15 groepen....!!!!", vbCritical
   End If
If TextBox11 > 12 And ComboBox2 = "VSKO" Then
   MsgBox "PAS OP!!, VSKO-verdeler gaat maar tot 12 groepen....!!!!", vbCritical
End If
If TextBox11 > 16 And ComboBox2 = "HERZ" And (frmGroeptekst.OptionButton7.Value = False Or frmGroeptekst.OptionButton8.Value = False) Then
   MsgBox "PAS OP!!, HERZ-verdeler gaat maar tot 16 groepen....!!!!", vbCritical
   End If
End Sub
Sub Combolijst2()
ComboBox2.AddItem "HERZ"
ComboBox2.AddItem "KMV"
ComboBox2.AddItem "LT"
ComboBox2.AddItem "LT-N"
ComboBox2.AddItem "LTS" 'aangepast
ComboBox2.AddItem "LTS-N" 'aangepast
ComboBox2.AddItem "LT-VK" 'aangepast
ComboBox2.AddItem "RINGLEIDING"
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
ComboBox2.AddItem "VSKO"
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
 If OptionButton2.Value = True Then
 'ComboBox4.AddItem " "
 ComboBox4.AddItem "PE-RT 14*2 mm"
 ComboBox4.AddItem "PE-RT 16*2 mm"
 ComboBox4.AddItem "ALUFLEX 14*2 mm"
 ComboBox4.AddItem "ALUFLEX 16*2 mm"
 ComboBox4.AddItem "ALUFLEX 18*2 mm"
 ComboBox4.AddItem "ALUFLEX 20*2 mm"
 ComboBox4.AddItem "PE-RT 10*1,25 mm"
 'ComboBox4.AddItem "PE-RT 16*2,7 mm"
  ComboBox4.ListIndex = 1
 'ComboBox4.Text = ComboBox4.List(1)
 End If
 
 If OptionButton6.Value = True Then
 ComboBox4.AddItem "ALUFLEX 14*2 mm"
 ComboBox4.AddItem "ALUFLEX 16*2 mm"
 ComboBox4.AddItem "ALUFLEX 18*2 mm"
 ComboBox4.AddItem "ALUFLEX 20*2 mm"
 ComboBox4.ListIndex = 1
 'ComboBox4.Text = ComboBox4.List(1)
 End If
 
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
If ComboBox4 = "" And OptionButton6 = True Then
    MsgBox "Je bent 'Type buis' vergeten in te vullen..!!"
    ComboBox4.SetFocus
    Exit Sub
    End If
If ComboBox4 = "" And OptionButton2 = True Then
    MsgBox "Je bent 'Type buis' vergeten in te vullen..!!"
    ComboBox4.SetFocus
    Exit Sub
    End If
frmGroeptekst.Hide
TextBox9.Locked = False
    Set newLayer = ThisDrawing.Layers.Add("BLOKLOGO")
    ThisDrawing.ActiveLayer = newLayer
    Update
    ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "*" & vbCr & vbCr
    Update
Call Schaal(scaal)
If (OptionButton1.Value = True Or OptionButton9.Value = True) And (TextBox11 = "0" And TextBox25 <> "0") Then
bestand = "c:\acad2002\dwg\Mat_spe_FLEX.dwg"
'MsgBox bestand
Else
 If OptionButton1.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_ZD.dwg"
 If OptionButton9.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_ZD_1627.dwg"
End If

If OptionButton2.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_PE.dwg"
If OptionButton6.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_ALU.dwg"
If OptionButton1.Value = True And (frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True) Then bestand = "c:\acad2002\dwg\Mat_spe_ZDringleiding.dwg"
If OptionButton9.Value = True And (frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True) Then bestand = "c:\acad2002\dwg\Mat_spe_ZDringleiding.dwg"
If OptionButton2.Value = True And (frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True) Then bestand = "c:\acad2002\dwg\Mat_spe_PEringleiding.dwg"
If OptionButton6.Value = True And (frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True) Then bestand = "c:\acad2002\dwg\Mat_spe_ALUringleiding.dwg"
If frmGroeptekst.CheckBox8.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_PE800.dwg"
If frmGroeptekst.CheckBox2.Value = True And OptionButton9.Value = True Then bestand = "c:\acad2002\dwg\Mat_spe_ZD_1627500.dwg"
bestand20 = "c:\acad2002\dwg\Mat_spe_FLEX_Aankoppel.dwg"

If frmGroeptekst.TextBox29 <> "" Then scaal = frmGroeptekst.TextBox29
Dim pb1 As Variant
pb1 = ThisDrawing.Utility.GetPoint(, "Plaats startpunt....")
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb1, bestand, scaal, scaal, 1, 0)
If TextBox11 <> "0" And TextBox25 <> "0" Then Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb1, bestand20, scaal, scaal, 1, 0)
Update
    If Err Then
    frmGroeptekst.show
    Exit Sub
    End If
    
Dim element2 As Object
If (OptionButton1.Value = True Or OptionButton9.Value = True) Then
aantal_groepen = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox4)) + (Val(TextBox5)) + _
(Val(TextBox6)) + (Val(TextBox7)) + (Val(TextBox15)) + (Val(TextBox16)) + (Val(TextBox17))
If (Val(TextBox1)) = 0 Then TextBox1 = "-"
If (Val(TextBox2)) = 0 Then TextBox2 = "-"
If (Val(TextBox3)) = 0 Then TextBox3 = "-"
If (Val(TextBox4)) = 0 Then TextBox4 = "-"
If (Val(TextBox5)) = 0 Then TextBox5 = "-"
If (Val(TextBox6)) = 0 Then TextBox6 = "-"
If (Val(TextBox7)) = 0 Then TextBox7 = "-"
If (Val(TextBox15)) = 0 Then TextBox15 = "-"
If (Val(TextBox16)) = 0 Then TextBox16 = "-"
aantal_groepen = aantal_groepen + (Val(TextBox26))

If ComboBox3.Enabled = False Then
REGELTYPE = ComboBox2 & " " & aantal_groepen 'ZONDER REGELING
Else
REGELTYPE = ComboBox2 & " " & aantal_groepen & "/" & ComboBox3 'met regeling
End If
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Then REGELTYPE = ComboBox2

If ComboBox2 = "RUW-Groot" Or ComboBox2 = "RUW-Klein" Then
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
      If UCase(element2.Name) = "MAT_SPE_ZD" Or element2.Name = "Mat_spe_ZDringleiding" _
      Or element2.Name = "Mat_spe_ZD_1627" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "RNU" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'REGELUNITNUMMER
        If attribuut.TagString = "WTHZD" And OptionButton1.Value = True And attribuut.textstring = "" Then attribuut.textstring = "WTH-ZD 20*3,4 mm" 'TYPE BUIS
        If attribuut.TagString = "WTHZD" And OptionButton9.Value = True And attribuut.textstring = "" Then attribuut.textstring = "WTH-ZD 16*2,7 mm" 'TYPE BUIS
        If attribuut.TagString = "WTH250" And attribuut.textstring = "" Then attribuut.textstring = TextBox1  '250 METER
        If attribuut.TagString = "WTH165" And attribuut.textstring = "" Then attribuut.textstring = TextBox2  '165 METER
        If attribuut.TagString = "WTH125" And attribuut.textstring = "" Then attribuut.textstring = TextBox3  '125 METER
        If attribuut.TagString = "WTH105" And attribuut.textstring = "" Then attribuut.textstring = TextBox4  '105 METER
        If attribuut.TagString = "WTH90" And attribuut.textstring = "" Then attribuut.textstring = TextBox5  '90 METER
        If attribuut.TagString = "WTH75" And attribuut.textstring = "" Then attribuut.textstring = TextBox6  '75 METER
        If attribuut.TagString = "WTH63" And attribuut.textstring = "" Then attribuut.textstring = TextBox7  '63 METER
        If attribuut.TagString = "WTH50" And attribuut.textstring = "" Then attribuut.textstring = TextBox15  '40 METER
        If attribuut.TagString = "WTH40" And attribuut.textstring = "" Then attribuut.textstring = TextBox16  '50 METER
        If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring = "" Then attribuut.textstring = REGELTYPE  'TYPE REGELUNIT
        If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox1  'BEVESTIGING
        If attribuut.TagString = "ROLGROTER250" And CheckBox1.Value = True Then attribuut.textstring = TextBox14 & " meter :" 'ROL GROTER DAN 250 METER
       Next I
       
      End If
      End If
      End If
  Next element2
 End If
 Update
 
 
 For Each element13 In ThisDrawing.ModelSpace
      If element13.ObjectName = "AcDbBlockReference" Then
      If UCase(element13.Name) = "MAT_SPE_ZD_1627500" Then
      Set symbool = element13
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
         For w = LBound(attributen) To UBound(attributen)
         Set attribuut = attributen(w)
         If attribuut.TagString = "RNU" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'REGELUNITNUMMER
         If attribuut.TagString = "PE" And attribuut.textstring = "" Then attribuut.textstring = "WTH-ZD 16*2,7 mm"
         If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring = "" Then attribuut.textstring = REGELTYPE  'TYPE REGELUNIT
         If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox1  'BEVESTIGING
         If attribuut.TagString = "LMETER" And attribuut.textstring = "" Then attribuut.textstring = Round(Val(TextBox27), 1)
         Next w
        
        End If
      End If
      End If
  Next element13
 
 
 For Each element6 In ThisDrawing.ModelSpace
      If element6.ObjectName = "AcDbBlockReference" Then
      If UCase(element6.Name) = "MAT_SPE_FLEX" Or element6.Name = "Mat_spe_FLEX_Aankoppel" Then
      Set symbool = element6
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "RNU" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'REGELUNITNUMMER
        If attribuut.TagString = "FLEX_BUIS" And attribuut.textstring = "" Then attribuut.textstring = "WTH-ZD 16*2,7 mm" 'REGELUNITNUMMER
        If attribuut.TagString = "FLEX_METERS" And attribuut.textstring = "" Then attribuut.textstring = TextBox25  'AANTAL METER
        If attribuut.TagString = "FLEX_MATTEN" And attribuut.textstring = "" Then attribuut.textstring = TextBox26  'AANTAL METER
        If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring = "" Then attribuut.textstring = REGELTYPE  'TYPE REGELUNIT
        If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox1  'BEVESTIGING
       Next I
       
      End If
      End If
      End If
  Next element6

 Update
 
 Dim element3 As Object
If OptionButton2.Value = True Then
aantal_groepen = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox17))
If (Val(TextBox1)) = 0 Then TextBox1 = "-"
If (Val(TextBox2)) = 0 Then TextBox2 = "-"
If (Val(TextBox3)) = 0 Then TextBox3 = "-"
If ComboBox3.Enabled = False Then
REGELTYPE = ComboBox2 & " " & aantal_groepen
Else
REGELTYPE = ComboBox2 & " " & aantal_groepen & "/" & ComboBox3
End If
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Then REGELTYPE = ComboBox2

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
        If attribuut.TagString = "RNU" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'REGELUNITNUMMER
        If attribuut.TagString = "PE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox4 'TYPE BUIS
        If CheckBox8.Value = False Then If attribuut.TagString = "PE120" And attribuut.textstring = "" Then attribuut.textstring = TextBox1  '120 METER
        If CheckBox8.Value = False Then If attribuut.TagString = "PE90" And attribuut.textstring = "" Then attribuut.textstring = TextBox2  '90 METER
        If CheckBox8.Value = False Then If attribuut.TagString = "PE60" And attribuut.textstring = "" Then attribuut.textstring = TextBox3  '60 METER
        If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring = "" Then attribuut.textstring = REGELTYPE  'TYPE REGELUNIT
        If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox1  'BEVESTIGING
        If attribuut.TagString = "LMETER" And attribuut.textstring = "" Then attribuut.textstring = Round(Val(TextBox27), 1)
    Next j
       
        End If
      End If
      End If
  Next element3
  
 End If
  Update
 

  
  

  
   Dim element4 As Object
If OptionButton6.Value Then
aantal_groepen = (Val(TextBox1)) + (Val(TextBox2)) + (Val(TextBox3)) + (Val(TextBox17))
If (Val(TextBox1)) = 0 Then TextBox1 = "-"
If (Val(TextBox2)) = 0 Then TextBox2 = "-"
If (Val(TextBox3)) = 0 Then TextBox3 = "-"
If ComboBox3.Enabled = False Then
REGELTYPE = ComboBox2 & " " & aantal_groepen
Else
REGELTYPE = ComboBox2 & " " & aantal_groepen & "/" & ComboBox3
End If
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Then REGELTYPE = ComboBox2

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

For Each element4 In ThisDrawing.ModelSpace
      If element4.ObjectName = "AcDbBlockReference" Then
      If UCase(element4.Name) = "MAT_SPE_ALU" Or element4.Name = "Mat_spe_ALUringleiding" Then
      Set symbool = element4
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "RNU" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'REGELUNITNUMMER
        If attribuut.TagString = "ALU" And attribuut.textstring = "" Then attribuut.textstring = ComboBox4 'TYPE BUIS
        If attribuut.TagString = "ALU200" And attribuut.textstring = "" Then attribuut.textstring = TextBox1  '120 METER
        If attribuut.TagString = "ALU100" And attribuut.textstring = "" Then attribuut.textstring = TextBox2  '90 METER
        If attribuut.TagString = "ALU50" And attribuut.textstring = "" Then attribuut.textstring = TextBox3  '60 METER
        If attribuut.TagString = "REGELUNITTYPE" And attribuut.textstring = "" Then attribuut.textstring = REGELTYPE  'TYPE REGELUNIT
        If attribuut.TagString = "BEVESTIGINGSTYPE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox1  'BEVESTIGING
        If attribuut.TagString = "PEALU200" And OptionButton6 = True Then attribuut.textstring = "200 meter :" 'Aluflex 200 METER
        If attribuut.TagString = "PEALU100" And OptionButton6 = True Then attribuut.textstring = "100 meter :" 'Aluflex 100 METER
        If attribuut.TagString = "PEALU50" And OptionButton6 = True Then attribuut.textstring = " 50 meter :" 'Aluflex 100 METER
       Next j
       
        End If
      End If
      End If
  Next element4
 End If
  Update
  
  
  
  Dim pb2(0 To 2) As Double
  If OptionButton1.Value = True Or OptionButton9.Value = True Then zakken = 210
  If OptionButton2.Value = True Or OptionButton6.Value = True Then zakken = 179
  If OptionButton9.Value = True And CheckBox2.Value = True Then zakken = 177
  If TextBox11 = "0" And TextBox25 <> "0" Then zakken = 179
  
  pb2(0) = pb1(0) - (scaal * 460)
  pb2(1) = pb1(1) + (scaal * zakken) '179) '177'210)
  pb2(2) = pb1(2)
  
  'juiste regelunitblokje.dwg inserten in de tekening
  bestand2 = ComboBox2 & ".dwg"
  If ComboBox2 = "RUW-Groot" Then bestand2 = "RUW" & ".dwg"
  If ComboBox2 = "RUW-Klein" Then bestand2 = "RUW" & ".dwg"
  If ComboBox2 = "RUB-R" And aantal_groepen > 4 Then bestand2 = "RUH-R" & ".dwg"
  If ComboBox2 = "RUB-RT" And aantal_groepen > 4 Then bestand2 = "RUH-RT" & ".dwg"
  If ComboBox2 = "RUB-S" And aantal_groepen > 0 Then bestand2 = "RUH-S" & ".dwg"
  If ComboBox2 = "VSKO" Then bestand2 = "VSKO-B" & ".dwg"
  bestand3 = "c:\acad2002\dwg\" & ComboBox2 & ".txt"  'tekstbestand met afmetingen
  If ComboBox2 = "VSKO" Then bestand3 = "c:\acad2002\dwg\" & ComboBox2 & "-B.txt"  'tekstbestand met afmetingen
  
  Update
  
  Call Unitblok.Unitblok(aantal_groepen, pb2, bestand2, bestand3, scaal)
  Call RESET
  Unload Me
  'frmGroeptekst.Show
  'Unload Me
End Sub

Sub Schaal(scaal)
frmGroeptekst.Hide
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


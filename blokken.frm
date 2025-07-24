VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} blokken 
   Caption         =   "BLOKKEN"
   ClientHeight    =   6684
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7632
   OleObjectBlob   =   "blokken.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "blokken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'14-08-2002 M.Bosch & G.C.Haak
'bijbehorende blokken:
'bltekor.dwg (1 radiator)
'bltekor1.dwg (2-3 radiator)
'bltekor2.dwg (4-6 radiator)
'bltekor3.dwg (7-9 radiator)
Dim nummer As Integer
Dim pbegin
Dim myValue As Integer

'Option Explicit

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
blokken.Hide
Dim bestand As String
bestand = "C:\acad2002\dwg\Bl-hr.dwg"
Dim pb1 As Variant
pb1 = ThisDrawing.Utility.GetPoint(, "Plaats startpunt....")
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb1, bestand, 1, 1, 1, 0)
Unload Me
End Sub


Private Sub UserForm_Initialize()

'begin sluiten toets uitschakelen
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
'einde sluiten toets uitschakelen

For I = 0 To MultiPage2.Count - 1 'met de 1e page laten beginnen
MultiPage2.Value = I - 1
Next I

nummer = 0

Call ThisDrawing.Combolijst1
Call ThisDrawing.Combolijst2
Call ThisDrawing.Combolijst3
Call ThisDrawing.Combolijst4
Call ThisDrawing.Combolijst5
Call ThisDrawing.Combolijst6
Call ThisDrawing.Combolijst7
Call ThisDrawing.Combolijst8
Call ThisDrawing.Combolijst9

Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)

If lognaam = "GERARD" Then blokken.StartUpPosition = 0

End Sub
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
End If

If CheckBox1.Value = True Then
blokken.Hide
frmbouwtek.Show
End If
End Sub
Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
End If

If CheckBox2.Value = True Then
blokken.Hide
frmBouwnummer.Show
End If
End Sub

Private Sub TextBox1_Change()
Dim a As Double
On Error Resume Next
a = TextBox1.Text
If Err Then
   TextBox1 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox2_Change()
Dim a As Double
On Error Resume Next
a = TextBox2.Text
If Err Then
   TextBox2 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox3_Change()
Dim a As Double
On Error Resume Next
a = TextBox3.Text
If Err Then
   TextBox3 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox4_Change()
Dim a As Double
On Error Resume Next
a = TextBox4.Text
If Err Then
   TextBox4 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox5_Change()
Dim a As Double
On Error Resume Next
a = TextBox5.Text
If Err Then
   TextBox5 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox6_Change()
Dim a As Double
On Error Resume Next
a = TextBox6.Text
If Err Then
   TextBox6 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox7_Change()
Dim a As Double
On Error Resume Next
a = TextBox7.Text
If Err Then
   TextBox7 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox8_Change()
Dim a As Double
On Error Resume Next
a = TextBox8.Text
If Err Then
   TextBox8 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox9_Change()
Dim a As Double
On Error Resume Next
a = TextBox9.Text
If Err Then
   TextBox9 = Clear
  Exit Sub
  End If
End Sub
Private Sub TextBox10_Change()
Dim a As Double
On Error Resume Next
a = TextBox10.Text
If a > 0 Then TextBox10 = Clear
End Sub
Private Sub TextBox11_Change()
Dim a As Double
On Error Resume Next
a = TextBox11.Text
If a > 0 Then TextBox11 = Clear
End Sub
Private Sub TextBox12_Change()
Dim a As Double
On Error Resume Next
a = TextBox12.Text
If a > 0 Then TextBox12 = Clear
End Sub
Private Sub TextBox13_Change()
Dim a As Double
On Error Resume Next
a = TextBox13.Text
If a > 0 Then TextBox13 = Clear
End Sub
Private Sub TextBox14_Change()
Dim a As Double
On Error Resume Next
a = TextBox14.Text
If a > 0 Then TextBox14 = Clear
End Sub
Private Sub TextBox15_Change()
Dim a As Double
On Error Resume Next
a = TextBox15.Text
If a > 0 Then TextBox15 = Clear
End Sub
Private Sub TextBox16_Change()
Dim a As Double
On Error Resume Next
a = TextBox16.Text
If a > 0 Then TextBox16 = Clear
End Sub
Private Sub TextBox17_Change()
Dim a As Double
On Error Resume Next
a = TextBox17.Text
If a > 0 Then TextBox17 = Clear
End Sub
Private Sub TextBox18_Change()
Dim a As Double
On Error Resume Next
a = TextBox18.Text
If a > 0 Then TextBox18 = Clear
End Sub
Private Sub OptionButton1_Click()
OptionButton10.Enabled = True
If OptionButton1.Value = True And OptionButton10.Value = False Then
ComboBox1.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = False: TextBox3.Visible = False
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = False: TextBox12.Visible = False
TextBox13.Visible = False: TextBox14.Visible = False: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = False: ComboBox3.Visible = False
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton1.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = False: ComboBox3.Visible = False
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1"
End If
End Sub
Private Sub OptionButton2_Click()
OptionButton10.Enabled = True
If OptionButton2.Value = True And OptionButton10.Value = False Then
ComboBox2.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = False
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = False
TextBox13.Visible = False: TextBox14.Visible = False: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = False
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton2.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = False
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
End If

End Sub
Private Sub OptionButton3_Click()
OptionButton10.Enabled = True
If OptionButton3.Value = True And OptionButton10.Value = False Then
ComboBox3.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = True
TextBox13.Visible = False: TextBox14.Visible = False: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton3.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3"
End If
End Sub
Private Sub OptionButton4_Click()
OptionButton10.Enabled = True
If OptionButton4.Value = True And OptionButton10.Value = False Then
ComboBox4.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True
TextBox4.Visible = True: TextBox5.Visible = False: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = True
TextBox13.Visible = True: TextBox14.Visible = False: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton4.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = False: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
End If
End Sub
Private Sub OptionButton5_Click()
OptionButton10.Enabled = True
If OptionButton5.Value = True And OptionButton10.Value = False Then
ComboBox5.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True
TextBox4.Visible = True: TextBox5.Visible = True: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = True
TextBox13.Visible = True: TextBox14.Visible = True: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton5.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = False
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5"
End If

End Sub
Private Sub OptionButton6_Click()
OptionButton10.Enabled = True
If OptionButton6.Value = True And OptionButton10.Value = False Then
ComboBox6.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True
TextBox4.Visible = True: TextBox5.Visible = True: TextBox6.Visible = True
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = True
TextBox13.Visible = True: TextBox14.Visible = True: TextBox15.Visible = True
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton6.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
End If

End Sub
Private Sub OptionButton7_Click()
OptionButton10.Enabled = True
If OptionButton7.Value = True And OptionButton10.Value = False Then
ComboBox7.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True
TextBox4.Visible = True: TextBox5.Visible = True: TextBox6.Visible = True
TextBox7.Visible = True: TextBox8.Visible = False: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = True
TextBox13.Visible = True: TextBox14.Visible = True: TextBox15.Visible = True
TextBox16.Visible = True: TextBox17.Visible = False: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = True: ComboBox8.Visible = False: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton7.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = True: ComboBox8.Visible = False: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
OptionButton7.Caption = "RUIMTE 7"
End If

End Sub
Private Sub OptionButton8_Click()
OptionButton10.Enabled = True
If OptionButton8.Value = True And OptionButton10.Value = False Then
ComboBox8.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True
TextBox4.Visible = True: TextBox5.Visible = True: TextBox6.Visible = True
TextBox7.Visible = True: TextBox8.Visible = True: TextBox9.Visible = False

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = True
TextBox13.Visible = True: TextBox14.Visible = True: TextBox15.Visible = True
TextBox16.Visible = True: TextBox17.Visible = True: TextBox18.Visible = False

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = True: ComboBox8.Visible = True: ComboBox9.Visible = False

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton8.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = True: ComboBox8.Visible = True: ComboBox9.Visible = False
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
OptionButton7.Caption = "RUIMTE 7": OptionButton8.Caption = "RUIMTE 8"
End If

End Sub
Private Sub OptionButton9_Click()
OptionButton10.Enabled = True
If OptionButton9.Value = True And OptionButton10.Value = False Then
ComboBox9.Left = 210: Label2.Left = 210: CmdButton1.Enabled = True
TextBox1.Visible = True: TextBox2.Visible = True: TextBox3.Visible = True
TextBox4.Visible = True: TextBox5.Visible = True: TextBox6.Visible = True
TextBox7.Visible = True: TextBox8.Visible = True: TextBox9.Visible = True

TextBox10.Visible = True: TextBox11.Visible = True: TextBox12.Visible = True
TextBox13.Visible = True: TextBox14.Visible = True: TextBox15.Visible = True
TextBox16.Visible = True: TextBox17.Visible = True: TextBox18.Visible = True

ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = True: ComboBox8.Visible = True: ComboBox9.Visible = True

TextBox1.SetFocus: Label1.Visible = True: Label2.Visible = True
End If

If OptionButton9.Value = True And OptionButton10.Value = True Then
ComboBox1.Visible = True: ComboBox2.Visible = True: ComboBox3.Visible = True
ComboBox4.Visible = True: ComboBox5.Visible = True: ComboBox6.Visible = True
ComboBox7.Visible = True: ComboBox8.Visible = True: ComboBox9.Visible = True
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
OptionButton7.Caption = "RUIMTE 7": OptionButton8.Caption = "RUIMTE 8"
OptionButton9.Caption = "RUIMTE 9"
End If
End Sub
Private Sub OptionButton10_Click()
Label1.Visible = False
Label2.Visible = True: Label2.Left = 84
ComboBox1.Left = 84: ComboBox2.Left = 84: ComboBox3.Left = 84
ComboBox4.Left = 84: ComboBox5.Left = 84: ComboBox6.Left = 84
ComboBox7.Left = 84: ComboBox8.Left = 84: ComboBox9.Left = 84
TextBox1.Visible = False: TextBox2.Visible = False: TextBox3.Visible = False
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False
TextBox10.Visible = False: TextBox11.Visible = False: TextBox12.Visible = False
TextBox13.Visible = False: TextBox14.Visible = False: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False

If OptionButton1.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox1.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton2.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox2.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton3.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox3.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton4.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox4.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton5.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox5.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton6.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox6.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton7.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox7.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
OptionButton7.Caption = "RUIMTE 7": Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton8.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox8.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
OptionButton7.Caption = "RUIMTE 7": OptionButton8.Caption = "RUIMTE 8"
Frame3.Caption = "Aanv.Verwarming"
End If
If OptionButton9.Value = True And OptionButton10.Value = True Then
CmdButton1.Enabled = True: ComboBox9.Visible = True: ComboBox1.SetFocus
OptionButton1.Caption = "RUIMTE 1": OptionButton2.Caption = "RUIMTE 2"
OptionButton3.Caption = "RUIMTE 3": OptionButton4.Caption = "RUIMTE 4"
OptionButton5.Caption = "RUIMTE 5": OptionButton6.Caption = "RUIMTE 6"
OptionButton7.Caption = "RUIMTE 7": OptionButton8.Caption = "RUIMTE 8"
OptionButton9.Caption = "RUIMTE 9": Frame3.Caption = "Aanv.Verwarming"
End If
End Sub
Private Sub blok1_Click()
   If blok1.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok2_Click()
If blok2.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok3_Click()
If blok3.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok4_Click()
If blok4.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok5_Click()
If blok5.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok6_Click()
If blok6.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok7_Click()
If blok7.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok8_Click()
If blok8.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok9_Click()
If blok9.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok10_Click()
If blok10.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok12_Click()
If blok12.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok13_Click()
  If blok13.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub

Private Sub blok100_Click()
If blok100.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok110_Click()
If blok110.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok120_Click()
If blok120.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok130_Click()
If blok130.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok140_Click()
If blok140.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
End Sub
Private Sub blok150_Click()
If blok150.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
If blok150.Value = True Then
blok160.Enabled = False
Else
blok160.Enabled = True
End If
End Sub
Private Sub blok160_Click()
If blok160.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
   
If blok160.Value = True Then
blok150.Enabled = False
Else
blok150.Enabled = True
End If
   
End Sub
Private Sub blok170_Click()
If blok170.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
   
If blok170.Value = True Then
blok180.Enabled = False
Else
blok180.Enabled = True
End If
   
End Sub
Private Sub blok180_Click()
If blok180.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
   End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If
   
If blok180.Value = True Then
blok170.Enabled = False
Else
blok170.Enabled = True
End If
   
End Sub
Private Sub blok190_Click()
If blok190.Value = True Then
      nummer = nummer + 1
   Else
      nummer = nummer - 1
 End If
   If nummer > 0 Then
      CmdButton1.Enabled = True
   Else
      CmdButton1.Enabled = False
   End If

End Sub
Private Sub CmdButton1_Click()
Dim newLayer As AcadLayer
On Error Resume Next
Set newLayer = ThisDrawing.Layers.Add("3")
ThisDrawing.ActiveLayer = newLayer
Update


If TextBox1.Text = "" And TextBox1.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox1.SetFocus
Exit Sub
End If

If ComboBox1.Value = "" And ComboBox1.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox1.SetFocus
Exit Sub
End If

If TextBox2.Text = "" And TextBox2.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox2.SetFocus
Exit Sub
End If

If ComboBox2.Value = "" And ComboBox2.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox2.SetFocus
Exit Sub
End If

If TextBox3.Text = "" And TextBox3.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox3.SetFocus
Exit Sub
End If

If ComboBox3.Value = "" And ComboBox3.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox3.SetFocus
Exit Sub
End If

If TextBox4.Text = "" And TextBox4.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox4.SetFocus
Exit Sub
End If

If ComboBox4.Value = "" And ComboBox4.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox4.SetFocus
Exit Sub
End If

If TextBox5.Text = "" And TextBox5.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox5.SetFocus
Exit Sub
End If

If ComboBox5.Value = "" And ComboBox5.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox5.SetFocus
Exit Sub
End If

If TextBox6.Text = "" And TextBox6.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox6.SetFocus
Exit Sub
End If

If ComboBox6.Value = "" And ComboBox6.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox6.SetFocus
Exit Sub
End If

If TextBox7.Text = "" And TextBox7.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox7.SetFocus
Exit Sub
End If

If ComboBox7.Value = "" And ComboBox7.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox7.SetFocus
Exit Sub
End If

If TextBox8.Text = "" And TextBox8.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox8.SetFocus
Exit Sub
End If

If ComboBox8.Value = "" And ComboBox8.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox8.SetFocus
Exit Sub
End If

If TextBox9.Text = "" And TextBox9.Visible = True Then
MsgBox "Je bent vergeten een tekort in te vullen...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
TextBox9.SetFocus
Exit Sub
End If

If ComboBox9.Value = "" And ComboBox9.Visible = True Then
MsgBox "Je bent vergeten een ruimtenaam te selecteren...!!!"
For I = 0 To MultiPage2.Count - 1
MultiPage2.Value = I
Next I
ComboBox9.SetFocus
Exit Sub
End If

'blokken.Hide


Call Schaal(scaal)
ThisDrawing.SetVariable "osmode", 1
Dim PBEGINZEROnul As Variant
Dim pbeginzero(0 To 2) As Double

If blok1.Value = True Or blok2.Value = True Or blok3.Value = True Or _
   blok4.Value = True Or blok5.Value = True Or blok6.Value = True Or blok7.Value = True Or _
   blok8.Value = True Or blok9.Value = True Or blok10.Value = True Or _
   blok12.Value = True Or blok13.Value = True Then
   blokken.Hide
   PBEGINZEROnul = ThisDrawing.Utility.GetPoint(, "Plaats STARTPUNT")
   pbeginzero(0) = PBEGINZEROnul(0)
   pbeginzero(1) = PBEGINZEROnul(1) - (77 * scaal)
    If Err Then
    'Call reset2
    blokken.Show
    Exit Sub
    End If
End If
     
      
       

If blok5.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok51(pbeginzero)
       End If
       
If blok13.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok131(pbeginzero)
       End If
       
If blok1.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok11(pbeginzero)
       End If

If blok2.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok21(pbeginzero)
       End If
If blok6.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok61(pbeginzero)
       End If
If blok7.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok71(pbeginzero)
       End If

'If blok12.Value = True Then
'       PBEGINZERO(0) = PBEGINZERO(0)
'       PBEGINZERO(1) = PBEGINZERO(1) + (77 * scaal)
'       Call blok121(PBEGINZERO)
'       End If
       

Dim attributen As Variant
Dim element As AcadEntity
Dim attribuut As AcadAttributeReference
Dim symbool As AcadBlockReference

ThisDrawing.SetVariable "osmode", 1
blokken.Hide
On Error Resume Next

If OptionButton1 = True Then
     If blok1.Value = False And blok2.Value = False And blok3.Value = False And _
       blok4.Value = False And blok5.Value = False And blok6.Value = False And blok7.Value = False And _
       blok8.Value = False And blok9.Value = False And blok10.Value = False And blok12.Value = False Then
       pBEGINZEROstart = ThisDrawing.Utility.GetPoint(, "Plaats STARTPUNT")
       pbeginzero(0) = pBEGINZEROstart(0)
       pbeginzero(1) = pBEGINZEROstart(1)
          
     Else
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
    End If
    
If OptionButton10.Value = False Then
Tbestand = "C:\ACAD2002\DWG\bltekor.dwg"
Else
Tbestand = "C:\ACAD2002\DWG\avbltekor.dwg"
End If
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, Tbestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Call optie1
End If

If OptionButton2 = True Or OptionButton3 = True Then
       If blok1.Value = False And blok2.Value = False And blok3.Value = False And _
       blok4.Value = False And blok5.Value = False And blok6.Value = False And blok7.Value = False And _
       blok8.Value = False And blok9.Value = False And blok10.Value = False And blok12.Value = False Then
       pBEGINZEROstart = ThisDrawing.Utility.GetPoint(, "Plaats STARTPUNT")
       pbeginzero(0) = pBEGINZEROstart(0)
       pbeginzero(1) = pBEGINZEROstart(1)
          
     Else
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
    End If

If OptionButton10.Value = False Then
Tbestand = "C:\ACAD2002\DWG\bltekor1.dwg"
Else
Tbestand = "C:\ACAD2002\DWG\avbltekor1.dwg"
End If

Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, Tbestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Call optie2
End If

If OptionButton4 = True Or OptionButton5 = True Or OptionButton6 = True Then
       If blok1.Value = False And blok2.Value = False And blok3.Value = False And _
       blok4.Value = False And blok5.Value = False And blok6.Value = False And blok7.Value = False And _
       blok8.Value = False And blok9.Value = False And blok10.Value = False And blok12.Value = False Then
       pBEGINZEROstart = ThisDrawing.Utility.GetPoint(, "Plaats STARTPUNT")
       pbeginzero(0) = pBEGINZEROstart(0)
       pbeginzero(1) = pBEGINZEROstart(1)
          
     Else
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
    End If
    
If OptionButton10.Value = False Then
Tbestand = "C:\ACAD2002\DWG\bltekor2.dwg"
Else
Tbestand = "C:\ACAD2002\DWG\avbltekor2.dwg"
End If
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, Tbestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Call optie3
End If

If OptionButton7 = True Or OptionButton8 = True Or OptionButton9 = True Then
       If blok1.Value = False And blok2.Value = False And blok3.Value = False And _
       blok4.Value = False And blok5.Value = False And blok6.Value = False And blok7.Value = False And _
       blok8.Value = False And blok9.Value = False And blok10.Value = False And blok12.Value = False Then
       pBEGINZEROstart = ThisDrawing.Utility.GetPoint(, "Plaats STARTPUNT")
       pbeginzero(0) = pBEGINZEROstart(0)
       pbeginzero(1) = pBEGINZEROstart(1)
     Else
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
    End If
    
If OptionButton10.Value = False Then
Tbestand = "C:\ACAD2002\DWG\bltekor3.dwg"
Else
Tbestand = "C:\ACAD2002\DWG\avbltekor3.dwg"
End If
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, Tbestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Call optie4
End If
Update
'--------------DEFINITIEF OF REVISIE-OF VOORLOPIG OF TER GOEDKEURING OF NAREGELEN--of UITVOERING-----------------------------------------
checkdef = 0
If blok3.Value = True And OptionButton1 = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (113.7 * scaal)
       checkdef = 1
       Call blok31(pbeginzero) 'blok definitief
       End If

If OptionButton2 = True Or OptionButton3 = True Then tekorten23 = 1
If blok3.Value = True And tekorten23 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (162.7 * scaal) '162.7
       checkdef = 1
       Call blok31(pbeginzero) 'blok definitief
       End If

If OptionButton4 = True Or OptionButton5 = True Or OptionButton6 = True Then tekorten456 = 1
If blok3.Value = True And tekorten456 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (238.7 * scaal)
       checkdef = 1
       Call blok31(pbeginzero) 'blok definitief
       End If
       
If OptionButton7 = True Or OptionButton8 = True Or OptionButton9 = True Then tekorten789 = 1
If blok3.Value = True And tekorten789 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (305.7 * scaal)
       checkdef = 1
       Call blok31(pbeginzero) 'blok definitief
       End If

If checkdef = 0 And blok3.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok31(pbeginzero) 'blok definitief
       End If
'----------------------------------------------------------------------------
checkrev = 0
If blok4.Value = True And OptionButton1 = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (113.7 * scaal)
       checkrev = 1
       Call blok41(pbeginzero) 'blok revisie
       End If

If blok4.Value = True And tekorten23 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (162.7 * scaal)
       checkrev = 1
       Call blok41(pbeginzero) 'blok revisie
       End If

If blok4.Value = True And tekorten456 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (238.7 * scaal)
       checkrev = 1
       Call blok41(pbeginzero) 'blok revisie
       End If
       
If blok4.Value = True And tekorten789 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (305.7 * scaal)
       checkrev = 1
       Call blok41(pbeginzero) 'blok revisie
       End If

If checkrev = 0 And blok4.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok41(pbeginzero)
       End If
       
       
checkgoed = 0
If blok8.Value = True And OptionButton1 = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (113.7 * scaal)
       checkgoed = 1
       Call blok81(pbeginzero) 'blok Ter Goedkeuring
       End If

If OptionButton2 = True Or OptionButton3 = True Then tekorten23 = 1
If blok8.Value = True And tekorten23 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (162.7 * scaal) '162.7
       checkgoed = 1
       Call blok81(pbeginzero) 'blok Ter Goedkeuring
       End If

If OptionButton4 = True Or OptionButton5 = True Or OptionButton6 = True Then tekorten456 = 1
If blok8.Value = True And tekorten456 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (238.7 * scaal)
       checkgoed = 1
       Call blok81(pbeginzero) 'blok Ter Goedkeuring
       End If
       
If OptionButton7 = True Or OptionButton8 = True Or OptionButton9 = True Then tekorten789 = 1
If blok8.Value = True And tekorten789 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (305.7 * scaal)
       checkgoed = 1
       Call blok81(pbeginzero) 'blok Ter Goedkeuring
       End If

If checkgoed = 0 And blok8.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok81(pbeginzero) 'blok Ter Goedkeuring
       End If
       
       
checkvoor = 0
If blok9.Value = True And OptionButton1 = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (113.7 * scaal)
       checkvoor = 1
       Call blok91(pbeginzero) 'blok voorlopig
       End If

If OptionButton2 = True Or OptionButton3 = True Then tekorten23 = 1
If blok9.Value = True And tekorten23 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (162.7 * scaal) '162.7
       checkvoor = 1
       Call blok91(pbeginzero) 'blok voorlopig
       End If

If OptionButton4 = True Or OptionButton5 = True Or OptionButton6 = True Then tekorten456 = 1
If blok9.Value = True And tekorten456 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (238.7 * scaal)
       checkvoor = 1
       Call blok91(pbeginzero) 'blok voorlopig
       End If
       
If OptionButton7 = True Or OptionButton8 = True Or OptionButton9 = True Then tekorten789 = 1
If blok9.Value = True And tekorten789 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (305.7 * scaal)
       checkvoor = 1
       Call blok91(pbeginzero) 'blok voorlopig
       End If

If checkvoor = 0 And blok9.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok91(pbeginzero) 'blok voorlopig
       End If
       
checkuitvoering = 0
If blok12.Value = True And OptionButton1 = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (113.7 * scaal)
       checkuitvoering = 1
       Call blok121(pbeginzero) 'blok checkuitvoering
       End If

If OptionButton2 = True Or OptionButton3 = True Then tekorten23 = 1
If blok12.Value = True And tekorten23 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (162.7 * scaal) '162.7
       checkuitvoering = 1
       Call blok121(pbeginzero) 'blok checkuitvoering
       End If

If OptionButton4 = True Or OptionButton5 = True Or OptionButton6 = True Then tekorten456 = 1
If blok12.Value = True And tekorten456 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (238.7 * scaal)
       checkuitvoering = 1
       Call blok121(pbeginzero) 'blok checkuitvoering
       End If
       
If OptionButton7 = True Or OptionButton8 = True Or OptionButton9 = True Then tekorten789 = 1
If blok12.Value = True And tekorten789 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (305.7 * scaal)
       checkuitvoering = 1
       Call blok121(pbeginzero) 'blok checkuitvoering
       End If

If checkuitvoering = 0 And blok12.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok121(pbeginzero) 'blok checkuitvoering
       End If
       
       
If checkdef = 1 Or checkrev = 1 Or checkgoed = 1 Or checkvoor = 1 Or checkuitvoering = 1 Then hh = 1
checknaregelen = 0

If hh <> 1 Then
 If blok10.Value = True And OptionButton1 = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (113.7 * scaal)
       checknaregelen = 1
       Call blok101(pbeginzero) 'blok voorlopig
       End If

 If OptionButton2 = True Or OptionButton3 = True Then tekorten23 = 1
 If blok10.Value = True And tekorten23 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (162.7 * scaal) '162.7
       checknaregelen = 1
       Call blok101(pbeginzero) 'blok voorlopig
       End If

 If OptionButton4 = True Or OptionButton5 = True Or OptionButton6 = True Then tekorten456 = 1
 If blok10.Value = True And tekorten456 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (238.7 * scaal)
       checknaregelen = 1
       Call blok101(pbeginzero) 'blok voorlopig
       End If
       
 If OptionButton7 = True Or OptionButton8 = True Or OptionButton9 = True Then tekorten789 = 1
 If blok10.Value = True And tekorten789 = 1 Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (305.7 * scaal)
       checknaregelen = 1
       Call blok101(pbeginzero) 'blok voorlopig
       End If
End If

   
   If checknaregelen = 0 And blok10.Value = True Then
       pbeginzero(0) = pbeginzero(0)
       pbeginzero(1) = pbeginzero(1) + (77 * scaal)
       Call blok101(pbeginzero) 'blok voorlopig
       End If

'-----------------------------------------------------------------
'blok bl-bouwk vullen

If blokken.CheckBox1.Value = True Then
        Call bouwk
        Dim element10
        For Each element10 In ThisDrawing.ModelSpace
              If element10.ObjectName = "AcDbBlockReference" Then
              If UCase(element10.Name) = "BL-BOUWK" Then
              Set symbool = element10
                If symbool.HasAttributes Then
                attributen = symbool.GetAttributes
                For I = LBound(attributen) To UBound(attributen)
                Set attribuut = attributen(I)
                If attribuut.TagString = "BOUWDATUM" Then attribuut.textstring = UCase(blokken.TextBox19)
                If attribuut.TagString = "BOUWNAAM" Then attribuut.textstring = UCase(blokken.TextBox20)
                
               Next I
        
              End If
              End If
              End If
        Next element10
End If

'blok bouwnummer vullen
If blokken.CheckBox2.Value = True Then
      Call bouwn
      Dim element11
      For Each element11 In ThisDrawing.ModelSpace
          If element11.ObjectName = "AcDbBlockReference" Then
          If UCase(element11.Name) = "BOUWNUMMER" Then
          Set symbool = element11
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For k = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(k)
            If attribuut.TagString = "BOUWNUMMER" Then attribuut.textstring = UCase(blokken.TextBox21)
            If attribuut.TagString = "WONINGTYPE" Then attribuut.textstring = UCase(blokken.TextBox22)
            
           Next k
    
          End If
          End If
          End If
    Next element11
End If
Update
       
ThisDrawing.SetVariable "osmode", 0
If blok100.Value = True Then Call blok1000
If blok110.Value = True Then Call blok1100
If blok120.Value = True Then Call blok1200
If blok130.Value = True Then Call blok1300
If blok140.Value = True Then Call blok1400
If blok150.Value = True Then Call blok1500
If blok160.Value = True Then Call blok1600
If blok170.Value = True Then Call blok1700
If blok180.Value = True Then Call blok1800
If blok190.Value = True Then Call blok1900

Call reset
End Sub
Private Sub bouwk()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
ThisDrawing.SetVariable "osmode", 1
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Bouwkundige tekening van de klant- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\bl-bouwk.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
ThisDrawing.SetVariable "osmode", 0
End Sub
Private Sub bouwn()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
ThisDrawing.SetVariable "osmode", 1
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Bouwnummer ~ Type woning- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\bouwnummer.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
ThisDrawing.SetVariable "osmode", 0
End Sub
Private Sub CancelButton_Click()
Call reset
blokken.Hide
End Sub
Private Sub CommandButton1_Click()
Call reset2
End Sub
Sub reset()
ComboBox1.Value = Clear: ComboBox2.Value = Clear: ComboBox3.Value = Clear:
ComboBox4.Value = Clear: ComboBox5.Value = Clear: ComboBox6.Value = Clear:
ComboBox7.Value = Clear: ComboBox8.Value = Clear: ComboBox9.Value = Clear:
TextBox1.Visible = False: TextBox2.Visible = False: TextBox3.Visible = False
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False
TextBox10.Visible = False: TextBox11.Visible = False: TextBox12.Visible = False
TextBox13.Visible = False: TextBox14.Visible = False: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False
blok1.Value = False: blok2.Value = False: blok3.Value = False
blok4.Value = False: blok5.Value = False: blok6.Value = False: blok7.Value = False
blok8.Value = False: blok9.Value = False: blok10.Value = False
blok13.Value = False
blok100.Value = False: blok110.Value = False: blok120.Value = False
blok130.Value = False: blok140.Value = False: blok150.Value = False
blok160.Value = False: blok190.Value = False
OptionButton1 = False: OptionButton2 = False: OptionButton3 = False
OptionButton4 = False: OptionButton5 = False: OptionButton6 = False
OptionButton7 = False: OptionButton8 = False: OptionButton9 = False
CmdButton1.Enabled = False
TextBox1 = Clear: TextBox2 = Clear: TextBox3 = Clear: TextBox4 = Clear: TextBox5 = Clear
TextBox6 = Clear: TextBox7 = Clear: TextBox8 = Clear: TextBox9 = Clear: TextBox10 = Clear
TextBox11 = Clear: TextBox12 = Clear: TextBox13 = Clear: TextBox14 = Clear: TextBox15 = Clear
TextBox16 = Clear: TextBox17 = Clear: TextBox18 = Clear: TextBox19 = Clear: TextBox20 = Clear
TextBox21 = Clear: TextBox22 = Clear: Label1.Visible = False: Label2.Visible = False
ComboBox1.Visible = False: ComboBox2.Visible = False: ComboBox3.Visible = False:
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False:
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
ThisDrawing.SetVariable "osmode", 0
checkdef = 0: checkrev = 0
tekorten23 = 0: tekorten456 = 0: tekorten789 = 0
OptionButton10 = False
OptionButton10.Enabled = False
ComboBox1.Left = 210: Label2.Left = 210
ComboBox2.Left = 210: ComboBox3.Left = 210: ComboBox4.Left = 210
ComboBox5.Left = 210: ComboBox6.Left = 210: ComboBox7.Left = 210
ComboBox8.Left = 210: ComboBox9.Left = 210
blokken.CheckBox1 = False
blokken.CheckBox2 = False
Unload Me
'ThisDrawing.SendCommand "vbaunload" & vbCr & "blokken" & vbCr
End Sub
Sub reset2()
ComboBox1.Value = Clear: ComboBox2.Value = Clear: ComboBox3.Value = Clear:
ComboBox4.Value = Clear: ComboBox5.Value = Clear: ComboBox6.Value = Clear:
ComboBox7.Value = Clear: ComboBox8.Value = Clear: ComboBox9.Value = Clear:
TextBox1.Visible = False: TextBox2.Visible = False: TextBox3.Visible = False
TextBox4.Visible = False: TextBox5.Visible = False: TextBox6.Visible = False
TextBox7.Visible = False: TextBox8.Visible = False: TextBox9.Visible = False
TextBox10.Visible = False: TextBox11.Visible = False: TextBox12.Visible = False
TextBox13.Visible = False: TextBox14.Visible = False: TextBox15.Visible = False
TextBox16.Visible = False: TextBox17.Visible = False: TextBox18.Visible = False
blok1.Value = False: blok2.Value = False: blok3.Value = False
blok4.Value = False: blok5.Value = False: blok6.Value = False: blok7.Value = False
blok8.Value = False: blok9.Value = False: blok10.Value = False: blok12.Value = False
blok13.Value = False
blok100.Value = False: blok110.Value = False: blok120.Value = False
blok130.Value = False: blok140.Value = False: blok150.Value = False
blok160.Value = False: blok170.Value = False: blok180.Value = False: blok190.Value = False
OptionButton1 = False: OptionButton2 = False: OptionButton3 = False
OptionButton4 = False: OptionButton5 = False: OptionButton6 = False
OptionButton7 = False: OptionButton8 = False: OptionButton9 = False
CmdButton1.Enabled = False
TextBox1 = Clear: TextBox2 = Clear: TextBox3 = Clear: TextBox4 = Clear: TextBox5 = Clear
TextBox6 = Clear: TextBox7 = Clear: TextBox8 = Clear: TextBox9 = Clear: TextBox10 = Clear
TextBox11 = Clear: TextBox12 = Clear: TextBox13 = Clear: TextBox14 = Clear: TextBox15 = Clear
TextBox16 = Clear: TextBox17 = Clear: TextBox18 = Clear: TextBox19 = Clear: TextBox20 = Clear
TextBox21 = Clear: TextBox22 = Clear: Label1.Visible = False: Label2.Visible = False
ComboBox1.Visible = False: ComboBox2.Visible = False: ComboBox3.Visible = False:
ComboBox4.Visible = False: ComboBox5.Visible = False: ComboBox6.Visible = False:
ComboBox7.Visible = False: ComboBox8.Visible = False: ComboBox9.Visible = False
checkdef = 0: checkrev = 0
tekorten23 = 0: tekorten456 = 0: tekorten789 = 0
For Z = 0 To MultiPage2.Count - 1
MultiPage2.Value = Z - 1
Next Z
OptionButton10 = False
OptionButton10.Enabled = False
ComboBox1.Left = 210: Label2.Left = 210
ComboBox2.Left = 210: ComboBox3.Left = 210: ComboBox4.Left = 210
ComboBox5.Left = 210: ComboBox6.Left = 210: ComboBox7.Left = 210
ComboBox8.Left = 210: ComboBox9.Left = 210
OptionButton1.Caption = "1 Radiator": OptionButton2.Caption = "2 Radiatoren"
OptionButton3.Caption = "3 Radiatoren": OptionButton4.Caption = "4 Radiatoren"
OptionButton5.Caption = "5 Radiatoren": OptionButton6.Caption = "6 Radiatoren"
OptionButton7.Caption = "7 Radiatoren": OptionButton8.Caption = "8 Radiatoren"
OptionButton9.Caption = "9 Radiatoren"
Frame3.Caption = "Radiator blokken"
blokken.CheckBox1 = False
blokken.CheckBox2 = False
End Sub

Sub berekenpunt()

End Sub
Sub blok11(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim blockRefObj As Object
Dim bestand As String
bestand = "C:\acad2002\dwg\BL-hr.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)



If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok21(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String


bestand = "C:\acad2002\dwg\Bl-bouw.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok31(pbeginzero)
blokken.Hide

On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\definitief2.dwg"

Dim blockRefObj
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok41(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\Bl-revisie.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok51(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\Bl-krimp.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok61(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\Bl-exp.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok71(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\Bl-keuk.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok81(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\goedkeuring.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok91(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\voorlopig.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok101(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\allesnaregelen.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok121(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\uitvoering.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok131(pbeginzero)
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\bl-recht.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbeginzero, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1000()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Dilatatie [mantelbuis toepassen]")
Dim bestand As String
bestand = "C:\acad2002\dwg\dilatati.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1100()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Duoleiding- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\bl-duo.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1200()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
If blok110.Value = True And blok120.Value = True Then ThisDrawing.SetVariable "osmode", 1
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -H.O.H. 10 cm- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\bl-hoh10.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If blok110.Value = True And blok120.Value = True Then ThisDrawing.SetVariable "osmode", 0

If Err Then
    ThisDrawing.SetVariable "osmode", 0
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1300()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Nivoverschil- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\bl-nivo.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1400()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -schetstekening- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\schets.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1500()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -wandverwarming 2 meter hoog- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\wand1.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1600()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -wandverwarming 2,5 meter hoog- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\wand2.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1700()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Vriescel (Kruispatroon)- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\vriescel1.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1800()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim bestand As String
bestand = "C:\acad2002\dwg\vriescel2.dwg"
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Vriescel (volgens opritprincipe)- ")
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub blok1900()
blokken.Hide
On Error Resume Next
Call Schaal(scaal)
Dim pbegin7 As Variant
pbegin7 = ThisDrawing.Utility.GetPoint(, "Plaats het blok -Mantelbuis toepassen- ")
Dim bestand As String
bestand = "C:\acad2002\dwg\mantel.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin7, bestand, scaal, scaal, 1, 0)
If Err Then
    blokken.Show
    Exit Sub
    End If
Update
End Sub
Sub optie1()
On Error Resume Next
For Each element In ThisDrawing.ModelSpace
    If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "bltekor" Or element.Name = "avbltekor" Then
        Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
            If TextBox1 <> "" Then w1 = TextBox1 & " Watt" & " " & TextBox10 & " " & ComboBox1
            If OptionButton10.Value = True Then w1 = ComboBox1
            If attribuut.TagString = "RUIMTE(N)" And attribuut.textstring = "" Then attribuut.textstring = w1
            Next I
            End If
        End If
    End If
  Next element


If Err Then
    blokken.Show
    Exit Sub
End If
Update
'blok.Show
End Sub
Sub optie2()
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
        If element.Name = "bltekor1" Or element.Name = "avbltekor1" Then
        Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
            If TextBox1 <> "" Then w1 = TextBox1 & " Watt" & " " & TextBox10 & " " & ComboBox1
            If TextBox2 <> "" Then w2 = TextBox2 & " Watt" & " " & TextBox11 & " " & ComboBox2
            If TextBox3 <> "" Then w3 = TextBox3 & " Watt" & " " & TextBox12 & " " & ComboBox3
            If OptionButton10.Value = True Then
            w1 = ComboBox1
            w2 = ComboBox2
            w3 = ComboBox3
            End If
            If attribuut.TagString = "RUIMTE1" And attribuut.textstring = "" Then attribuut.textstring = w1
            If attribuut.TagString = "RUIMTE2" And attribuut.textstring = "" Then attribuut.textstring = w2
            If attribuut.TagString = "RUIMTE3" And attribuut.textstring = "" Then attribuut.textstring = w3
            Next I
            End If
        End If
      End If
Next element
Update
blok.Show
End Sub
Sub optie3()
For Each element In ThisDrawing.ModelSpace
    If element.ObjectName = "AcDbBlockReference" Then
        If element.Name = "bltekor2" Or element.Name = "avbltekor2" Then
        Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
            If TextBox1 <> "" Then w1 = TextBox1 & " Watt" & " " & TextBox10 & " " & ComboBox1
            If TextBox2 <> "" Then w2 = TextBox2 & " Watt" & " " & TextBox11 & " " & ComboBox2
            If TextBox3 <> "" Then w3 = TextBox3 & " Watt" & " " & TextBox12 & " " & ComboBox3
            If TextBox4 <> "" Then w4 = TextBox4 & " Watt" & " " & TextBox13 & " " & ComboBox4
            If TextBox5 <> "" Then w5 = TextBox5 & " Watt" & " " & TextBox14 & " " & ComboBox5
            If TextBox6 <> "" Then w6 = TextBox6 & " Watt" & " " & TextBox15 & " " & ComboBox6
            If OptionButton10.Value = True Then
            w1 = ComboBox1
            w2 = ComboBox2
            w3 = ComboBox3
            w4 = ComboBox4
            w5 = ComboBox5
            w6 = ComboBox6
            End If
            If attribuut.TagString = "RUIMTE1" And attribuut.textstring = "" Then attribuut.textstring = w1
            If attribuut.TagString = "RUIMTE2" And attribuut.textstring = "" Then attribuut.textstring = w2
            If attribuut.TagString = "RUIMTE3" And attribuut.textstring = "" Then attribuut.textstring = w3
            If attribuut.TagString = "RUIMTE4" And attribuut.textstring = "" Then attribuut.textstring = w4
            If attribuut.TagString = "RUIMTE5" And attribuut.textstring = "" Then attribuut.textstring = w5
            If attribuut.TagString = "RUIMTE6" And attribuut.textstring = "" Then attribuut.textstring = w6
            Next I
            End If
        End If
    End If
Next element
Update
blok.Show
End Sub
Sub optie4()
For Each element In ThisDrawing.ModelSpace
    If element.ObjectName = "AcDbBlockReference" Then
        If element.Name = "bltekor3" Or element.Name = "avbltekor3" Then
        Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For I = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(I)
            If TextBox1 <> "" Then w1 = TextBox1 & " Watt" & " " & TextBox10 & " " & ComboBox1
            If TextBox2 <> "" Then w2 = TextBox2 & " Watt" & " " & TextBox11 & " " & ComboBox2
            If TextBox3 <> "" Then w3 = TextBox3 & " Watt" & " " & TextBox12 & " " & ComboBox3
            If TextBox4 <> "" Then w4 = TextBox4 & " Watt" & " " & TextBox13 & " " & ComboBox4
            If TextBox5 <> "" Then w5 = TextBox5 & " Watt" & " " & TextBox14 & " " & ComboBox5
            If TextBox6 <> "" Then w6 = TextBox6 & " Watt" & " " & TextBox15 & " " & ComboBox6
            If TextBox7 <> "" Then w7 = TextBox7 & " Watt" & " " & TextBox16 & " " & ComboBox7
            If TextBox8 <> "" Then w8 = TextBox8 & " Watt" & " " & TextBox17 & " " & ComboBox8
            If TextBox9 <> "" Then w9 = TextBox9 & " Watt" & " " & TextBox18 & " " & ComboBox9
            If OptionButton10.Value = True Then
            w1 = ComboBox1
            w2 = ComboBox2
            w3 = ComboBox3
            w4 = ComboBox4
            w5 = ComboBox5
            w6 = ComboBox6
            w7 = ComboBox7
            w8 = ComboBox8
            w9 = ComboBox9
            End If
            If attribuut.TagString = "RUIMTE1" And attribuut.textstring = "" Then attribuut.textstring = w1
            If attribuut.TagString = "RUIMTE2" And attribuut.textstring = "" Then attribuut.textstring = w2
            If attribuut.TagString = "RUIMTE3" And attribuut.textstring = "" Then attribuut.textstring = w3
            If attribuut.TagString = "RUIMTE4" And attribuut.textstring = "" Then attribuut.textstring = w4
            If attribuut.TagString = "RUIMTE5" And attribuut.textstring = "" Then attribuut.textstring = w5
            If attribuut.TagString = "RUIMTE6" And attribuut.textstring = "" Then attribuut.textstring = w6
            If attribuut.TagString = "RUIMTE7" And attribuut.textstring = "" Then attribuut.textstring = w7
            If attribuut.TagString = "RUIMTE8" And attribuut.textstring = "" Then attribuut.textstring = w8
            If attribuut.TagString = "RUIMTE9" And attribuut.textstring = "" Then attribuut.textstring = w9
            Next I
        End If
      End If
    End If
Next element
Update
blok.Show
End Sub
Private Sub ComboBox1_Change()
If ComboBox1.Text = "Kantoor" Or ComboBox1.Text = "Toilet" Or _
ComboBox1.Text = "Zwembad" Or ComboBox1.Text = "Portaal" Or _
ComboBox1.Text = "Bureel" Then
  TextBox10.Value = "in het"
Else
  TextBox10.Value = "in de"
End If
If ComboBox1.Text = "Kantoor1" Or ComboBox1.Text = "Kantoor2" Or _
ComboBox1.Text = "Verblijfsruimte1" Or ComboBox1.Text = "Verblijfsruimte2" Or _
ComboBox1.Text = "Groepsruimte1" Or ComboBox1.Text = "Groepsruimte2" Or _
ComboBox1.Text = "Slaapkamer1" Or ComboBox1.Text = "Slaapkamer2" Or _
ComboBox1.Text = "Slaapkamer3" Or ComboBox1.Text = "Slaapkamer4" Or _
ComboBox1.Text = "Bedrijfsruimte1" Or ComboBox1.Text = "Bedrijfsruimte2" Then
  TextBox10.Value = "in"
End If
If ComboBox1.Text = "Verkoop" Or ComboBox1.Text = "Zolder" Or ComboBox1.Text = "Overloop" Then
  TextBox10.Value = "op de"
End If
If ComboBox1.Text = "" Then
  TextBox10.Value = Clear
  TextBox10.Locked = True
Else
  TextBox10.Locked = False
  End If
  
End Sub
Private Sub ComboBox2_Change()

If ComboBox2.Text = "Kantoor" Or ComboBox2.Text = "Toilet" Or _
ComboBox2.Text = "Zwembad" Or ComboBox2.Text = "Portaal" Or _
ComboBox2.Text = "Bureel" Then
  TextBox11.Value = "in het"
 Else
  TextBox11.Value = "in de"
End If
If ComboBox2.Text = "Kantoor1" Or ComboBox2.Text = "Kantoor2" Or _
ComboBox2.Text = "Verblijfsruimte1" Or ComboBox2.Text = "Verblijfsruimte2" Or _
ComboBox2.Text = "Groepsruimte1" Or ComboBox2.Text = "Groepsruimte2" Or _
ComboBox2.Text = "Slaapkamer1" Or ComboBox2.Text = "Slaapkamer2" Or _
ComboBox2.Text = "Slaapkamer3" Or ComboBox2.Text = "Slaapkamer4" Or _
ComboBox2.Text = "Bedrijfsruimte1" Or ComboBox2.Text = "Bedrijfsruimte2" Then
  TextBox11.Value = "in"
End If
If ComboBox2.Text = "Verkoop" Or ComboBox2.Text = "Zolder" Or ComboBox2.Text = "Overloop" Then
  TextBox11.Value = "op de"
End If
If ComboBox2.Text = "" Then
  TextBox11.Value = Clear
  TextBox11.Locked = True
Else
  TextBox11.Locked = False
  End If

End Sub
Private Sub ComboBox3_Change()

If ComboBox3.Text = "Kantoor" Or ComboBox3.Text = "Toilet" Or _
ComboBox3.Text = "Zwembad" Or ComboBox3.Text = "Portaal" Or _
ComboBox3.Text = "Bureel" Then
 TextBox12.Value = "in het"
 Else
  TextBox12.Value = "in de"
End If
If ComboBox3.Text = "Kantoor1" Or ComboBox3.Text = "Kantoor2" Or _
ComboBox3.Text = "Verblijfsruimte1" Or ComboBox3.Text = "Verblijfsruimte2" Or _
ComboBox3.Text = "Groepsruimte1" Or ComboBox3.Text = "Groepsruimte2" Or _
ComboBox3.Text = "Slaapkamer1" Or ComboBox3.Text = "Slaapkamer2" Or _
ComboBox3.Text = "Slaapkamer3" Or ComboBox3.Text = "Slaapkamer4" Or _
ComboBox3.Text = "Bedrijfsruimte1" Or ComboBox3.Text = "Bedrijfsruimte2" Then
  TextBox12.Value = "in"
End If
If ComboBox3.Text = "Verkoop" Or ComboBox3.Text = "Zolder" Or ComboBox3.Text = "Overloop" Then
  TextBox12.Value = "op de"
End If
If ComboBox3.Text = "" Then
  TextBox12.Value = Clear
  TextBox12.Locked = True
Else
  TextBox12.Locked = False
  End If
End Sub
Private Sub Combobox4_Change()

If ComboBox4.Text = "Kantoor" Or ComboBox4.Text = "Toilet" Or _
ComboBox4.Text = "Zwembad" Or ComboBox4.Text = "Portaal" Or _
ComboBox4.Text = "Bureel" Then
 TextBox13.Value = "in het"
 Else
  TextBox13.Value = "in de"
End If
If ComboBox4.Text = "Kantoor1" Or ComboBox4.Text = "Kantoor2" Or _
ComboBox4.Text = "Verblijfsruimte1" Or ComboBox4.Text = "Verblijfsruimte2" Or _
ComboBox4.Text = "Groepsruimte1" Or ComboBox4.Text = "Groepsruimte2" Or _
ComboBox4.Text = "Slaapkamer1" Or ComboBox4.Text = "Slaapkamer2" Or _
ComboBox4.Text = "Slaapkamer3" Or ComboBox4.Text = "Slaapkamer4" Or _
ComboBox4.Text = "Bedrijfsruimte1" Or ComboBox4.Text = "Bedrijfsruimte2" Then
  TextBox13.Value = "in"
End If
If ComboBox4.Text = "Verkoop" Or ComboBox4.Text = "Zolder" Or ComboBox4.Text = "Overloop" Then
  TextBox13.Value = "op de"
End If
If ComboBox4.Text = "" Then
  TextBox13.Value = Clear
  TextBox13.Locked = True
Else
  TextBox13.Locked = False
  End If

End Sub


Private Sub Combobox5_Change()

If ComboBox5.Text = "Kantoor" Or ComboBox5.Text = "Toilet" Or _
ComboBox5.Text = "Zwembad" Or ComboBox5.Text = "Portaal" Or _
ComboBox5.Text = "Bureel" Then
 TextBox14.Value = "in het"
 Else
  TextBox14.Value = "in de"
End If
If ComboBox5.Text = "Kantoor1" Or ComboBox5.Text = "Kantoor2" Or _
ComboBox5.Text = "Verblijfsruimte1" Or ComboBox5.Text = "Verblijfsruimte2" Or _
ComboBox5.Text = "Groepsruimte1" Or ComboBox5.Text = "Groepsruimte2" Or _
ComboBox5.Text = "Slaapkamer1" Or ComboBox5.Text = "Slaapkamer2" Or _
ComboBox5.Text = "Slaapkamer3" Or ComboBox5.Text = "Slaapkamer4" Or _
ComboBox5.Text = "Bedrijfsruimte1" Or ComboBox5.Text = "Bedrijfsruimte2" Then
  TextBox14.Value = "in"
End If
If ComboBox5.Text = "Verkoop" Or ComboBox5.Text = "Zolder" Or ComboBox5.Text = "Overloop" Then
  TextBox14.Value = "op de"
End If
If ComboBox5.Text = "" Then
  TextBox14.Value = Clear
  TextBox14.Locked = True
Else
  TextBox14.Locked = False
  End If

End Sub
Private Sub Combobox6_Change()

If ComboBox6.Text = "Kantoor" Or ComboBox6.Text = "Toilet" Or _
ComboBox6.Text = "Zwembad" Or ComboBox6.Text = "Portaal" Or _
ComboBox6.Text = "Bureel" Then
 TextBox15.Value = "in het"
 Else
  TextBox15.Value = "in de"
End If
If ComboBox6.Text = "Kantoor1" Or ComboBox6.Text = "Kantoor2" Or _
ComboBox6.Text = "Verblijfsruimte1" Or ComboBox6.Text = "Verblijfsruimte2" Or _
ComboBox6.Text = "Groepsruimte1" Or ComboBox6.Text = "Groepsruimte2" Or _
ComboBox6.Text = "Slaapkamer1" Or ComboBox6.Text = "Slaapkamer2" Or _
ComboBox6.Text = "Slaapkamer3" Or ComboBox6.Text = "Slaapkamer4" Or _
ComboBox6.Text = "Bedrijfsruimte1" Or ComboBox6.Text = "Bedrijfsruimte2" Then
  TextBox15.Value = "in"
End If
If ComboBox6.Text = "Verkoop" Or ComboBox6.Text = "Zolder" Or ComboBox6.Text = "Overloop" Then
  TextBox15.Value = "op de"
End If
If ComboBox6.Text = "" Then
  TextBox15.Value = Clear
  TextBox15.Locked = True
Else
  TextBox15.Locked = False
  End If

End Sub
Private Sub Combobox7_Change()

If ComboBox7.Text = "Kantoor" Or ComboBox7.Text = "Toilet" Or _
ComboBox7.Text = "Zwembad" Or ComboBox7.Text = "Portaal" Or _
ComboBox7.Text = "Bureel" Then
 TextBox16.Value = "in het"
 Else
  TextBox16.Value = "in de"
End If
If ComboBox7.Text = "Kantoor1" Or ComboBox7.Text = "Kantoor2" Or _
ComboBox7.Text = "Verblijfsruimte1" Or ComboBox7.Text = "Verblijfsruimte2" Or _
ComboBox7.Text = "Groepsruimte1" Or ComboBox7.Text = "Groepsruimte2" Or _
ComboBox7.Text = "Slaapkamer1" Or ComboBox7.Text = "Slaapkamer2" Or _
ComboBox7.Text = "Slaapkamer3" Or ComboBox7.Text = "Slaapkamer4" Or _
ComboBox7.Text = "Bedrijfsruimte1" Or ComboBox7.Text = "Bedrijfsruimte2" Then
  TextBox16.Value = "in"
End If
If ComboBox7.Text = "Verkoop" Or ComboBox7.Text = "Zolder" Or ComboBox7.Text = "Overloop" Then
  TextBox16.Value = "op de"
End If
If ComboBox7.Text = "" Then
  TextBox16.Value = Clear
  TextBox16.Locked = True
Else
  TextBox16.Locked = False
  End If

End Sub
Private Sub Combobox8_Change()

If ComboBox8.Text = "Kantoor" Or ComboBox8.Text = "Toilet" Or _
ComboBox8.Text = "Zwembad" Or ComboBox8.Text = "Portaal" Or _
ComboBox8.Text = "Bureel" Then
 TextBox17.Value = "in het"
 Else
  TextBox17.Value = "in de"
End If
If ComboBox8.Text = "Kantoor1" Or ComboBox8.Text = "Kantoor2" Or _
ComboBox8.Text = "Verblijfsruimte1" Or ComboBox8.Text = "Verblijfsruimte2" Or _
ComboBox8.Text = "Groepsruimte1" Or ComboBox8.Text = "Groepsruimte2" Or _
ComboBox8.Text = "Slaapkamer1" Or ComboBox8.Text = "Slaapkamer2" Or _
ComboBox8.Text = "Slaapkamer3" Or ComboBox8.Text = "Slaapkamer4" Or _
ComboBox8.Text = "Bedrijfsruimte1" Or ComboBox8.Text = "Bedrijfsruimte2" Then
  TextBox17.Value = "in"
End If
If ComboBox8.Text = "Verkoop" Or ComboBox8.Text = "Zolder" Or ComboBox8.Text = "Overloop" Then
  TextBox17.Value = "op de"
End If
If ComboBox8.Text = "" Then
  TextBox17.Value = Clear
  TextBox17.Locked = True
Else
  TextBox17.Locked = False
  End If

End Sub
Private Sub Combobox9_Change()
If ComboBox9.Text = "Kantoor" Or ComboBox9.Text = "Toilet" Or _
ComboBox9.Text = "Zwembad" Or ComboBox9.Text = "Portaal" Or _
ComboBox9.Text = "Bureel" Then
 TextBox18.Value = "in het"
 Else
  TextBox18.Value = "in de"
End If
If ComboBox9.Text = "Kantoor1" Or ComboBox9.Text = "Kantoor2" Or _
ComboBox9.Text = "Verblijfsruimte1" Or ComboBox9.Text = "Verblijfsruimte2" Or _
ComboBox9.Text = "Groepsruimte1" Or ComboBox9.Text = "Groepsruimte2" Or _
ComboBox9.Text = "Slaapkamer1" Or ComboBox9.Text = "Slaapkamer2" Or _
ComboBox9.Text = "Slaapkamer3" Or ComboBox9.Text = "Slaapkamer4" Or _
ComboBox9.Text = "Bedrijfsruimte1" Or ComboBox9.Text = "Bedrijfsruimte2" Then
  TextBox18.Value = "in"
End If
If ComboBox9.Text = "Verkoop" Or ComboBox9.Text = "Zolder" Or ComboBox9.Text = "Overloop" Then
  TextBox18.Value = "op de"
End If
If ComboBox9.Text = "" Then
  TextBox18.Value = Clear
  TextBox18.Locked = True
Else
  TextBox18.Locked = False
  End If
End Sub
Sub Schaal(scaal)
blokken.Hide
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


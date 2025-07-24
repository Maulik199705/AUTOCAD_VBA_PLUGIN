VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitsen 
   Caption         =   "OPSPLITSEN VAN DE ROLLEN"
   ClientHeight    =   4860
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6072
   OleObjectBlob   =   "frmSplitsen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplitsen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Opsplitsen van rollengtes
'Michel Bosch en Gerard Haak
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

Dim newLayer As AcadLayer
Set newLayer = ThisDrawing.Layers.Add("OPSPLITSEN")
ThisDrawing.ActiveLayer = newLayer
Update
Call Combolijst1
Call Combolijst2
OptionButton1.Value = True
End Sub
Private Sub CancelButton1_Click()
frmSplitsen.Hide
Unload Me
End Sub
Private Sub CmdButton_Click()

 If ComboBox1 = ComboBox2 Then
 Label5 = "Groepsnummers mogen niet gelijk zijn"
 Label5.BackColor = &HFFFF&
 Exit Sub
 End If

teller = ListBox2.ListCount
CmdButton2.Enabled = True
If OptionButton1 = True Then d = " (WTH-ZD) "
If OptionButton2 = True Then d = " (PE-RT) "
c = "1 rol van " & ComboBox3 & " opsplitsen t.b.v. " & "Groep " & ComboBox1 & " en " & "Groep " & ComboBox2 & d
t = "1 rol van " & ComboBox3 & " opsplitsen t.b.v. " & "Groep " & ComboBox2 & " en " & "Groep " & ComboBox1 & d
E = "Groep " & ComboBox1 & " en " & "Groep " & ComboBox2
f = "Groep " & ComboBox2 & " en " & "Groep " & ComboBox1


For I = 0 To teller - 1
textstring = ListBox2.List(I)
If E = textstring Or f = textstring Then
 Label5 = "Deze combinatie heb je al gebruikt"
 Label5.BackColor = &HFFFF&
 Exit Sub
 End If
 Next I

ListBox1.AddItem c
ListBox2.AddItem E
ComboBox1.Text = ComboBox1.List(0)
ComboBox2.Text = ComboBox2.List(0)
Label5 = Clear
Label5.BackColor = &HC0C0C0
End Sub
Private Sub CmdButton1_Click()
frmSplitsen.Hide
End Sub
Private Sub CmdButton2_Click()
  On Error Resume Next
    Dim textobj As AcadText
    Dim textstring As String
    Dim pecht(0 To 2) As Double
    Dim pb0(0 To 2) As Double
    frmSplitsen.Hide
    ThisDrawing.SetVariable "osmode", 0
    pbzero = ThisDrawing.Utility.GetPoint(, "Plaats beginpunt") '---
    teller = ListBox1.ListCount
    omhoog = teller * 30
    
    Dim pb1(0 To 2) As Double
    
     pb1(0) = pbzero(0) - 910
     pb1(1) = pbzero(1) + (omhoog + 30)
     pb1(2) = pbzero(2)
    'pb1 = ThisDrawing.Utility.GetPoint(, "Plaats beginpunt")
    
  If Err Then
    frmSplitsen.Show
    Exit Sub
  End If
     
     pb0(0) = pb1(0)
     pb0(1) = pb1(1)
     pb0(2) = pb1(2)
   
    For I = 0 To teller - 1
    'Define the text object
    textstring = ListBox1.List(I)
     
     pecht(0) = pb1(0)             'Point(pBegin.pX, PY)
     pecht(1) = pb1(1) - 30
     pecht(2) = pb1(2)
       
     Set textobj = ThisDrawing.ModelSpace.AddText(textstring, pecht, 14.5)
     pb1(0) = pecht(0)             'Point(pBegin.pX, PY)
     pb1(1) = pecht(1)
     pb1(2) = pecht(2)

    Next I
    Update
    ' Create the text object in model space
    ZoomAll


 Dim LijnObj As AcadPolyline
' Dim Lijnobj As Object
 Dim pb2(0 To 2) As Double 'punt van rechthoek
 Dim pb3(0 To 2) As Double 'punt van rechthoek
 Dim pb4(0 To 2) As Double 'punt van rechthoek
 Dim pb5(0 To 2) As Double 'punt voor regelunittekst
 
 pb2(0) = pb0(0) - 15
 pb2(1) = pb0(1) + 15
 pb2(2) = 0
 
 pb3(0) = pb2(0) + 925
 pb3(1) = pb2(1)
 pb3(2) = 0
 
 pb4(0) = pb3(0)
 pb4(1) = pecht(1) - 30
 pb4(2) = 0
 
 pb5(0) = pb2(0)
 pb5(1) = pb4(1)
 pb5(2) = 0
   
Dim points(0 To 14) As Double
points(0) = pb2(0): points(1) = pb2(1): points(2) = pb2(2)
points(3) = pb3(0): points(4) = pb3(1): points(5) = pb3(2)
points(6) = pb4(0): points(7) = pb4(1): points(8) = pb4(2)
points(9) = pb5(0): points(10) = pb5(1): points(11) = pb5(2)
points(12) = pb2(0): points(13) = pb2(1): points(14) = pb2(2)
 
 Set LijnObj = ThisDrawing.ModelSpace.AddPolyline(points)
 LijnObj.Offset (3)
 LijnObj.Update
End Sub

Private Sub CmdButton3_Click()
ListBox1.Clear
ListBox2.Clear
CmdButton2.Enabled = False
ComboBox1.Text = ComboBox1.List(0)
ComboBox2.Text = ComboBox2.List(0)
Label5 = Clear
Label5.BackColor = &HC0C0C0
End Sub
Private Sub CmdButton4_Click()
    'Ensure ListBox contains list items
    ListBox2.ListIndex = ListBox1.ListIndex
    If ListBox1.ListCount >= 1 And ListBox2.ListCount >= 1 Then
        'If no selection, choose last list item.
        If ListBox1.ListIndex = -1 And ListBox2.ListCount >= -1 Then
            ListBox1.ListIndex = _
                     ListBox1.ListCount - 1
            ListBox2.ListIndex = _
                     ListBox2.ListCount - 1
        End If
        ListBox1.RemoveItem (ListBox1.ListIndex)
        ListBox2.RemoveItem (ListBox2.ListIndex)
    End If
End Sub
Private Sub ComboBox1_DropButtonClick()
 ComboBox1.SetFocus
 Label5 = Clear
 Label5.BackColor = &HC0C0C0
 End Sub
Private Sub ComboBox1_Enter()
 ComboBox1.SetFocus
 Label5 = Clear
 Label5.BackColor = &HC0C0C0
End Sub
Private Sub ComboBox2_DropButtonClick()
 ComboBox2.SetFocus
 Label5 = Clear
 Label5.BackColor = &HC0C0C0
End Sub

Private Sub ComboBox2_Enter()
 ComboBox2.SetFocus
 Label5 = Clear
 Label5.BackColor = &HC0C0C0
End Sub

Private Sub ComboBox3_Change()
  CmdButton.SetFocus
End Sub
Private Sub OptionButton1_Click()
Call Combolijst3
End Sub

Private Sub OptionButton2_Click()
Call Combolijst4
End Sub

Sub Combolijst1()
ComboBox1.AddItem "01.01"
ComboBox1.AddItem "02.01"
ComboBox1.AddItem "03.01"
ComboBox1.AddItem "04.01"
ComboBox1.AddItem "05.01"
ComboBox1.AddItem "06.01"
ComboBox1.AddItem "07.01"
ComboBox1.AddItem "08.01"
ComboBox1.AddItem "09.01"
ComboBox1.AddItem "10.01"
ComboBox1.Text = ComboBox1.List(0)
End Sub
Sub Combolijst2()
ComboBox2.AddItem "01.01"
ComboBox2.AddItem "02.01"
ComboBox2.AddItem "03.01"
ComboBox2.AddItem "04.01"
ComboBox2.AddItem "05.01"
ComboBox2.AddItem "06.01"
ComboBox2.AddItem "07.01"
ComboBox2.AddItem "08.01"
ComboBox2.AddItem "09.01"
ComboBox2.AddItem "10.01"
ComboBox2.Text = ComboBox2.List(0)
End Sub
Sub Combolijst3()
ComboBox3.Clear
ComboBox3.AddItem "125 meter"
ComboBox3.AddItem "105 meter"
ComboBox3.AddItem "90 meter"
ComboBox3.AddItem "75 meter"
ComboBox3.AddItem "63 meter"
ComboBox3.AddItem "50 meter"
ComboBox3.AddItem "40 meter"
ComboBox3.Text = ComboBox3.List(0)
End Sub
Sub Combolijst4()
ComboBox3.Clear
ComboBox3.AddItem "120 meter"
ComboBox3.AddItem "90 meter"
ComboBox3.Text = ComboBox3.List(0)
End Sub


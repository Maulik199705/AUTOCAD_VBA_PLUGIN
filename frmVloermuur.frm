VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVloermuur 
   Caption         =   "VLOERSPARING OF MUURSPARING"
   ClientHeight    =   2550
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6120
   OleObjectBlob   =   "frmVloermuur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVloermuur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'10-09-2002 Plaatsen van vloer of muursparing
'M.Bosch en G.C.Haak
'blmuur.dwg en bl-vloerspar


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
Private Sub OKButton1_Click()
On Error Resume Next
Dim newLayer As AcadLayer
Set newLayer = ThisDrawing.Layers.Add("3")
ThisDrawing.ActiveLayer = newLayer
Update
ThisDrawing.SetVariable "osmode", 0

Dim teller As Integer
Dim element As Object
frmVloermuur.Hide
If OptionButton1.Value = True Then
Call Schaal(scaal)
pbegin = ThisDrawing.Utility.GetPoint(, "Plaats punt -Blok Muursparing-")
Dim bestand As String
bestand = "C:\ACAD2002\DWG\blmuur.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin, bestand, scaal, scaal, 1, 0)
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "BLMUUR" Or "blmuur" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If TextBox1 = 1 Then w1 = "Muursparing 10 * 5 cm"
        If TextBox1 = 2 Then w1 = "Muursparing 15 * 5 cm"
        If TextBox1 = 3 Then w1 = "Muursparing 25 * 5 cm"
        If TextBox1 = 4 Then w1 = "Muursparing 30 * 5 cm"
        If TextBox1 = 5 Then w1 = "Muursparing 35 * 5 cm"
        If TextBox1 = 6 Then w1 = "Muursparing 45 * 5 cm"
        If TextBox1 = 7 Then w1 = "Muursparing 50 * 5 cm"
        If TextBox1 = 8 Then w1 = "Muursparing 60 * 5 cm"
        If TextBox1 = 9 Then w1 = "Muursparing 65 * 5 cm"
        If TextBox1 = 10 Then w1 = "Muursparing 75 * 5 cm"
        If TextBox1 = 11 Then w1 = "Muursparing 80 * 5 cm"
        If TextBox1 = 12 Then w1 = "Muursparing 85 * 5 cm"
        If TextBox1 = 13 Then w1 = "Muursparing 95 * 5 cm"
        If TextBox1 = 14 Then w1 = "Muursparing 100 * 5 cm"
        If TextBox1 = 15 Then w1 = "Muursparing 105 * 5 cm"
        If TextBox1 = 16 Then w1 = "Muursparing 115 * 5 cm"
        If TextBox1 = 17 Then w1 = "Muursparing 120 * 5 cm"
        If TextBox1 = 18 Then w1 = "Muursparing 130 * 5 cm"
        If TextBox1 = 19 Then w1 = "Muursparing 135 * 5 cm"
        If TextBox1 = 20 Then w1 = "Muursparing 145 * 5 cm"
        If attribuut.TagString = "MUURSPARING" And attribuut.textstring = "" Then attribuut.textstring = w1
        Next I
        End If
      End If
      End If
  Next element
  End If
Update

If OptionButton2.Value = True Then
Call Schaal(scaal)
pbegin = ThisDrawing.Utility.GetPoint(, "Plaats punt -Blok Vloersparing-")
Dim bestand2 As String
bestand2 = "C:\ACAD2002\DWG\bl-vloerspar.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pbegin, bestand2, scaal, scaal, 1, 0)
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "SPARING" Or "sparing" Or "ETAGE" Or "etage" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If TextBox2 = 1 Then w2 = "Vloersparing 10 * 10 cm"
        If TextBox2 = 2 Then w2 = "Vloersparing 15 * 10 cm"
        If TextBox2 = 3 Then w2 = "Vloersparing 25 * 10 cm"
        If TextBox2 = 4 Then w2 = "Vloersparing 30 * 10 cm"
        If TextBox2 = 5 Then w2 = "Vloersparing 35 * 10 cm"
        If TextBox2 = 6 Then w2 = "Vloersparing 45 * 10 cm"
        If TextBox2 = 7 Then w2 = "Vloersparing 50 * 10 cm"
        If TextBox2 = 8 Then w2 = "Vloersparing 60 * 10 cm"
        If TextBox2 = 9 Then w2 = "Vloersparing 65 * 10 cm"
        If TextBox2 = 10 Then w2 = "Vloersparing 75 * 10 cm"
        If TextBox2 = 11 Then w2 = "Vloersparing 80 * 10 cm"
        If TextBox2 = 12 Then w2 = "Vloersparing 85 * 10 cm"
        If TextBox2 = 13 Then w2 = "Vloersparing 95 * 10 cm"
        If TextBox2 = 14 Then w2 = "Vloersparing 100 * 10 cm"
        If TextBox2 = 15 Then w2 = "Vloersparing 105 * 10 cm"
        If TextBox2 = 16 Then w2 = "Vloersparing 115 * 10 cm"
        If TextBox2 = 17 Then w2 = "Vloersparing 120 * 10 cm"
        If TextBox2 = 18 Then w2 = "Vloersparing 130 * 10 cm"
        If TextBox2 = 19 Then w2 = "Vloersparing 135 * 10 cm"
        If TextBox2 = 20 Then w2 = "Vloersparing 145 * 10 cm"
        If attribuut.TagString = "ETAGE" And attribuut.textstring = "" Then attribuut.textstring = ComboBox1
        If attribuut.TagString = "SPARING" And attribuut.textstring = "" Then attribuut.textstring = w2
        Next I
        End If
      End If
      End If
  Next element
  End If
Update

TextBox1.Value = Clear
TextBox2.Value = Clear
ComboBox1.Text = ComboBox1.List(2)
If OptionButton1.Value = True Then TextBox1.SetFocus
If OptionButton2.Value = True Then TextBox2.SetFocus

If Err Then
    If OptionButton1.Value = True Then TextBox1.SetFocus
    If OptionButton2.Value = True Then TextBox2.SetFocus
    frmVloermuur.Show
    Exit Sub
    End If
Update
End Sub

Private Sub OptionButton1_Click()
If OptionButton1.Value = True Then
  TextBox1.Visible = True
  TextBox1.SetFocus
  OptionButton2.Value = False
  TextBox2.Visible = False
  ComboBox1.Visible = False
  Label1.Visible = False
  'AcPreview1.Preview = "C:\ACAD2002\DWG\blmuur" & ".dwg"
  End If
  
End Sub
Private Sub OptionButton2_Click()
If OptionButton2.Value = True Then
  OptionButton1.Value = False
  TextBox2.Visible = True
  TextBox1.Visible = False
  TextBox2.SetFocus
  ComboBox1.Visible = True
  Label1.Visible = True
  Call combolijst
  End If
End Sub
Private Sub ResetButton1_Click()
Call RESET
End Sub
  Sub RESET()
   OptionButton1.Value = False: OptionButton2.Value = False
   ComboBox1.Visible = False
   OKButton1.Enabled = False
   Label1.Visible = False
   TextBox1.Visible = False: TextBox2.Visible = False
  End Sub
Private Sub TextBox1_Change()
Dim a As Double
On Error Resume Next
If TextBox1.Value <> "" Then OKButton1.Enabled = True Else OKButton1.Enabled = False
a = TextBox1.Text

If Err Then
   TextBox1.Text = Clear
   Exit Sub
  End If
  
If a > 20 Then
   MsgBox "Groter dan 20 groepen is niet toegestaan..!!!!", vbCritical
   TextBox1 = Clear
End If
  
If a < 1 Then
   MsgBox "Kleiner dan een 1 groeps muurparing is niet mogelijk..!!!!", vbCritical
   TextBox1 = Clear
End If
  
End Sub

Private Sub TextBox2_Change()
Dim b As Double
On Error Resume Next
If TextBox2.Value <> "" Then OKButton1.Enabled = True Else OKButton1.Enabled = False
b = TextBox2.Text
If Err Then
   TextBox2.Text = Clear
   Exit Sub
  End If
  
If b > 20 Then
   MsgBox "Groter dan 20 groepen is niet toegestaan..!!!!", vbCritical
   TextBox2 = Clear
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
frmVloermuur.TextBox1.SetFocus
  End Sub
   Sub combolijst()
   ComboBox1.AddItem "in de kruipruimte."
   ComboBox1.AddItem "in de leidingkoker."
   ComboBox1.AddItem "op de begane grond."
   ComboBox1.AddItem "op de 1e Verdieping."
   ComboBox1.AddItem "op de 2e Verdieping."
   ComboBox1.AddItem "op de 3e Verdieping."
   ComboBox1.AddItem "op de 4e Verdieping."
   ComboBox1.AddItem "op de zolder."
   ComboBox1.AddItem "in de kelder."
   'ComboBox1.Text = ComboBox1.List(2)
   ComboBox1.MatchEntry = fmMatchEntryFirstLetter
  End Sub
Sub Schaal(scaal)
frmVloermuur.Hide
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

Private Sub CancelButton1_Click()
frmVloermuur.Hide
Unload Me
' ThisDrawing.SetVariable "osmode" = Z
End Sub

Private Sub CommandButton1_Click()
Unload Me
ThisDrawing.SendCommand "arrow" & vbCr
End Sub

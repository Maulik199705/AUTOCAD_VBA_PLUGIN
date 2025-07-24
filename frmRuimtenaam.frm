VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRuimtenaam 
   Caption         =   "Ruimtenaam plaatsen"
   ClientHeight    =   4095
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5568
   OleObjectBlob   =   "frmRuimtenaam.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRuimtenaam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'05-11-2004 RUIMTENAAM PLAATSEN
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
Call uitlezenTXT
End Sub

Private Sub CommandButton1_Click()
waarde2 = 0
If TextBox2 = "" Then
samenvoegbox = TextBox1 & Space(1)
samenvoegbox = Split(samenvoegbox, " ")
samenvoegbox = samenvoegbox(0)
Else
samenvoegbox = TextBox1 & Space(1)
samenvoegbox = Split(samenvoegbox, " ")
samenvoegbox = samenvoegbox(0) & Space(1) & TextBox2
End If

For I = 0 To ListBox1.ListCount - 1
waarde1 = ListBox1.List(I)
If samenvoegbox = waarde1 Then
MsgBox "Staat al in de lijst"
waarde2 = 1
End If
Next I

If waarde2 <> 1 Then
ListBox1.AddItem (samenvoegbox)
bestand10 = "g:\tekeningen\ruimtenaam.txt"
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim fs, a, vullistbox1
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(bestand10, ForAppending, -2)
    a.write samenvoegbox
    a.write Chr(13) + Chr(10)
    a.Close 'sluiten van tekstbestand
End If
TextBox1 = Clear
TextBox2 = Clear

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
ListBox1.SetFocus

  End Sub
Private Sub TextBox1_Change()
'Dim ab As String
'ab = TextBox1.Text
'ab =
TextBox1 = StrConv(TextBox1, vbProperCase)
If TextBox1 <> "" Then CommandButton1.Enabled = True
If TextBox1 = "" Then CommandButton1.Enabled = False
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ListBox1.MatchEntry = fmMatchEntryFirstLetter
Call plaatstekst
End Sub
Private Sub txtplaats_Click()
Call plaatstekst
End Sub
Sub plaatstekst()
On Error Resume Next
Call Schaal(scaal)
Dim returnpnt As Variant
Dim textobj As AcadText
Dim layerObj As AcadLayer
Set laagoud = ThisDrawing.ActiveLayer
Set layerObj = ThisDrawing.Layers.Add("RUIMTENAAM")
ThisDrawing.ActiveLayer = layerObj
If TextBox1 <> "" Then
waarde1 = TextBox1 'de waarde invullen van de textbox
Else
waarde1 = ListBox1.Value 'de waarde invullen van listbox1
End If
frmRuimtenaam.Hide 'het formulier verbergen
returnpnt = ThisDrawing.Utility.GetPoint(, "Geef begin punt op : ") 'beginpun opgeven

If Err Then
    frmRuimtenaam.Show
    Exit Sub
    End If
textobj.Alignment = returnpnt.acAlignmentCenter
Set textobj = ThisDrawing.ModelSpace.AddText(waarde1, returnpnt, scaal) 'text plaatsen

ThisDrawing.ActiveLayer = laagoud
Update
TextBox1 = ""
frmRuimtenaam.Show
End Sub
Sub uitlezenTXT()
bestand10 = "g:\tekeningen\ruimtenaam.txt"
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim fs, a, vullistbox1
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(bestand10, ForReading, False)


Do While a.AtEndOfLine <> True
    vullistbox1 = a.ReadLine
    ListBox1.AddItem (vullistbox1)
Loop
a.Close 'sluiten van tekstbestand

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
End Sub
Sub Schaal(scaal)
frmRuimtenaam.Hide
On Error Resume Next
vartab = ThisDrawing.GetVariable("EXTMAX")
If vartab(0) >= 2145 And vartab(0) <= 2155 Or vartab(0) >= 4095 And vartab(0) <= 4105 _
    Or vartab(0) >= 5835 And vartab(0) <= 5845 Or vartab(0) >= 8355 And vartab(0) <= 8365 _
    Or vartab(0) >= 11785 And vartab(0) <= 11795 Or vartab(0) >= 15685 And vartab(0) <= 15695 Then
    scaal = 30
Else
    scaal = 15
End If
If vartab(0) >= 16715 And vartab(0) <= 16725 Or vartab(0) >= 23575 And vartab(0) <= 23585 _
    Or vartab(0) >= 31375 And vartab(0) <= 31385 Then
    scaal = 60
End If
Update
End Sub
Private Sub cmdAfsluiten_Click()
Unload Me
End Sub

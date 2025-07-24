VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmbouwtek 
   Caption         =   "Bouwkundige tekening....."
   ClientHeight    =   2595
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   4668
   OleObjectBlob   =   "frmbouwtek.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmbouwtek"
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
frmbouwtek.TextBox1.SetFocus

End Sub
Private Sub CancelButton_Click()

If (frmbouwtek.TextBox1 <> "" And frmbouwtek.TextBox1.TextLength < 2) Or frmbouwtek.TextBox1 > "31" Then
    MsgBox "Je moet minimaal 2 getallen invoeren.. (Aantal dagen van de maand)" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) & _
           "Dit moet een getal zijn die ligt tussen  0 en 32 " & (Chr(13) & Chr(10)) & _
           "Voor de getallen 1 tot en met 9 een nul vermelden.", vbExclamation
    frmbouwtek.TextBox1 = Clear
    frmbouwtek.TextBox1.SetFocus
    Exit Sub
End If

If (frmbouwtek.TextBox3 <> "" And frmbouwtek.TextBox3.TextLength < 2) Or frmbouwtek.TextBox3 > "12" Then
    MsgBox "Je moet minimaal 2 getallen invoeren.. (Aantal maanden van het jaar)" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) & _
           "Dit moet een getal zijn die ligt tussen  0 en 13" & (Chr(13) & Chr(10)) & _
           "Voor de getallen 1 tot en met 9 een nul vermelden.", vbExclamation
    frmbouwtek.TextBox3 = Clear
    frmbouwtek.TextBox3.SetFocus
    Exit Sub
End If

If frmbouwtek.TextBox4 <> "" And frmbouwtek.TextBox4.TextLength < 4 Then
    MsgBox "Je moet minimaal 4 getallen invoeren..(Jaartal)", vbExclamation
    frmbouwtek.TextBox4 = Clear
    frmbouwtek.TextBox4.SetFocus
    Exit Sub
End If



If frmbouwtek.TextBox1 <> "" And frmbouwtek.TextBox2 <> "" And frmbouwtek.TextBox3 <> "" And frmbouwtek.TextBox4 <> "" Then
   blokken.TextBox19 = frmbouwtek.TextBox1 & "-" & frmbouwtek.TextBox3 & "-" & frmbouwtek.TextBox4  'datum
   blokken.TextBox20 = frmbouwtek.TextBox2 'naam
End If

If TextBox1 = "" Or TextBox2 = "" Or TextBox3 = "" Or TextBox4 = "" Then blokken.CheckBox1 = False
Unload Me
blokken.Show

End Sub
Private Sub TextBox1_Change()
On Error Resume Next
Dim b As Double

b = frmbouwtek.TextBox1
If Err Then
   frmbouwtek.TextBox1.Value = Clear
   Exit Sub
End If

If frmbouwtek.TextBox1.TextLength = 2 Then frmbouwtek.TextBox3.SetFocus


End Sub
Private Sub TextBox3_Change()
On Error Resume Next
Dim b As Double
b = frmbouwtek.TextBox3

If Err Then
   frmbouwtek.TextBox3.Value = Clear
   Exit Sub
End If

If frmbouwtek.TextBox3.TextLength = 2 Then frmbouwtek.TextBox4.SetFocus


End Sub
Private Sub TextBox4_Change()
On Error Resume Next
Dim b As Double

b = frmbouwtek.TextBox4
If Err Then
   frmbouwtek.TextBox4.Value = Clear
   Exit Sub
End If

If frmbouwtek.TextBox4.TextLength = 4 Then frmbouwtek.TextBox2.SetFocus


End Sub

Private Sub CommandButton1_Click()
frmbouwtek.TextBox1 = Clear: frmbouwtek.TextBox2 = Clear: frmbouwtek.TextBox3 = Clear: frmbouwtek.TextBox4 = Clear
frmbouwtek.TextBox1.SetFocus
End Sub

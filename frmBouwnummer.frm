VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBouwnummer 
   Caption         =   "Bouwnummer~Type woning"
   ClientHeight    =   2208
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   5064
   OleObjectBlob   =   "frmBouwnummer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBouwnummer"
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
frmBouwnummer.TextBox3.SetFocus
End Sub
Private Sub CancelButton_Click()
If frmBouwnummer.TextBox3 <> "" Then blokken.TextBox21 = "BOUWNUMMER: " & frmBouwnummer.TextBox3 'bouwnummer
If frmBouwnummer.TextBox2 <> "" Then blokken.TextBox22 = "TYPE: " & frmBouwnummer.TextBox2 'type
If TextBox3 = "" And TextBox2 = "" Then blokken.CheckBox2 = False
Unload Me
blokken.Show
End Sub
Private Sub CommandButton1_Click()
frmBouwnummer.TextBox2 = Clear: frmBouwnummer.TextBox3 = Clear
frmBouwnummer.TextBox3.SetFocus
End Sub


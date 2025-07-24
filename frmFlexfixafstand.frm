VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFlexfixafstand 
   Caption         =   "Leidingafstand & radius hoek"
   ClientHeight    =   2520
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3672
   OleObjectBlob   =   "frmFlexfixafstand.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFlexfixafstand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'13-09-2005 Flexfixafstand
'M.Bosch en G.C.Haak

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As LongPtr) As Long

    Private Declare PtrSafe Function GetMenuItemCount Lib "user32" _
    (ByVal hMenu As LongPtr) As Long

    Private Declare PtrSafe Function GetSystemMenu Lib "user32" _
    (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr

    Private Declare PtrSafe Function RemoveMenu Lib "user32" _
    (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As Long) As Long

    Private Declare Function GetMenuItemCount Lib "user32" _
    (ByVal hMenu As Long) As Long

    Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hWnd As Long, ByVal bRevert As Long) As Long

    Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
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
        RemoveMenu lngMenu, lngCnt - 1, MF_REMOVE Or MF_BYPOSITION
        DrawMenuBar lngHwnd
    End If
End Sub

Private Sub cmdAfsluiten_Click()
    Unload Me
    If OptionButton1.Value = True Then ThisDrawing.SetVariable "filletrad", 10
    If OptionButton2.Value = True Then ThisDrawing.SetVariable "filletrad", 12.5
    If OptionButton3.Value = True Then ThisDrawing.SetVariable "filletrad", 15
    If OptionButton4.Value = True Then ThisDrawing.SetVariable "filletrad", 20
    If OptionButton5.Value = True Then ThisDrawing.SetVariable "filletrad", 25
    If OptionButton6.Value = True Then ThisDrawing.SetVariable "filletrad", 30
End Sub



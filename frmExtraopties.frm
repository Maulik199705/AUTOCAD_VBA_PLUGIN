VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExtraopties 
   Caption         =   "Extra Opties"
   ClientHeight    =   1875
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   5352
   OleObjectBlob   =   "frmExtraopties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExtraopties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'23-2-2005 Extra Opties
'4-7-2005 wijziging
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


frmExtraopties.TextBox2.Value = 1
frmExtraopties.TextBox3.SetFocus
End Sub
Private Sub CommandButton1_Click()
On Error Resume Next
z1 = (Val(frmExtraopties.TextBox2))
z2 = (Val(frmExtraopties.TextBox3))

If z2 < z1 Then
  MsgBox "De laatste waarde moet groter zijn dan de 1e....!!!!", vbExclamation, "Let op"
  Exit Sub
End If

cc = frmExtraopties.TextBox1 & "_" & frmExtraopties.TextBox2
If (Val(frmExtraopties.TextBox2)) > 0 And (Val(frmExtraopties.TextBox2)) < 10 Then cc = frmExtraopties.TextBox1 & "_" & "0" & frmExtraopties.TextBox2

bb = (Val(frmExtraopties.TextBox2))

For I = (Val(frmExtraopties.TextBox2)) To (Val(frmExtraopties.TextBox3))
 
 C = frmExtraopties.TextBox1 & "_" & I

 If I > 0 And I < 10 Then C = frmExtraopties.TextBox1 & "_" & "0" & I
 
 Set newLayer = ThisDrawing.Layers.Add(C)
 ThisDrawing.ActiveLayer = newLayer


        rr = ThisDrawing.ActiveLayer.TrueColor.ColorIndex  '
        'MsgBox C & " -" & rr
        
        If rr = 2 Or rr = 7 Then
        
        Dim col As New AcadAcCmColor
        col.ColorMethod = AutoCAD.acColorMethodForeground

        Dim layColor As New AcadAcCmColor
        'Set layColor = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.16")
        Call layColor.SetRGB(255, 255, 0)

        ThisDrawing.ActiveLayer.TrueColor = layColor
        End If
        
        
 Next I

frmExtraopties.TextBox2 = bb
Set newLayer = ThisDrawing.Layers.Add(cc)
ThisDrawing.ActiveLayer = newLayer
Unload Me
End Sub
Private Sub TextBox2_Change()
Dim b As Double
On Error Resume Next
b = frmExtraopties.TextBox2.Text

If Err Then
   frmExtraopties.TextBox2 = Clear
   PlaatsButton.Enabled = False
  Exit Sub
  End If
  
If b > -1 And b < 1 Then
    frmExtraopties.TextBox2 = Clear
    frmExtraopties.TextBox2.SetFocus
End If
End Sub

Private Sub TextBox3_Change()
Dim a As Double
On Error Resume Next
a = frmExtraopties.TextBox3.Text

If Err Then
   frmExtraopties.TextBox3 = Clear
   PlaatsButton.Enabled = False
  Exit Sub
  End If
  
If a > -1 And a < 1 Then
    frmExtraopties.TextBox3 = Clear
    frmExtraopties.TextBox3.SetFocus
End If
End Sub
Private Sub cmdAfsluiten_Click()
Unload Me
End Sub

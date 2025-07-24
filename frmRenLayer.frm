VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRenLayer 
   Caption         =   "Layer(s) hernummeren "
   ClientHeight    =   2790
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3588
   OleObjectBlob   =   "frmRenLayer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRenLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mappAcad As AcadApplication
Private mstrPath As String

'09-09-2005 Legrichting bepalen
'M.Bosch en G.C.Haak

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As LongPtr) As Long
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
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
Dim box1 As Variant

lngHwnd = FindWindow(vbNullString, Me.Caption)
lngMenu = GetSystemMenu(lngHwnd, 0)

If lngMenu Then
    lngCnt = GetMenuItemCount(lngMenu)
    RemoveMenu lngMenu, lngCnt - 1, MF_REMOVE Or MF_BYPOSITION
    DrawMenuBar lngHwnd
End If

box1 = ThisDrawing.GetVariable("angdir")

If box1 = 0 Then OptionButton1.Value = True
If box1 = 1 Then OptionButton2.Value = True
End Sub

Private Sub cmdAfsluiten_Click()
If OptionButton1.Value = True Then ThisDrawing.SetVariable "angdir", 0
If OptionButton2.Value = True Then ThisDrawing.SetVariable "angdir", 1
Unload Me
End Sub




Private Sub CommandButton1_Click()
On Error Resume Next
ListBox1.Clear
Dim laagobj As Object
Dim mystr
Dim MYSTR2
Dim MYSTR3
Dim MYSTR4
Dim UNITTEL
Dim UNITTEL2
Dim UNITTEL3
Dim UNITTEL5


UNITTEL = frmRenLayer.TextBox1
If UNITTEL > 0 And UNITTEL < 10 Then
   frmRenLayer.TextBox1 = "0" & frmRenLayer.TextBox1
End If
UNITTEL2 = frmRenLayer.TextBox2
UNITTEL3 = frmRenLayer.TextBox2
If UNITTEL2 > 0 And UNITTEL2 < 10 Then
   UNITTEL3 = "0" & frmRenLayer.TextBox2
   frmRenLayer.TextBox2 = "0" & frmRenLayer.TextBox2
End If



For Each laagobj In ThisDrawing.Layers
mystr = Split(laagobj.Name, " ")
'MYSTR(0) = GROEP
'MYSTR(1) = 01.01

If UCase(mystr(0)) = "GROEP" Then

    MYSTR3 = Split(mystr(1), ".")
    'MYSTR3(0) = 01
    'MYSTR3(1) = 01
    

    If MYSTR3(0) = Val(TextBox1) Then
    MYSTR4 = "groep " & UNITTEL3 & "." & MYSTR3(1)
    ListBox1.AddItem (MYSTR4)
    laagobj.Name = MYSTR4
    
    End If
End If
Next laagobj

Dim element2 As Object
 For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "GROEPTEKSTBLOKNEW" Then
      Set SYMBOOL = element2
       If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
        If UCase(ATTRIBUUT.TagString) = "UNITNUMMER" Then
               rr = ATTRIBUUT.TextString
               If rr = frmRenLayer.TextBox1 Then
                      For Z = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                      Set ATTRIBUUT = ATTRIBUTEN(Z)
                        If ATTRIBUUT.TagString = "GROEPTEKST" Then
                        TT = ATTRIBUUT.TextString
                        hh = Split(TT, ".")
                        ee = "groep " & frmRenLayer.TextBox2 & "." & hh(1)
                        ATTRIBUUT.TextString = ee
                               For Q = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                               Set ATTRIBUUT = ATTRIBUTEN(Q)
                               If ATTRIBUUT.TagString = "UNITNUMMER" Then ATTRIBUUT.TextString = frmRenLayer.TextBox2
                               Next Q
                        End If
                      Next Z
               End If
         End If
     Next i
End If
End If
End If
Next element2
'Unload Me

Dim element22 As Object
 For Each element22 In ThisDrawing.ModelSpace
      If element22.ObjectName = "AcDbBlockReference" Then
      If UCase(element22.Name) = "MAT_SPE_ZD" Or UCase(element22.Name) = "MAT_SPE_PE" Or UCase(element22.Name) = "MAT_SPE_PE800" _
      Or UCase(element22.Name) = "MAT_SPE_ALU" Or UCase(element22.Name) = "MAT_SPE_ZDRINGLEIDING" _
      Or UCase(element22.Name) = "MAT_SPE_PERINGLEIDING" Or UCase(element22.Name) = "MAT_SPE_ALURINGLEIDING" _
      Or UCase(element22.Name) = "MAT_SPE_FLEX" Or UCase(element22.Name) = "MAT_SPE_FLEX_AANKOPPEL" Then
      Set SYMBOOL = element22
       If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For W = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
             Set ATTRIBUUT = ATTRIBUTEN(W)
               If UCase(ATTRIBUUT.TagString) = "RNU" Then
               rrs = ATTRIBUUT.TextString
               If rrs = frmRenLayer.TextBox1 Then ATTRIBUUT.TextString = frmRenLayer.TextBox2
               End If
        Next W
      End If
     End If
 End If
Next element22


For Each element222 In ThisDrawing.ModelSpace
      If element222.ObjectName = "AcDbBlockReference" Then
        If element222.Name = "HERZ" Or element222.Name = "RUH-R" Or element222.Name = "RUH-RT" _
        Or element222.Name = "RUB-R" Or element222.Name = "RUB-RT" Or element222.Name = "RUBK-R" Or element222.Name = "LT-VK" _
        Or element222.Name = "RUBK-RT" Or element222.Name = "LT" Or element222.Name = "LTS" Or element222.Name = "LT-N" Or element222.Name = "LTS-N" _
        Or element222.Name = "RUW" Or element222.Name = "RUV" Or element222.Name = "RUH-S" Or element222.Name = "VSKO-B" _
        Or element222.Name = "RUB-S" Or element222.Name = "KMV" Or element222.Name = "RUH-N" Or element222.Name = "RU-WW" Or element222.Name = "RINGLEIDING" _
        Or element222.Name = "RU-WWN" Or element222.Name = "RU-WWS" Or element222.Name = "RU-WKN" Or element222.Name = "RU-WKS" Then
      Set SYMBOOL = element222
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For u = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(u)
           If ATTRIBUUT.TagString = "UNITNUMMER" Then
               rrst = ATTRIBUUT.TextString
               If rrst = frmRenLayer.TextBox1 Then ATTRIBUUT.TextString = frmRenLayer.TextBox2
           End If
       
       Next u
      End If
     End If
     End If
   Next element222
Update

hern1 = "Regelunit " & TextBox1
hern2 = "Regelunit " & TextBox2
ThisDrawing.SendCommand "replacetext" & vbCr & hern1 & vbCr & hern2 & vbCr
frmRenLayer.Hide
Update
frmRenLayer.show

TextBox1 = Clear: TextBox2 = Clear
'''Dim Obj As AcadText
'''   For Each Obj In ThisDrawing.ModelSpace
'''        dd = "Regelunit " & frmRenLayer.TextBox1
'''        If Obj.TextString = dd Then Obj.TextString = "Regelunit " & frmRenLayer.TextBox2
'''   Next
End Sub
Private Sub TextBox1_Change()
If TextBox1.Text <> "" And TextBox2.Text <> "" Then
   CommandButton1.Enabled = True
   Else
   CommandButton1.Enabled = False
End If
End Sub

Private Sub TextBox2_Change()
If TextBox1.Text <> "" And TextBox2.Text <> "" Then
   CommandButton1.Enabled = True
   Else
   CommandButton1.Enabled = False
End If
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub


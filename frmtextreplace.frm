VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmtextreplace 
   Caption         =   "Vervang tekst in de tekening"
   ClientHeight    =   4170
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4848
   OleObjectBlob   =   "frmtextreplace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmtextreplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VBA Macro: Prepare DWG Text Replacement Tool (Updated for 64-bit)
' Fully compatible with AutoCAD 2020+ and modern systems
' Mimics exact behavior of original VBA logic
' Requires: Microsoft Office XX.0 Object Library enabled (Tools > References)

Option Explicit

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

    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Dim mappAcad As AcadApplication
Dim mstrPath As String

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
        Call RemoveMenu(lngMenu, lngCnt - 1, MF_REMOVE Or MF_BYPOSITION)
        Call DrawMenuBar(lngHwnd)
    End If
    Set mappAcad = GetObject(, "AutoCAD.Application")
End Sub

Private Sub TextBox1_Change()
    cmdStart.Enabled = (TextBox1 <> "" And TextBox2 <> "")
End Sub

Private Sub TextBox2_Change()
    cmdStart.Enabled = (TextBox1 <> "" And TextBox2 <> "")
End Sub

Private Sub CommandButton1_Click()
    lstDrawings.Clear
    ListBox1.Clear
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Set mappAcad = Nothing
End Sub

Private Sub cmdSelect_Click()
    Dim fd As FileDialog
    Dim selectedFile As Variant
    Dim strFileName As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select DWG Files"
        .Filters.Clear
        .Filters.Add "AutoCAD Drawings", "*.dwg"
        .AllowMultiSelect = True
        .InitialFileName = "F:\\Fserver2\\Gegevens\\Projecten\\"
        If .show = -1 Then
            lstDrawings.Clear
            ListBox1.Clear
            For Each selectedFile In .SelectedItems
                mstrPath = Left(selectedFile, InStrRev(selectedFile, "\\"))
                strFileName = Mid(selectedFile, InStrRev(selectedFile, "\\") + 1)
                lstDrawings.AddItem strFileName
                ListBox1.AddItem mstrPath
            Next selectedFile
        End If
    End With
End Sub

Private Sub cmdStart_Click()
    Dim i As Integer
    If lstDrawings.ListCount = 0 Then Exit Sub
    For i = 0 To lstDrawings.ListCount - 1
        Call PrepareAsBackGround(ListBox1.List(i) & lstDrawings.List(i))
    Next i
    Unload Me
End Sub

Sub PrepareAsBackGround(strDrawingFullname As String)
    Dim objDocument As AcadDocument
    Dim ori As String, verv As String

    If FileExists(strDrawingFullname) Then
        Set objDocument = mappAcad.Documents.Open(strDrawingFullname)
        ori = TextBox1
        verv = TextBox2
        ThisDrawing.SendCommand "srxtext" & vbCr & "R" & vbCr & ori & vbCr & verv & vbCr & "A" & vbCr & "A" & vbCr & "R" & vbCr
        objDocument.Close SaveChanges:=True
        Set objDocument = Nothing
    End If
End Sub

Private Function FileExists(strFullFilename As String) As Boolean
    If Len(strFullFilename) <> 0 Then
        FileExists = (Len(Dir(strFullFilename, vbNormal)) <> 0)
    End If
End Function



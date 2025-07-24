VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChange2BackGround 
   Caption         =   "Alles naar 1 layer"
   ClientHeight    =   3690
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3984
   OleObjectBlob   =   "frmChange2BackGround.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChange2BackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VBA Macro: Prepare Selected Drawings as Background
' Fully compatible with 64-bit AutoCAD + modern systems
' Updated to behave exactly like the legacy VBA form version
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

Dim mstrPath As String
Dim mappAcad As AcadApplication

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

    With Me.cboColors
        .AddItem "242"
        .AddItem "250"
        .AddItem "251"
        .AddItem "252"
        .AddItem "253"
        .AddItem "254"
        .AddItem "255"
        .ListIndex = 5
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Set mappAcad = Nothing
End Sub

Private Sub cmdSelect_Click()
    Dim sh As Object
    Dim folder As Object
    Dim item As Object
    Dim path As String
    Dim strFileName As String

    Set sh = CreateObject("Shell.Application")
    Set folder = sh.BrowseForFolder(0, "Select Folder with DWG files", 0, 0)

    If Not folder Is Nothing Then
        path = folder.Self.path
        If Right(path, 1) <> "\" Then path = path & "\"
        mstrPath = path
        lstDrawings.Clear

        strFileName = Dir(path & "*.dwg")
        Do While strFileName <> ""
            lstDrawings.AddItem strFileName
            strFileName = Dir
        Loop
    End If
End Sub


Private Sub cmdStart_Click()
    Dim intDrawing As Integer
    If lstDrawings.ListCount = 0 Then
        MsgBox "No drawings selected.", vbExclamation
        Exit Sub
    End If
    For intDrawing = 0 To lstDrawings.ListCount - 1
        PrepareAsBackGround mstrPath & lstDrawings.List(intDrawing), cboColors.Value
    Next intDrawing
    Unload Me
    MsgBox "All drawings processed successfully.", vbInformation
End Sub

Private Sub CommandButton1_Click()
    Dim ACADPref As AcadPreferencesOpenSave
    Dim originalValue As Variant, DisplayValue As String

    Set ACADPref = ThisDrawing.Application.Preferences.OpenSave
    originalValue = ACADPref.SaveAsType

    GoSub GETVALUE
    MsgBox "Current SaveAsType: " & DisplayValue

    ACADPref.SaveAsType = ac2000_dwg
    GoSub GETVALUE
    MsgBox "Changed SaveAsType to: " & DisplayValue

    ACADPref.SaveAsType = originalValue
    GoSub GETVALUE
    MsgBox "Reset SaveAsType back to: " & DisplayValue

    Exit Sub

GETVALUE:
    DisplayValue = ACADPref.SaveAsType
    Select Case DisplayValue
        Case ac2000_dwg:  DisplayValue = "AutoCAD 2000 DWG (*.dwg)"
        Case ac2000_dxf:  DisplayValue = "AutoCAD 2000 DXF (*.dxf)"
        Case ac2000_Template: DisplayValue = "AutoCAD 2000 DWT"
        Case ac2004_dwg:  DisplayValue = "AutoCAD 2004 DWG (*.dwg)"
        Case ac2004_dxf:  DisplayValue = "AutoCAD 2004 DXF (*.dxf)"
        Case ac2004_Template: DisplayValue = "AutoCAD 2004 DWT"
        Case ac2007_dwg:  DisplayValue = "AutoCAD 2007 DWG (*.dwg)"
        Case acNative:    DisplayValue = "Latest Drawing Format"
        Case acUnknown:   DisplayValue = "Unknown Type"
    End Select
    Return
End Sub

Sub PrepareAsBackGround(strDrawingFullname As String, colorValue As String)
    Dim objDocument As AcadDocument
    Dim objBlock As AcadBlock
    Dim objEntity As AcadEntity
    Dim objLayer As AcadLayer
    Dim strSaveAs As String
    Dim lngColor As Long

    If Len(Dir(strDrawingFullname)) = 0 Then Exit Sub
    Set objDocument = mappAcad.Documents.Open(strDrawingFullname)

    With objDocument
        strSaveAs = .path & "\Calculatie_" & .Name
        .SaveAs strSaveAs, acNative
        .AuditInfo True
        .PurgeAll
        ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr

        For Each objBlock In .Blocks
            If objBlock.IsXRef Then objBlock.Bind True
        Next objBlock

        Set objLayer = .Layers.Add("Bouwkundig")
        objLayer.color = 254
        objLayer.Linetype = "Continuous"
        objLayer.Lineweight = acLnWtByLwDefault
        objLayer.Lock = False

        lngColor = CLng(colorValue)
        For Each objLayer In .Layers
            objLayer.color = lngColor
            objLayer.Lineweight = acLnWtByLwDefault
            objLayer.Lock = False
        Next objLayer

        For Each objEntity In .ModelSpace
            objEntity.layer = "Bouwkundig"
            objEntity.color = acByLayer
            objEntity.Lineweight = acLnWtByLwDefault
        Next objEntity

        For Each objEntity In .ModelSpace
            Select Case objEntity.ObjectName
                Case "AcDbDimension", "AcDbHatch", "AcDbSolid", "AcDbImage", _
                     "AcDbRotatedDimension", "AcDbAlignedDimension", "AcDbAngularDimension"
                    objEntity.Delete
            End Select
        Next objEntity

        .PurgeAll
        .Close True
    End With
    Set objDocument = Nothing
End Sub



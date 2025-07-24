VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSchtek 
   Caption         =   "UserForm1"
   ClientHeight    =   2796
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5352
   OleObjectBlob   =   "frmSchtek.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSchtek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton5_Click()
Dim element2

For Each element2 In ThisDrawing.ModelSpace
      If element2.Name = UCase("EXTERNAL REFERENCE") Then MsgBox "YEP"
Next element2
End Sub

Private Sub UserForm_Initialize()

ComboBox1.AddItem ("BOUWKUNDIG")
ComboBox1.AddItem ("ALGEMEEN")
ComboBox1.ListIndex = 0
  
End Sub

Private Sub CommandButton1_Click()
On Error Resume Next
UserForm1.Hide

Dim controle
Dim laagobj As AcadLayer
For Each laagobj In ThisDrawing.Layers
controle = InStr(1, laagobj.Name, UCase("STRAMIEN"), vbTextCompare)
If controle <> 0 Then
   laagobj.Name = "Stramien"
   laagobj.color = acRed
End If
Next laagobj

End Sub
Private Sub CommandButton3_Click()
Dim element2
For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = UCase("LWPOLYLINE") Then MsgBox "YEP"
Next element2
End Sub

Private Sub CommandButton4_Click()
Dim aa
aa = UCase(ComboBox1)
UserForm1.Hide





Dim controle
Dim laagobj As AcadLayer
For Each laagobj In ThisDrawing.Layers
controle = InStr(1, laagobj.Name, aa, vbTextCompare)
If controle <> 0 Then
   laagobj.Name = "Bouwkundig"
   laagobj.color = acByLayer
End If
Next laagobj


ThisDrawing.SendCommand "-layer" & vbCr & "Make" & vbCr & "3" & vbCr & vbCr
ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr
ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr
ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr
Update

Dim layer
Dim entHandle As String
    Dim entry As AcadEntity
    For Each entry In ThisDrawing.ModelSpace
        entHandle = entry.ObjectName
        entry.Highlight (True)
        'MsgBox "The handle of this object is " & entHandle, vbInformation, "Handle Example"
        entry.Highlight (False)
        If entHandle = "AcDbDimension" Then entry.Delete
        If entHandle = "AcDbHatch" Then entry.Delete
        If entHandle = "AcDbSolid" Then entry.Delete
        If entHandle = "AcDbText" Then
            entry.layer = 3
            entry.color = acWhite
        End If
        If entHandle = "AcDbMText" Then
            entry.layer = 3
            entry.color = acWhite
        End If
        If entHandle = "AcDbImage" Then entry.Delete
        Update
    Next

ThisDrawing.SendCommand "-purge" & vbCr & "all" & vbCr & vbCr & "N" & vbCr
ThisDrawing.SendCommand "wthlayer" & vbCr
Unload Me
End Sub


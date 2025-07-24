VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmverwijder 
   Caption         =   "Verwijder"
   ClientHeight    =   3060
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   3852
   OleObjectBlob   =   "frmverwijder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmverwijder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim layer
On Error Resume Next

Dim entHandle As String
    Dim entry As AcadEntity
    For Each entry In ThisDrawing.ModelSpace
        entHandle = entry.ObjectName
        'entry.Highlight (True)
  
        
        If entHandle = "AcDbText" Then
            If InStr(1, entry.TextString, TextBox1, vbBinaryCompare) Then entry.Delete
            If InStr(1, entry.TextString, UCase(TextBox1), vbBinaryCompare) Then entry.Delete
        End If
        
        If entHandle = "AcDbMText" Then
            If InStr(1, entry.TextString, TextBox1, vbBinaryCompare) Then entry.Delete
            If InStr(1, entry.TextString, UCase(TextBox1), vbBinaryCompare) Then entry.Delete
        End If
        
        Update

    Next
    
TextBox1 = Clear
TextBox1.SetFocus
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

'''''''''Private Sub CommandButton3_Click()
'''''''''
'''''''''frmverwijder.Hide
'''''''''
'''''''''Dim sset As AcadSelectionSet
'''''''''Set sset = ThisDrawing.SelectionSets.Add("SS3")
'''''''''sset.SelectOnScreen
'''''''''
'''''''''Dim ent As Object
'''''''''For Each ent In sset
'''''''''   If ent.ObjectName = "AcDbBlockreference" Then
'''''''''        b = ent.Object.Name
'''''''''      ListBox1.AddItem (b)
'''''''''   End If
'''''''''  ent.Update
'''''''''
'''''''''Next ent
'''''''''sset.Delete
'''''''''
'''''''''frmverwijder.Show
'''''''''
'''''''''End Sub

Private Sub CommandButton4_Click()
Dim layer
Dim entHandle As String
    Dim entry As AcadEntity
    For Each entry In ThisDrawing.ModelSpace
        entHandle = entry.ObjectName
        entry.Highlight (True)
        'MsgBox "The handle of this object is " & entHandle, vbInformation, "Handle Example"
        entry.Highlight (False)
        If entHandle = "AcDbBlockReference" Then ListBox1.AddItem (entry.Name)
        Update
    Next

End Sub

Private Sub CommandButton5_Click()

Dim layer
Dim entHandle As String
    Dim entry As AcadEntity
    For Each entry In ThisDrawing.ModelSpace
        entHandle = entry.ObjectName
        
        If entHandle = "AcDbText" Then
            
            If InStr(1, entry.TextString, "Keuken", vbBinaryCompare) Then entry.TextString = "Kitchen"
        End If
        
        If entHandle = "AcDbMText" Then
            If InStr(1, entry.TextString, "H.O.H. ", vbBinaryCompare) Then entry.TextString = " c.t.c. "
           
        End If
        
        Update
    Next

End Sub

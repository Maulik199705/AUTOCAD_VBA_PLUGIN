VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmopschoon 
   Caption         =   "UserForm1"
   ClientHeight    =   3120
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4704
   OleObjectBlob   =   "frmopschoon.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmopschoon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgo_Click()
'''    Dim drawing As AcadDocument
'''    For Each drawing In ThisDrawing.Application.Documents
'''    drawing.Activate
On Error Resume Next
ThisDrawing.SendCommand "audit" & vbCr & "y" & vbCr
        
    Dim objBlock As AcadBlock
    Dim objEntity As AcadEntity
    Dim objLayer As AcadLayer
    
''''      Dim mypos
''''
''''      For Each objBlock In ThisDrawing.Blocks
''''        For Each objEntity In objBlock
''''          mypos = InStr(1, objEntity, "STRAMIEN")
''''          objEntity.layer = "STRAMIEN"
''''          objEntity.color = acByLayer
''''          objEntity.Lineweight = acLnWtByLwDefault
''''        Next objEntity
''''      Next objBlock
         
Dim olddata As String
olddata = ThisDrawing.GetVariable("clayer")
ThisDrawing.SendCommand "-layer" & vbCr & "M" & vbCr & "BOUWKUNDIG" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "M" & vbCr & "3" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Set" & vbCr & olddata & vbCr & vbCr
Update




  
      'alles naar layer bouwkundig
      Dim object3
        For Each object3 In ThisDrawing.ModelSpace
           object3.layer = "bouwkundig"
        Next
       
       
'set all object colors to bylayer
      For Each objBlock In ThisDrawing.Blocks
        For Each objEntity In objBlock
          objEntity.layer = "bouwkundig"
          objEntity.color = acByLayer
          objEntity.Lineweight = acLnWtByLwDefault
        Next objEntity
      Next objBlock
      
      ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr
      
      
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
        If entHandle = "AcDbRotatedDimension" Then entry.Delete
        If entHandle = "AcDbAlignedDimension" Then entry.Delete
        If entHandle = "AcDbAngularDimension" Then entry.Delete
        
        If entHandle = "AcDbText" Then
            entry.layer = "3"
            entry.color = acByLayer
            If InStr(1, entry.TextString, "merk", vbBinaryCompare) Then entry.Delete
            If InStr(1, entry.TextString, "MERK", vbBinaryCompare) Then entry.Delete
        End If
        If entHandle = "AcDbMText" Then
    
            If InStr(1, entry.TextString, "merk", vbBinaryCompare) Then entry.Delete
            If InStr(1, entry.TextString, "MERK", vbBinaryCompare) Then entry.Delete
            If UCase(entry.TextString) = "K." Then entry.Delete
            If UCase(entry.TextString) = "MV" Then entry.Delete
            If UCase(entry.TextString) = "HWA" Then entry.Delete
            If UCase(entry.TextString) = "H.W.A." Then entry.Delete
            If UCase(entry.TextString) = "WINDVERBAND" Then entry.Delete
            entry.layer = "3"
            entry.color = acByLayer
        End If
        If entHandle = "AcDbImage" Then entry.Delete
        Update
    Next



ThisDrawing.PurgeAll
ThisDrawing.SendCommand "wthlayer" & vbCr
ThisDrawing.Save

'''    Next drawing

End Sub

Attribute VB_Name = "Checklayer2"
Sub Checklayer2()
On Error Resume Next

'als groepslayer leeg is dan verwijderen
'uitgevoerd na knop indrukken
Dim layerObj As Object
For Each layerObj In ThisDrawing.Layers
  mystr = Left(layerObj.Name, 5)
   If mystr = "groep" Or mystr = "GROEP" Or mystr = "wand" Or mystr = "WAND" Then
    If layerObj.Length = 0 Then layerObj.Delete
   End If 'mystr
Next layerObj
Update

'frmGroeptekst.Hide
ThisDrawing.SendCommand "-purge" & vbCr & "all" & vbCr & "*" & vbCr & "N" & vbCr
ThisDrawing.SendCommand "wthlayer" & vbCr
'frmGroeptekst.Show
End Sub


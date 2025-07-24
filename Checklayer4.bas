Attribute VB_Name = "Checklayer4"
Sub Checklayer4()
On Error Resume Next

'als groepslayer leeg is dan verwijderen i.v.m. dubbele layernamen 1.01 en 1.01h bijvoorbeeld
'uitgevoerd na knop indrukken
Dim layerObj As Object
For Each layerObj In ThisDrawing.Layers
  mystr = Left(layerObj.Name, 5)
   If mystr = "groep" Or mystr = "GROEP" Or mystr = "wand" Or mystr = "WAND" Then
    If layerObj.Length = 0 Then layerObj.Delete
   End If 'mystr
Next layerObj
Update

End Sub


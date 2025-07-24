Attribute VB_Name = "Checklayer"
Sub Checklayer()
On Error Resume Next

'als groeps- of wandlayer leeg is dan verwijderen
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


Attribute VB_Name = "Checklayer1"
Sub Checklayer1()
On Error Resume Next

'Dim varData As Variant
'varData = ThisDrawing.GetVariable("users5")

'If varData <> "check" Then
'als groepslayer leeg is dan verwijderen
'wordt de 1e keer automatisch uitgevoerd
Dim layerObj As Object
For Each layerObj In ThisDrawing.Layers
  mystr = Left(layerObj.Name, 4)
  If mystr = "groep" Or mystr = "GROEP" Or mystr = "wand" Or mystr = "WAND" Then
  If layerObj.Length = 0 Then layerObj.Delete
 End If 'mystr
Next layerObj
Update
'End If

'ThisDrawing.SetVariable "users5", "check"
End Sub

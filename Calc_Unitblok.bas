Attribute VB_Name = "Calc_Unitblok"
Sub Unitblok(scaal, pb2, bestand2)
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb2, bestand2, scaal, scaal, 1, 0)
  Update
End Sub

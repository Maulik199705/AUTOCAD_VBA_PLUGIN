Attribute VB_Name = "Check_Kaderlogo"
'checken of er meerdere kaderlogo's in de tekening staan
Sub Check_Kaderlogo()

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then a = a + 1
      End If
Next element
      
      
If a > 1 Then

For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
       If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then
       Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" And attribuut.textstring = "" Then element.Erase
        Next I
       End If
     End If
   End If
  Next element

End If



Dim pe2(0 To 2) As Double
Dim element2
Dim insp

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
         If element2.Name = "ba3" Or element2.Name = "ba2" Or element2.Name = "ba1" _
         Or element2.Name = "ba0" Or element2.Name = "ba0+" Then
          insp = element2.InsertionPoint

          pe2(0) = insp(0)
          pe2(1) = insp(1)
          pe2(2) = 0
          If pe2(0) <> 0 Or pe2(1) <> 0 Then
          MsgBox "Je kader staat niet goed...!!!", vbCritical
          element2.Erase
          Exit Sub
          End If
          'MsgBox pe2(0) & " - " & pe2(1) & " - " & pe2(2) & " - "

          End If

      End If
Next element2

End Sub

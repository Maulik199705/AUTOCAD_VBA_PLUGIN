Attribute VB_Name = "Unitblok"
Sub Unitblok(aantal_groepen, pb2, bestand2, bestand3, scaal)

  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb2, bestand2, scaal, scaal, 1, 0)
  Update
  
If frmGroeptekst.OptionButton7 = True Or frmGroeptekst.OptionButton8 = True Or ComboBox2 = "RINGLEIDING" Then
 afmunitRING = "RINGLEIDING"
Else


 'open text bestand om afmetingen van de unit uit te lezen
  aantal_groepen2 = aantal_groepen
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fs, a, afmunit
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(bestand3, ForReading, False)
Do While aantal_groepen2 <> 0
    afmunit = a.ReadLine
    aantal_groepen2 = aantal_groepen2 - 1
Loop
a.Close 'sluiten van tekstbestand
End If
 


unittel = frmGroeptekst.TextBox9
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 0 And unittel < 10 Then unitonder10 = "0" & frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = False Then
  If unittel > 9 Then unitonder10 = frmGroeptekst.TextBox9
End If
If frmGroeptekst.CheckBox3.Value = True Then
   unitonder10 = frmGroeptekst.TextBox9
End If


Dim element4 As Object
unit = frmGroeptekst.ComboBox2
If frmGroeptekst.ComboBox2 = "RUW-Groot" Then unit = "RUW"
If frmGroeptekst.ComboBox2 = "RUW-Klein" Then unit = "RUW"
If frmGroeptekst.ComboBox2 = "RUB-R" And aantal_groepen > 4 Then unit = "RUH-R"
If frmGroeptekst.ComboBox2 = "RUB-RT" And aantal_groepen > 4 Then unit = "RUH-RT"
If frmGroeptekst.ComboBox2 = "RUB-S" And aantal_groepen > 0 Then unit = "RUH-S"
If frmGroeptekst.ComboBox2 = "VSKO" Then unit = "VSKO-B"

For Each element4 In ThisDrawing.ModelSpace
      If element4.ObjectName = "AcDbBlockReference" Then
      If UCase(element4.Name) = unit Then
      Set symbool = element4
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If unit = "RINGLEIDING" Then
        If attribuut.TagString = "AFMETINGEN" And attribuut.textstring = "" Then attribuut.textstring = afmunitRING 'AFMETING VAN DEUNIT
        Else
        If attribuut.TagString = "AFMETINGEN" And attribuut.textstring = "" Then attribuut.textstring = afmunit 'AFMETING VAN DEUNIT
        End If
        If attribuut.TagString = "UNITNUMMER" And attribuut.textstring = "" Then attribuut.textstring = unitonder10 'UNITNUMMER
        Next I
       End If
      End If
      End If
 Next element4
End Sub

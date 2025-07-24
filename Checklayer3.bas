Attribute VB_Name = "Checklayer3"
Sub Checklayer3(a)

'kijken of de groepslayer al gebruikt is.
'a = 0
Dim element As Object
Dim layerObj As AcadLayer

optel = frmGroeptekst.TextBox10
  If optel > 0 And optel < 10 Then groeponder10 = "0"
unittel = frmGroeptekst.TextBox9
    
  If frmGroeptekst.CheckBox3.Value = False Then
    If unittel > 0 And unittel < 10 Then unitonder10 = "0"
    groepsnummer = "groep " & unitonder10 & frmGroeptekst.TextBox9 & "." & groeponder10 & optel    'tekst samenvoegen
    End If
  If frmGroeptekst.CheckBox3.Value = True Then
    groepsnummer = "groep " & frmGroeptekst.TextBox9 & "." & groeponder10 & optel    'tekst samenvoegen
  End If
  
'groepsnummer = "groep " & frmGroeptekst.TextBox9 & "." & groeponder10 & optel  'tekst samenvoegen

For Each layerObj In ThisDrawing.Layers
     If layerObj.Name = groepsnummer Then
       '  fixlayer = "groep " & unitonder10 & frmGroeptekst.TextBox9 & "." & groeponder10 & optel & "h"
       '  layerObj.Name = fixlayer
       ' MsgBox "Layernaam bestaat al.!!! --> Eerst de layernaam hernummeren voordat je verder gaat.!!!", vbExclamation
      a = 1
      frmGroeptekst.TextBox12 = frmGroeptekst.TextBox9
      'frmGroeptekst.TextBox9 = Clear
      'frmGroeptekst.TextBox10 = Clear
      'frmGroeptekst.TextBox12.SetFocus
      'Exit Sub
     End If
Next 'layerobj

End Sub



              
            

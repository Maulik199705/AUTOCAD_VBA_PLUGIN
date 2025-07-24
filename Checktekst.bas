Attribute VB_Name = "Checktekst"
Sub Checktekst(a)

'kijken of de groeptekst al gebruikt is.
a = 0
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

For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
        If element.Name = "GROEPTEKSTBLOK" Or element.Name = "groeptekstblok" Then
        Set symbool = element
         If symbool.HasAttributes Then
           attributen = symbool.GetAttributes
           For I = LBound(attributen) To UBound(attributen)
           Set attribuut = attributen(I)
            If attribuut.TagString = "GROEPTEKST" Then sjektekst = attribuut.textstring
             If groepsnummer = sjektekst Then
               MsgBox "Groeptekst al gebruikt", vbExclamation
               a = 1
               frmGroeptekst.TextBox9 = Clear
               frmGroeptekst.TextBox10 = Clear
               frmGroeptekst.TextBox9.SetFocus
               Exit Sub
            End If
           Next I
        End If
      End If
    End If
  Next element
End Sub

Attribute VB_Name = "Zoek_blok"
Sub zoek_blok()

unittel = frmGroeptekst.TextBox9

If unittel > 0 And unittel < 10 Then
  unitonder10 = "0" & frmGroeptekst.TextBox9
Else
  unitonder10 = frmGroeptekst.TextBox9
End If


For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_PE" Or element.Name = "Mat_spe_ZD" Or element.Name = "Mat_spe_ZD_1627" Or element.Name = "Mat_spe_ALU" Or _
      element.Name = "Mat_spe_PEringleiding" Or element.Name = "Mat_spe_ZDringleiding" Or _
      element.Name = "Mat_spe_ALUringleiding" Or element.Name = "Mat_spe_PE800" Or element.Name = "Mat_spe_FLEX" Then
      
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "RNU" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
            If bb = unitonder10 Then
              For k = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(k)
       
                 If attribuut.TagString = "PE" Then
                 frmGroeptekst.OptionButton2.Value = True
                 frmGroeptekst.ComboBox4.Value = attribuut.textstring
                 frmGroeptekst.CheckBox8.Visible = True
                 End If
                 If attribuut.TagString = "PE" And Left(attribuut.textstring, 7) = "ALUFLEX" Then
                 frmGroeptekst.OptionButton6.Value = True
                 frmGroeptekst.ComboBox4.Value = attribuut.textstring
                 End If
                 If attribuut.TagString = "WTHZD" And attribuut.textstring = "WTH-ZD 16 * 2,7 mm" Then OptionButton9 = True
                 If attribuut.TagString = "WTHZD" And OptionButton1 = True Then frmGroeptekst.ComboBox4.Value = attribuut.textstring
                 If attribuut.TagString = "ALU" Then frmGroeptekst.ComboBox4.Value = attribuut.textstring
                 If attribuut.TagString = "BEVESTIGINGSTYPE" Then frmGroeptekst.ComboBox1.Value = attribuut.textstring
                 
                 If attribuut.TagString = "REGELUNITTYPE" Then
                   RT = attribuut.textstring
                 
                  If RT <> "RINGLEIDING" Then
                  trimstring = Split(RT, (" "))
                  Dim mystr As Variant
                  mystr = Len(trimstring(1))
                  frmGroeptekst.ComboBox2 = trimstring(0)
                         
                   If mystr > 2 Then
                   trimstring2 = Split(trimstring(1), ("/"))
                   ComboBox3 = trimstring2(1)
                   End If
                 
                  End If
                 If RT = "RINGLEIDING" Then frmGroeptekst.OptionButton7.Value = True
        
                 
                 End If
             Next k
             bb = ""
             End If
       Next I
       
        End If
      End If
      End If
  Next element


End Sub

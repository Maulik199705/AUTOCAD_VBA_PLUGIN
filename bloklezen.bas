Attribute VB_Name = "bloklezen"
Sub blokl()


Dim element As Object
Dim layerObj As AcadLayer
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_ZD" Or element.Name = "Mat_spe_ZDringleiding" Or element.Name = "Mat_spe_PE" Or _
      element.Name = "Mat_spe_PEringleiding" Then
      Set SYMBOOL = element
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
               'groepen tijdelijk hernummeren
                If ATTRIBUUT.TagString = "RNU" Then unitnummer = ATTRIBUUT.textstring
                    If ATTRIBUUT.TagString = "REGELUNITTYPE" Then
                      tiepe = ATTRIBUUT.textstring
                        If unitnummer = "01" Then frmAcadNavi.TextBox1 = tiepe
                        If unitnummer = "02" Then frmAcadNavi.TextBox2 = tiepe
                        If unitnummer = "03" Then frmAcadNavi.TextBox3 = tiepe
                        If unitnummer = "04" Then frmAcadNavi.TextBox4 = tiepe
                        If unitnummer = "05" Then frmAcadNavi.TextBox5 = tiepe
                        If unitnummer = "06" Then frmAcadNavi.TextBox6 = tiepe
                        If unitnummer = "07" Then frmAcadNavi.TextBox7 = tiepe
                        If unitnummer = "08" Then frmAcadNavi.TextBox8 = tiepe
                        If unitnummer = "09" Then frmAcadNavi.TextBox9 = tiepe
                        If unitnummer = "10" Then frmAcadNavi.TextBox10 = tiepe
                        If unitnummer = "11" Then frmAcadNavi.TextBox11 = tiepe
                        If unitnummer = "12" Then frmAcadNavi.TextBox12 = tiepe
                        If unitnummer = "13" Then frmAcadNavi.TextBox13 = tiepe
                        If unitnummer = "14" Then frmAcadNavi.TextBox14 = tiepe
                        If unitnummer = "15" Then frmAcadNavi.TextBox15 = tiepe
                        If unitnummer = "16" Then frmAcadNavi.TextBox16 = tiepe
                        If unitnummer = "17" Then frmAcadNavi.TextBox17 = tiepe
                        If unitnummer = "18" Then frmAcadNavi.TextBox18 = tiepe
                        If unitnummer = "19" Then frmAcadNavi.TextBox19 = tiepe
                        If unitnummer = "20" Then frmAcadNavi.TextBox20 = tiepe
                    End If
                    
                             
                If ATTRIBUUT.TagString = "BEVESTIGINGSTYPE" Then
                        If unitnummer = "01" Then frmAcadNavi.TextBox200 = ATTRIBUUT.textstring
                        If unitnummer = "02" Then frmAcadNavi.TextBox201 = ATTRIBUUT.textstring
                        If unitnummer = "03" Then frmAcadNavi.TextBox202 = ATTRIBUUT.textstring
                        If unitnummer = "04" Then frmAcadNavi.TextBox203 = ATTRIBUUT.textstring
                        If unitnummer = "05" Then frmAcadNavi.TextBox204 = ATTRIBUUT.textstring
                        If unitnummer = "06" Then frmAcadNavi.TextBox205 = ATTRIBUUT.textstring
                        If unitnummer = "07" Then frmAcadNavi.TextBox206 = ATTRIBUUT.textstring
                        If unitnummer = "08" Then frmAcadNavi.TextBox207 = ATTRIBUUT.textstring
                        If unitnummer = "09" Then frmAcadNavi.TextBox208 = ATTRIBUUT.textstring
                        If unitnummer = "10" Then frmAcadNavi.TextBox209 = ATTRIBUUT.textstring
                        If unitnummer = "11" Then frmAcadNavi.TextBox210 = ATTRIBUUT.textstring
                        If unitnummer = "12" Then frmAcadNavi.TextBox211 = ATTRIBUUT.textstring
                        If unitnummer = "13" Then frmAcadNavi.TextBox212 = ATTRIBUUT.textstring
                        If unitnummer = "14" Then frmAcadNavi.TextBox213 = ATTRIBUUT.textstring
                        If unitnummer = "15" Then frmAcadNavi.TextBox214 = ATTRIBUUT.textstring
                        If unitnummer = "16" Then frmAcadNavi.TextBox215 = ATTRIBUUT.textstring
                        If unitnummer = "17" Then frmAcadNavi.TextBox216 = ATTRIBUUT.textstring
                        If unitnummer = "18" Then frmAcadNavi.TextBox217 = ATTRIBUUT.textstring
                        If unitnummer = "19" Then frmAcadNavi.TextBox218 = ATTRIBUUT.textstring
                        If unitnummer = "20" Then frmAcadNavi.TextBox219 = ATTRIBUUT.textstring
                        
                        
                        If tiepe <> "" And ATTRIBUUT.textstring = "Vlechtdraad" Then
                        tiepe1 = Split(tiepe, " ")
                        'MsgBox tiepe1(1)
                        frmAcadNavi.TextBox220 = frmAcadNavi.TextBox220 + (Val(tiepe1(1)))
                        End If
                        
                        If tiepe <> "" And ATTRIBUUT.textstring = "Witmarmerbeugels" Then
                            For a = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                            Set ATTRIBUUT = ATTRIBUTEN(a)
                               If ATTRIBUUT.TagString = "PE" And (ATTRIBUUT.textstring = "PE-RT 16*2 mm" Or ATTRIBUUT.textstring = "PE-RT 14*2 mm") Then
                                    tiepe1 = Split(tiepe, " ")
                                    frmAcadNavi.TextBox222 = frmAcadNavi.TextBox222 + (Val(tiepe1(1)))
                               End If
                               If ATTRIBUUT.TagString = "WTHZD" And ATTRIBUUT.textstring = "WTH-ZD 20 * 3,4 mm" Then
                                    tiepe1 = Split(tiepe, " ")
                                    frmAcadNavi.TextBox221 = frmAcadNavi.TextBox221 + (Val(tiepe1(1)))
                               End If
                            
                            Next a
                        End If
                        If tiepe <> "" And ATTRIBUUT.textstring = "Ty-raps" Then
                        tiepe1 = Split(tiepe, " ")
                        'MsgBox tiepe1(1)
                        frmAcadNavi.TextBox223 = frmAcadNavi.TextBox223 + (Val(tiepe1(1)))
                        End If
                        
                        If tiepe <> "" And ATTRIBUUT.textstring = "Isoclips" Then
                            For b = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                            Set ATTRIBUUT = ATTRIBUTEN(b)
                               If ATTRIBUUT.TagString = "PE" And ATTRIBUUT.textstring = "PE-RT 16*2 mm" Then
                                    tiepe1 = Split(tiepe, " ")
                                    frmAcadNavi.TextBox224 = frmAcadNavi.TextBox224 + (Val(tiepe1(1)))
                               End If
                               If ATTRIBUUT.TagString = "WTHZD" And ATTRIBUUT.textstring = "WTH-ZD 20 * 3,4 mm" Then
                                    tiepe1 = Split(tiepe, " ")
                                    frmAcadNavi.TextBox225 = frmAcadNavi.TextBox225 + (Val(tiepe1(1)))
                               End If
                            
                            Next b
                        End If
                        
                        If tiepe <> "" And ATTRIBUUT.textstring = "Varisoclips" Then
                                    tiepe1 = Split(tiepe, " ")
                                    frmAcadNavi.TextBox226 = frmAcadNavi.TextBox226 + (Val(tiepe1(1)))
                        End If

                        If tiepe <> "" And ATTRIBUUT.textstring = "Beugels/Nagels" Then
                        tiepe1 = Split(tiepe, " ")
                        frmAcadNavi.TextBox228 = frmAcadNavi.TextBox228 + (Val(tiepe1(1)))
                        End If

                        If tiepe <> "" And ATTRIBUUT.textstring = "IFD-Polystyreen" Then frmAcadNavi.CheckBox3.Value = True
                        If tiepe <> "" And ATTRIBUUT.textstring = "Keg" Then frmAcadNavi.CheckBox4.Value = True
                        If tiepe <> "" And ATTRIBUUT.textstring = "Montagestrip" Then frmAcadNavi.CheckBox5.Value = True
                        If tiepe <> "" And ATTRIBUUT.textstring = "Noppenplaat" Then frmAcadNavi.CheckBox6.Value = True
                        If tiepe <> "" And ATTRIBUUT.textstring = "Schietbeugels" Then frmAcadNavi.CheckBox7.Value = True
                    End If
                
                
                
                If ATTRIBUUT.TagString = "WTH250" And ATTRIBUUT.textstring <> "" Then _
                 frmAcadNavi.TextBox44 = frmAcadNavi.TextBox44 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH165" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox45 = frmAcadNavi.TextBox45 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH125" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox46 = frmAcadNavi.TextBox46 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH105" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox47 = frmAcadNavi.TextBox47 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH90" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox48 = frmAcadNavi.TextBox48 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH75" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox49 = frmAcadNavi.TextBox49 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH63" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox50 = frmAcadNavi.TextBox50 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH50" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox51 = frmAcadNavi.TextBox51 + (Val(ATTRIBUUT.textstring))
                If ATTRIBUUT.TagString = "WTH40" And ATTRIBUUT.textstring <> "" Then _
                frmAcadNavi.TextBox52 = frmAcadNavi.TextBox52 + (Val(ATTRIBUUT.textstring))
                
                              
                 If ATTRIBUUT.TagString = "PE" And ATTRIBUUT.textstring = "PE-RT 16*2 mm" Then
                  For L = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                     Set ATTRIBUUT = ATTRIBUTEN(L)
                  If ATTRIBUUT.TagString = "PE120" And ATTRIBUUT.textstring <> "" Then frmAcadNavi.TextBox101 = frmAcadNavi.TextBox101 + (Val(ATTRIBUUT.textstring))  '120 METER
                  If ATTRIBUUT.TagString = "PE90" And ATTRIBUUT.textstring <> "" Then frmAcadNavi.TextBox102 = frmAcadNavi.TextBox102 + (Val(ATTRIBUUT.textstring))  '90 METER
                  If ATTRIBUUT.TagString = "PE60" And ATTRIBUUT.textstring <> "" Then frmAcadNavi.TextBox103 = frmAcadNavi.TextBox103 + (Val(ATTRIBUUT.textstring)) '60 METER
                  Next L
                 End If
                
                If ATTRIBUUT.TagString = "PE" And ATTRIBUUT.textstring = "PE-RT 14*2 mm" Then
                  For m = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                     Set ATTRIBUUT = ATTRIBUTEN(m)
                  If ATTRIBUUT.TagString = "PE90" And ATTRIBUUT.textstring <> "" Then frmAcadNavi.TextBox104 = frmAcadNavi.TextBox104 + (Val(ATTRIBUUT.textstring)) '90 METER
                  If ATTRIBUUT.TagString = "PE60" And ATTRIBUUT.textstring <> "" Then frmAcadNavi.TextBox105 = frmAcadNavi.TextBox105 + (Val(ATTRIBUUT.textstring)) '60 METER
                  Next m
                End If
                
                
                
              Next i
       
        End If
      End If
      End If
Next element
''''''''  Or element.Name = "Mat_spe_PE" Or element.Name = "Mat_spe_PE800" _
''''''''      Or element.Name = "Mat_spe_ALU" Or element.Name = "Mat_spe_ZDringleiding" Or element.Name = "Mat_spe_PEringleiding" Or _
''''''''      element.Name = "Mat_spe_ALUringleiding" Or element.Name = "Mat_spe_FLEX" Or _
''''''''      element.Name = "Mat_spe_FLEX_Aankoppel"
 
Dim WD

For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If UCase(element.Name) = "KADERLOGO" Then
      Set SYMBOOL = element
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
               'groepen tijdelijk hernummeren
                If ATTRIBUUT.TagString = "BLAD" Then frmAcadNavi.TextBox43 = ATTRIBUUT.textstring
                If ATTRIBUUT.TagString = "DATUM" Then frmAcadNavi.ListBox1.AddItem (ATTRIBUUT.textstring)
                If ATTRIBUUT.TagString = "WIJZIGING1" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
                If ATTRIBUUT.TagString = "WIJZIGING2" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
                If ATTRIBUUT.TagString = "WIJZIGING3" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
                If ATTRIBUUT.TagString = "WIJZIGING4" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
                If ATTRIBUUT.TagString = "WIJZIGING5" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
                If ATTRIBUUT.TagString = "WIJZIGING6" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
                If ATTRIBUUT.TagString = "WIJZIGING7" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
                If ATTRIBUUT.TagString = "REVISIE" And ATTRIBUUT.textstring <> "" Then
                     WD = Split(ATTRIBUUT.textstring, "|")
                     frmAcadNavi.ListBox1.AddItem (WD(0))
                End If
         Next i
       
        End If
      End If
      End If
Next element
    
Dim teller
Dim textstring
 teller = frmAcadNavi.ListBox1.ListCount
 'MsgBox TELLER
For i = 0 To teller - 1
   'Define the text object
    textstring = frmAcadNavi.ListBox1.List(i)
Next i
frmAcadNavi.TextBox100 = textstring

End Sub
Sub uitlez() 'ruimtenaam

''''End Sub
''''Sub lstbox2()

frmnaregelblok.ListBox2.Clear
frmnaregelblok.ComboBox200.Clear
frmnaregelblok.ComboBox201.Clear
frmnaregelblok.ComboBox202.Clear
frmnaregelblok.ComboBox203.Clear
frmnaregelblok.ComboBox204.Clear
frmnaregelblok.ComboBox205.Clear
frmnaregelblok.ComboBox206.Clear
frmnaregelblok.ComboBox207.Clear
frmnaregelblok.ComboBox208.Clear
frmnaregelblok.ComboBox209.Clear
frmnaregelblok.ComboBox210.Clear
frmnaregelblok.ComboBox211.Clear
frmnaregelblok.ComboBox212.Clear
frmnaregelblok.ComboBox213.Clear
frmnaregelblok.ComboBox214.Clear
frmnaregelblok.ComboBox215.Clear
frmnaregelblok.ComboBox216.Clear
frmnaregelblok.ComboBox217.Clear
frmnaregelblok.ComboBox218.Clear
frmnaregelblok.ComboBox219.Clear


bestand10 = "g:\tekeningen\ruimtenaam.txt"
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim fs, a, vullistbox1
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(bestand10, ForReading, False)


Do While a.AtEndOfLine <> True
    vullistbox1 = a.ReadLine
    frmnaregelblok.ListBox2.AddItem (vullistbox1)
Loop
a.Close 'sluiten van tekstbestand

  'lijst rangschikken
  Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For i = 0 To frmnaregelblok.ListBox2.ListCount - 1
    textstring2 = frmnaregelblok.ListBox2.List(i)
    Veld(i) = textstring2
   
   Dim LB&, UB&, TEMP$, Pos&, x&
 
    LB = LBound(Veld)
    UB = UBound(Veld)
 
    While UB > LB
      Pos = LB
 
      For x = LB To UB - 1
        If Veld(x) > Veld(x + 1) Then
          TEMP = Veld(x + 1)
          Veld(x + 1) = Veld(x)
          Veld(x) = TEMP
          Pos = x
        End If
      Next x
 
      UB = Pos
    Wend
    
  Next i
  frmnaregelblok.ListBox2.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then frmnaregelblok.ListBox2.AddItem Veld(x)
  Next x

''''''================ eerst in de listbox sorteren en dan in de combobox plaatsen================
For i = 0 To frmnaregelblok.ListBox2.ListCount - 1
    textstring3 = frmnaregelblok.ListBox2.List(i)
    frmnaregelblok.ComboBox200.AddItem (textstring3)
    frmnaregelblok.ComboBox201.AddItem (textstring3)
    frmnaregelblok.ComboBox202.AddItem (textstring3)
    frmnaregelblok.ComboBox203.AddItem (textstring3)
    frmnaregelblok.ComboBox204.AddItem (textstring3)
    frmnaregelblok.ComboBox205.AddItem (textstring3)
    frmnaregelblok.ComboBox206.AddItem (textstring3)
    frmnaregelblok.ComboBox207.AddItem (textstring3)
    frmnaregelblok.ComboBox208.AddItem (textstring3)
    frmnaregelblok.ComboBox209.AddItem (textstring3)
    frmnaregelblok.ComboBox210.AddItem (textstring3)
    frmnaregelblok.ComboBox211.AddItem (textstring3)
    frmnaregelblok.ComboBox212.AddItem (textstring3)
    frmnaregelblok.ComboBox213.AddItem (textstring3)
    frmnaregelblok.ComboBox214.AddItem (textstring3)
    frmnaregelblok.ComboBox215.AddItem (textstring3)
    frmnaregelblok.ComboBox216.AddItem (textstring3)
    frmnaregelblok.ComboBox217.AddItem (textstring3)
    frmnaregelblok.ComboBox218.AddItem (textstring3)
    frmnaregelblok.ComboBox219.AddItem (textstring3)
    
    
Next i
End Sub

''''bestand10 = "g:\tekeningen\ruimtenaam.txt"
''''Const ForReading = 1, ForWriting = 2, ForAppending = 8
''''Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
''''
''''Dim fs, a, vullistbox2
''''Set fs = CreateObject("Scripting.FileSystemObject")
''''Set a = fs.OpenTextFile(bestand10, ForReading, False)
''''
''''
''''Do While a.AtEndOfLine <> True
''''    vullistbox2 = a.ReadLine
''''    frmnaregelblok.ComboBox200.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox201.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox202.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox203.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox204.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox205.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox206.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox207.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox208.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox209.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox210.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox211.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox212.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox213.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox214.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox215.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox216.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox217.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox218.AddItem (vullistbox2)
''''    frmnaregelblok.ComboBox219.AddItem (vullistbox2)
''''Loop
''''a.Close 'sluiten van tekstbestand

Sub view()
R = frmnaregelblok.groepinvoerBox1
If R = 1 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True

    frmnaregelblok.TextBox201.Visible = False: frmnaregelblok.ComboBox201.Visible = False: frmnaregelblok.CheckBox201.Visible = False
    frmnaregelblok.TextBox202.Visible = False: frmnaregelblok.ComboBox202.Visible = False: frmnaregelblok.CheckBox202.Visible = False
    frmnaregelblok.TextBox203.Visible = False: frmnaregelblok.ComboBox203.Visible = False: frmnaregelblok.CheckBox203.Visible = False
    frmnaregelblok.TextBox204.Visible = False: frmnaregelblok.ComboBox204.Visible = False: frmnaregelblok.CheckBox204.Visible = False
    frmnaregelblok.TextBox205.Visible = False: frmnaregelblok.ComboBox205.Visible = False: frmnaregelblok.CheckBox205.Visible = False
    frmnaregelblok.TextBox206.Visible = False: frmnaregelblok.ComboBox206.Visible = False: frmnaregelblok.CheckBox206.Visible = False
    frmnaregelblok.TextBox207.Visible = False: frmnaregelblok.ComboBox207.Visible = False: frmnaregelblok.CheckBox207.Visible = False
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 2 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    
    frmnaregelblok.TextBox202.Visible = False: frmnaregelblok.ComboBox202.Visible = False: frmnaregelblok.CheckBox202.Visible = False
    frmnaregelblok.TextBox203.Visible = False: frmnaregelblok.ComboBox203.Visible = False: frmnaregelblok.CheckBox203.Visible = False
    frmnaregelblok.TextBox204.Visible = False: frmnaregelblok.ComboBox204.Visible = False: frmnaregelblok.CheckBox204.Visible = False
    frmnaregelblok.TextBox205.Visible = False: frmnaregelblok.ComboBox205.Visible = False: frmnaregelblok.CheckBox205.Visible = False
    frmnaregelblok.TextBox206.Visible = False: frmnaregelblok.ComboBox206.Visible = False: frmnaregelblok.CheckBox206.Visible = False
    frmnaregelblok.TextBox207.Visible = False: frmnaregelblok.ComboBox207.Visible = False: frmnaregelblok.CheckBox207.Visible = False
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 3 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    
    frmnaregelblok.TextBox203.Visible = False: frmnaregelblok.ComboBox203.Visible = False: frmnaregelblok.CheckBox203.Visible = False
    frmnaregelblok.TextBox204.Visible = False: frmnaregelblok.ComboBox204.Visible = False: frmnaregelblok.CheckBox204.Visible = False
    frmnaregelblok.TextBox205.Visible = False: frmnaregelblok.ComboBox205.Visible = False: frmnaregelblok.CheckBox205.Visible = False
    frmnaregelblok.TextBox206.Visible = False: frmnaregelblok.ComboBox206.Visible = False: frmnaregelblok.CheckBox206.Visible = False
    frmnaregelblok.TextBox207.Visible = False: frmnaregelblok.ComboBox207.Visible = False: frmnaregelblok.CheckBox207.Visible = False
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
    
End If
If R = 4 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    
    frmnaregelblok.TextBox204.Visible = False: frmnaregelblok.ComboBox204.Visible = False: frmnaregelblok.CheckBox204.Visible = False
    frmnaregelblok.TextBox205.Visible = False: frmnaregelblok.ComboBox205.Visible = False: frmnaregelblok.CheckBox205.Visible = False
    frmnaregelblok.TextBox206.Visible = False: frmnaregelblok.ComboBox206.Visible = False: frmnaregelblok.CheckBox206.Visible = False
    frmnaregelblok.TextBox207.Visible = False: frmnaregelblok.ComboBox207.Visible = False: frmnaregelblok.CheckBox207.Visible = False
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 5 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    
    frmnaregelblok.TextBox205.Visible = False: frmnaregelblok.ComboBox205.Visible = False: frmnaregelblok.CheckBox205.Visible = False
    frmnaregelblok.TextBox206.Visible = False: frmnaregelblok.ComboBox206.Visible = False: frmnaregelblok.CheckBox206.Visible = False
    frmnaregelblok.TextBox207.Visible = False: frmnaregelblok.ComboBox207.Visible = False: frmnaregelblok.CheckBox207.Visible = False
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 6 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    
   frmnaregelblok.TextBox206.Visible = False: frmnaregelblok.ComboBox206.Visible = False: frmnaregelblok.CheckBox206.Visible = False
    frmnaregelblok.TextBox207.Visible = False: frmnaregelblok.ComboBox207.Visible = False: frmnaregelblok.CheckBox207.Visible = False
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 7 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    
    frmnaregelblok.TextBox207.Visible = False: frmnaregelblok.ComboBox207.Visible = False: frmnaregelblok.CheckBox207.Visible = False
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 8 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    
    frmnaregelblok.TextBox208.Visible = False: frmnaregelblok.ComboBox208.Visible = False: frmnaregelblok.CheckBox208.Visible = False
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 9 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    
    frmnaregelblok.TextBox209.Visible = False: frmnaregelblok.ComboBox209.Visible = False: frmnaregelblok.CheckBox209.Visible = False
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 10 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    
    frmnaregelblok.TextBox210.Visible = False: frmnaregelblok.ComboBox210.Visible = False: frmnaregelblok.CheckBox210.Visible = False
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 11 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    
    frmnaregelblok.TextBox211.Visible = False: frmnaregelblok.ComboBox211.Visible = False: frmnaregelblok.CheckBox211.Visible = False
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 12 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    
    frmnaregelblok.TextBox212.Visible = False: frmnaregelblok.ComboBox212.Visible = False: frmnaregelblok.CheckBox212.Visible = False
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 13 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    
    frmnaregelblok.TextBox213.Visible = False: frmnaregelblok.ComboBox213.Visible = False: frmnaregelblok.CheckBox213.Visible = False
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 14 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    
    frmnaregelblok.TextBox214.Visible = False: frmnaregelblok.ComboBox214.Visible = False: frmnaregelblok.CheckBox214.Visible = False
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 15 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    frmnaregelblok.TextBox214.Visible = True: frmnaregelblok.ComboBox214.Visible = True: frmnaregelblok.CheckBox214.Visible = True
    
    frmnaregelblok.TextBox215.Visible = False: frmnaregelblok.ComboBox215.Visible = False: frmnaregelblok.CheckBox215.Visible = False
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 16 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    frmnaregelblok.TextBox214.Visible = True: frmnaregelblok.ComboBox214.Visible = True: frmnaregelblok.CheckBox214.Visible = True
    frmnaregelblok.TextBox215.Visible = True: frmnaregelblok.ComboBox215.Visible = True: frmnaregelblok.CheckBox215.Visible = True
    
    frmnaregelblok.TextBox216.Visible = False: frmnaregelblok.ComboBox216.Visible = False: frmnaregelblok.CheckBox216.Visible = False
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 17 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    frmnaregelblok.TextBox214.Visible = True: frmnaregelblok.ComboBox214.Visible = True: frmnaregelblok.CheckBox214.Visible = True
    frmnaregelblok.TextBox215.Visible = True: frmnaregelblok.ComboBox215.Visible = True: frmnaregelblok.CheckBox215.Visible = True
    frmnaregelblok.TextBox216.Visible = True: frmnaregelblok.ComboBox216.Visible = True: frmnaregelblok.CheckBox216.Visible = True
    
    frmnaregelblok.TextBox217.Visible = False: frmnaregelblok.ComboBox217.Visible = False: frmnaregelblok.CheckBox217.Visible = False
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 18 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    frmnaregelblok.TextBox214.Visible = True: frmnaregelblok.ComboBox214.Visible = True: frmnaregelblok.CheckBox214.Visible = True
    frmnaregelblok.TextBox215.Visible = True: frmnaregelblok.ComboBox215.Visible = True: frmnaregelblok.CheckBox215.Visible = True
    frmnaregelblok.TextBox216.Visible = True: frmnaregelblok.ComboBox216.Visible = True: frmnaregelblok.CheckBox216.Visible = True
    frmnaregelblok.TextBox217.Visible = True: frmnaregelblok.ComboBox217.Visible = True: frmnaregelblok.CheckBox217.Visible = True
    
    frmnaregelblok.TextBox218.Visible = False: frmnaregelblok.ComboBox218.Visible = False: frmnaregelblok.CheckBox218.Visible = False
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 19 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    frmnaregelblok.TextBox214.Visible = True: frmnaregelblok.ComboBox214.Visible = True: frmnaregelblok.CheckBox214.Visible = True
    frmnaregelblok.TextBox215.Visible = True: frmnaregelblok.ComboBox215.Visible = True: frmnaregelblok.CheckBox215.Visible = True
    frmnaregelblok.TextBox216.Visible = True: frmnaregelblok.ComboBox216.Visible = True: frmnaregelblok.CheckBox216.Visible = True
    frmnaregelblok.TextBox217.Visible = True: frmnaregelblok.ComboBox217.Visible = True: frmnaregelblok.CheckBox217.Visible = True
    frmnaregelblok.TextBox218.Visible = True: frmnaregelblok.ComboBox218.Visible = True: frmnaregelblok.CheckBox218.Visible = True
    
    frmnaregelblok.TextBox219.Visible = False: frmnaregelblok.ComboBox219.Visible = False: frmnaregelblok.CheckBox219.Visible = False
    
End If
If R = 20 Then
    frmnaregelblok.TextBox200.Visible = True: frmnaregelblok.ComboBox200.Visible = True: frmnaregelblok.CheckBox200.Visible = True
    frmnaregelblok.TextBox201.Visible = True: frmnaregelblok.ComboBox201.Visible = True: frmnaregelblok.CheckBox201.Visible = True
    frmnaregelblok.TextBox202.Visible = True: frmnaregelblok.ComboBox202.Visible = True: frmnaregelblok.CheckBox202.Visible = True
    frmnaregelblok.TextBox203.Visible = True: frmnaregelblok.ComboBox203.Visible = True: frmnaregelblok.CheckBox203.Visible = True
    frmnaregelblok.TextBox204.Visible = True: frmnaregelblok.ComboBox204.Visible = True: frmnaregelblok.CheckBox204.Visible = True
    frmnaregelblok.TextBox205.Visible = True: frmnaregelblok.ComboBox205.Visible = True: frmnaregelblok.CheckBox205.Visible = True
    frmnaregelblok.TextBox206.Visible = True: frmnaregelblok.ComboBox206.Visible = True: frmnaregelblok.CheckBox206.Visible = True
    frmnaregelblok.TextBox207.Visible = True: frmnaregelblok.ComboBox207.Visible = True: frmnaregelblok.CheckBox207.Visible = True
    frmnaregelblok.TextBox208.Visible = True: frmnaregelblok.ComboBox208.Visible = True: frmnaregelblok.CheckBox208.Visible = True
    frmnaregelblok.TextBox209.Visible = True: frmnaregelblok.ComboBox209.Visible = True: frmnaregelblok.CheckBox209.Visible = True
    frmnaregelblok.TextBox210.Visible = True: frmnaregelblok.ComboBox210.Visible = True: frmnaregelblok.CheckBox210.Visible = True
    frmnaregelblok.TextBox211.Visible = True: frmnaregelblok.ComboBox211.Visible = True: frmnaregelblok.CheckBox211.Visible = True
    frmnaregelblok.TextBox212.Visible = True: frmnaregelblok.ComboBox212.Visible = True: frmnaregelblok.CheckBox212.Visible = True
    frmnaregelblok.TextBox213.Visible = True: frmnaregelblok.ComboBox213.Visible = True: frmnaregelblok.CheckBox213.Visible = True
    frmnaregelblok.TextBox214.Visible = True: frmnaregelblok.ComboBox214.Visible = True: frmnaregelblok.CheckBox214.Visible = True
    frmnaregelblok.TextBox215.Visible = True: frmnaregelblok.ComboBox215.Visible = True: frmnaregelblok.CheckBox215.Visible = True
    frmnaregelblok.TextBox216.Visible = True: frmnaregelblok.ComboBox216.Visible = True: frmnaregelblok.CheckBox216.Visible = True
    frmnaregelblok.TextBox217.Visible = True: frmnaregelblok.ComboBox217.Visible = True: frmnaregelblok.CheckBox217.Visible = True
    frmnaregelblok.TextBox218.Visible = True: frmnaregelblok.ComboBox218.Visible = True: frmnaregelblok.CheckBox218.Visible = True
    frmnaregelblok.TextBox219.Visible = True: frmnaregelblok.ComboBox219.Visible = True: frmnaregelblok.CheckBox219.Visible = True
    
End If
End Sub


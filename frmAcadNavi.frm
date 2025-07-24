VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAcadNavi 
   Caption         =   "Exporting Autocad to Navision"
   ClientHeight    =   84
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3600
   OleObjectBlob   =   "frmAcadNavi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAcadNavi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Activate()
Call bloklezen.blokl
Dim teknaam
teknaam = ThisDrawing.GetVariable("dwgname")    'teknaam = ThisDrawing.GetVariable("dwgname")

         Dim mystr As Variant
         Dim teknaam6
         Dim over
         mystr = Len(teknaam)
         over = mystr - 4 'aantal karakters
         teknaam6 = Left(teknaam, over)

TextBox41 = teknaam6
TextBox42 = ThisDrawing.GetVariable("dwgprefix")
Call But1
End Sub

Private Sub But1()
Call checkboxen.checkbox
Call printtekst
Unload Me
End Sub

Sub printtekst()
Dim q
Call varitel(q)
'MsgBox q
Dim bestandnaam
bestandnaam = frmAcadNavi.TextBox42 & frmAcadNavi.TextBox41 & ".txt"

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fs, f
Dim s1 As AcadSelectionSet
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile(bestandnaam, ForWriting, -2)
f.Close
    
    Dim xt
   'Define the text object
    
    Set f = fs.OpenTextFile(bestandnaam, ForAppending, -2) '.txt"
    If frmAcadNavi.TextBox1 <> "-" Then
    f.write frmAcadNavi.TextBox21 & ";" & frmAcadNavi.TextBox1 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox2 <> "-" Then
    f.write frmAcadNavi.TextBox22 & ";" & frmAcadNavi.TextBox2 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox3 <> "-" Then
    f.write frmAcadNavi.TextBox23 & ";" & frmAcadNavi.TextBox3 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox4 <> "-" Then
    f.write frmAcadNavi.TextBox24 & ";" & frmAcadNavi.TextBox4 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox5 <> "-" Then
    f.write frmAcadNavi.TextBox25 & ";" & frmAcadNavi.TextBox5 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox6 <> "-" Then
    f.write frmAcadNavi.TextBox26 & ";" & frmAcadNavi.TextBox6 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox7 <> "-" Then
    f.write frmAcadNavi.TextBox27 & ";" & frmAcadNavi.TextBox7 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox8 <> "-" Then
    f.write frmAcadNavi.TextBox28 & ";" & frmAcadNavi.TextBox8 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox9 <> "-" Then
    f.write frmAcadNavi.TextBox29 & ";" & frmAcadNavi.TextBox9 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox10 <> "-" Then
    f.write frmAcadNavi.TextBox30 & ";" & frmAcadNavi.TextBox10 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox11 <> "-" Then
    f.write frmAcadNavi.TextBox31 & ";" & frmAcadNavi.TextBox11 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox12 <> "-" Then
    f.write frmAcadNavi.TextBox32 & ";" & frmAcadNavi.TextBox12 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox13 <> "-" Then
    f.write frmAcadNavi.TextBox33 & ";" & frmAcadNavi.TextBox13 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox14 <> "-" Then
    f.write frmAcadNavi.TextBox34 & ";" & frmAcadNavi.TextBox14 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox15 <> "-" Then
    f.write frmAcadNavi.TextBox35 & ";" & frmAcadNavi.TextBox15 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox16 <> "-" Then
    f.write frmAcadNavi.TextBox36 & ";" & frmAcadNavi.TextBox16 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox17 <> "-" Then
    f.write frmAcadNavi.TextBox37 & ";" & frmAcadNavi.TextBox17 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox18 <> "-" Then
    f.write frmAcadNavi.TextBox38 & ";" & frmAcadNavi.TextBox18 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox19 <> "-" Then
    f.write frmAcadNavi.TextBox39 & ";" & frmAcadNavi.TextBox19 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox20 <> "-" Then f.write frmAcadNavi.TextBox40 & ";" & frmAcadNavi.TextBox20 & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
'    f.write Chr(13) + Chr(10)
    
    ' buislengtes
    If frmAcadNavi.TextBox44 <> "0" Then
    f.write frmAcadNavi.TextBox44 & ";" & "WTH-ZD-20*3,4-250 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox45 <> "0" Then
    f.write frmAcadNavi.TextBox45 & ";" & "WTH-ZD-20*3,4-165 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox46 <> "0" Then
    f.write frmAcadNavi.TextBox46 & ";" & "WTH-ZD-20*3,4-125 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox47 <> "0" Then
    f.write frmAcadNavi.TextBox47 & ";" & "WTH-ZD-20*3,4-105 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox48 <> "0" Then
    f.write frmAcadNavi.TextBox48 & ";" & "WTH-ZD-20*3,4-90 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox49 <> "0" Then
    f.write frmAcadNavi.TextBox49 & ";" & "WTH-ZD-20*3,4-75 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox50 <> "0" Then
    f.write frmAcadNavi.TextBox50 & ";" & "WTH-ZD-20*3,4-63 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox51 <> "0" Then
    f.write frmAcadNavi.TextBox51 & ";" & "WTH-ZD-20*3,4-50 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox52 <> "0" Then
    f.write frmAcadNavi.TextBox52 & ";" & "WTH-ZD-20*3,4-40 METER" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    'pe-rt
    If frmAcadNavi.TextBox101 <> "0" Then
    f.write frmAcadNavi.TextBox101 & ";" & "PE-RT 16/2 120" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox102 <> "0" Then
    f.write frmAcadNavi.TextBox102 & ";" & "PE-RT 16/2 90" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox103 <> "0" Then
    f.write frmAcadNavi.TextBox103 & ";" & "PE-RT 16/2 60" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox104 <> "0" Then
    f.write frmAcadNavi.TextBox104 & ";" & "PE-RT 14/2 90" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox105 <> "0" Then
    f.write frmAcadNavi.TextBox105 & ";" & "PE-RT 14/2 60" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    
    'bevestiging
    Dim gg
    Dim controle
    If frmAcadNavi.TextBox220 <> "0" Then
    f.write ((Val(frmAcadNavi.TextBox220)) * 2) & ";" & "Vlechtdraad - 100 stuks" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox221 <> "0" Then
            gg = ((Val(frmAcadNavi.TextBox221)) * 1.5)
            controle = InStr(1, gg, ".", vbBinaryCompare) 'staat er een komma in??
             If controle <> 0 Then gg = gg + 0.5
             f.write gg & ";" & "Kunststof slagbeugels 20 mm" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox222 <> "0" Then
            gg = ((Val(frmAcadNavi.TextBox222)) * 1.5)
            controle = InStr(1, gg, ".", vbBinaryCompare) 'staat er een komma in??
             If controle <> 0 Then gg = gg + 0.5
             f.write gg & ";" & "Kunststof slagbeugels 16 mm" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox223 <> "0" Then
    f.write ((Val(frmAcadNavi.TextBox223)) * 2) & ";" & "Ty-Rap - 100 stuks" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox224 <> "0" Then
            gg = ((Val(frmAcadNavi.TextBox224)) * 1.5)
            controle = InStr(1, gg, ".", vbBinaryCompare) 'staat er een komma in??
             If controle <> 0 Then gg = gg + 0.5
             f.write gg & ";" & "Isoclip schroefbeugel 16 mm" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.TextBox225 <> "0" Then
            gg = ((Val(frmAcadNavi.TextBox225)) * 1.5)
            controle = InStr(1, gg, ".", vbBinaryCompare) 'staat er een komma in??
             If controle <> 0 Then gg = gg + 0.5
             f.write gg & ";" & "Isoclip schroefbeugel 20 mm" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    
    If frmAcadNavi.TextBox226 <> "0" Then
    f.write q & ";" & "Varisoclip zwart" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
    f.write Chr(13) + Chr(10)
    End If
 
    If frmAcadNavi.TextBox228 <> "0" Then
            gg = ((Val(frmAcadNavi.TextBox228)) * 1.5)
            controle = InStr(1, gg, ".", vbBinaryCompare) 'staat er een komma in??
             If controle <> 0 Then gg = gg + 0.5
             f.write gg & ";" & "Beugels/Nagels" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    
    If frmAcadNavi.CheckBox3 = True Then
             f.write "0" & ";" & "IFD-Polystyreen" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.CheckBox4 = True Then
             f.write "0" & ";" & "Keg" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.CheckBox5 = True Then
             f.write "0" & ";" & "Montagestrip" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.CheckBox6 = True Then
             f.write "0" & ";" & "Noppenplaat" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    If frmAcadNavi.CheckBox7 = True Then
             f.write "0" & ";" & "Schietbeugels" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
             f.write Chr(13) + Chr(10)
    End If
    
 Dim element
 Dim SYMBOOL
 Dim G1: Dim G2: Dim G3: Dim G4: Dim G5: Dim G6
 Dim N1: Dim N2: Dim N3: Dim N4: Dim N5: Dim N6
 Dim ATTRIBUTEN
 Dim i
 Dim ATTRIBUUT
 Dim elcheck
 
     elcheck = 0 'check op naregelblokken
 For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If UCase(element.Name) = "NAREGELBLOK6" Then
      frmAcadNavi.CheckBox1.Value = True
      Set SYMBOOL = element
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
                If ATTRIBUUT.TagString = "GROEPSNUMMER1" Then G1 = ATTRIBUUT.textstring
                If ATTRIBUUT.TagString = "NAGEREGELD1" Then
                           If ATTRIBUUT.textstring = "NAGEREGELD" Then frmAcadNavi.ListBox2.AddItem (G1)
                End If
                If ATTRIBUUT.TagString = "GROEPSNUMMER2" Then G2 = ATTRIBUUT.textstring
                If ATTRIBUUT.TagString = "NAGEREGELD2" Then
                           If ATTRIBUUT.textstring = "NAGEREGELD" Then frmAcadNavi.ListBox2.AddItem (G2)
                End If
                If ATTRIBUUT.TagString = "GROEPSNUMMER3" Then G3 = ATTRIBUUT.textstring
                If ATTRIBUUT.TagString = "NAGEREGELD3" Then
                           If ATTRIBUUT.textstring = "NAGEREGELD" Then frmAcadNavi.ListBox2.AddItem (G3)
                End If
                If ATTRIBUUT.TagString = "GROEPSNUMMER4" Then G4 = ATTRIBUUT.textstring
                If ATTRIBUUT.TagString = "NAGEREGELD4" Then
                           If ATTRIBUUT.textstring = "NAGEREGELD" Then frmAcadNavi.ListBox2.AddItem (G4)
                End If
                If ATTRIBUUT.TagString = "GROEPSNUMMER5" Then G5 = ATTRIBUUT.textstring
                If ATTRIBUUT.TagString = "NAGEREGELD5" Then
                           If ATTRIBUUT.textstring = "NAGEREGELD" Then frmAcadNavi.ListBox2.AddItem (G5)
                End If
                If ATTRIBUUT.TagString = "GROEPSNUMMER6" Then G6 = ATTRIBUUT.textstring
                If ATTRIBUUT.TagString = "NAGEREGELD6" Then
                           If ATTRIBUUT.textstring = "NAGEREGELD" Then frmAcadNavi.ListBox2.AddItem (G6)
                End If
                
         Next i
       
        End If
      End If
      End If
      elcheck = 1
 Next element
 
    'If frmAcadNavi.CheckBox1.Value = True Then f.write "Nageregelde groepen" & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43
    f.Close

  If elcheck = 1 Then
 
    Dim teller: Dim j
    Dim textstring: Dim textstring1
    Dim kenter
    teller = frmAcadNavi.ListBox2.ListCount
    kenter = 0
    
    For j = 0 To teller - 1
        textstring = frmAcadNavi.ListBox2.List(j)
'''''        textstring1 = Left(textstring, 8)
'''''
'''''            If textstring1 = frmAcadNavi.TextBox227 Then
'''''                  kenter = 1
'''''                  Else
'''''                  kenter = 2
'''''            End If
'''''        frmAcadNavi.TextBox227 = textstring1
        
        Set f = fs.OpenTextFile(bestandnaam, ForAppending, -2)
            
'            If kenter = 2 Then
            f.write textstring & ";" & frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43 & ";" & frmAcadNavi.TextBox100
            f.write Chr(13) + Chr(10)
'            End If
        'f.write textstring & ";"
        f.Close
        kenter = 0
    Next j
''' Set f = fs.OpenTextFile(bestandnaam, ForAppending, -2)
'''  f.write frmAcadNavi.TextBox41 & ";" & frmAcadNavi.TextBox43
''' f.Close
End If
End Sub
Sub varitel(q)
'tot 400 groepen per tekening  varisoclips per 1400
Dim qwe
qwe = Val(frmAcadNavi.TextBox226)
q = qwe * 200
 If q <= 1400 Then q = 1
 If q > 1400 And q <= 2800 Then q = 2
 If q > 2800 And q <= 4200 Then q = 3
 If q > 4200 And q <= 5600 Then q = 4
 If q > 5600 And q <= 7000 Then q = 5
 If q > 7000 And q <= 8400 Then q = 6
 If q > 8400 And q <= 9800 Then q = 7
 If q > 9800 And q <= 11200 Then q = 8
 If q > 11200 And q <= 12600 Then q = 9
 If q > 12600 And q <= 14000 Then q = 10
 If q > 14000 And q <= 15400 Then q = 11
 If q > 15400 And q <= 16800 Then q = 12
 If q > 16800 And q <= 18200 Then q = 13
 If q > 18200 And q <= 19600 Then q = 14
 If q > 19600 And q <= 21000 Then q = 15
 If q > 21000 And q <= 22400 Then q = 16
 If q > 22400 And q <= 23800 Then q = 17
 If q > 23800 And q <= 25200 Then q = 18
 If q > 25200 And q <= 26600 Then q = 19
 If q > 26600 And q <= 28000 Then q = 20
 If q > 28000 And q <= 29400 Then q = 21
 If q > 29400 And q <= 30800 Then q = 22
 If q > 30800 And q <= 32200 Then q = 23
 If q > 32200 And q <= 33600 Then q = 24
 If q > 33600 And q <= 35000 Then q = 25
 If q > 35000 And q <= 36400 Then q = 26
 If q > 36400 And q <= 37800 Then q = 27
 If q > 37800 And q <= 39200 Then q = 28
 If q > 39200 And q <= 40600 Then q = 29
 If q > 40600 And q <= 42000 Then q = 30
 If q > 42000 And q <= 43400 Then q = 31
 If q > 43400 And q <= 44800 Then q = 32
 If q > 44800 And q <= 46200 Then q = 33
 If q > 46200 And q <= 47600 Then q = 34
 If q > 47600 And q <= 49000 Then q = 35
 If q > 49000 And q <= 50400 Then q = 36
 If q > 50400 And q <= 51800 Then q = 37
 If q > 51800 And q <= 53200 Then q = 38
 If q > 53200 And q <= 54600 Then q = 39
 If q > 54600 And q <= 56000 Then q = 40
 If q > 56000 And q <= 57400 Then q = 41
 If q > 57400 And q <= 58800 Then q = 42
 If q > 58800 And q <= 60200 Then q = 43
 If q > 60200 And q <= 61600 Then q = 44
 If q > 61600 And q <= 63000 Then q = 45
 If q > 63000 And q <= 64400 Then q = 46
 If q > 64400 And q <= 65800 Then q = 47
 If q > 65800 And q <= 67200 Then q = 48
 If q > 67200 And q <= 68600 Then q = 49
 If q > 68600 And q <= 70000 Then q = 50
 If q > 70000 And q <= 71400 Then q = 51
 If q > 71400 And q <= 72800 Then q = 52
 If q > 72800 And q <= 74200 Then q = 53
 If q > 74200 And q <= 75600 Then q = 54
 If q > 75600 And q <= 77000 Then q = 55
 If q > 77000 And q <= 78400 Then q = 56
 If q > 78400 And q <= 79800 Then q = 57
 If q > 79800 And q <= 81200 Then q = 58
End Sub

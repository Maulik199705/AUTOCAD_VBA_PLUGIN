VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmregelstaat 
   Caption         =   "Projectrealisatie werkelijke lengtes (Niet voor Flexfix.!!!)"
   ClientHeight    =   1170
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5292
   HelpContextID   =   5
   OleObjectBlob   =   "frmregelstaat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmregelstaat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProgressBar1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub UserForm_Activate()
Call kaderuit2  'schaal kader uitlezen
 TextBox20 = ThisDrawing.GetVariable("dwgprefix")
 Call extreemtek2
 
 'Kill ("c:\acad2002\Inregelstaat.txt")
 If TextBox14 > 30 Then Call importuserr21
 If TextBox14 < 31 Then TextBox13 = ThisDrawing.GetVariable("USERR2")
 
 ComboBox1.AddItem "2.5"
 ComboBox1.AddItem "2"
 ComboBox1.Text = ComboBox1.List(0)
'Call wandvv1
Call cmdLayers_Click1
'Call Checklayer2.Checklayer2
End Sub
Private Sub wandvv1()
    For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOK" Then
                Set SYMBOOL = element
                If SYMBOOL.HasAttributes Then
                    ATTRIBUTEN = SYMBOOL.GetAttributes
                    For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                         Set ATTRIBUUT = ATTRIBUTEN(i)
                         'If attribuut.TagString = "GROEPTEKST" Then GRP = attribuut.textstring
                         If ATTRIBUUT.TagString = "HOHAFSTAND" Then WH = ATTRIBUUT.textstring
                         
                    Next i
                End If
            End If
        End If
    Next element
    
   If WH = "Wandverwarming" Then

     MsgBox "Er is wandverwarming in de tekening aanwezig.!!!!" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) & _
            "Vul de hoogte van de wandverwarming in. (standaard staat ie op 2,5 meter)", vbExclamation ' & (Chr(13) & Chr(10)) & _
            '"Als je meerdere hoogte's heb vul dan de grootste waarde in, of een gemiddelde waarde.", vbExclamation
    frmregelstaat.Height = 134
    frmregelstaat.cmdLayers.top = 78
    frmregelstaat.cmdAfsluiten.top = 78
    frmregelstaat.Label39.Visible = True
    frmregelstaat.ComboBox1.Visible = True
   End If
End Sub
Private Sub kaderuit2()

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If UCase(element.Name) = "KADERLOGO" Or UCase(element.Name) = "LOGOTGH" Then
      Set SYMBOOL = element
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(i)
        If ATTRIBUUT.TagString = "OPDRACHTGEVER" Then ListBox4.AddItem ("OPDRACHTGEVER" & "#" & ATTRIBUUT.textstring)
        If ATTRIBUUT.TagString = "PLAATS" Then ListBox4.AddItem ("PLAATS" & "#" & ATTRIBUUT.textstring)
        If ATTRIBUUT.TagString = "PROJECTNAAM" Then ListBox4.AddItem ("PROJECTNAAM" & "#" & ATTRIBUUT.textstring)
        If ATTRIBUUT.TagString = "MONTAGEADRES" Then ListBox4.AddItem ("MONTAGEADRES" & "#" & ATTRIBUUT.textstring)
        If ATTRIBUUT.TagString = "MONTAGEPLAATS" Then ListBox4.AddItem ("MONTAGEPLAATS" & "#" & ATTRIBUUT.textstring)
        If ATTRIBUUT.TagString = "PROJECTNUMMER" Then ListBox4.AddItem ("PROJECTNUMMER" & "#" & ATTRIBUUT.textstring)
        If ATTRIBUUT.TagString = "BLAD" Then ListBox4.AddItem ("BLAD" & "#" & ATTRIBUUT.textstring)
       
        Next i
       End If
     End If
   End If
  Next element
End Sub
Sub exportuserr2()
TT = ThisDrawing.GetVariable("USERR2")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fg, fh
'Dim s1 As AcadSelectionSet
Set fg = CreateObject("Scripting.FileSystemObject")

Set fh = fg.OpenTextFile("c:\acad2002\userr2.txt", ForWriting, -2)
    fh.write TT
    fh.Close
End Sub
Sub importuserr21()
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fl, m, TS
Set fl = CreateObject("Scripting.FileSystemObject")
Set m = fl.OpenTextFile("c:\acad2002\userr2.txt", ForReading, False)
    TS = m.ReadLine
    TextBox13 = TS
m.Close 'sluiten van tekstbestand
End Sub
Private Sub TextBox1_Change()
If TextBox1 <> "" Then CommandButton1.Locked = False
End Sub
Private Sub cmdLayers_Click1()

ListBox1.Clear
ListBox2.Clear
Update
On Error Resume Next
  
 Dim cirkel As Object
 Dim element As Object
 Dim Lengte As Double
 Dim laagobj As Object
 
 Dim minaantal As Integer
 Dim maxaantal As Integer
 Dim i As Integer
 i = 0
 minaantal = 0
 maxaantal = ThisDrawing.Layers.Count
 For Each laagobj In ThisDrawing.Layers
    i = i + 1
    ProgressBar1.Min = minaantal
    ProgressBar1.Max = maxaantal
    ProgressBar1.Value = i
 
      mystr = Left(laagobj.Name, 5)
      mystcontr = Len(laagobj.Name)
      
      If mystr = "GROEP" Then mystr = "groep"


      If mystr = "groep" And mystcontr = 11 Then

'''''7-2-2007
'''''      If Not mystr = "groep" Then
'''''      GoTo wand
'''''      Else
        
        
        For Each element In ThisDrawing.ModelSpace
          If element.layer = laagobj.Name Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
          End If 'elementlayer
            
        Next element
              
              
              For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.layer = laagobj.Name Then
                If cirkel.EntityName = "AcDbCircle" Then
                Z = Z + 1
                'wvaanwezig = " (WV)"
                
                zlengte = (Z * (ComboBox1.Text * 100)) + 100    '
                End If
                End If
              Next cirkel
         
                
                
'''''7-2-2007
'''''wand:
'''''        If mystr = "WAND_" Then mystr = "wand_"
'''''        If mystr = "wand_" Then
'''''             For Each element In ThisDrawing.ModelSpace
'''''             If element.Layer = laagobj.Name Then
'''''                'BEREKENEN TOTALE LENGTE
'''''                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
'''''                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
'''''
'''''             End If 'elementlayer
'''''             Next element
'''''
'''''                For Each cirkel In ThisDrawing.ModelSpace
'''''                If cirkel.Layer = laagobj.Name Then
'''''                If cirkel.EntityName = "AcDbCircle" Then
'''''                z = z + 1
'''''                wvaanwezig = " (WV)"
'''''                End If
'''''                End If
'''''                Next cirkel
'''''                zlengte = (z * (ComboBox1.Text * 100)) + 100
'''''                End If '2e mystr


                
    Lengte = (Lengte * TextBox19) + zlengte
    'Lengte = Lengte + zlengte
    Lengte = Lengte / 100
    Lengte = Round(Lengte, 1)   'deze
    'Lengte = Fix(Lengte)  'deze
    'Lengte = Lengte + 1   'deze
    'totalrollen = totalrollen + Lengte
  TextBox22.Value = Val(TextBox22.Value) + Lengte

  If mystr = "groep" Or mystr = "GROEP" Then
  t = Split(laagobj.Name, " ")
  
  
  d = t(1) & "#" & Lengte 'laagobj.Name
  'D = laagobj.Name & "#" & Lengte 'laagobj.Name
  ListBox1.AddItem (d)
  End If

    Lengte = 0 'Lengte leeggooien voordat de volgende groep wordt gemeten
    zlengte = 0
    Z = 0
'''''    wvaanwezig = ""
'''''7-2-2007'''''End If  'end if  mystr

   End If 'mystr + mystrcontr
 Next laagobj
 ProgressBar1.Value = minaantal
 Update
 

 Cmdprint.Enabled = True

   'lijst rangschikken
  Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For i = 0 To ListBox1.ListCount - 1
    textstring2 = ListBox1.List(i)
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
  ListBox1.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then ListBox1.AddItem Veld(x)
  Next x
 
'''' ONDERSTAANDE REGELS LATEN STAAN
'''' Dim textstring3
'''' teller = ListBox1.ListCount
'''' For j = 0 To teller - 1
''''    textstring2 = ListBox1.List(j)
''''    textstring3 = Split(textstring2, ("#"))
''''    Dim a
''''    a = textstring3(0)
''''    Dim b
''''    b = textstring3(1)
''''    'MsgBox textstring3(0)
''''    ListBox2.AddItem (a)
''''    ListBox3.AddItem (b)
''''
'''' Next j

 Call Cmdprint_Click1
End Sub
Private Sub Cmdprint_Click1()
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fs, f
Dim s1 As AcadSelectionSet
Set fs = CreateObject("Scripting.FileSystemObject")
teknaam = ThisDrawing.GetVariable("dwgname")
pad = ThisDrawing.GetVariable("dwgprefix")
usernaam = ThisDrawing.GetVariable("loginname")
Dim MyDate
MyDate = DateValue(Date)    ' Return a date.


Dim xlsnaam
xlsnaam = TextBox20 & TextBox17 & ".xls"
Dim textstring
Dim textstring2



Set f = fs.OpenTextFile(xlsnaam, ForWriting, -2) ' "c:\acad2002\Inregelstaat.xls"
    teller3 = ListBox4.ListCount
    For k = 0 To teller3 - 1
       'Define the text object
        textstring5 = ListBox4.List(k)
        textstring6 = Split(textstring5, ("#"))
        f.write textstring6(0)
        f.write Chr(9)
        f.write textstring6(1)
        f.write Chr(13) + Chr(10)
     Next k
     f.write Chr(13) + Chr(10)
f.Close


    teller = ListBox1.ListCount
    For i = 0 To teller - 1
       'Define the text object
        textstring = ListBox1.List(i)
        textstring2 = Split(textstring, ("#"))
        Set f = fs.OpenTextFile(xlsnaam, ForAppending, -2)
        f.write textstring2(0)
        f.write Chr(9)
        f.write textstring2(1)
        f.write Chr(13) + Chr(10)
        f.Close
    Next i
    
    Set f = fs.OpenTextFile(xlsnaam, ForAppending, -2)
        f.write Chr(13) + Chr(10)
        f.write Chr(9)
        f.write TextBox22.Value
    f.Close
        

    If TextBox14 > 30 Then ThisDrawing.Close
    If TextBox14 > 30 Then
    bestandnaam3 = TextBox16 & TextBox17 & "-meten" & ".dwg"
    Kill (bestandnaam3)
End If

Unload Me
End Sub
Private Sub cmdAfsluiten_Click()

If TextBox18 <> "" Then
 checkopexport = Split(TextBox18, ".")
 checkopexport1 = Right(checkopexport(0), 5)
 If checkopexport1 = "meten" Then ThisDrawing.Close
 bestandnaam3 = TextBox16 & TextBox17 & "-meten" & ".dwg"
 Kill (bestandnaam3)
End If

Unload Me
End Sub
Sub extreemtek2()

TextBox14 = Clear
For Each layerObj In ThisDrawing.Layers
     If Left(layerObj.Name, 5) = "groep" Then TextBox14 = (Val(TextBox14)) + 1  'aantal groepen
Next 'layerobj
If TextBox14 = "" Then TextBox14 = 0

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "KADERLOGO" Or UCase(element2.Name) = "LOGOTGH" Then
      Set SYMBOOL = element2
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For k = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(k)
        If ATTRIBUUT.TagString = "SCHAAL" Then schaalfaktor = ATTRIBUUT.textstring
        Next k
       
      End If
      End If
      End If
Next element2

vartab = ThisDrawing.GetVariable("EXTMAX")

' A3 T/M A0+
If vartab(0) >= 2045 And vartab(0) <= 2055 Or vartab(0) >= 2915 And vartab(0) <= 2925 _
    Or vartab(0) >= 4175 And vartab(0) <= 4185 Or vartab(0) >= 5890 And vartab(0) <= 5900 _
    Or vartab(0) >= 7840 And vartab(0) <= 7850 Then
    maxwaarde = 1
Else
    maxwaarde = 2
End If
If vartab(0) >= 16715 And vartab(0) <= 16725 Or vartab(0) >= 23575 And vartab(0) <= 23585 _
    Or vartab(0) >= 31375 And vartab(0) <= 31385 Then
    maxwaarde = 4
End If
Update

If schaalfaktor = "1:50" And maxwaarde = 1 Then scaal = 1
If schaalfaktor = "1:100" And maxwaarde = 1 Then scaal = 2
If schaalfaktor = "1:200" And maxwaarde = 1 Then scaal = 4

If schaalfaktor = "1:50" And maxwaarde = 2 Then scaal = 0.5
If schaalfaktor = "1:100" And maxwaarde = 2 Then scaal = 1
If schaalfaktor = "1:200" And maxwaarde = 2 Then scaal = 2

If schaalfaktor = "1:200" And maxwaarde = 4 Then scaal = 1
TextBox19.Value = scaal

teknaam = ThisDrawing.GetVariable("dwgname")    'teknaam = ThisDrawing.GetVariable("dwgname")
TextBox15 = teknaam
TextBox16 = "c:\acad2002\"  'pad = ThisDrawing.GetVariable("dwgprefix")
teknaam1 = Split(teknaam, ("."))                'teknaam1 = Split(teknaam, ("."))

         Dim mystr As Variant
         mystr = Len(teknaam)
         over = mystr - 4 'aantal karakters
         teknaam6 = Left(teknaam, over)

'TextBox18 = teknaam1(0)
TextBox17 = teknaam6


'aantal groepen groter dan 30 ????
If TextBox14 > 30 Then
Call exportuserr2
frmregelstaat.Hide

 
'bestandnaam2 = pad & teknaam1(0) & "-export" & ".dwg"
bestandnaam2 = TextBox16 & TextBox17 & "-meten" & ".dwg"
TextBox18 = bestandnaam2

'Kill (bestandnaam)
'Call zoekblad(bladnummer)

ThisDrawing.SendCommand "-layer" & vbCr & "unlock" & vbCr & "*" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "set" & vbCr & "0" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "*" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "groep*" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "gt" & vbCr & vbCr

ThisDrawing.SendCommand "_copyclip" & vbCr & "all" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "*" & vbCr & vbCr

templatenaam = ThisDrawing.GetVariable("acadver")
If templatenaam = "15.06s (LMS Tech)" Then ThisDrawing.Application.Documents.Add ("acad2000.dwt")
If templatenaam >= "16.0s (LMS Tech)" Then ThisDrawing.Application.Documents.Add ("autocad.dwt")

ThisDrawing.SendCommand "_pasteclip" & vbCr & "0,0" & vbCr

ThisDrawing.SaveAs (bestandnaam2)

End If

ZoomExtents


End Sub

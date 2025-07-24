VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInregelstaat 
   Caption         =   "Inregelstaat (Niet voor Flexfix.!!!)"
   ClientHeight    =   1665
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3864
   HelpContextID   =   5
   OleObjectBlob   =   "frmInregelstaat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInregelstaat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'24-01-2007 Inregelstaat genereren
'G.C.Haak
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
        ByVal hWnd As LongPtr) As Long

    Private Declare PtrSafe Function GetMenuItemCount Lib "user32" ( _
        ByVal hMenu As LongPtr) As Long

    Private Declare PtrSafe Function GetSystemMenu Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr

    Private Declare PtrSafe Function RemoveMenu Lib "user32" ( _
        ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function DrawMenuBar Lib "user32" ( _
        ByVal hWnd As Long) As Long

    Private Declare Function GetMenuItemCount Lib "user32" ( _
        ByVal hMenu As Long) As Long

    Private Declare Function GetSystemMenu Lib "user32" ( _
        ByVal hWnd As Long, ByVal bRevert As Long) As Long

    Private Declare Function RemoveMenu Lib "user32" ( _
        ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000
Private Sub UserForm_Initialize()
#If VBA7 Then
    Dim lngHwnd As LongPtr
    Dim lngMenu As LongPtr
#Else
    Dim lngHwnd As Long
    Dim lngMenu As Long
#End If
  Dim lngCnt As Long
  lngHwnd = FindWindow(vbNullString, Me.Caption)
  lngMenu = GetSystemMenu(lngHwnd, 0)
  If lngMenu Then
    lngCnt = GetMenuItemCount(lngMenu)
    Call RemoveMenu(lngMenu, lngCnt - 1, _
    MF_REMOVE Or MF_BYPOSITION)
    Call DrawMenuBar(lngHwnd)
  End If
 
 On Error Resume Next
 Call kaderuit  'schaal kader uitlezen
 TextBox20 = ThisDrawing.GetVariable("dwgprefix")
 Call extreemtek
 
 'Kill ("c:\acad2002\Inregelstaat.txt")
 If TextBox14 > 30 Then Call importuserr2
 If TextBox14 < 31 Then TextBox13 = ThisDrawing.GetVariable("USERR2")
 
 ComboBox1.AddItem "2.5"
 ComboBox1.AddItem "2"
 ComboBox1.Text = ComboBox1.List(0)
Call wandvv
'Call Checklayer2.Checklayer2
End Sub
Private Sub wandvv()
    For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOK" Then
                Set symbool = element
                If symbool.HasAttributes Then
                    attributen = symbool.GetAttributes
                    For I = LBound(attributen) To UBound(attributen)
                         Set attribuut = attributen(I)
                         'If attribuut.TagString = "GROEPTEKST" Then GRP = attribuut.textstring
                         If attribuut.TagString = "HOHAFSTAND" Then WH = attribuut.textstring
                         
                    Next I
                End If
            End If
        End If
    Next element
    
   If WH = "Wandverwarming" Then
   
     MsgBox "Er is wandverwarming in de tekening aanwezig.!!!!" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) & _
            "Vul de hoogte van de wandverwarming in. (standaard staat ie op 2,5 meter)", vbExclamation ' & (Chr(13) & Chr(10)) & _
            '"Als je meerdere hoogte's heb vul dan de grootste waarde in, of een gemiddelde waarde.", vbExclamation
    frmInregelstaat.Height = 134
    frmInregelstaat.cmdLayers.top = 78
    frmInregelstaat.cmdAfsluiten.top = 78
    frmInregelstaat.Label39.Visible = True
    frmInregelstaat.ComboBox1.Visible = True
   End If
End Sub
Private Sub kaderuit()

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If UCase(element.Name) = "KADERLOGO" Or UCase(element.Name) = "LOGOTGH" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" Then ListBox4.AddItem ("OPDRACHTGEVER" & "#" & attribuut.textstring)
        If attribuut.TagString = "PLAATS" Then ListBox4.AddItem ("PLAATS" & "#" & attribuut.textstring)
        If attribuut.TagString = "PROJECTNAAM" Then ListBox4.AddItem ("PROJECTNAAM" & "#" & attribuut.textstring)
        If attribuut.TagString = "MONTAGEADRES" Then ListBox4.AddItem ("MONTAGEADRES" & "#" & attribuut.textstring)
        If attribuut.TagString = "MONTAGEPLAATS" Then ListBox4.AddItem ("MONTAGEPLAATS" & "#" & attribuut.textstring)
        If attribuut.TagString = "PROJECTNUMMER" Then ListBox4.AddItem ("PROJECTNUMMER" & "#" & attribuut.textstring)
        If attribuut.TagString = "BLAD" Then ListBox4.AddItem ("BLAD" & "#" & attribuut.textstring)
       
        Next I
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
Sub importuserr2()
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
Private Sub cmdLayers_Click()

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
 Dim I As Integer
 I = 0
 minaantal = 0
 maxaantal = ThisDrawing.Layers.Count
 For Each laagobj In ThisDrawing.Layers
    I = I + 1
    ProgressBar1.Min = minaantal
    ProgressBar1.Max = maxaantal
    ProgressBar1.Value = I
 
      mystr = Left(laagobj.Name, 5)
      mystcontr = Len(laagobj.Name)
      
      If mystr = "GROEP" Then mystr = "groep"


      If mystr = "groep" And mystcontr = 11 Then

'''''7-2-2007
'''''      If Not mystr = "groep" Then
'''''      GoTo wand
'''''      Else
        
        
        For Each element In ThisDrawing.ModelSpace
          If element.Layer = laagobj.Name Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
          End If 'elementlayer
            
        Next element
              
              
              For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.Layer = laagobj.Name Then
                If cirkel.EntityName = "AcDbCircle" Then
                z = z + 1
                'wvaanwezig = " (WV)"
                
                zlengte = (z * (ComboBox1.Text * 100)) + 100    '
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
    Lengte = Fix(Lengte)
    Lengte = Lengte + 1
    'totalrollen = totalrollen + Lengte
  

  If mystr = "groep" Or mystr = "GROEP" Then
  T = Split(laagobj.Name, " ")
  
  
  D = T(1) & "#" & Lengte 'laagobj.Name
  'D = laagobj.Name & "#" & Lengte 'laagobj.Name
  ListBox1.AddItem (D)
  End If

    Lengte = 0 'Lengte leeggooien voordat de volgende groep wordt gemeten
    zlengte = 0
    z = 0
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
  
    For I = 0 To ListBox1.ListCount - 1
    textstring2 = ListBox1.List(I)
    Veld(I) = textstring2
   
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
    
  Next I
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

 Call Cmdprint_Click
End Sub
Private Sub Cmdprint_Click()
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
    For I = 0 To teller - 1
       'Define the text object
        textstring = ListBox1.List(I)
        textstring2 = Split(textstring, ("#"))
        Set f = fs.OpenTextFile(xlsnaam, ForAppending, -2)
        f.write textstring2(0)
        f.write Chr(9)
        f.write textstring2(1)
        f.write Chr(13) + Chr(10)
        f.Close
    Next I

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
Sub extreemtek()

TextBox14 = Clear
For Each layerObj In ThisDrawing.Layers
     If Left(layerObj.Name, 5) = "groep" Then TextBox14 = (Val(TextBox14)) + 1  'aantal groepen
Next 'layerobj
If TextBox14 = "" Then TextBox14 = 0

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "KADERLOGO" Or UCase(element2.Name) = "LOGOTGH" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For k = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(k)
        If attribuut.TagString = "SCHAAL" Then schaalfaktor = attribuut.textstring
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
frmInregelstaat.Hide

 
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

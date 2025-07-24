VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportXLS 
   Caption         =   "Rollengte exporteren naar Excel"
   ClientHeight    =   8700.001
   ClientLeft      =   48
   ClientTop       =   540
   ClientWidth     =   11760
   HelpContextID   =   5
   OleObjectBlob   =   "frmExportXLS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'26-01-2004 export werk lengte's naar xls-bestand
'M.Bosch en G.C.Haak
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
frmExportXLS.Width = 285
frmExportXLS.Height = 88
On Error Resume Next

 'Kill ("c:\acad2002\layerlijst.txt")
Call exptek
End Sub
Sub exptek()

Call zoekblad(bladnummer)
TextBox14 = bladnummer

 If TextBox14 = "" Then
     Exit Sub
 End If


TextBox13 = Clear
For Each layerObj In ThisDrawing.Layers
     If Left(layerObj.Name, 5) = "groep" Then TextBox13 = (Val(TextBox13)) + 1  'aantal groepen
Next 'layerobj


For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "KADERLOGO" Or UCase(element2.Name) = "LOGOTGH" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For k = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(k)
        'If attribuut.TagString = "FORMAAT" Then FRMNO = attribuut.TextString
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
TextBox19 = scaal


teknaam = ThisDrawing.GetVariable("dwgname")    'teknaam = ThisDrawing.GetVariable("dwgname")
TextBox16 = teknaam
TextBox17 = ThisDrawing.GetVariable("dwgprefix")  'pad = ThisDrawing.GetVariable("dwgprefix")
teknaam1 = Split(teknaam, ("."))                'teknaam1 = Split(teknaam, ("."))

         Dim mystr As Variant
         mystr = Len(teknaam)
         over = mystr - 4 'aantal karakters
         teknaam6 = Left(teknaam, over)

'TextBox18 = teknaam1(0)
TextBox18 = teknaam6


'aantal groepen groter dan 30 ????
If TextBox13 > 30 Then
frmExportXLS.Hide


'bestandnaam2 = pad & teknaam1(0) & "-export" & ".dwg"
bestandnaam2 = TextBox17 & TextBox18 & "-export" & ".dwg"
TextBox20 = bestandnaam2

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
'If templatenaam = "16.0s (LMS Tech)" Then ThisDrawing.Application.Documents.Add ("acad2004.dwt")
If templatenaam >= "16.0s (LMS Tech)" Then ThisDrawing.Application.Documents.Add ("acad2004.dwt")

ThisDrawing.SendCommand "_pasteclip" & vbCr & "0,0" & vbCr

ThisDrawing.SaveAs (bestandnaam2)
'frmExportXLS.Show
End If

End Sub
Sub zoekblad(bladnummer)
For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "KADERLOGO" Or UCase(element2.Name) = "LOGOTGH" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "BLAD" Then bladnummer = attribuut.textstring
        Next I
       
      End If
      End If
      End If
  Next element2
 Update
End Sub
Private Sub CommandButton2_Click()
'voorbeeld groep 1.01 = 17.51 meter. --> wordt een rol van 40 meter
'waar je op scheidt valt weg
'trimst(0)-> groep 1.01
'trimst(1) -> 17.51 meter. --> wordt een rol van 40 meter
'cmdAfsluiten.Enabled = False
If TextBox14 = "" Then
   MsgBox "Één van de onderstaande fouten is geconstateerd:" & Chr(13) & Chr(10) & "1 - Het kaderlogo is niet aanwezig in de tekening," & Chr(13) & "2 - Of het bladnummer is niet ingevuld," & Chr(13) & "3 - Of er zijn 2 Kaderlogo's aanwezig.", vbExclamation, "Foutmelding"
   Unload Me
   Exit Sub
End If
 

ListBox1.Clear
ListBox3.Clear
TextBox12 = ""

Call cmdLayers_Click

'teknaam = TextBox16
'pad = ThisDrawing.GetVariable("dwgprefix")
teknaam1 = TextBox18
'Split(teknaam, ("."))

bestandnaam = TextBox17 & TextBox18 & "-lh" & ".xls"
'teknaam1 (0) & "-lh" & ".xls"


unitstrim1 = 1
 
teller = ListBox1.ListCount
For I = 0 To teller - 1
   'Define the text object
    textstring = ListBox1.List(I)
    trimst = Split(textstring, ("=")) 'groepsnaam en de rest scheiden
    trimst1 = Split(trimst(0), (" "))
    trimst2 = Split(trimst(1), (" "))
    
       
    unitstrim2 = Split(trimst1(1), ("."))
    If (unitstrim2(0) > unitstrim1) Then
      trimst3 = " - "
      ListBox3.AddItem (trimst3)
    End If
    
    trimst3 = "[" & TextBox14 & "] " & trimst(0) & "-" & (trimst2(1))
'bladnummer
    ListBox3.AddItem (trimst3) 'groepsnaam invullen
    unitstrim1 = unitstrim2(0)
Next I

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fs, f
Dim s1 As AcadSelectionSet
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile(bestandnaam, ForWriting, -2)
f.Close

teller = ListBox3.ListCount
For I = 0 To teller - 1
   'Define the text object
    xtstring = ListBox3.List(I)
    xtstring2 = Split(xtstring, ("-"))
    Set f = fs.OpenTextFile(bestandnaam, ForAppending, -2) '.xls"
    f.write xtstring2(0)
    f.write Chr(9)
    f.write xtstring2(1)
    'f.write Chr(9)
    'f.write xtstring2(2)
    f.write Chr(13) + Chr(10)
    f.Close
Next I


If TextBox13 > 30 Then ThisDrawing.Close
 If TextBox13 > 30 Then
  bestandnaam3 = TextBox17 & TextBox18 & "-export" & ".dwg"
  Kill (bestandnaam3)
 End If
Unload Me
End Sub
Sub cmdLayers_Click()

ListBox1.Clear
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
      If mystr = "GROEP" Then mystr = "groep"
      If mystr = "groep" Then
        For Each element In ThisDrawing.ModelSpace
          If element.Layer = laagobj.Name Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.Length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
             End If 'elementlayer
             Next element
             
              For Each cirkel In ThisDrawing.ModelSpace
                If cirkel.Layer = laagobj.Name Then
                If cirkel.EntityName = "AcDbCircle" Then
                
                
                For Each element In ThisDrawing.ModelSpace
                    If element.ObjectName = "AcDbBlockReference" Then
                    If UCase(element.Name) = "GROEPTEKSTBLOKNEW" Then
                    Set symbool = element
                    If symbool.HasAttributes Then
                     attributen = symbool.GetAttributes
                     For I = LBound(attributen) To UBound(attributen)
                         Set attribuut = attributen(I)
                         If attribuut.TagString = "GROEPTEKST" Then gpt = attribuut.textstring
                         If gpt = laagobj.Name Then
                          If attribuut.TagString = "WANDHOOGTE" Then wdh = attribuut.textstring
                         End If 'gpt
                         Next I
                        End If
                      End If
                   End If
                 Next element
                      
                wdh1 = Split(wdh, " ")
                z = z + 1
                 zlengte = (z * (wdh1(0) * 100)) + 100
                wvaanwezig = " (WV)"
                
                End If
                End If
              Next cirkel
               
              '  zlengte = zlengte - Lengte
        '  End If '2e mystr
                
    Lengte = (Lengte * TextBox19) + zlengte
    Lengte = Lengte / 100
    'Lengte = Lengte * scaal
    Lengte = Round(Lengte, 0)
    totalrollen = totalrollen + Lengte
 
    If mystr = "groep" Then
    s = " = "
    Else
    s = "  = "
    End If
    
    
    zw = Len(laagobj.Name)
    
    If zw > 9 And (mystr = "groep" Or mystr = "GROEP") Then
    
    D = laagobj.Name & s & Lengte & " meter. "
    D = D & " " & wvaanwezig
    ListBox1.AddItem (D)
    End If
    'End If
    
     
    Lengte = 0 'Lengte leeggooien voordat de volgende groep wordt gemeten
    zlengte = 0
    z = 0
    wvaanwezig = ""
  End If  'end if  mystr
 Next laagobj
 ProgressBar1.Value = minaantal
 'Update
  

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
 TextBox12 = totalrollen
 
 End Sub

Private Sub cmdAfsluiten_Click()
'checkopexport = Right(TextBox20, 5)
If TextBox20 <> "" Then
 checkopexport = Split(TextBox20, ".")
 checkopexport1 = Right(checkopexport(0), 6)
 If checkopexport1 = "export" Then ThisDrawing.Close
 bestandnaam3 = TextBox17 & TextBox18 & "-export" & ".dwg"
 Kill (bestandnaam3)
End If

Unload Me
End Sub

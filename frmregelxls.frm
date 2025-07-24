VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmregelxls 
   Caption         =   "Inregelxls"
   ClientHeight    =   6660
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   15084
   HelpContextID   =   5
   OleObjectBlob   =   "frmregelxls.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmregelxls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' G:\EXCEL\Calculatiesheet!\inregelen\Inregelstaat t.b.v. tekenkamer\digital.xlt
' vba- tool om inregelstaten te genereren
' maakt gebruik van excel
' G.C.Haak

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As LongPtr) As Long
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000


Private Sub UserForm_Initialize()
Dim lngHwnd As LongPtr
  Dim lngMenu As LongPtr
  Dim lngCnt As Long
  lngHwnd = FindWindow(vbNullString, Me.Caption)
  lngMenu = GetSystemMenu(lngHwnd, 0)
  If lngMenu Then
    lngCnt = GetMenuItemCount(lngMenu)
    Call RemoveMenu(lngMenu, lngCnt - 1, _
    MF_REMOVE Or MF_BYPOSITION)
    Call DrawMenuBar(lngHwnd)
  End If



 TextBox20 = ThisDrawing.GetVariable("dwgprefix")
 Call extreemtek2
 Call kaderuit2  'schaal kader uitlezen        '   29-04
 Call unitlogoos_uitlezen
 
 'If TextBox14 > 30 Then Call importuserr21
 'If TextBox14 < 31 Then TextBox13 = ThisDrawing.GetVariable("USERR2")
'
 frmregelxls.Height = 140
 frmregelxls.Width = 270
 
        frmregelxls.ComboBox2.AddItem ("10")
        frmregelxls.ComboBox2.AddItem ("12")
        frmregelxls.ComboBox2.AddItem ("14")
        frmregelxls.ComboBox2.AddItem ("15")
        frmregelxls.ComboBox2.AddItem ("16")
        frmregelxls.ComboBox2.AddItem ("18")
        frmregelxls.ComboBox2.AddItem ("20")
        frmregelxls.ComboBox2.AddItem ("21")
        frmregelxls.ComboBox2.AddItem ("22")
        frmregelxls.ComboBox2.AddItem ("24")
        frmregelxls.ComboBox2.ListIndex = 6
        

 ComboBox1.AddItem "2.5"
 ComboBox1.AddItem "2"
 ComboBox1.Text = ComboBox1.List(0)
 cmdLayers.SetFocus

End Sub
Private Sub kaderuit2()





Dim element1 As Object
For Each element1 In ThisDrawing.ModelSpace
      If element1.ObjectName = "AcDbBlockReference" Then
         If element1.Name = "Kaderlogo" Or element1.Name = "logotgh" Then a = a + 1
      End If
Next element1
      
If a > 1 Then

        For Each element2 In ThisDrawing.ModelSpace
              If element2.ObjectName = "AcDbBlockReference" Then
               If element2.Name = "Kaderlogo" Or element2.Name = "logotgh" Then
               Set SYMBOOL = element2
                If SYMBOOL.HasAttributes Then
                ATTRIBUTEN = SYMBOOL.GetAttributes
                For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                Set ATTRIBUUT = ATTRIBUTEN(i)
                If ATTRIBUUT.TagString = "OPDRACHTGEVER" And ATTRIBUUT.textstring = "" Then element2.Erase
                Next i
               End If
             End If
           End If
          Next element2

End If

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If UCase(element.Name) = "KADERLOGO" Or UCase(element.Name) = "LOGOTGH" Then
      Set SYMBOOL = element
        If SYMBOOL.HasAttributes Then
        ATTRIBUTEN = SYMBOOL.GetAttributes
        For j = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
        Set ATTRIBUUT = ATTRIBUTEN(j)
            If ATTRIBUUT.TagString = "OPDRACHTGEVER" Then ListBox4.AddItem ("OPDRACHTGEVER" & "#" & ATTRIBUUT.textstring)
            If ATTRIBUUT.TagString = "PLAATS" Then ListBox4.AddItem ("PLAATS" & "#" & ATTRIBUUT.textstring)
            If ATTRIBUUT.TagString = "PROJECTNAAM" Then ListBox4.AddItem ("PROJECTNAAM" & "#" & ATTRIBUUT.textstring)
            If ATTRIBUUT.TagString = "MONTAGEADRES" Then ListBox4.AddItem ("MONTAGEADRES" & "#" & ATTRIBUUT.textstring)
            If ATTRIBUUT.TagString = "MONTAGEPLAATS" Then ListBox4.AddItem ("MONTAGEPLAATS" & "#" & ATTRIBUUT.textstring)
            If ATTRIBUUT.TagString = "PROJECTNUMMER" Then                                             '   29-04
               ListBox4.AddItem ("PROJECTNUMMER" & "#" & TextBox17)   ' & ATTRIBUUT.textstring)        '   29-04
               bewaar = ATTRIBUUT.textstring                                                            '   29-04
            End If                                                                                      '   29-04
            If ATTRIBUUT.TagString = "BLAD" Then ListBox4.AddItem ("BLAD" & "#" & ATTRIBUUT.textstring)
       
        Next j
       End If
     End If
   End If
  Next element

  ListBox4.AddItem (bewaar)   '   29-04
End Sub
Sub extreemtek2()

'''TextBox14 = Clear
'''For Each layerObj In ThisDrawing.Layers
'''     If Left(layerObj.Name, 5) = "groep" Then TextBox14 = (Val(TextBox14)) + 1  'aantal groepen
'''Next 'layerobj
'''If TextBox14 = "" Then TextBox14 = 0

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

TextBox17 = teknaam6

'aantal groepen groter dan 30 ????
''''''''''If TextBox14 > 30 Then
''''''''''Call exportuserr2
''''''''''frmregelxls.Hide
''''''''''
''''''''''bestandnaam2 = TextBox16 & TextBox17 & "-meten" & ".dwg"
''''''''''TextBox18 = bestandnaam2
''''''''''
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "unlock" & vbCr & "*" & vbCr & vbCr
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "set" & vbCr & "0" & vbCr & vbCr
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "*" & vbCr & vbCr
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "groep*" & vbCr & vbCr
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "gt" & vbCr & vbCr
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "bloklogo" & vbCr & vbCr
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "3" & vbCr & vbCr
''''''''''
''''''''''ThisDrawing.SendCommand "_copyclip" & vbCr & "all" & vbCr & vbCr
''''''''''ThisDrawing.SendCommand "-layer" & vbCr & "Thaw" & vbCr & "*" & vbCr & vbCr
''''''''''ThisDrawing.Application.Documents.Add ("autocad.dwt")
''''''''''ThisDrawing.SendCommand "_pasteclip" & vbCr & "0,0" & vbCr
''''''''''
''''''''''ThisDrawing.SaveAs (bestandnaam2)

''''''''''End If

ZoomExtents
End Sub
Private Sub unitlogoos_uitlezen()
For Each element10 In ThisDrawing.ModelSpace
      If element10.ObjectName = "AcDbBlockReference" Then
      If element10.Name = "Mat_spe_ZD" Or element10.Name = "Mat_spe_PE" Or element10.Name = "Mat_spe_PE800" _
      Or element10.Name = "Mat_spe_ALU" Or element10.Name = "Mat_spe_ZDringleiding" Or element10.Name = "Mat_spe_PEringleiding" Or _
      element10.Name = "Mat_spe_ALUringleiding" Or element10.Name = "Mat_spe_FLEX" Or element10.Name = "Mat_spe_ZD_1627" Or _
      element10.Name = "Mat_spe_FLEX_Aankoppel" Or element10.Name = "Mat_spe_ZD_1627500" Then
        Set SYMBOOL = element10
        If SYMBOOL.HasAttributes Then
                   ATTRIBUTEN = SYMBOOL.GetAttributes
                    For j = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                    Set ATTRIBUUT = ATTRIBUTEN(j)
                      If ATTRIBUUT.TagString = "RNU" Then
                          oo = ATTRIBUUT.textstring
                          frmregelxls.ListBox3.AddItem (oo)
                      End If
                    Next j
        End If
      End If
      End If
  Next element10
  
  
  'lijst rangschikken
  Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For i = 0 To ListBox3.ListCount - 1
    textstring2 = ListBox3.List(i)
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
  ListBox3.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then ListBox3.AddItem Veld(x)
  Next x
  
  
End Sub

Private Sub cmdLayers_Click()
Call cmdLayers_Click1
Call herindeel
Call cmdAfsluiten_Click

Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "GERARD" Then ThisDrawing.Close (False)
    
End Sub

Private Sub cmdLayers_Click1()
ListBox5.Clear
Update
On Error Resume Next
  
 Dim cirkel As Object
 Dim element As Object
 Dim Lengte As Double
 Dim laagobj As Object
 
 Dim minaantal As Integer
 Dim maxaantal As Integer
 Dim y As Integer
 y = 0
 minaantal = 0
 maxaantal = ThisDrawing.Layers.Count
 For Each laagobj In ThisDrawing.Layers
    y = y + 1
    ProgressBar1.Min = minaantal
    ProgressBar1.Max = maxaantal
    ProgressBar1.Value = y
 
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
  ListBox5.AddItem (d)
  End If

    Lengte = 0 'Lengte leeggooien voordat de volgende groep wordt gemeten
    zlengte = 0
    Z = 0
   End If 'mystr + mystrcontr
 
 Next laagobj
ProgressBar1.Value = maxaantal
 Update
 

 Cmdprint.Enabled = True

   'lijst rangschikken
  Dim Veld(0 To 500)
  Dim textstring2 As String
  
    For i = 0 To ListBox5.ListCount - 1
    textstring2 = ListBox5.List(i)
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
  ListBox5.Clear
 
  For x = 0 To UBound(Veld)
  If Veld(x) <> "" Then ListBox5.AddItem Veld(x)
  Next x

End Sub
Private Sub herindeel()
teller3 = frmregelxls.ListBox3.ListCount
teller4 = frmregelxls.ListBox5.ListCount
              
          For k = 0 To teller3 - 1
          
            'Define the text object
            text5 = frmregelxls.ListBox3.List(k) 'unitnummer
            frmregelxls.ListBox1.Clear
            frmregelxls.Caption = "Unit: " & text5 & " (totaal: " & teller3 & " units)"
                    For u = 0 To teller4 - 1
                    'Define the text object
                    text7 = frmregelxls.ListBox5.List(u)
                    text6 = Split(text7, ".")
                    If text6(0) = text5 Then frmregelxls.ListBox1.AddItem (text7)
                    Next u
                Call lengte_typeunit
                Call Module2.excel_schrijven1(text5)
                Update
          Next k
End Sub
Private Sub cmdAfsluiten_Click()

If frmregelxls.TextBox18 <> "" Then
 checkopexport = Split(frmregelxls.TextBox18, ".")
 checkopexport1 = Right(checkopexport(0), 5)
 If checkopexport1 = "meten" Then ThisDrawing.Close
 bestandnaam3 = frmregelxls.TextBox16 & frmregelxls.TextBox17 & "-meten" & ".dwg"
 Kill (bestandnaam3)
End If

Unload Me
End Sub
Private Sub lengte_typeunit()

frmregelxls.ListBox2.Clear
Dim teller3
          Dim k
          Dim rl
          Dim gp
          Dim gg
          
          Dim textstring5
          Dim textstring6
          teller3 = frmregelxls.ListBox1.ListCount
''''''' Dim minaantal As Integer
''''''' Dim maxaantal As Integer
''''''' Dim I As Integer
''''''' I = 0
''''''' minaantal = 0
''''''' maxaantal = ThisDrawing.Layers.Count
          
          For k = 0 To teller3 - 1
            
            

''''''''    I = I + 1
''''''''    ProgressBar1.Min = minaantal
''''''''    ProgressBar1.Max = maxaantal
''''''''    ProgressBar1.Value = I
''''''''
            
           'Define the text object
            textstring5 = frmregelxls.ListBox1.List(k)
            textstring6 = Split(textstring5, ("#"))
           
             gg = "groep " & textstring6(0) 'groepbenaming samenvoegen
             
                   For Each element In ThisDrawing.ModelSpace
                      If element.ObjectName = "AcDbBlockReference" Then
                       If UCase(element.Name) = "GROEPTEKSTBLOK" Then
                           Set SYMBOOL = element
                           If SYMBOOL.HasAttributes Then
                               ATTRIBUTEN = SYMBOOL.GetAttributes
                               For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                                    Set ATTRIBUUT = ATTRIBUTEN(i)
                                    If ATTRIBUUT.TagString = "GROEPTEKST" Then gp = ATTRIBUUT.textstring
                                    If ATTRIBUUT.TagString = "ROLLENGTE" Then rolleesff = ATTRIBUUT.textstring
                                    If ATTRIBUUT.TagString = "HOHAFSTAND" Then hh = ATTRIBUUT.textstring
                                    If ATTRIBUUT.TagString = "WANDHOOGTE" And ATTRIBUUT.textstring <> " " Then
                                         whoogte = Split(ATTRIBUUT.textstring, " ")
                                    End If 'wandhoogte
                               Next i
                                           If rolleesff <> " " Then
                                                If (gp = gg) Then
                                                  If hh = "Wandverwarming" Then hh = "wv 15"
                                                hh = Split(hh, " ")
                                                textstring7 = textstring5 & "#" & hh(1)    'hh1 = hoh
                                                'frmregelxls.ListBox2.AddItem (textstring7)
'                                           rolleesff = ""
'                                           gp = ""
'                                           gg = ""
'                                           hh = ""
'                                           textstring7 = ""
                                                End If
                                           End If
                               End If
                           End If
                       End If
                               
                    Next element
                  
          

            'regelunitnummer
            TEXTSTRING61 = Split(textstring5, ("."))
            'MsgBox TEXTSTRING61(0)


Dim zdrt As String
For Each element10 In ThisDrawing.ModelSpace
      If element10.ObjectName = "AcDbBlockReference" Then
      If element10.Name = "Mat_spe_ZD" Or element10.Name = "Mat_spe_PE" Or element10.Name = "Mat_spe_PE800" _
      Or element10.Name = "Mat_spe_ALU" Or element10.Name = "Mat_spe_ZDringleiding" Or element10.Name = "Mat_spe_PEringleiding" Or _
      element10.Name = "Mat_spe_ALUringleiding" Or element10.Name = "Mat_spe_FLEX" Or element10.Name = "Mat_spe_ZD_1627" Or _
      element10.Name = "Mat_spe_FLEX_Aankoppel" Or element10.Name = "Mat_spe_ZD_1627500" Then
        Set SYMBOOL = element10
        If SYMBOOL.HasAttributes Then
                   ATTRIBUTEN = SYMBOOL.GetAttributes
                    For j = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                    Set ATTRIBUUT = ATTRIBUTEN(j)
                      If ATTRIBUUT.TagString = "RNU" Then RNUS = ATTRIBUUT.textstring
                            
                            If TEXTSTRING61(0) = RNUS Then
                                    
                                    For p = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                                    Set ATTRIBUUT = ATTRIBUTEN(p)
                                    If ATTRIBUUT.TagString = "REGELUNITTYPE" And ATTRIBUUT.textstring <> "" Then rt = ATTRIBUUT.textstring
                                    If ATTRIBUUT.TagString = "WTHZD" And ATTRIBUUT.textstring <> "" Then zdrt = ATTRIBUUT.textstring
                                    If ATTRIBUUT.TagString = "PE" And ATTRIBUUT.textstring <> "" Then zdrt = ATTRIBUUT.textstring
                                    If ATTRIBUUT.TagString = "ALU" And ATTRIBUUT.textstring <> "" Then zdrt = ATTRIBUUT.textstring
                                    If ATTRIBUUT.TagString = "FLEX_BUIS" And ATTRIBUUT.textstring <> "" Then zdrt = ATTRIBUUT.textstring
                                    If zdrt = "WTH-ZD 20 * 3,4 mm" Then zdrt = "WTH-ZD 20*3,4 mm"
                                    If zdrt = "WTH-ZD 16 * 2,7 mm" Then zdrt = "WTH-ZD 16*2,7 mm"
                                    If zdrt <> "" Then
                                       zdrts = Split(zdrt, " ")
                                       zdrts2 = Split(zdrts(1), "*")
                                       zdrts3 = zdrts2(0) & "/" & zdrts2(1)
                                    End If
                                    controle10 = InStr(1, rt, "/", vbBinaryCompare) 'staat er een / in??
                                    If controle10 <> 0 Then
                                              oplos10 = Split(rt, "/")
                                              oplos10 = Split(oplos10(0), " ")
                                              Else
                                              oplos10 = Split(rt, " ")
                                    End If
                                    
                                   Next p
                                   
                            
                            End If

           
                     Next j
              End If
              
                  
                    If TEXTSTRING61(0) = RNUS Then textstring8 = textstring7 & "#" & zdrts3 & "#" & oplos10(0)
                    If TEXTSTRING61(0) = RNUS Then frmregelxls.ListBox2.AddItem (textstring8)
             
     
      End If
      End If
  Next element10

Next k
''''''' ProgressBar1.Value = minaantal
 Update

End Sub
'''''''''''''''''''Private Sub wandvv1()
'''''''''''''''''''MsgBox "hallo"
'''''''''''''''''''    For Each element In ThisDrawing.ModelSpace
'''''''''''''''''''        If element.ObjectName = "AcDbBlockReference" Then
'''''''''''''''''''            If UCase(element.Name) = "GROEPTEKSTBLOK" Then
'''''''''''''''''''                Set SYMBOOL = element
'''''''''''''''''''                If SYMBOOL.HasAttributes Then
'''''''''''''''''''                    ATTRIBUTEN = SYMBOOL.GetAttributes
'''''''''''''''''''                    For I = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
'''''''''''''''''''                         Set ATTRIBUUT = ATTRIBUTEN(I)
'''''''''''''''''''                         'If attribuut.TagString = "GROEPTEKST" Then GRP = attribuut.textstring
'''''''''''''''''''                         If ATTRIBUUT.TagString = "HOHAFSTAND" Then WH = ATTRIBUUT.textstring
'''''''''''''''''''
'''''''''''''''''''                    Next I
'''''''''''''''''''                End If
'''''''''''''''''''            End If
'''''''''''''''''''        End If
'''''''''''''''''''    Next element
'''''''''''''''''''
'''''''''''''''''''   If WH = "Wandverwarming" Then
'''''''''''''''''''
'''''''''''''''''''     MsgBox "Er is wandverwarming in de tekening aanwezig.!!!!" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) & _
'''''''''''''''''''            "Vul de hoogte van de wandverwarming in. (standaard staat ie op 2,5 meter)", vbExclamation ' & (Chr(13) & Chr(10)) & _
'''''''''''''''''''            '"Als je meerdere hoogte's heb vul dan de grootste waarde in, of een gemiddelde waarde.", vbExclamation
'''''''''''''''''''    frmregelxls.Height = 134
'''''''''''''''''''    frmregelxls.cmdLayers.top = 78
'''''''''''''''''''    frmregelxls.cmdAfsluiten.top = 78
'''''''''''''''''''    frmregelxls.Label39.Visible = True
'''''''''''''''''''    frmregelxls.ComboBox1.Visible = True
'''''''''''''''''''   End If
'''''''''''''''''''End Sub

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

'''''Private Sub TextBox1_Change()
'''''If TextBox1 <> "" Then CommandButton1.Locked = False
'''''End Sub





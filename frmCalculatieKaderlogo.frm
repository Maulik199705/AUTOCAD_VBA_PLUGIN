VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalculatieKaderlogo 
   Caption         =   "Kaderlogo Calculatie"
   ClientHeight    =   2568
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   13248
   OleObjectBlob   =   "frmCalculatieKaderlogo.frx":0000
End
Attribute VB_Name = "frmCalculatieKaderlogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'29-09-2003 Plaatsen van Kaderlogo
'M.Bosch en G.C.Haak
'kruisje uitschakelen
'kaderlogocalc.dwg
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
Dim blockObj As AcadBlock
Call Calc_Check_Kaderlogo.Check_kaderlogocalc
 
TextBox3.SetFocus
Call combolijst1
Call Combolijst2
Call Schaal(scaal)
If scaal = 2 Then
schaalv = "1:100"
Else
schaalv = "1:50"
End If
If scaal = 4 Then schaalv = "1:200"


Call loginnaam(lognaam) 'Loginnaam oproepen
Call datum(kdate)


Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If UCase(element.Name) = "KADERLOGOCALC" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "PROJECTNAAM" Then TextBox3 = attribuut.textstring
        If attribuut.TagString = "MONTAGEADRES" Then TextBox4 = attribuut.textstring
        If attribuut.TagString = "MONTAGEPLAATS" Then TextBox5 = attribuut.textstring
        If attribuut.TagString = "PROJECTNUMMER" Then TextBox6 = Left(ThisDrawing.GetVariable("dwgname"), 9) 'attribuut.TextString
        If attribuut.TagString = "BLAD" Then ComboBox1 = attribuut.textstring
        If attribuut.TagString = "BLAD" And attribuut.textstring = "" Then ComboBox1 = "1"
        If attribuut.TagString = "FORMAAT" Then TextBox8 = attribuut.textstring
        If attribuut.TagString = "TEKENAAR" And attribuut.textstring = "" Then TextBox9 = lognaam
        If attribuut.TagString = "TEKENAAR" And attribuut.textstring <> "" Then TextBox9 = attribuut.textstring
        If attribuut.TagString = "SCHAAL" And attribuut.textstring = "" Then TextBox10 = schaalv
        If attribuut.TagString = "SCHAAL" And attribuut.textstring <> "" Then TextBox10 = attribuut.textstring
        If attribuut.TagString = "DATUM" And attribuut.textstring = "" Then TextBox11 = kdate ' & "|" & lognaam
        If attribuut.TagString = "DATUM" And attribuut.textstring <> "" Then TextBox11 = attribuut.textstring
        If attribuut.TagString = "WIJZIGING1" And attribuut.textstring <> "" Then TextBox12 = attribuut.textstring
        If attribuut.TagString = "WIJZIGING2" And attribuut.textstring <> "" Then TextBox13 = attribuut.textstring
        If attribuut.TagString = "WIJZIGING3" And attribuut.textstring <> "" Then TextBox14 = attribuut.textstring
        If attribuut.TagString = "WIJZIGING4" And attribuut.textstring <> "" Then TextBox15 = attribuut.textstring
        Next I
       End If
     End If
   End If
  Next element
  
Dim element10 As Object
For Each element10 In ThisDrawing.ModelSpace
      If element10.ObjectName = "AcDbBlockReference" Then
      If UCase(element10.Name) = "VLRWTH" Then
      Set symbool = element10
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "TEKENINGNUMMER" Then TextBox21 = attribuut.textstring
        If attribuut.TagString = "BLAD_BWK" Then TextBox22 = attribuut.textstring
        If attribuut.TagString = "DATUM_BWK" Then TextBox23 = attribuut.textstring
        If attribuut.TagString = "PROJEKTLEIDER" Then ComboBox2 = attribuut.textstring
        If attribuut.TagString = "PROJEKTLEIDER" And attribuut.textstring = "" Then ComboBox2 = "F.v/d Corput"
        Next j
       End If
     End If
   End If
  Next element10

If TextBox6.Value <> "" Then TextBox6.Locked = True
If TextBox8.Value <> "" Then TextBox8.Locked = True
If TextBox10.Value <> "" Then TextBox10.Locked = True
Update
End Sub
Sub datum(kdate)
datumacad1 = ThisDrawing.GetVariable("cdate")
datumacad = Left(datumacad1, 8)
dag = Right(datumacad, 2)
maand = Left(datumacad, 6)
maand2 = Right(maand, 2)
jaar = Left(datumacad, 4)

kdate = dag & "-" & maand2 & "-" & jaar

End Sub
Private Sub CheckBox2_Click()
Call loginnaam(lognaam)
Call datum(kdate)
If CheckBox2.Value = True Then
TextBox12 = kdate & "|" & lognaam
Else
TextBox12 = ""
End If
End Sub

Private Sub CheckBox3_Click()
Call loginnaam(lognaam)
Call datum(kdate)
If CheckBox3.Value = True Then
TextBox13 = kdate & "|" & lognaam
Else
TextBox13 = ""
End If
End Sub

Private Sub CheckBox4_Click()
Call loginnaam(lognaam)
Call datum(kdate)
If CheckBox4.Value = True Then
TextBox14 = kdate & "|" & lognaam
Else
TextBox14 = ""
End If
End Sub
Private Sub CheckBox5_Click()
Call loginnaam(lognaam)
Call datum(kdate)
If CheckBox5.Value = True Then
TextBox15 = kdate & "|" & lognaam
Else
TextBox15 = ""
End If
End Sub
Private Sub CheckBox1_Click()
Call loginnaam(lognaam)
Call datum(kdate)
If CheckBox1.Value = True Then
TextBox19 = kdate & "|" & lognaam
Else
TextBox19 = ""
End If
End Sub
Private Sub TextBox10_Change()
If TextBox10.Value <> "" Then TextBox10.Locked = True
End Sub
Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox3 = ""
End Sub
Private Sub TextBox4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox4 = ""
End Sub
Private Sub TextBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox5 = ""
End Sub
Private Sub CmdUpdate_Click()
Grootklein = Left(ThisDrawing.GetVariable("dwgname"), 2) 'nummer checken
scaal = frmCalculatieKaderlogo.TextBox20

   For Each element10 In ThisDrawing.ModelSpace
      If element10.ObjectName = "AcDbBlockReference" Then
      If UCase(element10.Name) = "KADERLOGOCALC" Then
      Set symbool = element10
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "PROJECTNAAM" Then attribuut.textstring = TextBox3
        If attribuut.TagString = "MONTAGEADRES" Then attribuut.textstring = TextBox4
        If attribuut.TagString = "MONTAGEPLAATS" Then attribuut.textstring = TextBox5
        If attribuut.TagString = "PROJECTNUMMER" And Grootklein = "p0" Then attribuut.textstring = LCase(TextBox6)
        If attribuut.TagString = "PROJECTNUMMER" And Grootklein <> "p0" Then attribuut.textstring = UCase(TextBox6)
        If attribuut.TagString = "BLAD" Then attribuut.textstring = ComboBox1
        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = TextBox8
        If attribuut.TagString = "TEKENAAR" Then attribuut.textstring = TextBox9
        If attribuut.TagString = "SCHAAL" Then attribuut.textstring = TextBox10
        If attribuut.TagString = "DATUM" Then attribuut.textstring = TextBox11
        If attribuut.TagString = "WIJZIGING1" Then attribuut.textstring = TextBox12
        If attribuut.TagString = "WIJZIGING2" Then attribuut.textstring = TextBox13
        If attribuut.TagString = "WIJZIGING3" Then attribuut.textstring = TextBox14
        If attribuut.TagString = "WIJZIGING4" Then attribuut.textstring = TextBox15
     Next I

      End If
      End If
      End If
      Next element10
Update
Call BLOK2
End Sub
Private Sub BLOK2()
    For Each element11 In ThisDrawing.ModelSpace
      If element11.ObjectName = "AcDbBlockReference" Then
      If UCase(element11.Name) = "VLRWTH" Then
      Set symbool = element11
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
       For k = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(k)
        If attribuut.TagString = "TEKENINGNUMMER" Then attribuut.textstring = TextBox21
        If attribuut.TagString = "BLAD_BWK" Then attribuut.textstring = TextBox22
        If attribuut.TagString = "DATUM_BWK" Then attribuut.textstring = TextBox23
        If attribuut.TagString = "PROJEKTLEIDER" Then attribuut.textstring = ComboBox2
       Next k

      End If
      End If
      End If
     Next element11
Update
Call cmdAfsluiten_Click
End Sub
Private Sub cmdAfsluiten_Click()
Unload Me
End Sub
Sub Schaal(scaal)
frmCalculatieKaderlogo.Hide
On Error Resume Next
vartab = ThisDrawing.GetVariable("EXTMAX")
If vartab(0) >= 2145 And vartab(0) <= 2155 Or vartab(0) >= 4095 And vartab(0) <= 4105 _
    Or vartab(0) >= 5835 And vartab(0) <= 5845 Or vartab(0) >= 8355 And vartab(0) <= 8365 _
    Or vartab(0) >= 11785 And vartab(0) <= 11795 Or vartab(0) >= 15685 And vartab(0) <= 15695 Then
    scaal = 2
Else
    scaal = 1
End If
If vartab(0) >= 16715 And vartab(0) <= 16725 Or vartab(0) >= 23575 And vartab(0) <= 23585 _
    Or vartab(0) >= 31375 And vartab(0) <= 31385 Then
    scaal = 4
End If
TextBox20 = scaal
Update
End Sub
Sub loginnaam(lognaam)
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "CBOOI" Then lognaam = "CB"
If lognaam = "DVERB" Then lognaam = "DV"
If lognaam = "EHOUW" Then lognaam = "EH"
If lognaam = "GWIJK" Then lognaam = "GW"
If lognaam = "LARS" Then lognaam = "LvD"
If lognaam = "JZILV" Then lognaam = "JZ"
If lognaam = "KREITERJ" Then lognaam = "JK"
If lognaam = "LIESBETH" Then lognaam = "LR"
If lognaam = "RRAAT" Then lognaam = "RR"
If lognaam = "JHAM" Then lognaam = "JH"
If lognaam = "GERARD" Then lognaam = "GCH"
End Sub
Sub combolijst1()

ComboBox1.AddItem "1"
ComboBox1.AddItem "W00"
ComboBox1.AddItem "W01"
ComboBox1.AddItem "W02"
ComboBox1.AddItem "W03"
ComboBox1.AddItem "W04"
ComboBox1.AddItem "W05"
ComboBox1.AddItem "W06"
ComboBox1.AddItem "W07"
ComboBox1.AddItem "W08"
ComboBox1.ListIndex = 0
ComboBox1.MatchEntry = fmMatchEntryFirstLetter
End Sub

Sub Combolijst2()

ComboBox2.AddItem "A. Schenkel"
ComboBox2.AddItem "D. Heron"
ComboBox2.AddItem "F. v/d Corput"
ComboBox2.AddItem "F. Prenen"
ComboBox2.ListIndex = 2
ComboBox2.MatchEntry = fmMatchEntryFirstLetter
End Sub

''''''''Private Sub CommandButton1_Click()
'''''''''TextBox3 = StrConv(TextBox3, vbProperCase)
'''''''''TextBox4 = StrConv(TextBox4, vbProperCase)
'''''''''TextBox5 = StrConv(TextBox5, vbProperCase)
''''''''
''''''''Grootklein = Left(ThisDrawing.GetVariable("dwgname"), 2) 'nummer checken
''''''''
''''''''Dim element10 As Object
''''''''If CheckBox9.Value = False Then
''''''''     For Each element10 In ThisDrawing.ModelSpace
''''''''      If element10.ObjectName = "AcDbBlockReference" Then
''''''''      If element10.Name = "Kaderlogo" And CheckBox9.Value = False Then
''''''''      Set symbool = element10
''''''''        If symbool.HasAttributes Then
''''''''        attributen = symbool.GetAttributes
''''''''        For I = LBound(attributen) To UBound(attributen)
''''''''        Set attribuut = attributen(I)
''''''''        If attribuut.TagString = "OPDRACHTGEVER" Then attribuut.textstring = TextBox1
''''''''        If attribuut.TagString = "PLAATS" Then attribuut.textstring = UCase(TextBox2)
''''''''        If attribuut.TagString = "PROJECTNAAM" Then attribuut.textstring = TextBox3
''''''''        If attribuut.TagString = "MONTAGEADRES" Then attribuut.textstring = TextBox4
''''''''        If attribuut.TagString = "MONTAGEPLAATS" Then attribuut.textstring = TextBox5
''''''''        If attribuut.TagString = "PROJECTNUMMER" And Grootklein = "p0" Then attribuut.textstring = LCase(TextBox6)
''''''''        If attribuut.TagString = "PROJECTNUMMER" And Grootklein <> "p0" Then attribuut.textstring = UCase(TextBox6)
''''''''        If attribuut.TagString = "BLAD" Then attribuut.textstring = ComboBox1
''''''''        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = TextBox8
''''''''        If attribuut.TagString = "TEKENAAR" Then attribuut.textstring = TextBox9
''''''''        If attribuut.TagString = "SCHAAL" Then attribuut.textstring = TextBox10
''''''''        If attribuut.TagString = "DATUM" Then attribuut.textstring = TextBox11
''''''''        If attribuut.TagString = "WIJZIGING1" Then attribuut.textstring = TextBox12
''''''''        If attribuut.TagString = "WIJZIGING2" Then attribuut.textstring = TextBox13
''''''''        If attribuut.TagString = "WIJZIGING3" Then attribuut.textstring = TextBox14
''''''''        If attribuut.TagString = "WIJZIGING4" Then attribuut.textstring = TextBox15
''''''''        If attribuut.TagString = "WIJZIGING5" Then attribuut.textstring = TextBox16
''''''''        If attribuut.TagString = "WIJZIGING6" Then attribuut.textstring = TextBox17
''''''''        If attribuut.TagString = "WIJZIGING7" Then attribuut.textstring = TextBox18
''''''''        If attribuut.TagString = "REVISIE" Then attribuut.textstring = TextBox19
''''''''       Next I
''''''''
''''''''      End If
''''''''      End If
''''''''      End If
''''''''      Next element10
''''''''End If
'''''''''Else
''''''''
''''''''
'''''''' ' For Each GG In ThisDrawing.ModelSpace
'''''''' '   If GG.ObjectName = "AcDbBlockReference" Then
'''''''' '      If GG.Name = "Kaderlogo" Then
'''''''' '         a = GG.InsertionPoint
'''''''' '         'GG.Delete
''''''''          'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(a, "C:\ACAD2002\DWG\bl-bouw.dwg", 1, 1, 1, 0)
'''''''' '       End If
'''''''' '   End If
'''''''' ' Next
''''''''
''''''''
''''''''If CheckBox9 = True And CheckBox10 = False Then
''''''''Dim element As Object
''''''''Dim layerObj As AcadLayer
'''''''''Kaderlogo verwijderen
''''''''     For Each element In ThisDrawing.ModelSpace
''''''''      If element.ObjectName = "AcDbBlockReference" Then
''''''''         If element.Name = "Kaderlogo" Then
''''''''          engkaderlogo = element.InsertionPoint
''''''''          element.Erase
''''''''         End If
''''''''      Update
''''''''      End If
''''''''     Next element
''''''''
''''''''
''''''''         On Error Resume Next
''''''''         For Each element In ThisDrawing.ModelSpace
''''''''         If element.ObjectName = "AcDbBlockReference" Then
''''''''         If element.Name <> "KaderlogoEngels" Then
''''''''
''''''''
''''''''    If Err Then
''''''''    CheckBox10.Value = False
''''''''    frmCalculatieKaderlogo.Show
''''''''    Exit Sub
''''''''    End If
''''''''
''''''''         Dim bestand91 As String
''''''''         bestand91 = "C:\ACAD2002\DWG\KaderlogoEngels.dwg"
''''''''         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(engkaderlogo, bestand91, 1, 1, 1, 0)
''''''''         End If
''''''''         End If
''''''''         Next element
''''''''         Update
''''''''         CheckBox10 = True
''''''''End If
''''''''
''''''''
''''''''
''''''''Dim element11 As Object
''''''''For Each element11 In ThisDrawing.ModelSpace
''''''''     If element11.ObjectName = "AcDbBlockReference" Then
''''''''      If element11.Name = "KaderlogoEngels" And CheckBox9.Value = True Then
''''''''      Set symbool = element11
''''''''        If symbool.HasAttributes Then
''''''''        attributen = symbool.GetAttributes
''''''''        For I = LBound(attributen) To UBound(attributen)
''''''''        Set attribuut = attributen(I)
''''''''        If attribuut.TagString = "OPDRACHTGEVER" Then attribuut.textstring = TextBox1
''''''''        If attribuut.TagString = "PLAATS" Then attribuut.textstring = UCase(TextBox2)
''''''''        If attribuut.TagString = "PROJECTNAAM" Then attribuut.textstring = TextBox3
''''''''        If attribuut.TagString = "MONTAGEADRES" Then attribuut.textstring = TextBox4
''''''''        If attribuut.TagString = "MONTAGEPLAATS" Then attribuut.textstring = TextBox5
''''''''        If attribuut.TagString = "PROJECTNUMMER" And Grootklein = "p0" Then attribuut.textstring = LCase(TextBox6)
''''''''        If attribuut.TagString = "PROJECTNUMMER" And Grootklein <> "p0" Then attribuut.textstring = UCase(TextBox6)
''''''''        If attribuut.TagString = "BLAD" Then attribuut.textstring = ComboBox1
''''''''        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = TextBox8
''''''''        If attribuut.TagString = "TEKENAAR" Then attribuut.textstring = TextBox9
''''''''        If attribuut.TagString = "SCHAAL" Then attribuut.textstring = TextBox10
''''''''        If attribuut.TagString = "DATUM" Then attribuut.textstring = TextBox11
''''''''        If attribuut.TagString = "WIJZIGING1" Then attribuut.textstring = TextBox12
''''''''        If attribuut.TagString = "WIJZIGING2" Then attribuut.textstring = TextBox13
''''''''        If attribuut.TagString = "WIJZIGING3" Then attribuut.textstring = TextBox14
''''''''        If attribuut.TagString = "WIJZIGING4" Then attribuut.textstring = TextBox15
''''''''        If attribuut.TagString = "WIJZIGING5" Then attribuut.textstring = TextBox16
''''''''        If attribuut.TagString = "WIJZIGING6" Then attribuut.textstring = TextBox17
''''''''        If attribuut.TagString = "WIJZIGING7" Then attribuut.textstring = TextBox18
''''''''        If attribuut.TagString = "REVISIE" Then attribuut.textstring = TextBox19
''''''''
''''''''       Next I
''''''''
''''''''
''''''''
''''''''        End If
''''''''      End If
''''''''    End If
''''''''Next element11
'''''''''End If
'''''''''End If 'checkbox9
''''''''Update
''''''''
''''''''
''''''''End Sub

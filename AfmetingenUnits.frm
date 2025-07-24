VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AfmetingenUnits 
   Caption         =   "AFMETINGEN VAN DE UNITS"
   ClientHeight    =   5724
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5748
   OleObjectBlob   =   "AfmetingenUnits.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AfmetingenUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'afmetingen units aangepast op 17-11-2006 G.C.Haak
'maakt gebruik van HERZ.txt ' KMV.txt ' LT.txt  ' LT-N.txt ' LTVK.txt ' LTS.txt ' LTS-N.txt ' RUBK-R.txt
' RUB-R.txt ' RU-EE.txt  ' RUH-R.txt ' RUH-N.txt ' RUV.txt ' RU-WK.txt ' RU-WW.txt ' VSKO.txt

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
'EntryCount = 0
ListBox2.Clear
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Melding.Visible = False
ListBox1.AddItem ("HERZ")
ListBox1.AddItem ("KMV")
ListBox1.AddItem ("LT")
ListBox1.AddItem ("LT-N")
ListBox1.AddItem ("LT-VK")
ListBox1.AddItem ("LTS")
ListBox1.AddItem ("LTS-N")
ListBox1.AddItem ("RUBK-R") '
ListBox1.AddItem ("RUB-R/T")
ListBox1.AddItem ("RUB-S")
ListBox1.AddItem ("RU-EE")
ListBox1.AddItem ("RUH-R/T")
ListBox1.AddItem ("RUH-S")
ListBox1.AddItem ("RUH-N")
ListBox1.AddItem ("RUV")
ListBox1.AddItem ("RU-WK")
ListBox1.AddItem ("RU-WW")
ListBox1.AddItem ("VSKO")

ListBox3.AddItem ("1 groeps")
ListBox3.AddItem ("2 groepen")
ListBox3.AddItem ("3 groepen")
ListBox3.AddItem ("4 groepen")
ListBox3.AddItem ("5 groepen")
ListBox3.AddItem ("6 groepen")
ListBox3.AddItem ("7 groepen")
ListBox3.AddItem ("8 groepen")
ListBox3.AddItem ("9 groepen")
ListBox3.AddItem ("10 groepen")
ListBox3.AddItem ("11 groepen")
ListBox3.AddItem ("12 groepen")
ListBox3.AddItem ("13 groepen")
ListBox3.AddItem ("14 groepen")
ListBox3.AddItem ("15 groepen")
ListBox3.AddItem ("16 groepen")
ListBox3.AddItem ("17 groepen")
ListBox3.AddItem ("18 groepen")
ListBox3.AddItem ("19 groepen")
ListBox3.AddItem ("20 groepen")
End Sub
Private Sub CancelButton1_Click()
Unload Me
End Sub
Private Sub ListBox1_Click()

If ListBox1.Value = "HERZ" Then bestand10 = "c:\acad2002\dwg\herz.txt" 'Call herz
If ListBox1.Value = "KMV" Then bestand10 = "c:\acad2002\dwg\kmv.txt" ' Then Call kmv
If ListBox1.Value = "LT" Then bestand10 = "c:\acad2002\dwg\lt.txt" 'Then Call lt
If ListBox1.Value = "LT-N" Then bestand10 = "c:\acad2002\dwg\lt-n.txt" 'Then Call lt
If ListBox1.Value = "LT-VK" Then bestand10 = "c:\acad2002\dwg\lt-vk.txt" 'Then Call ltvk
If ListBox1.Value = "LTS" Then bestand10 = "c:\acad2002\dwg\lts.txt" 'Then Call lts
If ListBox1.Value = "LTS-N" Then bestand10 = "c:\acad2002\dwg\lts-n.txt" 'Then Call ltsn
If ListBox1.Value = "RUBK-R" Then bestand10 = "c:\acad2002\dwg\rubk-r.txt" 'Then Call rubk
If ListBox1.Value = "RUB-R/T" Then bestand10 = "c:\acad2002\dwg\rub-r.txt" 'Then Call rubr
If ListBox1.Value = "RUB-S" Then bestand10 = "c:\acad2002\dwg\rub-s.txt" 'Then Call rubr
If ListBox1.Value = "RU-EE" Then bestand10 = "c:\acad2002\dwg\ru-ee.txt" 'Then Call ruee
If ListBox1.Value = "RUH-R/T" Then bestand10 = "c:\acad2002\dwg\ruh-r.txt" 'Then Call ruhr
If ListBox1.Value = "RUH-S" Then bestand10 = "c:\acad2002\dwg\ruh-s.txt" 'Then Call ruhr
If ListBox1.Value = "RUH-N" Then bestand10 = "c:\acad2002\dwg\ruh-n.txt" 'Then Call ruhn
If ListBox1.Value = "RUV" Then bestand10 = "c:\acad2002\dwg\ruv.txt" 'Then Call ruv
If ListBox1.Value = "RU-WK" Then bestand10 = "c:\acad2002\dwg\ru-wk.txt" 'Then Call ruwk
If ListBox1.Value = "RU-WW" Then
    bestand10 = "c:\acad2002\dwg\ru-ww.txt"
    Label5.Caption = "B352 t/m B370"
    Label7.Caption = "50-40"
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Melding.Visible = True
End If
If ListBox1.Value = "VSKO" Then bestand10 = "c:\acad2002\dwg\vsko.txt" 'Then Call vsko
If Not ListBox3.ListIndex < 0 Then ListBox2.ListIndex = ListBox3.ListIndex

Call lees(bestand10)
End Sub
Sub lbvals()
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Melding.Visible = False
End Sub
Private Sub ListBox3_Click()
   If Not ListBox1.ListIndex < 0 Then
       Melding.Visible = False
       ListBox2.ListIndex = ListBox3.ListIndex
   Else
       Melding.Visible = True
   End If
End Sub
Sub lees(bestand10)

ListBox2.Clear
If bestand10 <> "c:\acad2002\dwg\ru-ww.txt" Then Call lbvals
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim fs, a, vullistbox2
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(bestand10, ForReading, False)


Do While a.AtEndOfLine <> True
    vullistbox2 = a.ReadLine
   AfmetingenUnits.ListBox2.AddItem (vullistbox2)
Loop
a.Close 'sluiten van tekstbestand
End Sub


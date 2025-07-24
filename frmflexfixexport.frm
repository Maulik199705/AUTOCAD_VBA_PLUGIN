VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmflexfixexport 
   Caption         =   "Export Flexfix....."
   ClientHeight    =   1365
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   6912
   OleObjectBlob   =   "frmflexfixexport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmflexfixexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mappAcad As AcadApplication
Private mstrPath As String


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
  'create acadapplication object
  Set mappAcad = GetObject(Class:="AutoCAD.Application")
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub CommandButton1_Click()
''For Each ELEMENT4 In ThisDrawing.ModelSpace
''        If ELEMENT4.ObjectName = "AcDbBlockReference" Then
''            If UCase(ELEMENT4.Name) = "GROEPTEKSTBLOKNEW" Then we = we + 1
''        End If
''Next
Dim element
Dim SYMBOOL
Dim ATTRIBUTEN: Dim ATTRIBUUT: Dim zz: Dim rol
Dim i: Dim J: Dim K: Dim aa: Dim bb: Dim cc: Dim dd: Dim ee:
Dim ff: Dim gg: Dim hh: Dim ii: Dim jj: Dim kk: Dim ll
Dim x
For Each element In ThisDrawing.ModelSpace
        
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOKNEW" Then
                Set SYMBOOL = element
                If SYMBOOL.HasAttributes Then
                    ATTRIBUTEN = SYMBOOL.GetAttributes
                    For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                    Set ATTRIBUUT = ATTRIBUTEN(i)
                      If ATTRIBUUT.TagString = "FLEXFIX" And UCase(ATTRIBUUT.TextString) = "JA" Then
                                   For x = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                                   Set ATTRIBUUT = ATTRIBUTEN(x)
                                   If ATTRIBUUT.TagString = "ROLLENGTE" And ATTRIBUUT.TextString <> " " Then
                                        For J = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                                        Set ATTRIBUUT = ATTRIBUTEN(J)
                                             If ATTRIBUUT.TagString = "UNITNUMMER" And ATTRIBUUT.TextString <> " " Then aa = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "MATNUMMER_FLEXFIX" And ATTRIBUUT.TextString <> " " Then bb = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "FLEXFIX_RETOUR" And ATTRIBUUT.TextString <> " " Then cc = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "FLEXFIX_AANVOER" And ATTRIBUUT.TextString <> " " Then dd = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "AANTAL_SLINGERS_FLEXFIX" And ATTRIBUUT.TextString <> " " Then ee = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "BREEDTE_FLEXFIX" And ATTRIBUUT.TextString <> " " Then ff = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "ROLLENGTE" And ATTRIBUUT.TextString <> " " Then gg = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "GROEPTEKST" And ATTRIBUUT.TextString <> " " Then hh = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "HOHAFSTAND" And ATTRIBUUT.TextString <> " " Then ii = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "SL_RETOUR" And ATTRIBUUT.TextString <> " " Then jj = ATTRIBUUT.TextString
                                             If ATTRIBUUT.TagString = "SL_AANVOER" And ATTRIBUUT.TextString <> " " Then kk = ATTRIBUUT.TextString
                                        Next J
                                   End If
                      
                                   Next x
                        End If
                                    
                    Next i
                End If
            End If
        End If
                    If aa <> "" And bb <> "" And cc <> "" And dd <> "" And ee <> "" And ff <> "" And gg <> "" And hh <> "" _
                    And ii <> "" And jj <> "" And kk <> "" Then
                    
                    ll = aa & "+" & bb & "+" & cc & "+" & dd & "+" & ee & "+" & ff & "+" & gg & "+" & hh & "+" & ii & "+" & jj & "+" & kk
                    If ll <> "" Then frmflexfixexport.ListBox1.AddItem (ll)
                    ll = "": aa = "":  bb = "": cc = "": dd = "": ee = "": ff = "": gg = "": hh = "": ii = "": jj = "": kk = ""
                    End If
                    
          
         
    Next element

Call ffxs

End Sub
Sub ffxs()
Dim teknaam
Dim pad
Dim usernaam
Dim teknaam2
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fs, f
Dim s1 As AcadSelectionSet
Set fs = CreateObject("Scripting.FileSystemObject")
teknaam = ThisDrawing.GetVariable("dwgname")
teknaam2 = Split(teknaam, ".")
pad = ThisDrawing.GetVariable("dwgprefix")
usernaam = pad & teknaam2(0) & "-flexfix.xls"
Dim MyDate
MyDate = DateValue(Date)    ' Return a date.


Set f = fs.OpenTextFile(usernaam, ForWriting, -2)
    f.write "Unitnummer"
    f.write Chr(9)
    f.write "Matnummer_flexfix"
    f.write Chr(9)
    f.write "Flexfix_retour"
    f.write Chr(9)
    f.write "Flexfix_aanvoer"
    f.write Chr(9)
    f.write "Aantal_slingers_flexfix"
    f.write Chr(9)
    f.write "Breedte_flexfix"
    f.write Chr(9)
    f.write "Rollengte"
    f.write Chr(9)
    f.write "Groeptekst"
    f.write Chr(9)
    f.write "HOHafstand"
    f.write Chr(9)
    f.write "SL_retour"
    f.write Chr(9)
    f.write "SL_aanvoer"
    f.write Chr(10) + Chr(13)
    f.Close


teller = ListBox1.ListCount
    For i = 0 To teller - 1
       'Define the text object
        TextString = ListBox1.List(i)
        t2 = Split(TextString, "+")
        Set f = fs.OpenTextFile(usernaam, ForAppending, -2)
        f.write t2(0)
        f.write Chr(9)
        f.write t2(1)
        f.write Chr(9)
        f.write t2(2)
        f.write Chr(9)
        f.write t2(3)
        f.write Chr(9)
        f.write t2(4)
        f.write Chr(9)
        f.write t2(5)
        f.write Chr(9)
        f.write t2(6)
        f.write Chr(9)
        f.write t2(7)
        f.write Chr(9)
        f.write t2(8)
        f.write Chr(9)
        f.write t2(9)
        f.write Chr(9)
        f.write t2(10)
        f.write Chr(9)
        f.write Chr(13)
        f.Close
    Next i
Unload Me
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmplotten 
   Caption         =   "Plotten & Printen"
   ClientHeight    =   6780
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10152
   OleObjectBlob   =   "frmplotten.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmplotten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mappAcad As AcadApplication
Private mstrPath As String

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
  CheckBox1.ForeColor = &HFF&
  CheckBox1.BackColor = &H80FFFF   'lichtgeel
  CheckBox1.Value = True
  
  If lstDrawings.ListCount = 0 Then
     cmdStart.Enabled = False
     CheckBox5.Enabled = False
  End If
  frmplotten.Height = 363
     frmplotten.Width = 205
  frmplotten.StartUpPosition = 0
  
Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "GERARD" Then
   CheckBox5.Visible = True
   CheckBox6.Visible = True
   CheckBox12.Visible = False
   frmplotten.StartUpPosition = 0
   
''''ThisDrawing.SetVariable "users4", lognaam

End If


'Kill ("c:\acad2002\printlijst.txt")
End Sub
Private Sub CommandButton1_Click()
Call reset
End Sub
Sub reset()
lstDrawings.Clear  'reset button
ListBox1.Clear
CheckBox1.Value = True
cmdStart.Enabled = False
CommandButton2.Enabled = True
CheckBox2.Value = False
CheckBox2.ForeColor = &H0&
CheckBox2.BackColor = &HC0C0C0
CheckBox3.Value = False
CheckBox3.ForeColor = &H0&
CheckBox3.BackColor = &HC0C0C0
CheckBox4.Value = False
CheckBox4.ForeColor = &H0&
CheckBox4.BackColor = &HC0C0C0
CheckBox10.Value = False
CheckBox10.ForeColor = &H0&
CheckBox10.BackColor = &HC0C0C0
CheckBox11.Value = False
CheckBox11.ForeColor = &H0&
CheckBox11.BackColor = &HC0C0C0
Label2.Caption = ""
CheckBox5.Value = False
CheckBox5.Enabled = False
CheckBox14.Enabled = False
End Sub
Private Sub CommandButton2_Click()

frmplotten.Hide   'uitvoeren
Dim a
Dim b

Dim G1
Dim G2
Dim G3
G1 = ThisDrawing.GetVariable("dwgprefix")
G2 = ThisDrawing.GetVariable("dwgname")
G3 = G1 & G2
    
Dim foutcontrole
foutcontrole = ThisDrawing.GetVariable("users5")

    If CheckBox1.Value = True Then
    ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "gt" & vbCr & "ON" & vbCr & "gt" & vbCr & "T" & vbCr & "gt" & vbCr & vbCr
    ThisDrawing.SendCommand "pf" & vbCr 'plotfile aanmaken oce 5150
    If foutcontrole = "foutplot" Then MsgBox "Kader van -| " & G2 & " |- staat niet goed.", vbCritical
    End If
    If CheckBox2.Value = True Then
    ThisDrawing.SendCommand "a3plot" & vbCr 'A3 plot scale to fit
    End If
    If CheckBox3.Value = True Then
    ThisDrawing.SendCommand "a4laser" & vbCr 'A4 printen scale to fit
    End If
    If CheckBox4.Value = True Then
    ThisDrawing.SendCommand "-purge" & vbCr & "All" & vbCr & "*" & vbCr & "N" & vbCr
    ThisDrawing.SendCommand "inpak" & vbCr 'inpakken (etransmit)
    End If
    If CheckBox10.Value = True Then
    ThisDrawing.SendCommand "dwgpdf" & vbCr 'dwg to pdf
    End If
    If CheckBox11.Value = True Then
    ThisDrawing.SendCommand "d9400" & vbCr 'direct naar de 9400
    End If
    
    'If CheckBox5.Value = True Then
    'ThisDrawing.SendCommand "a4kleur" & vbCr 'A4 printen scale to fit in kleur
    'End If
    
Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)

'If frmplotten.CheckBox13.Value = False Then Call ThisDrawing.regelstaat
'If lognaam = "GERARD" And CheckBox12.Value = True Then Call ThisDrawing.acadnavi
If lognaam = "GERARD" Then ThisDrawing.Close (True)

  
Unload Me
End Sub
Private Sub Checkbox1_Click()
If CheckBox1.Value = True Then
  CheckBox1.ForeColor = &HFF& 'rood
  CheckBox1.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox1.ForeColor = &H0&
  CheckBox1.BackColor = &HC0C0C0
End If
End Sub
Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
  CheckBox2.ForeColor = &HFF& 'rood
  CheckBox2.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox2.ForeColor = &H0&
  CheckBox2.BackColor = &HC0C0C0
End If
End Sub
Private Sub Checkbox3_Click()
If CheckBox3.Value = True Then
  CheckBox3.ForeColor = &HFF& 'rood
  CheckBox3.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox3.ForeColor = &H0&
  CheckBox3.BackColor = &HC0C0C0
End If
End Sub

Private Sub Checkbox4_Click()
If CheckBox4.Value = True Then
  CheckBox4.ForeColor = &HFF& 'rood
  CheckBox4.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox4.ForeColor = &H0&
  CheckBox4.BackColor = &HC0C0C0
End If
End Sub
Private Sub CheckBox5_Click()
If CheckBox3.Value = True Then
  CheckBox3.ForeColor = &HFF& 'rood
  CheckBox3.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox3.ForeColor = &H0&
  CheckBox3.BackColor = &HC0C0C0
End If
End Sub
Private Sub CheckBox10_Click()
If CheckBox10.Value = True Then
  CheckBox10.ForeColor = &HFF& 'rood
  CheckBox10.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox10.ForeColor = &H0&
  CheckBox10.BackColor = &HC0C0C0
End If
End Sub

Private Sub CheckBox14_Click()
If CheckBox14.Value = True Then
  CheckBox14.ForeColor = &HFF& 'rood
  CheckBox14.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox14.ForeColor = &H0&
  CheckBox14.BackColor = &HC0C0C0
End If
End Sub
Private Sub UserForm_Terminate()

  'clean up
  Set mappAcad = Nothing

End Sub
Private Sub cmdClose_Click()
 
  ThisDrawing.SendCommand "doskopie" & vbCr
  'close this form
  Call Unload(Me)

End Sub

Private Sub cmdSelect_Click()
  Dim strFileNames As String
  Dim varFileNames As Variant
  Dim intFileName As Integer
  Dim strFileName As String
  
  Dim vardata
  Dim sysvarname
  sysvarname = "users1"
  vardata = ThisDrawing.GetVariable(sysvarname)
  If vardata = "" Then vardata = "f:\\Fserver2\\Gegevens\\Projecten"
  'MsgBox vardata
  
  'ask user to select one or more files
  With Me.CommonDialog
    .MaxFileSize = 2000
    .InitDir = vardata
    .FileName = ""
    .Filter = "AutoCAD Drawing [*.dwg]|*.dwg"
    .Flags = &H200 Or &H80000
    .DefaultExt = ".dwg"
    '.InitDir = "f:\\Fserver2\\Gegevens\\Projecten" 'mstrPath
    .ShowOpen
    strFileNames = .FileName
  End With
  
  'check returned files
  If Len(strFileNames) <> 0 Then
    'multiple files returned, split returned string
    If InStr(1, strFileNames, Chr(0), vbTextCompare) <> 0 Then
      varFileNames = Split(strFileNames, Chr(0), -1, vbTextCompare)
      'ListBox1.AddItem (varFileNames(0) & "\")
      With Me.lstDrawings
        'store path where selected fiels reside
        mstrPath = varFileNames(0) & "\"
        'ListBox1.AddItem (mstrPath)
        'show selected files in list box
        '.Clear  uitgezet gerard 8-3-2005
        For intFileName = 1 To UBound(varFileNames)
          Call .AddItem(varFileNames(intFileName))
          ListBox1.AddItem (varFileNames(0) & "\")
        Next intFileName
      End With
    Else
      'just one filename returned
      With Me.lstDrawings
        'split path and filename
        Call SplitFullFilename(strFileNames, mstrPath, strFileName)
        'show selected file in listbox
        '.Clear  uitgezet gerard 8-3-2005
        Call .AddItem(strFileName)
        
      End With
    End If
  
  End If
  If lstDrawings.ListCount <> 0 Then
      cmdStart.Enabled = True
      CommandButton2.Enabled = False
      CheckBox5.Enabled = True
      CheckBox14.Enabled = True
  End If
  
  Label2.Caption = lstDrawings.ListCount
End Sub
Private Sub SplitFullFilename(ByVal strFullFilename As String, ByRef strPath As String, ByRef strFile As String)
  Dim intLen As Integer
  Dim intTmp As Integer
    
  'check given full filename length
  intLen = Len(strFullFilename)
  If intLen <> 0 Then
    'check whether full filename contains a path seperator
    If InStr(1, strFullFilename, "\", vbTextCompare) <> 0 Then
      'path seperator found, split path and filename
      For intTmp = intLen To 0 Step -1
        If Mid(strFullFilename, intTmp, 1) = "\" Then
          strFile = Right(strFullFilename, intLen - intTmp)
          strPath = Left(strFullFilename, intTmp) 'path waar de file staat
          ListBox1.AddItem (strPath)
          Exit For
        End If
      Next intTmp
    Else
      'return filename only
      strFile = strFullFilename  'filenaam+dwg
    End If
  End If
'MsgBox strPath
Dim sysvarname
sysvarname = "users1"
ThisDrawing.SetVariable sysvarname, strPath

End Sub
Private Function FileExists(strFullFilename As String) As Boolean
  
  'check given filename string
  If Len(strFullFilename) <> 0 Then
    'check whether given filename exists
    If Len(Dir(strFullFilename, vbNormal)) <> 0 Then
      FileExists = True
    End If
  End If

End Function
Private Sub cmdStart_Click()

  Dim intDrawing As Integer
  Dim intDrawings As Integer

  'check whether anyt valid files are selected
  If Len(mstrPath) <> 0 Then
   
    intDrawings = Me.lstDrawings.ListCount - 1
    If intDrawings >= 0 Then
      'iterate through drawings
      For intDrawing = 0 To intDrawings
        'prepare each drawing
        'Call PrepareAsBackGround(mstrPath & Me.lstDrawings.List(intDrawing))
        Call PrepareAsBackGround(ListBox1.List(intDrawing) & Me.lstDrawings.List(intDrawing))
          
      Next intDrawing
    End If
    
  End If
    

    Call Unload(Me)


End Sub

Private Sub PrepareAsBackGround(strDrawingFullname As String)
frmplotten.Hide
  Dim objDocument As AcadDocument
  Dim objBlock As AcadBlock
  Dim objEntity As AcadEntity
  Dim objLayer As AcadLayer
  Dim strSaveAs As String
  Dim lngColor As Long
  Dim a
  Dim b
  
    Dim minaantal As Integer
    Dim maxaantal As Integer
    Dim i As Integer
    i = 0
    minaantal = 0
    maxaantal = Label2
 
     
     
    Dim G1
    Dim G2
    Dim G3
   
'    Dim lognaam
'    lognaam = ThisDrawing.GetVariable("loginname")
'    lognaam = UCase(lognaam)
'
  
  'check whether given filename is valid
  If FileExists(strDrawingFullname) Then
    
    i = i + 1
    ProgressBar1.Min = minaantal
    ProgressBar1.Max = maxaantal
    ProgressBar1.Value = i
    
    
    
    
    
    
    'open document to prepare as background
    Set objDocument = mappAcad.Documents.Open(strDrawingFullname)
    If CheckBox1.Value = True Then
'''    a = 1
    ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "gt" & vbCr & "ON" & vbCr & "gt" & vbCr & "T" & vbCr & "gt" & vbCr & vbCr
    ThisDrawing.SendCommand "pf" & vbCr 'plotfile aanmaken oce TCS400
'''    Call checkallesgoed(a, b)
'''    a = 0
        G1 = ThisDrawing.GetVariable("dwgprefix")
        G2 = ThisDrawing.GetVariable("dwgname")
        G3 = G1 & G2
            
        Dim foutcontrole
        foutcontrole = ThisDrawing.GetVariable("users5")
        If foutcontrole = "foutplot" Then MsgBox "Kader van -| " & G2 & " |- staat niet goed.", vbCritical
    
    End If
    If CheckBox2.Value = True Then
'''    a = 1
    ThisDrawing.SendCommand "a3plot" & vbCr 'A3 plot scale to fit
'''    Call checkallesgoed(a, b)
'''    a = 0
    End If
    If CheckBox3.Value = True Then
    ThisDrawing.SendCommand "a4laser" & vbCr 'A4 printen scale to fit
    End If
    If CheckBox4.Value = True Then
    ThisDrawing.SendCommand "-purge" & vbCr & "All" & vbCr & "*" & vbCr & "N" & vbCr
    ThisDrawing.SendCommand "inpak" & vbCr 'inpakken (etransmit)
    End If
    If CheckBox5.Value = True Then
    ThisDrawing.SendCommand "recoverit" & vbCr 'recover tekening
    End If
    If CheckBox10.Value = True Then
    ThisDrawing.SendCommand "dwgpdf" & vbCr 'dwg to pdf
    End If
    If CheckBox11.Value = True Then
    ThisDrawing.SendCommand "d9400" & vbCr 'direct naar de 9400
    End If
    
    If CheckBox14.Value = True Then
        CheckBox13.Value = True
        ThisDrawing.SendCommand "-vbarun" & vbCr & "regelxls" & vbCr
        SendKeys "{ENTER}", True
    End If
     
    With objDocument
  
  '  If frmplotten.CheckBox13.Value = False Then Call ThisDrawing.regelstaat
''''    a = ThisDrawing.GetVariable("users4")
''''    If a = "GERARD" Then ThisDrawing.Close (False)
    
    ThisDrawing.Close (False)
    'If CheckBox12.Value = True Then Call ThisDrawing.acadnavi
    'If foutcontrole <> "foutplot" Then Call .Close(SaveChanges:=False)
    End With
    Label2 = Label2 - 1
    
    'clean up
    Set objLayer = Nothing
    Set objEntity = Nothing
    Set objBlock = Nothing
    Set objDocument = Nothing

  End If

End Sub
Sub checkallesgoed(a, b)
Dim mypath
Dim myname
mypath = "c:\plot\"
myname = Dir(mypath, vbDirectory)     ' Retrieve the first entry.

 

'Do While MyName <> ""    ' Start the loop.
'    ' Ignore the current directory and the encompassing directory.
'
'
'    If MyName <> "." And MyName <> ".." Then
'        ' Use bitwise comparison to make sure MyName is a directory.
'        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
'
'            ListBox1.AddItem (MyPath & MyName)
'            'ListBox2.AddItem (MyPath2 & MyName2)
'
'        End If    ' it represents a directory.
'    End If
'
'    MyName = Dir    ' Get next entry.
'
'Loop
Dim teller
Dim i
Dim tstring
Dim tstring2
teller = lstDrawings.ListCount
For i = 0 To teller - 1
   'Define the text object
    tstring = Split(lstDrawings.List(i), ".")
    tstring2 = tstring(0) & ".plt"
     mypath = tstring2  ' Set the path.
     myname = Dir(mypath, vbDirectory)     ' Retrieve the first entry.
      
     Do While myname <> ""    ' Start the loop.
    ' Ignore the current directory and the encompassing directory.
      If myname <> "." And myname <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
'        If (GetAttr(mypath & myname) And vbDirectory) = vbDirectory Then
'
'            lstDrawings.AddItem (mypath & myname & "\Tekeningen\Bouwkundig")
'            'ListBox2.AddItem (MyPath2 & MyName2)
'
'        End If    ' it represents a directory.
     End If

     myname = Dir    ' Get next entry.
     Loop
Next i

If a = 1 Then
' Dim a
 If FileExists("c:\plot\" & tstring2) Then
 a = 0
 b = 1
 Else
 MsgBox "Plotfile is niet aangemaakt", vbCritical
 End If
End If
End Sub
Private Sub CheckBox6_Click()
If CheckBox6.Value = True Then
frmplotten.Height = 536
frmplotten.Width = 489
CheckBox7.Visible = True
Else
frmplotten.Height = 350
frmplotten.Width = 205
CheckBox7.Visible = False
End If
End Sub
Private Sub CheckBox11_Click()
If CheckBox11.Value = True Then
  CheckBox11.ForeColor = &HFF& 'rood
  CheckBox11.BackColor = &H80FFFF   'lichtgeel
  Else
  CheckBox11.ForeColor = &H0&
  CheckBox11.BackColor = &HC0C0C0
End If
End Sub
Private Sub CommandButton3_Click()
CommandButton5.Visible = True
  Dim intDrawing As Integer
  Dim intDrawings As Integer

  'check whether anyt valid files are selected
  If Len(mstrPath) <> 0 Then
   
    intDrawings = Me.ListBox2.ListCount - 1
    If intDrawings >= 0 Then
      'iterate through drawings
      For intDrawing = 0 To intDrawings
        'prepare each drawing
        'Call PrepareAsBackGround(mstrPath & Me.lstDrawings.List(intDrawing))
        Call PrepareAsBackGround2(ListBox3.List(intDrawing) & Me.ListBox2.List(intDrawing))
          
      Next intDrawing
    End If
         
    
  End If
    
  If CheckBox7.Value = True Then
  'ThisDrawing.SendCommand "doskopie" & vbCr
  Call Unload(Me)
  End If
End Sub
Private Sub PrepareAsBackGround2(strDrawingFullname As String)
  Dim objDocument As AcadDocument
  Dim objBlock As AcadBlock
  Dim objEntity As AcadEntity
  Dim objLayer As AcadLayer
  Dim strSaveAs As String
  Dim lngColor As Long
  Dim a
  Dim b
  
'''    Dim minaantal As Integer
'''    Dim maxaantal As Integer
'''    Dim I As Integer
'''    I = 0
'''    minaantal = 0
'''    maxaantal = Label2
 
     
  
  
  
  'check whether given filename is valid
  If FileExists(strDrawingFullname) Then
    
'''    I = I + 1
'''    ProgressBar1.Min = minaantal
'''    ProgressBar1.Max = maxaantal
'''    ProgressBar1.Value = I
    
    'open document to prepare as background
    Set objDocument = mappAcad.Documents.Open(strDrawingFullname)
     Dim xx
     xx = strDrawingFullname
     'MsgBox xx
     Dim mystr
     mystr = Right(xx, 3)
     'MsgBox mystr
     
     Dim cc1
     Dim dd1
     Dim cc
     Dim dd
     cc = ThisDrawing.GetVariable("EXTMIN")
     cc1 = Round(cc(0), 2) & " [] " & Round(cc(1), 2)
     ListBox4.AddItem cc1
     dd = ThisDrawing.GetVariable("EXTMAX")
     dd1 = Round(dd(0), 2) & " [] " & Round(dd(1), 2)
     ListBox5.AddItem dd1
     
     If CheckBox9.Value = True Then
     ThisDrawing.SendCommand "scale" & vbCr & "all" & vbCr & vbCr & "0,0" & vbCr & Val(TextBox1) & vbCr
     End If
          
     If CheckBox9.Value = False Then
     With objDocument
        If mystr <> "dwg" Then Call .Close(SaveChanges:=True)
        If mystr = "dwg" Then Call .Close(SaveChanges:=False)
     End With
     End If
     
     If CheckBox9.Value = True Then
     With objDocument
        If mystr <> "dwg" Then Call .Close(SaveChanges:=True)
        If mystr = "dwg" Then Call .Close(SaveChanges:=True)
     End With
     End If
    
'''    Label2 = Label2 - 1
    
    'clean up
    Set objLayer = Nothing
    Set objEntity = Nothing
    Set objBlock = Nothing
    Set objDocument = Nothing

  End If

End Sub
Private Sub CommandButton4_Click()
 Dim strFileNames As String
  Dim varFileNames As Variant
  Dim intFileName As Integer
  Dim strFileName As String
  
  'ask user to select one or more files
  With Me.CommonDialog
    .MaxFileSize = 3000
    .FileName = ""
    .Filter = ""
    .Flags = &H200 Or &H80000
    .DefaultExt = ".dwg"
    .InitDir = "f:\\Fserver2\\Gegevens\\Projecten" 'mstrPath
    .ShowOpen
    strFileNames = .FileName
  End With
  
  'check returned files
  If Len(strFileNames) <> 0 Then
    'multiple files returned, split returned string
    If InStr(1, strFileNames, Chr(0), vbTextCompare) <> 0 Then
      varFileNames = Split(strFileNames, Chr(0), -1, vbTextCompare)
      'ListBox1.AddItem (varFileNames(0) & "\")
      With Me.ListBox2
        'store path where selected fiels reside
        mstrPath = varFileNames(0) & "\"
        'ListBox1.AddItem (mstrPath)
        'show selected files in list box
        '.Clear  uitgezet gerard 8-3-2005
        For intFileName = 1 To UBound(varFileNames)
          Call .AddItem(varFileNames(intFileName))
          ListBox3.AddItem (varFileNames(0) & "\")
        Next intFileName
      End With
    Else
      'just one filename returned
      With Me.ListBox2
        'split path and filename
        Call SplitFullFilename2(strFileNames, mstrPath, strFileName)
        'show selected file in listbox
        '.Clear  uitgezet gerard 8-3-2005
        Call .AddItem(strFileName)
        
      End With
    End If
  
  End If
  If ListBox2.ListCount <> 0 Then
      cmdStart.Enabled = True
      CommandButton2.Enabled = False
      CheckBox5.Enabled = True
  End If
End Sub
Private Sub SplitFullFilename2(ByVal strFullFilename As String, ByRef strPath As String, ByRef strFile As String)
  Dim intLen As Integer
  Dim intTmp As Integer
    
  'check given full filename length
  intLen = Len(strFullFilename)
  If intLen <> 0 Then
    'check whether full filename contains a path seperator
    If InStr(1, strFullFilename, "\", vbTextCompare) <> 0 Then
      'path seperator found, split path and filename
      For intTmp = intLen To 0 Step -1
        If Mid(strFullFilename, intTmp, 1) = "\" Then
          strFile = Right(strFullFilename, intLen - intTmp)
          strPath = Left(strFullFilename, intTmp) 'path waar de file staat
          ListBox3.AddItem (strPath)
          Exit For
        End If
      Next intTmp
    Else
      'return filename only
      strFile = strFullFilename  'filenaam+dwg
    End If
  End If

End Sub
Private Sub CommandButton5_Click()
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim pad
Dim usernaam
Dim fs, f
Dim s1 As AcadSelectionSet
Set fs = CreateObject("Scripting.FileSystemObject")
pad = ThisDrawing.GetVariable("dwgprefix")
usernaam = ThisDrawing.GetVariable("loginname")
Dim MyDate
MyDate = DateValue(Date)    ' Return a date.


Set f = fs.OpenTextFile("c:\acad2002\printlijst.txt", ForAppending, -2)
    f.write "Tekenaar: " & usernaam & " |Datum: " & MyDate
    f.write Chr(13) + Chr(10)
    f.write Chr(13) + Chr(10)
    f.write pad
    f.write Chr(13) + Chr(10)
    f.write Chr(13) + Chr(10)
    f.Close
    
    Dim teller
    Dim i
    Dim textstring2
    Dim textstring3
    Dim textstring4
    
    teller = ListBox3.ListCount
    For i = 0 To teller - 1
       'Define the text object
        Set f = fs.OpenTextFile("c:\acad2002\printlijst.txt", ForAppending, -2)
        textstring3 = ListBox4.List(i)
        f.write "| EXTMIN: " & textstring3
        textstring4 = ListBox5.List(i)
        f.write "  | EXTMAX: " & textstring4
        textstring2 = ListBox2.List(i)
        f.write " | " & textstring2
        f.write Chr(13) + Chr(10)
        f.Close
    Next i
    
Dim RetVal
RetVal = Shell("C:\acad2002\vba\printlijst.bat", 1)    ' uitprinten textfile.


End Sub


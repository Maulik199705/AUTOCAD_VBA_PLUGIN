VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmvervangen 
   Caption         =   "Vervang Blok & Datum"
   ClientHeight    =   3576
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9012.001
   OleObjectBlob   =   "frmvervangen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmvervangen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'— Late-bind the AutoCAD application so you don't need a specific reference —
Private mappAcad As Object
Private mstrPath As String

#If VBA7 Then
  Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
      ByVal lpClassName As String, _
      ByVal lpWindowName As String _
  ) As LongPtr

  Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
      ByVal hWnd As LongPtr _
  ) As Long

  Private Declare PtrSafe Function GetMenuItemCount Lib "user32" ( _
      ByVal hMenu As LongPtr _
  ) As Long

  Private Declare PtrSafe Function GetSystemMenu Lib "user32" ( _
      ByVal hWnd As LongPtr, _
      ByVal bRevert As Long _
  ) As LongPtr

  Private Declare PtrSafe Function RemoveMenu Lib "user32" ( _
      ByVal hMenu As LongPtr, _
      ByVal nPosition As Long, _
      ByVal wFlags As Long _
  ) As Long
#Else
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
      ByVal lpClassName As String, _
      ByVal lpWindowName As String _
  ) As Long

  Private Declare Function DrawMenuBar Lib "user32" ( _
      ByVal hWnd As Long _
  ) As Long

  Private Declare Function GetMenuItemCount Lib "user32" ( _
      ByVal hMenu As Long _
  ) As Long

  Private Declare Function GetSystemMenu Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal bRevert As Long _
  ) As Long

  Private Declare Function RemoveMenu Lib "user32" ( _
      ByVal hMenu As Long, _
      ByVal nPosition As Long, _
      ByVal wFlags As Long _
  ) As Long
#End If

Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000
Private Sub CommandButton3_Click()
    Dim acadApp     As Object
    Dim doc         As Object
    Dim fdlCol      As Object
    Dim fdl         As Object
    Dim info        As String

    ' 1) Get the running AutoCAD (or fail)
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "Could not find a running instance of AutoCAD.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' 2) Open the drawing and zoom all
    Set doc = acadApp.Documents.Open("C:\acad2002\p08-01072.dwg")
    acadApp.ZoomAll

    ' 3) Late-bind the FileDependencies property
    On Error Resume Next
    Set fdlCol = CallByName(doc, "FileDependencies", VbGet)
    On Error GoTo 0

    If fdlCol Is Nothing Then
        MsgBox "FileDependencies is not supported in this AutoCAD version.", vbInformation
        Exit Sub
    End If

    ' Show how many entries there are
    MsgBox "The number of entries in the File Dependency List is " & _
           CallByName(fdlCol, "Count", VbGet) & "."

    ' 4) Loop and display each dependency
    For Each fdl In fdlCol
        info = ""
        info = info & "Affects graphics?: " & vbTab & CallByName(fdl, "AffectsGraphics", VbGet) & vbCrLf
        info = info & "Feature:         " & vbTab & CallByName(fdl, "Feature", VbGet) & vbCrLf
        info = info & "FileName:        " & vbTab & CallByName(fdl, "FileName", VbGet) & vbCrLf
        info = info & "FileSize:        " & vbTab & CallByName(fdl, "FileSize", VbGet) & vbCrLf
        info = info & "Fingerprint GUID:" & vbTab & CallByName(fdl, "FingerprintGuid", VbGet) & vbCrLf
        info = info & "FoundPath:       " & vbTab & CallByName(fdl, "FoundPath", VbGet) & vbCrLf
        info = info & "FullFileName:    " & vbTab & CallByName(fdl, "FullFileName", VbGet) & vbCrLf
        info = info & "Index:           " & vbTab & CallByName(fdl, "Index", VbGet) & vbCrLf
        info = info & "Modified?:       " & vbTab & CallByName(fdl, "IsModified", VbGet) & vbCrLf
        info = info & "ReferenceCount:  " & vbTab & CallByName(fdl, "ReferenceCount", VbGet) & vbCrLf
        info = info & "Timestamp:       " & vbTab & CallByName(fdl, "TimeStamp", VbGet) & vbCrLf
        info = info & "Version GUID:    " & vbTab & CallByName(fdl, "VersionGuid", VbGet)

        MsgBox info, vbInformation, "Dependency #" & CallByName(fdl, "Index", VbGet)
    Next
End Sub




Private Sub OptionButton1_Click()
  ComboBox1.Enabled = True
  ComboBox1.ListIndex = 0
End Sub

Private Sub OptionButton2_Click()
   ComboBox1.Enabled = True
  ComboBox1.ListIndex = 8
End Sub

Private Sub OptionButton3_Click()
   ComboBox1.Enabled = True
  ComboBox1.ListIndex = 0
End Sub

Private Sub OptionButton4_Click()
   ComboBox1.Enabled = True
  ComboBox1.ListIndex = 0
End Sub

Private Sub OptionButton5_Click()
   ComboBox1.Enabled = True
  ComboBox1.ListIndex = 0
End Sub

Private Sub OptionButton6_Click()
   ComboBox1.Enabled = True
   ComboBox1.ListIndex = 0
End Sub

''''
''''Sub ERASEBLOK()
''''Dim layerelement As Object
''''For Each layerelement In ThisDrawing.ModelSpace
''''  If layerelement.layer = "MONTAGEBLOK" Then layerelement.Erase
''''Update
''''Next layerelement
''''End Sub

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
  'Set mappAcad = GetObject(Class:="AutoCAD.Application")
  On Error Resume Next
    Set mappAcad = GetObject(Class:="AutoCAD.Application")
    If mappAcad Is Nothing Then
        MsgBox "Could not get AutoCAD.Application"
        Unload Me
        Exit Sub
    End If
    On Error GoTo 0
  
 ' now populate your combo-box exactly as before…
    With Me.ComboBox1
      .AddItem "0"
      .AddItem "WIJZIGING1"
      .AddItem "WIJZIGING2"
      .AddItem "WIJZIGING3"
      .AddItem "WIJZIGING4"
      .AddItem "WIJZIGING5"
      .AddItem "WIJZIGING6"
      .AddItem "WIJZIGING7"
      .AddItem "REVISIE"
      .ListIndex = 0
    End With
End Sub
Private Sub CommandButton1_Click()
lstDrawings.Clear  'reset button
ListBox1.Clear
OptionButton1.Value = False
OptionButton2.Value = False
OptionButton3.Value = False
OptionButton4.Value = False
OptionButton5.Value = False
OptionButton6.Value = False
ComboBox1.ListIndex = 0
End Sub
Private Sub UserForm_Terminate()
  'clean up
  Set mappAcad = Nothing

End Sub
Private Sub cmdClose_Click()

  'close this form
  Call Unload(Me)

End Sub

Private Sub cmdSelect_Click()
    Dim cd          As Object
    Dim strFileNames As String
    Dim varFiles    As Variant
    Dim i           As Long
    
    ' Late-bind the VB6 CommonDialog control
    On Error Resume Next
    Set cd = CreateObject("MSComDlg.CommonDialog")
    If cd Is Nothing Then
        MsgBox "MSComDlg32.ocx not available. Please register the Common Dialog control.", _
               vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Configure and show the multi-select Open dialog
    With cd
        .Filter = "AutoCAD Drawings (*.dwg)|*.dwg"
        .FilterIndex = 1
        .MaxFileSize = 6000
        ' &H200 = OFN_ALLOWMULTISELECT, &H80000 = OFN_EXPLORER
        .Flags = &H200 Or &H80000
        .InitDir = mstrPath               ' remember last folder if you like
        .ShowOpen
        strFileNames = .FileName
    End With
    
    ' Nothing picked? bail out.
    If Len(strFileNames) = 0 Then Exit Sub
    
    ' Split the result on vbNullChar ? first element is folder, the rest are file names
    varFiles = Split(strFileNames, vbNullChar)
    
    ' Clear your two listboxes and fill them again
    Me.lstDrawings.Clear
    Me.ListBox1.Clear                  ' if you used a second ListBox for the path
    
    For i = 1 To UBound(varFiles)
        Me.lstDrawings.AddItem varFiles(i)
        Me.ListBox1.AddItem varFiles(0) & "\"     ' show the path next to each file
    Next i
    
    ' Update your TextBox1 with how many you picked
    Me.TextBox1 = Me.lstDrawings.ListCount
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
ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "*" & vbCr & vbCr

'ThisDrawing.SendCommand "setvar" & vbCr & "acadlspasdoc" & vbCr & "0" & vbCr
  Dim intDrawing As Integer
  Dim intDrawings As Integer
  
  'check whether anyt valid files are selected
  If Len(mstrPath) <> 0 Then
    intDrawings = Me.lstDrawings.ListCount - 1
    If intDrawings >= 0 Then

      For intDrawing = 0 To intDrawings
      Call PrepareAsBackGround(ListBox1.List(intDrawing) & Me.lstDrawings.List(intDrawing))
      Next intDrawing
    End If
 
  End If


  Call Unload(Me)
'ThisDrawing.SendCommand "setvar" & vbCr & "acadlspasdoc" & vbCr & "1" & vbCr
End Sub
Sub ERASEBLOK()
 Dim element2
 For Each element2 In ThisDrawing.ModelSpace
      If element2.layer = "MONTAGEBLOK" Then
         element2.Erase
      End If
      Update
     Next element2
End Sub
Private Sub PrepareAsBackGround(strDrawingFullname As String)
    On Error GoTo ErrHandler
    
    '– local variables –
    Dim objDocument As Object                ' AcadDocument
    Dim element    As Object
    Dim attribs    As Variant
    Dim attrib     As Object
    Dim naamblok   As String
    Dim vINS As Doubl
    Dim i          As Long
    
    '– choose which logo/block to insert based on the radio buttons –
    If Me.OptionButton1.Value Then
        naamblok = "C:\ACAD2002\DWG\definitief2.dwg"
    ElseIf Me.OptionButton2.Value Then
        naamblok = "C:\ACAD2002\DWG\bl-revisie.dwg"
    ElseIf Me.OptionButton3.Value Then
        naamblok = "C:\ACAD2002\DWG\goedkeuring.dwg"
    ElseIf Me.OptionButton4.Value Then
        naamblok = "C:\ACAD2002\DWG\voorlopig.dwg"
    ElseIf Me.OptionButton5.Value Then
        naamblok = "C:\ACAD2002\DWG\uitvoering.dwg"
    ElseIf Me.OptionButton6.Value Then
        naamblok = "C:\ACAD2002\DWG\allesnaregelen.dwg"
    Else
        ' no option selected: nothing to do
        Exit Sub
    End If
    
    '– only proceed if the file actually exists –
    If Not FileExists(strDrawingFullname) Then Exit Sub
    
    '– open it as a background document –
    Set objDocument = mappAcad.Documents.Open(strDrawingFullname)
    
    '– update the date/block attributes in the new document –
    Call datum
    
    '– switch to layer GT in the current drawing –
    ThisDrawing.SendCommand "-layer" & vbCr & "Set" & vbCr & "GT" & vbCr & vbCr
    
    '– optionally erase existing montage-blocks –
    If Me.CheckBox1.Value Then ERASEBLOK
    
    '– purge all unused elements –
    ThisDrawing.SendCommand "-purge" & vbCr & "All" & vbCr & "*" & vbCr & "N" & vbCr
    
    '– remove any old logo-blocks and insert the new one at the same point –
    For Each element In objDocument.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            Select Case LCase(element.Name)
                Case "allesnaregelen", "definitief2", "definitief", _
                     "goedkeuring", "voorlopig", "bl-revisie", "uitvoering"
                     
                     ' capture the insertion point, erase the old block…
                     PE2 = element.InsertionPoint
                     element.Erase
                     
                     ' nudge up by 77 and insert new block
                     PE2(1) = PE2(1) + 77
                     objDocument.ModelSpace.InsertBlock PE2, naamblok, 1, 1, 1, 0
            End Select
        End If
    Next
    
    '– fill any “XXGROEPEN” blocks with default numeric attributes –
    For Each element In objDocument.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If element.Name Like "##GROEPEN" Then
                attribs = element.GetAttributes
                For i = LBound(attribs) To UBound(attribs)
                    Set attrib = attribs(i)
                    Select Case attrib.TagString
                        Case "01", "02", "03", "04", "05", _
                             "06", "07", "08", "09", "10", _
                             "11", "12", "13", "14", "15", _
                             "16", "17", "18", "19", "20"
                             If attrib.TextString = "" Then
                                 attrib.TextString = CStr(Val(attrib.TagString))
                             End If
                    End Select
                Next
            End If
        End If
    Next
    
    '– save and close the background document –
    objDocument.Close SaveChanges:=True
    
    '– decrement your counter on the form –
    Me.frmvervangen.TextBox1 = Val(Me.frmvervangen.TextBox1) - 1
    If Val(Me.TextBox1) <> 0 Then
        Me.frmvervangen.Caption = _
          "Vervang Blok & Datum...nog " & Me.frmvervangen.TextBox1 & " tekening(en)."
    End If

Cleanup:
    Set objDocument = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Sub datum()
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "JMASI" Then lognaam = "JM"
If lognaam = "GERARD" Then lognaam = "GCH"
If lognaam = "ILONA" Then lognaam = "IK"
If lognaam = "BJORN" Then lognaam = "BC"
If lognaam = "DENNIS" Then lognaam = "DvdW"
If lognaam = "ZILVERSCHOONJ" Then lognaam = "JZ"
If lognaam = "BGOUW" Then lognaam = "BG"
If lognaam = "OYILM" Then lognaam = "OY"
If lognaam = "SNABI" Then lognaam = "SN"
If lognaam = "DLALI" Then lognaam = "DL"
If lognaam = "GLUII" Then lognaam = "GL"
If lognaam = "DWILS" Then lognaam = "DW"
If lognaam = "JPRINS" Then lognaam = "JP"
datumacad1 = ThisDrawing.GetVariable("cdate")
datumacad = Left(datumacad1, 8)

dag = Right(datumacad, 2)
maand = Left(datumacad, 6)
maand2 = Right(maand, 2)
jaar = Left(datumacad, 4)

kdate = dag & "-" & maand2 & "-" & jaar & "|" & lognaam

If ComboBox1 <> "0" Then
 wyz = ComboBox1

    Dim element As Object
    For Each element In ThisDrawing.ModelSpace
          If element.ObjectName = "AcDbBlockReference" Then
          If element.Name = "Kaderlogo" Then
          Set SYMBOOL = element
            If SYMBOOL.HasAttributes Then
            ATTRIBUTEN = SYMBOOL.GetAttributes
            For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
            Set ATTRIBUUT = ATTRIBUTEN(i)
                 If ATTRIBUUT.TagString = ComboBox1 Then ATTRIBUUT.TextString = kdate
            Next i
    
            End If
          End If
          End If
      Next element
End If
End Sub
Private Sub CommandButton2_Click()
Dim controle1
Dim controle2
Dim controle3
Dim controle4

frmvervangen.Hide
Dim pe1(0 To 2) As Double
Dim PE2(0 To 2) As Double
Dim pe6(0 To 2) As Double
Dim insp
Dim pe3 As String
Dim pe4 As String

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
          controle1 = InStr(1, element.Name, "h0.kl.1", vbTextCompare)
          controle2 = InStr(1, element.Name, "h1.kl.1", vbTextCompare)
          controle3 = InStr(1, element.Name, "h2.kl.1", vbTextCompare)
          controle4 = InStr(1, element.Name, "h3.kl.1", vbTextCompare)
           
          
          If controle1 <> 0 Or controle2 <> 0 Or controle3 <> 0 Or controle4 <> 0 Then
            insp = element.InsertionPoint
          pe1(0) = insp(0) + 7000
          pe1(1) = insp(1) + 10000
          pe1(2) = 0
          
          PE2(0) = insp(0) - 7000
          PE2(1) = insp(1) - 10000
          PE2(2) = 0
          
          pe6(0) = insp(0) + 2000
          pe6(1) = insp(1) - 9800
          pe6(2) = 0
          
          
          'MsgBox pe2
          
          pe3 = pe1(0) & "," & pe1(1)
          pe4 = PE2(0) & "," & PE2(1)
          'MsgBox pe3
          dwgnm = ThisDrawing.GetVariable("dwgname")
          Dim textObj As AcadText
          Set textObj = ThisDrawing.ModelSpace.AddText(dwgnm, pe6, 100)

          'ThisDrawing.SendCommand "text" & vbCr & pe4 & vbCr & "100" & vbCr & "90" & vbCr & dwgnm & vbCr
          ThisDrawing.SendCommand "-plot" & vbCr & "Y" & vbCr & "Model" & vbCr & "\\FSERVER2\P39 L2100TN" & vbCr & "A4" & vbCr & "M" & vbCr & "P" & vbCr & "N" & vbCr & "W" & vbCr & pe3 & vbCr & pe4 & vbCr & "F" & vbCr & "0,0" & vbCr & "Y" & vbCr & "wth.ctb" & vbCr & "Y" & vbCr & "N" & vbCr & "N" & vbCr & "Y" & vbCr & "Y" & vbCr
         
          
          
                                                                                                                                                                                                                                          

''          Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(insp, naamblok, 1, 1, 1, 0)
          End If
        
        End If
     Next element

End Sub




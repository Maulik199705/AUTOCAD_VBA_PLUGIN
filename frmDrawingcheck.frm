VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDrawingcheck 
   Caption         =   "Controleer de tekening"
   ClientHeight    =   1215
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   4032
   OleObjectBlob   =   "frmDrawingcheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDrawingcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
'begin sluiten toets uitschakelen
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
End Sub
Private Sub cmdAfsluiten_Click()
Unload Me
End Sub

Private Sub go_Click()
Dim element500 As Object
AZ = 0
For Each element500 In ThisDrawing.ModelSpace
      If element500.ObjectName = "AcDbBlockReference" Then
         If UCase(element500.Name) = "MAT_SPE_ZD" Or UCase(element500.Name) = "MAT_SPE_PE" _
         Or UCase(element500.Name) = "MAT_SPE_ZD_1627" Or UCase(element500.Name) = "MAT_SPE_PE800" Then AZ = AZ + 1
      End If
Next element500
On Error Resume Next

For I = 1 To AZ

Dim j As String
j = I
If j > 0 And j < 10 Then k = "0" & j
If j > 9 Then k = j

Dim TRIMSTRING
For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If UCase(element2.Name) = "MAT_SPE_ZD" Or UCase(element2.Name) = "MAT_SPE_PE" _
      Or UCase(element2.Name) = "MAT_SPE_ZD_1627" Or UCase(element2.Name) = "MAT_SPE_PE800" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For O = LBound(attributen) To UBound(attributen)
         Set attribuut = attributen(O)
               If attribuut.TagString = "RNU" Then
                   hk = attribuut.textstring
                   If hk = k Then
                        b0 = element2.InsertionPoint
                        c0 = b0
                        For y = LBound(attributen) To UBound(attributen)
                           Set attribuut = attributen(y)
                             If hk = k Then
                              If attribuut.TagString = "REGELUNITTYPE" Then
                              RT = attribuut.textstring
                     
                               TRIMSTRING = Split(RT, (" "))
                               mystr = Len(TRIMSTRING(1))
                                   If mystr > 2 Then
                                   trimstring2 = Split(TRIMSTRING(1), ("/"))
                                   TextBox1 = trimstring2(0)
                                   Else
                                   TextBox1 = TRIMSTRING(1)
                                   End If
                              End If
                             End If
                             
    'aanvulling under construction--------------------------------------------------------------
             If attribuut.TagString = "WTH250" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 250
             End If
             If attribuut.TagString = "WTH165" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 165
             End If
             If attribuut.TagString = "WTH125" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 125
             End If
             If attribuut.TagString = "WTH105" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 105
             End If
             If attribuut.TagString = "WTH90" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 90
             End If
             If attribuut.TagString = "WTH75" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 75
             End If
             If attribuut.TagString = "WTH63" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 63
             End If
             If attribuut.TagString = "WTH50" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 50
             End If
             If attribuut.TagString = "WTH40" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 40
             End If
             
             ' pe-rt 120,90,60
             If attribuut.TagString = "PE120" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 120
             End If
             If attribuut.TagString = "PE90" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 90
             End If
             If attribuut.TagString = "PE60" And attribuut.textstring <> "-" Then
             yt = attribuut.textstring
             TextBox4 = Val(TextBox4) + yt * 60
             End If
       'aanvulling under construction--------------------------------------------------------------
                       Next y
                   End If
               End If
       Next O
       
      End If
      End If
      End If
 Next element2
     


For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOKNEW" Then
                Set symbool = element
                If symbool.HasAttributes Then
                    attributen = symbool.GetAttributes
                    For m = LBound(attributen) To UBound(attributen)
                         Set attribuut = attributen(m)
                         If attribuut.TagString = "UNITNUMMER" Then
                            GH = attribuut.textstring
                            If GH = k Then
                             For z = LBound(attributen) To UBound(attributen)
                               Set attribuut = attributen(z)
                               If GH = k Then
                                If attribuut.TagString = "ROLLENGTE" And attribuut.textstring <> " " Then
                                TextBox2 = TextBox2 + 1
                                qwer = Split(attribuut.textstring)
                                TextBox3 = Val(TextBox3) + qwer(0)
                                End If
                               End If
                             Next z
                            End If
                         End If
                    Next m
                End If
            End If
        End If

    Next element



If TextBox1 <> TextBox2 Then
   MsgBox "Het totaal aantal groepen in het unit bloklogo" & (Chr(13) & Chr(10)) & _
          "en het totaal aantal groeptekstblokken" & (Chr(13) & Chr(10)) & _
          "van " & "UNIT: " & k & " wijken af.", vbExclamation
          
  Dim b1(0 To 2) As Double
  b1(0) = b0(0) - 1500 '1000
  b1(1) = b0(1) - 400  '500
  b1(2) = 0
 
  Dim b2(0 To 2) As Double
  b2(0) = b0(0) + 500
  b2(1) = b0(1) + 1000 '750
  b2(2) = 0
  ZoomWindow b1, b2
          
          
   Unload Me
   Exit Sub
End If

If TextBox3 <> TextBox4 Then
   MsgBox "Het aantal meters in het unit bloklogo" & (Chr(13) & Chr(10)) & _
          "en het aantal meters van de groeptekstblokken" & (Chr(13) & Chr(10)) & _
          "van " & "UNIT: " & k & " wijken af.", vbCritical
          
  Dim c1(0 To 2) As Double
  c1(0) = c0(0) - 1500 '1000
  c1(1) = c0(1) - 400  '500
  c1(2) = 0
 
  Dim c2(0 To 2) As Double
  c2(0) = c0(0) + 500
  c2(1) = c0(1) + 1000 '750
  c2(2) = 0
  ZoomWindow c1, c2
          
          
   Unload Me
   Exit Sub
End If


If TextBox1 = TextBox2 Then
  TextBox1 = Clear
  TextBox2 = "0"
End If
If TextBox3 = TextBox4 Then
  TextBox3 = "0"
  TextBox4 = "0"
End If



Next I

Unload Me
End Sub


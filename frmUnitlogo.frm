VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnitlogo 
   Caption         =   "Bloklogo's uitlezen"
   ClientHeight    =   7212
   ClientLeft      =   48
   ClientTop       =   540
   ClientWidth     =   15432
   OleObjectBlob   =   "frmUnitlogo.frx":0000
End
Attribute VB_Name = "frmUnitlogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'19-08-2003
'M.Bosch en G.C.Haak
'BESTANDEN:
'unitlogo_wthzd.dwg
'unitlogo_pert.dwg
'unitlogo_alu.dwg
'unitlogo_extragroepen


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

Private Sub CheckBox4_Click()
If CheckBox4.Value = False Then TextBox70 = "0"
End Sub
Private Sub CheckBox5_Click()
If CheckBox5.Value = False Then TextBox71 = "0"
End Sub

Private Sub TextBox70_Change()
If frmUnitlogo.TextBox70 <> "0" Then frmUnitlogo.CheckBox4 = True
End Sub

Private Sub TextBox71_Change()
If frmUnitlogo.TextBox71 <> "0" Then CheckBox5.Value = True
End Sub

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
 frmUnitlogo.Height = 250 '282
 'frmUnitlogo.Width = 698
 Label28.Caption = Clear: Label30.Caption = Clear
 Label25.Caption = Clear: Label29.Caption = Clear
 Label34.Caption = Clear
TextBox1.SetFocus


Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "GERARD" Then
    CheckBox2.Visible = True
    CommandButton1.SetFocus
'    CheckBox3.Visible = True
End If

End Sub
Private Sub cmdschets_Click()
On Error Resume Next
Call Schaal(scaal)
ThisDrawing.SetVariable "osmode", 1
frmUnitlogo.Hide


Dim PBEGIN(0 To 2) As Double
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "ENG-Kaderlogo" Then
          insp = element.InsertionPoint
          PBEGIN(0) = insp(0) - (940 * scaal)
          PBEGIN(1) = insp(1) + (298.7 * scaal)
          PBEGIN(2) = insp(2)
          End If
      Update
      End If
     Next element


'PBEGIN = ThisDrawing.Utility.GetPoint(, "Plaats beginpunt -Materiaalstaat-")
Dim bestand93 As String
bestand93 = "C:\ACAD2002\DWG\unitlogo_WTHZD.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand93, scaal, scaal, 1, 0)

Dim pb3(0 To 2) As Double

 pb3(0) = PBEGIN(0)
 pb3(1) = PBEGIN(1) + (254.5 * scaal)
 pb3(2) = 0

 Dim bestand94 As String
bestand94 = "C:\ACAD2002\DWG\bl-krimp.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb3, bestand94, scaal, scaal, 1, 0)
If Err Then
    frmUnitlogo.Show
    Exit Sub
    End If
ThisDrawing.SetVariable "osmode", 0
End Sub

Private Sub CommandButton1_Click()
Call invullen
If CheckBox1.Value = True Then frmUnitlogo.Width = 780
If TextBox40 <> "0" Or TextBox46 <> "0" Or TextBox50 <> "0" Or TextBox54 <> "0" Or TextBox55 <> "0" _
Or TextBox56 <> "0" Or TextBox60 <> "0" Or TextBox69 <> "0" Or TextBox70 <> "0" Or TextBox71 <> "0" Then
Logoplaats.Enabled = True
cmdschets.Enabled = False
Else
Logoplaats.Enabled = False
cmdschets.Enabled = True
End If

If Label51 = "PE-RT16" Then Frame6.Enabled = False

Dim lognaam
lognaam = ThisDrawing.GetVariable("loginname")
lognaam = UCase(lognaam)
If lognaam = "GERARD" And CheckBox2.Value = True Then
    Call Logoplaats_Click
End If

End Sub
Sub invullen()
On Error Resume Next
Dim element As Object
a = 0
b = 0
grtwaarde = 0
'----WTH-ZD-----20 * 3,4 ------------------------------------------------------
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_ZD" Or element.Name = "Mat_spe_ZDringleiding" Then
      Set symbool = element
      b = b + 1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)

        If attribuut.TagString = "RNU" Then a = attribuut.textstring
        If a > 10 Then grtwaarde = 1
        If a > 16 Then message = 1
        
        If attribuut.TagString = "WTH250" And attribuut.textstring <> "" Then _
        TextBox17 = TextBox17 + (Val(attribuut.textstring))  '250 METER
        If attribuut.TagString = "WTH165" And attribuut.textstring <> "" Then _
        TextBox18 = TextBox18 + (Val(attribuut.textstring))  '165 METER
        If attribuut.TagString = "WTH125" And attribuut.textstring <> "" Then _
        TextBox19 = TextBox19 + (Val(attribuut.textstring))  '125 METER
        If attribuut.TagString = "WTH105" And attribuut.textstring <> "" Then _
        TextBox20 = TextBox20 + (Val(attribuut.textstring))  '105 METER
        If attribuut.TagString = "WTH90" And attribuut.textstring <> "" Then _
        TextBox21 = TextBox21 + (Val(attribuut.textstring))  '90 METER
        If attribuut.TagString = "WTH75" And attribuut.textstring <> "" Then _
        TextBox22 = TextBox22 + (Val(attribuut.textstring))  '75 METER
        If attribuut.TagString = "WTH63" And attribuut.textstring <> "" Then _
        TextBox23 = TextBox23 + (Val(attribuut.textstring))  '63 METER
        If attribuut.TagString = "WTH50" And attribuut.textstring <> "" Then _
        TextBox41 = TextBox41 + (Val(attribuut.textstring)) '50 METER
        If attribuut.TagString = "WTH40" And attribuut.textstring <> "" Then _
        TextBox42 = TextBox42 + (Val(attribuut.textstring))   '40 METER
        If attribuut.TagString = "WTHZD" Then
         Label50 = "WTH-ZD"
         'Label30 = "ZD"
        End If
        
     
        Next j
       
        End If
      End If
      End If
  Next element


'----WTH-ZD--16 * 2,7--+ 500 meter rollen werkelijke lengte -------------------------------------------------------
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_ZD_1627" Or element.Name = "Mat_spe_ZD_1627500" Then
      Set symbool = element
      b = b + 1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)

        If attribuut.TagString = "RNU" Then a = attribuut.textstring
        If a > 10 Then grtwaarde = 1
        If a > 16 Then message = 1
        
        If attribuut.TagString = "WTH105" And attribuut.textstring <> "" Then _
        TextBox65 = TextBox65 + (Val(attribuut.textstring))  '105 METER
        If attribuut.TagString = "WTH90" And attribuut.textstring <> "" Then _
        TextBox66 = TextBox66 + (Val(attribuut.textstring))  '90 METER
        If attribuut.TagString = "WTH75" And attribuut.textstring <> "" Then _
        TextBox67 = TextBox67 + (Val(attribuut.textstring))  '75 METER
        If attribuut.TagString = "WTH63" And attribuut.textstring <> "" Then _
        TextBox68 = TextBox68 + (Val(attribuut.textstring))  '63 METER
        
        If attribuut.TagString = "WTHZD" Then
         Label71 = "WTH-ZD_1627"
         'Label30 = "ZD"
        End If
               
        If attribuut.TagString = "LMETER" And attribuut.textstring <> "" Then C = attribuut.textstring
             
        Next j
        
        frmUnitlogo.TextBox71 = Val(frmUnitlogo.TextBox71) + C
       
        End If
      End If
      End If
  Next element


'----FLEXFIX-----------------------------------------------------------
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_FLEX" Or element.Name = "Mat_spe_FLEX_Aankoppel" Then
      Set symbool = element
      b = b + 1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)

        If attribuut.TagString = "RNU" Then a = attribuut.textstring
        If a > 10 Then grtwaarde = 1
        If a > 16 Then message = 1
        
        If attribuut.TagString = "FLEX_MATTEN" And attribuut.textstring <> "" Then _
        TextBox55 = TextBox55 + (Val(attribuut.textstring))  'MATTEN
        If attribuut.TagString = "FLEX_METERS" And attribuut.textstring <> "" Then _
        TextBox56 = TextBox56 + (Val(attribuut.textstring))  'METERS
      
        Next j
       
        End If
      End If
      End If
  Next element

'----PE-RT----16*2--14*2-----------------------------------------------------
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_PE" Or element.Name = "Mat_spe_PEringleiding" Or element.Name = "Mat_spe_PE800" Then
      Set symbool = element
      b = b + 1
      If element.Name = "Mat_spe_PE800" Then CheckBox1.Value = True
      If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For k = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(k)
        If attribuut.TagString = "RNU" Then a = attribuut.textstring
        If a > 10 Then grtwaarde = 1
        If a > 16 Then message = 1
        
        
       If attribuut.TagString = "PE" And attribuut.textstring = "PE-RT 16*2 mm" Then
        For l = LBound(attributen) To UBound(attributen)
           Set attribuut = attributen(l)
        If attribuut.TagString = "PE120" And attribuut.textstring <> "" Then TextBox43 = TextBox43 + (Val(attribuut.textstring))  '120 METER
        If attribuut.TagString = "PE90" And attribuut.textstring <> "" Then TextBox44 = TextBox44 + (Val(attribuut.textstring))  '90 METER
        If attribuut.TagString = "PE60" And attribuut.textstring <> "" Then TextBox45 = TextBox45 + (Val(attribuut.textstring))  '60 METER
        Label51 = "PE-RT16"
        Next l
       End If
      
      If attribuut.TagString = "PE" And attribuut.textstring = "PE-RT 14*2 mm" Then
        For m = LBound(attributen) To UBound(attributen)
           Set attribuut = attributen(m)
        If attribuut.TagString = "PE90" And attribuut.textstring <> "" Then TextBox52 = TextBox52 + (Val(attribuut.textstring))  '90 METER
        If attribuut.TagString = "PE60" And attribuut.textstring <> "" Then TextBox53 = TextBox53 + (Val(attribuut.textstring))  '60 METER
        Label52 = "PE-RT14"
        Next m
      End If
        
      If attribuut.TagString = "LMETER" And attribuut.textstring <> "" Then b = attribuut.textstring
          
 
          
       
    Next k
       frmUnitlogo.TextBox70 = Val(frmUnitlogo.TextBox70) + b
   End If
   End If
   End If
  Next element
  
  
  


'----ALUFLEX-----------------------------------------------------------
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "Mat_spe_ALU" Or element.Name = "Mat_spe_ALUringleiding" Then
      Set symbool = element
      b = b + 1
      If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For t = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(t)
        If attribuut.TagString = "RNU" Then a = attribuut.textstring
        If a > 10 Then grtwaarde = 1
        If a > 16 Then message = 1
        
        If attribuut.TagString = "ALU200" And attribuut.textstring <> "" Then _
        TextBox47 = TextBox47 + (Val(attribuut.textstring))  '200 METER
        If attribuut.TagString = "ALU100" And attribuut.textstring <> "" Then _
        TextBox48 = TextBox48 + (Val(attribuut.textstring))  '100 METER
        If attribuut.TagString = "ALU50" And attribuut.textstring <> "" Then _
        TextBox49 = TextBox49 + (Val(attribuut.textstring))  '50 METER
        
        
        If attribuut.TagString = "ALU" Then
          Label53 = "ALUFLEX"
          Label29 = attribuut.textstring
        End If
        
        Next t
       
   End If
   End If
   End If
  Next element

'frmUnitlogo.Caption = "Bloklogo's uitlezen  " & L29 & L30 & L31 & L32 'leidingsoort
Label28.Caption = a
Label34.Caption = grtwaarde
If a > 10 Or b > 10 Or grtwaarde = 1 Then
   frmUnitlogo.Height = 385
End If
  
   
'If rol250 = 0 Then rol250 = "-"
'If rol165 = 0 Then rol165 = "-"
'If rol125 = 0 Then rol125 = "-"
'If rol105 = 0 Then rol105 = "-"
'If rol90 = 0 Then rol90 = "-"
'If rol75 = 0 Then rol75 = "-"
'If rol63 = 0 Then rol63 = "-"
'If rol50 = 0 Then rol50 = "-"
'If rol40 = 0 Then rol40 = "-"

'If PE120 = 0 Then PE120 = "-"
'If PE90 = 0 Then PE90 = "-"
'If PE60 = 0 Then PE60 = "-"

'---WTH-ZD----------------------------------------------------
For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If element.Name = "Mat_spe_ZD" Or element.Name = "Mat_spe_PE" Or element.Name = "Mat_spe_PE800" _
            Or element.Name = "Mat_spe_ALU" Or element.Name = "Mat_spe_ZDringleiding" Or element.Name = "Mat_spe_PEringleiding" _
            Or element.Name = "Mat_spe_ALUringleiding" Or element.Name = "Mat_spe_FLEX" Or element.Name = "Mat_spe_ZD_1627" _
            Or element.Name = "Mat_spe_ZD_1627500" Then
                Set symbool = element
                If symbool.HasAttributes Then
                    attributen = symbool.GetAttributes
                    For I = LBound(attributen) To UBound(attributen)
                         Set attribuut = attributen(I)
                '--------------------------unit 1 t/m 10
                If TextBox1 = "" Then
                  If attribuut.TagString = "RNU" Then t1 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t1 = "1" Or t1 = "01") Then TextBox1 = attribuut.textstring
                End If
                If TextBox2 = "" Then
                  If attribuut.TagString = "RNU" Then t2 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t2 = "2" Or t2 = "02") Then TextBox2 = attribuut.textstring
                End If
                If TextBox3 = "" Then
                  If attribuut.TagString = "RNU" Then t3 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t3 = "3" Or t3 = "03") Then TextBox3 = attribuut.textstring
                End If
                If TextBox4 = "" Then
                  If attribuut.TagString = "RNU" Then t4 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t4 = "4" Or t4 = "04") Then TextBox4 = attribuut.textstring
                End If
                If TextBox5 = "" Then
                  If attribuut.TagString = "RNU" Then t5 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t5 = "5" Or t5 = "05") Then TextBox5 = attribuut.textstring
                End If
                If TextBox6 = "" Then
                  If attribuut.TagString = "RNU" Then t6 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t6 = "6" Or t6 = "06") Then TextBox6 = attribuut.textstring
                End If
                 If TextBox7 = "" Then
                  If attribuut.TagString = "RNU" Then t7 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t7 = "7" Or t7 = "07") Then TextBox7 = attribuut.textstring
                End If
                If TextBox8 = "" Then
                  If attribuut.TagString = "RNU" Then t8 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t8 = "8" Or t8 = "08") Then TextBox8 = attribuut.textstring
                End If
                'regelunit 9
                If TextBox24 = "" Then
                  If attribuut.TagString = "RNU" Then t9 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And (t9 = "9" Or t9 = "09") Then TextBox24 = attribuut.textstring
                End If
                'regelunit 10
                If TextBox25 = "" Then
                  If attribuut.TagString = "RNU" Then t10 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And t10 = "10" Then TextBox25 = attribuut.textstring
                End If
                
                If TextBox9 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t1 = "1" Or t1 = "01") Then TextBox9 = attribuut.textstring
                End If
                If TextBox10 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t2 = "2" Or t2 = "02") Then TextBox10 = attribuut.textstring
                End If
                If TextBox11 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t3 = "3" Or t3 = "03") Then TextBox11 = attribuut.textstring
                End If
                If TextBox12 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t4 = "4" Or t4 = "04") Then TextBox12 = attribuut.textstring
                End If
                If TextBox13 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t5 = "5" Or t5 = "05") Then TextBox13 = attribuut.textstring
                End If
                If TextBox14 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t6 = "6" Or t6 = "06") Then TextBox14 = attribuut.textstring
                End If
                If TextBox15 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t7 = "7" Or t7 = "07") Then TextBox15 = attribuut.textstring
                End If
                If TextBox16 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t8 = "8" Or t8 = "08") Then TextBox16 = attribuut.textstring
                End If
                '--------------------------unit 1 t/m 10
                 
                '--------------------------unit 11 t/m 16
                'If TextBox24 = "" Then
                '  If attribuut.TagString = "RNU" Then t9 = attribuut.TextString
                '  If attribuut.TagString = "REGELUNITTYPE" And (t9 = "9" Or t9 = "09") Then TextBox24 = attribuut.TextString
                'End If
                'If TextBox25 = "" Then
                '  If attribuut.TagString = "RNU" Then t10 = attribuut.TextString
                '  If attribuut.TagString = "REGELUNITTYPE" And t10 = "10" Then TextBox25 = attribuut.TextString
                'End If
                If TextBox26 = "" Then
                  If attribuut.TagString = "RNU" Then t11 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And t11 = "11" Then TextBox26 = attribuut.textstring
                End If
                If TextBox27 = "" Then
                  If attribuut.TagString = "RNU" Then t12 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And t12 = "12" Then TextBox27 = attribuut.textstring
                End If
                If TextBox28 = "" Then
                  If attribuut.TagString = "RNU" Then t13 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And t13 = "13" Then TextBox28 = attribuut.textstring
                End If
                If TextBox29 = "" Then
                  If attribuut.TagString = "RNU" Then t14 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And t14 = "14" Then TextBox29 = attribuut.textstring
                End If
                If TextBox30 = "" Then
                  If attribuut.TagString = "RNU" Then t15 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And t15 = "15" Then TextBox30 = attribuut.textstring
                End If
                If TextBox31 = "" Then
                  If attribuut.TagString = "RNU" Then t16 = attribuut.textstring
                  If attribuut.TagString = "REGELUNITTYPE" And t16 = "16" Then TextBox31 = attribuut.textstring
                End If
                If TextBox32 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And (t9 = "9" Or t9 = "09") Then TextBox32 = attribuut.textstring
                End If
                If TextBox33 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And t10 = "10" Then TextBox33 = attribuut.textstring
                End If
                If TextBox34 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And t11 = "11" Then TextBox34 = attribuut.textstring
                End If
                If TextBox35 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And t12 = "12" Then TextBox35 = attribuut.textstring
                End If
                If TextBox36 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And t13 = "13" Then TextBox36 = attribuut.textstring
                End If
                If TextBox37 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And t14 = "14" Then TextBox37 = attribuut.textstring
                End If
                If TextBox38 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And t15 = "15" Then TextBox38 = attribuut.textstring
                End If
                If TextBox39 = "" Then
                If attribuut.TagString = "BEVESTIGINGSTYPE" And t16 = "16" Then TextBox39 = attribuut.textstring
                End If
           
           Next I
              End If
           End If
        End If
    Next element
    
TextBox69 = (Val(TextBox65)) + (Val(TextBox66)) + (Val(TextBox67)) + (Val(TextBox68))
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
TextBox46 = (Val(TextBox43)) + (Val(TextBox44)) + (Val(TextBox45))
TextBox50 = (Val(TextBox47)) + (Val(TextBox48)) + (Val(TextBox49))
CommandButton1.Enabled = False

If CheckBox1.Value = True Then
   TextBox57 = TextBox43
   TextBox58 = TextBox44
   TextBox59 = TextBox45
   TextBox43 = "0"
   TextBox44 = "0"
   TextBox45 = "0"
   TextBox46 = "0"
   Label64 = 800
End If

If TextBox1 = "" Then TextBox1 = "-"
If TextBox2 = "" Then TextBox2 = "-"
If TextBox3 = "" Then TextBox3 = "-"
If TextBox4 = "" Then TextBox4 = "-"
If TextBox5 = "" Then TextBox5 = "-"
If TextBox6 = "" Then TextBox6 = "-"
If TextBox7 = "" Then TextBox7 = "-"
If TextBox8 = "" Then TextBox8 = "-"
If TextBox9 = "" Then TextBox9 = "-"
If TextBox10 = "" Then TextBox10 = "-"
If TextBox11 = "" Then TextBox11 = "-"
If TextBox12 = "" Then TextBox12 = "-"
If TextBox13 = "" Then TextBox13 = "-"
If TextBox14 = "" Then TextBox14 = "-"
If TextBox15 = "" Then TextBox15 = "-"
If TextBox16 = "" Then TextBox16 = "-"
If TextBox24 = "" Then TextBox24 = "-"
If TextBox25 = "" Then TextBox25 = "-"
If TextBox26 = "" Then TextBox26 = "-"
If TextBox27 = "" Then TextBox27 = "-"
If TextBox28 = "" Then TextBox28 = "-"
If TextBox29 = "" Then TextBox29 = "-"
If TextBox30 = "" Then TextBox30 = "-"
If TextBox31 = "" Then TextBox31 = "-"
If TextBox32 = "" Then TextBox32 = "-"
If TextBox33 = "" Then TextBox33 = "-"
If TextBox34 = "" Then TextBox34 = "-"
If TextBox35 = "" Then TextBox35 = "-"
If TextBox36 = "" Then TextBox36 = "-"
If TextBox37 = "" Then TextBox37 = "-"
If TextBox38 = "" Then TextBox38 = "-"
If TextBox39 = "" Then TextBox39 = "-"
If TextBox41 = "" Then TextBox41 = "-"
If TextBox42 = "" Then TextBox42 = "-"
If TextBox43 = "" Then TextBox43 = "-"
If TextBox44 = "" Then TextBox44 = "-"
If TextBox45 = "" Then TextBox45 = "-"
If TextBox46 = "" Then TextBox46 = "-"
If TextBox47 = "" Then TextBox41 = "-"
If TextBox48 = "" Then TextBox42 = "-"
If TextBox49 = "" Then TextBox43 = "-"
If TextBox50 = "" Then TextBox44 = "-"
If TextBox52 = "" Then TextBox45 = "-"
If TextBox53 = "" Then TextBox46 = "-"
If TextBox54 = "" Then TextBox41 = "-"
If TextBox55 = "" Then TextBox42 = "-"
If TextBox56 = "" Then TextBox43 = "-"
If TextBox57 = "" Then TextBox44 = "-"
If TextBox58 = "" Then TextBox45 = "-"
If TextBox59 = "" Then TextBox46 = "-"
If TextBox60 = "" Then TextBox45 = "-"
If TextBox61 = "" Then TextBox46 = "-"
If TextBox57 = "" Then TextBox44 = "-"


If message = 1 Then MsgBox "Pas op - het aantal units is groter dan 16..!!!!:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "1 - Unit 17 en hoger komt NIET in het kaderlogo terecht.!!!" & Chr(13) & "2 - Het aantal groepen klopt wel.", vbExclamation, "Foutmelding"
End Sub
Private Sub TextBox1_Change()
If TextBox1 = "" Then TextBox1 = "-"
End Sub
Private Sub TextBox2_Change()
If TextBox2 = "" Then TextBox2 = "-"
End Sub
Private Sub TextBox3_Change()
If TextBox3 = "" Then TextBox3 = "-"
End Sub
Private Sub TextBox4_Change()
If TextBox4 = "" Then TextBox4 = "-"
End Sub
Private Sub TextBox5_Change()
If TextBox5 = "" Then TextBox5 = "-"
End Sub
Private Sub TextBox6_Change()
If TextBox6 = "" Then TextBox6 = "-"
End Sub
Private Sub TextBox7_Change()
If TextBox7 = "" Then TextBox7 = "-"
End Sub
Private Sub TextBox8_Change()
If TextBox8 = "" Then TextBox8 = "-"
End Sub
Private Sub TextBox9_Change()
If TextBox9 = "" Then TextBox9 = "-"
End Sub
Private Sub TextBox10_Change()
If TextBox10 = "" Then TextBox10 = "-"
End Sub
Private Sub TextBox11_Change()
If TextBox11 = "" Then TextBox11 = "-"
End Sub
Private Sub TextBox12_Change()
If TextBox12 = "" Then TextBox12 = "-"
End Sub
Private Sub TextBox13_Change()
If TextBox13 = "" Then TextBox13 = "-"
End Sub
Private Sub TextBox14_Change()
If TextBox14 = "" Then TextBox14 = "-"
End Sub
Private Sub TextBox15_Change()
If TextBox15 = "" Then TextBox15 = "-"
End Sub
Private Sub TextBox16_Change()
If TextBox16 = "" Then TextBox16 = "-"
End Sub
Private Sub TextBox17_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox17 = "" Then TextBox17 = "0"
End Sub
Private Sub TextBox18_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox18 = "" Then TextBox18 = "0"
End Sub
Private Sub TextBox19_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox19 = "" Then TextBox19 = "0"
End Sub
Private Sub TextBox20_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox20 = "" Then TextBox20 = "0"
End Sub
Private Sub TextBox21_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox21 = "" Then TextBox21 = "0"
End Sub
Private Sub TextBox22_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox22 = "" Then TextBox22 = "0"
End Sub
Private Sub TextBox23_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox23 = "" Then TextBox23 = "0"
End Sub
Private Sub TextBox41_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox41 = "" Then TextBox41 = "0"
End Sub
Private Sub TextBox42_Change()
TextBox40 = (Val(TextBox17)) + (Val(TextBox18)) + (Val(TextBox19)) + (Val(TextBox20)) + (Val(TextBox21)) + (Val(TextBox22)) + (Val(TextBox23)) + (Val(TextBox41)) + (Val(TextBox42))
If TextBox42 = "" Then TextBox42 = "0"
End Sub
Private Sub TextBox24_Change()
If TextBox24 = "" Then TextBox24 = "-"
End Sub
Private Sub TextBox25_Change()
If TextBox25 = "" Then TextBox25 = "-"
End Sub
Private Sub TextBox26_Change()
If TextBox26 = "" Then TextBox26 = "-"
End Sub
Private Sub TextBox27_Change()
If TextBox27 = "" Then TextBox27 = "-"
End Sub
Private Sub TextBox28_Change()
If TextBox28 = "" Then TextBox28 = "-"
End Sub
Private Sub TextBox29_Change()
If TextBox29 = "" Then TextBox29 = "-"
End Sub
Private Sub TextBox30_Change()
If TextBox30 = "" Then TextBox30 = "-"
End Sub
Private Sub TextBox31_Change()
If TextBox31 = "" Then TextBox31 = "-"
End Sub
Private Sub TextBox32_Change()
If TextBox32 = "" Then TextBox32 = "-"
End Sub
Private Sub TextBox33_Change()
If TextBox33 = "" Then TextBox33 = "-"
End Sub
Private Sub TextBox34_Change()
If TextBox34 = "" Then TextBox34 = "-"
End Sub
Private Sub TextBox35_Change()
If TextBox35 = "" Then TextBox35 = "-"
End Sub
Private Sub TextBox36_Change()
If TextBox36 = "" Then TextBox36 = "-"
End Sub
Private Sub TextBox37_Change()
If TextBox37 = "" Then TextBox37 = "-"
End Sub
Private Sub TextBox38_Change()
If TextBox38 = "" Then TextBox38 = "-"
End Sub
Private Sub TextBox39_Change()
If TextBox39 = "" Then TextBox39 = "-"
End Sub
Private Sub TextBox43_Change()
TextBox46 = (Val(TextBox43)) + (Val(TextBox44)) + (Val(TextBox45))
If TextBox43 = "" Then TextBox43 = "0"
End Sub
Private Sub TextBox46_Change()
TextBox46 = (Val(TextBox43)) + (Val(TextBox44)) + (Val(TextBox45))
If TextBox46 = "" Then TextBox46 = "0"
End Sub
Private Sub TextBox50_Change()
TextBox50 = (Val(TextBox47)) + (Val(TextBox48)) + (Val(TextBox49))
If TextBox50 = "" Then TextBox50 = "0"
End Sub
Private Sub TextBox52_Change()
TextBox54 = (Val(TextBox52)) + (Val(TextBox53))
If TextBox52 = "" Then TextBox52 = "0"
End Sub
Private Sub TextBox53_Change()
TextBox54 = (Val(TextBox52)) + (Val(TextBox53))
If TextBox53 = "" Then TextBox53 = "0"
End Sub
Private Sub TextBox54_Change()
TextBox54 = (Val(TextBox52)) + (Val(TextBox53))
If TextBox54 = "" Then TextBox54 = "0"
Label52 = "PE-RT14"
End Sub
Private Sub TextBox60_Change()
TextBox60 = (Val(TextBox57)) + (Val(TextBox58)) + (Val(TextBox59)) + (Val(TextBox61))
If TextBox60 = "" Then TextBox60 = "0"
If TextBox60 <> "0" Then
Logoplaats.Enabled = True
cmdschets.Enabled = False
Label64 = 800
End If
End Sub
Private Sub TextBox61_Change()
TextBox60 = (Val(TextBox57)) + (Val(TextBox58)) + (Val(TextBox59)) + (Val(TextBox61))
If TextBox61 = "" Then TextBox61 = "0"
End Sub
Private Sub TextBox57_Change()
TextBox60 = (Val(TextBox57)) + (Val(TextBox58)) + (Val(TextBox59)) + (Val(TextBox61))
If TextBox57 = "" Then TextBox57 = "0"
End Sub
Private Sub TextBox58_Change()
TextBox60 = (Val(TextBox57)) + (Val(TextBox58)) + (Val(TextBox59)) + (Val(TextBox61))
If TextBox58 = "" Then TextBox58 = "0"
End Sub
Private Sub TextBox59_Change()
TextBox60 = (Val(TextBox57)) + (Val(TextBox58)) + (Val(TextBox59)) + (Val(TextBox61))
If TextBox60 = "" Then TextBox60 = "0"
End Sub
Private Sub TextBox65_Change()
TextBox69 = (Val(TextBox65)) + (Val(TextBox66)) + (Val(TextBox67)) + (Val(TextBox68))
If TextBox65 = "" Then TextBox65 = "0"
End Sub
Private Sub TextBox66_Change()
TextBox69 = (Val(TextBox65)) + (Val(TextBox66)) + (Val(TextBox67)) + (Val(TextBox68))
If TextBox66 = "" Then TextBox66 = "0"
End Sub
Private Sub TextBox67_Change()
TextBox69 = (Val(TextBox65)) + (Val(TextBox66)) + (Val(TextBox67)) + (Val(TextBox68))
If TextBox67 = "" Then TextBox67 = "0"
End Sub
Private Sub TextBox68_Change()
TextBox69 = (Val(TextBox65)) + (Val(TextBox66)) + (Val(TextBox67)) + (Val(TextBox68))
If TextBox68 = "" Then TextBox68 = "0"
End Sub
Private Sub Logoplaats_Click()
On Error Resume Next
 'Dim PBEGIN As Variant
 Dim PBEGIN(0 To 2) As Double
 Dim pb2(0 To 2) As Double
 Dim pb3(0 To 2) As Double
 Dim pb4(0 To 2) As Double
 Dim pb5(0 To 2) As Double
 Call Schaal(scaal)
 
frmUnitlogo.Hide
    Set newLayer = ThisDrawing.Layers.Add("UNITLOGO")
    ThisDrawing.ActiveLayer = newLayer
    Update
        
    ThisDrawing.SendCommand "-layer" & vbCr & "U" & vbCr & "*" & vbCr & vbCr
    ThisDrawing.SendCommand "osmode" & vbCr & "1" & vbCr
    Update

    For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "ENG-Kaderlogo" Then
          insp = element.InsertionPoint
          PBEGIN(0) = insp(0) - (940 * scaal)
          PBEGIN(1) = insp(1) + (298.7 * scaal)
          PBEGIN(2) = insp(2)
          End If
      Update
      End If
     Next element


'If Not INSP Then pbegin = ThisDrawing.Utility.GetPoint(, "Selecteer beginpunt ")


If TextBox55 <> "0" Then
extraafstand = 30
Dim bestand95 As String
bestand95 = "C:\ACAD2002\DWG\unitlogo_FLEX.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand95, scaal, scaal, 1, 0)
Else
extraafstand = 0
End If

If Label28.Caption > 10 Or Label34.Caption = 1 Then  ' meer dan 10 units
 pb2(0) = PBEGIN(0)
 pb2(1) = PBEGIN(1) + (extraafstand * scaal)
 pb2(2) = 0
 Dim bestand96 As String
 bestand96 = "C:\ACAD2002\DWG\unitlogo_extragroepen.dwg"
 Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb2, bestand96, scaal, scaal, 1, 0)
 PBEGIN(0) = pb2(0)
 PBEGIN(1) = pb2(1) + (130.7 * scaal)
 PBEGIN(2) = pb2(2)
 Else
 PBEGIN(0) = PBEGIN(0)
 PBEGIN(1) = PBEGIN(1) + (extraafstand * scaal)
 PBEGIN(2) = PBEGIN(2)
End If


 
'COMBINATIELOGO'S
Dim bestand100 As String
If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "PE-RT16" And Label52 = "" And Label53 = "" And Label64 = "" Then
  bestand100 = "C:\ACAD2002\DWG\unitlogo_ZDPE16.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZDPE14.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "PE-RT16" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZDPE14-16.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "" And Label52 = "" And Label53 = "ALUFLEX" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZDALU.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "" And Label51 = "" And Label52 = "PE-RT14" And Label53 = "ALUFLEX" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_PE14ALU.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If

'WTHZD 16*2,7
If Label71 = "WTH-ZD_1627" And Label50 = "WTH-ZD" And Label51 = "" And Label52 = "" And Label53 = "" And Label64 = "" Then
  bestand100 = "C:\ACAD2002\DWG\unitlogo_WTHZD2034_1627.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "WTH-ZD_1627" And Label50 = "" And Label51 = "" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "" Then
  bestand100 = "C:\ACAD2002\DWG\unitlogo_ZD1627PE14.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "WTH-ZD_1627" And Label50 = "" And Label51 = "PE-RT16" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZD1627PE14-16.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "WTH-ZD_1627" And Label50 = "" And Label51 = "" And Label52 = "" And Label53 = "ALUFLEX" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZD1627ALU.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If

If Label71 = "WTH-ZD_1627" And Label50 = "" And Label51 = "PE-RT16" And Label52 = "" And Label53 = "" And Label64 = "800" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZD1627PE16_800.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "WTH-ZD_1627" And Label50 = "" And Label51 = "PE-RT16" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "800" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZD1627PE14-16_800.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If


    '------ 16 * 2 mm in c.b.n. met ander leiding
If Label71 = "" And Label50 = "" And Label51 = "PE-RT16" And Label52 = "" And Label53 = "ALUFLEX" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_PE16ALU.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "" And Label51 = "PE-RT16" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_PE14PE16.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
        '800 meter erbij
If Label71 = "" And Label50 = "" And Label51 = "PE-RT16" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "800" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_PE14PE16_800.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
'----wediplaat
If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "" And Label52 = "" And Label53 = "" And Label64 = "800" And CheckBox3.Value = True Then  '' wthzd i.c.m. wediplaat 29-5-2007
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZDwedi.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "" And Label51 = "" And Label52 = "" And Label53 = "" And Label64 = "800" And CheckBox3.Value = True Then  '' wthzd i.c.m. wediplaat 29-5-2007
bestand100 = "C:\ACAD2002\DWG\unitlogo_wedi.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
'----wediplaat

If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "PE-RT16" And Label52 = "" And Label53 = "" And Label64 = "800" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZDPE16_800.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "PE-RT16" And Label52 = "PE-RT14" And Label53 = "" And Label64 = "800" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ZDPE14-16_800.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "" And Label51 = "PE-RT16" And Label52 = "" And Label53 = "ALUFLEX" And Label64 = "800" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_PE16ALU_800.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If

'ENKELE LOGO'S ----------------------------------------------------------------------------------------------------------------

If Label71 = "" And Label50 = "" And Label51 = "" And Label52 = "" And Label53 = "" And Label64 = "" And TextBox55 <> "0" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_FLEXFIX.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label50 = "WTH-ZD" And Label51 = "" And Label52 = "" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_WTHZD.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "WTH-ZD_1627" And Label50 = "" And Label51 = "" And Label52 = "" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_WTHZD1627.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label51 = "PE-RT16" And Label50 = "" And Label52 = "" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_PERT16.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label51 = "" And Label50 = "" And Label52 = "" And Label53 = "" And Label64 = "" And CheckBox5.Value = True Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_WTHZD1627werkleng.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If


'------------------------800 meter----hier
If Label71 = "" And Label51 = "PE-RT16" And Label50 = "" And Label52 = "" And Label53 = "" And Label64 = "800" And CheckBox4 = False Then
  bestand100 = "C:\ACAD2002\DWG\unitlogo_PERT16_800.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label51 = "PE-RT16" And Label50 = "" And Label52 = "" And Label53 = "" And Label64 = "800" And CheckBox4 = True Then
  bestand100 = "C:\ACAD2002\DWG\unitlogo_PERT16_800wk.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
'------------------------800 meter




If Label71 = "" And Label52 = "PE-RT14" And Label50 = "" And Label51 = "" And Label53 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_PERT14.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If
If Label71 = "" And Label53 = "ALUFLEX" And Label50 = "" And Label51 = "" And Label52 = "" And Label64 = "" Then
bestand100 = "C:\ACAD2002\DWG\unitlogo_ALU.dwg"
  Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(PBEGIN, bestand100, scaal, scaal, 1, 0)
End If

'If TextBox55 <> "0" And (TextBox40 = "0" And TextBox46 = "0" And TextBox50 = "0" And TextBox54 = "0") Then
'xp = 0
'Else
xp = 254.5 * scaal
'End If

 pb3(0) = PBEGIN(0)
 pb3(1) = PBEGIN(1) + xp '(254.5 * scaal)
 pb3(2) = 0

Dim bestand101 As String
bestand101 = "C:\ACAD2002\DWG\bl-krimp.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb3, bestand101, scaal, scaal, 1, 0)

 pb4(0) = pb3(0)
 pb4(1) = pb3(1) + (77 * scaal)
 pb4(2) = 0
 
 
Dim bestand102 As String
bestand102 = "C:\ACAD2002\DWG\bl-recht.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb4, bestand102, scaal, scaal, 1, 0)

 pb5(0) = pb4(0)
 pb5(1) = pb4(1) + (77 * scaal)
 pb5(2) = 0


Dim bestand103 As String
bestand103 = "C:\ACAD2002\DWG\bl-hr.dwg"
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(pb5, bestand103, scaal, scaal, 1, 0)


If Err Then
    frmUnitlogo.Show
    Exit Sub
    End If

Call vulunitlogoin
Update
 ThisDrawing.SendCommand "osmode" & vbCr & "0" & vbCr
Unload Me

End Sub
Sub vulunitlogoin()
'bloklogo invullen WTH-ZD
        If TextBox17 = "0" Then TextBox17 = "-"
        If TextBox18 = "0" Then TextBox18 = "-"
        If TextBox19 = "0" Then TextBox19 = "-"
        If TextBox20 = "0" Then TextBox20 = "-"
        If TextBox21 = "0" Then TextBox21 = "-"
        If TextBox22 = "0" Then TextBox22 = "-"
        If TextBox23 = "0" Then TextBox23 = "-"
        If TextBox41 = "0" Then TextBox41 = "-"
        If TextBox42 = "0" Then TextBox42 = "-"
        
        If TextBox43 = "0" Then TextBox43 = "-"
        If TextBox44 = "0" Then TextBox44 = "-"
        If TextBox45 = "0" Then TextBox45 = "-"
        
        If TextBox47 = "0" Then TextBox47 = "-"
        If TextBox48 = "0" Then TextBox48 = "-"
        If TextBox49 = "0" Then TextBox49 = "-"
        
        If TextBox52 = "0" Then TextBox52 = "-"
        If TextBox53 = "0" Then TextBox53 = "-"
        
        If TextBox57 = "0" Then TextBox57 = "-"
        If TextBox58 = "0" Then TextBox58 = "-"
        If TextBox59 = "0" Then TextBox59 = "-"
        If TextBox61 = "0" Then TextBox61 = "-"

        If TextBox65 = "0" Then TextBox65 = "-"
        If TextBox66 = "0" Then TextBox66 = "-"
        If TextBox67 = "0" Then TextBox67 = "-"
        If TextBox68 = "0" Then TextBox68 = "-"
        
        

For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "unitlogo_WTHZD" Or element.Name = "unitlogo_ZDPE16" Or element.Name = "unitlogo_ZDPE16_800" _
        Or element.Name = "unitlogo_ZDPE14" Or element.Name = "unitlogo_ZDPE14-16" Or element.Name = "unitlogo_ZDPE14-16_800" _
        Or element.Name = "unitlogo_ZDALU" Or element.Name = "unitlogo_FLEXFIX" Or element.Name = "unitlogo_ZDwedi" _
        Or element.Name = "unitlogo_wedi" _
        Or element.Name = "unitlogo_WTHZD1627" Or element.Name = "unitlogo_WTHZD2034_1627" Or element.Name = "unitlogo_ZD1627PE16" _
        Or element.Name = "unitlogo_ZD1627PE16_800" Or element.Name = "unitlogo_ZD1627ALU" Or element.Name = "unitlogo_ZD1627PE14-16" _
        Or element.Name = "unitlogo_ZD1627PE14-16_800" Or element.Name = "unitlogo_ZD1627PE14" Or element.Name = "unitlogo_WTHZD1627werkleng" Then
   
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "U_REGELUNITTYPE1" And attribuut.textstring = "" Then attribuut.textstring = TextBox1 'REGELUNIT1
        If attribuut.TagString = "U_REGELUNITTYPE2" And attribuut.textstring = "" Then attribuut.textstring = TextBox2 'REGELUNIT2
        If attribuut.TagString = "U_REGELUNITTYPE3" And attribuut.textstring = "" Then attribuut.textstring = TextBox3 'REGELUNIT3
        If attribuut.TagString = "U_REGELUNITTYPE4" And attribuut.textstring = "" Then attribuut.textstring = TextBox4 'REGELUNIT4
        If attribuut.TagString = "U_REGELUNITTYPE5" And attribuut.textstring = "" Then attribuut.textstring = TextBox5 'REGELUNIT5
        If attribuut.TagString = "U_REGELUNITTYPE6" And attribuut.textstring = "" Then attribuut.textstring = TextBox6 'REGELUNIT6
        If attribuut.TagString = "U_REGELUNITTYPE7" And attribuut.textstring = "" Then attribuut.textstring = TextBox7 'REGELUNIT7
        If attribuut.TagString = "U_REGELUNITTYPE8" And attribuut.textstring = "" Then attribuut.textstring = TextBox8 'REGELUNIT8
        If attribuut.TagString = "U_REGELUNITTYPE9" And attribuut.textstring = "" Then attribuut.textstring = TextBox24 'REGELUNIT9
        If attribuut.TagString = "U_REGELUNITTYPE10" And attribuut.textstring = "" Then attribuut.textstring = TextBox25 'REGELUNIT10
       
        If attribuut.TagString = "U_BEVESTIGINGSTYPE1" And attribuut.textstring = "" Then attribuut.textstring = TextBox9 'BEVESTIGING1
        If attribuut.TagString = "U_BEVESTIGINGSTYPE2" And attribuut.textstring = "" Then attribuut.textstring = TextBox10 'BEVESTIGING2
        If attribuut.TagString = "U_BEVESTIGINGSTYPE3" And attribuut.textstring = "" Then attribuut.textstring = TextBox11 'BEVESTIGING3
        If attribuut.TagString = "U_BEVESTIGINGSTYPE4" And attribuut.textstring = "" Then attribuut.textstring = TextBox12 'BEVESTIGING4
        If attribuut.TagString = "U_BEVESTIGINGSTYPE5" And attribuut.textstring = "" Then attribuut.textstring = TextBox13 'BEVESTIGING5
        If attribuut.TagString = "U_BEVESTIGINGSTYPE6" And attribuut.textstring = "" Then attribuut.textstring = TextBox14 'BEVESTIGING6
        If attribuut.TagString = "U_BEVESTIGINGSTYPE7" And attribuut.textstring = "" Then attribuut.textstring = TextBox15 'BEVESTIGING7
        If attribuut.TagString = "U_BEVESTIGINGSTYPE8" And attribuut.textstring = "" Then attribuut.textstring = TextBox16 'BEVESTIGING8
        If attribuut.TagString = "U_BEVESTIGINGSTYPE9" And attribuut.textstring = "" Then attribuut.textstring = TextBox32 'BEVESTIGING9
        If attribuut.TagString = "U_BEVESTIGINGSTYPE10" And attribuut.textstring = "" Then attribuut.textstring = TextBox33 'BEVESTIGING10
        

        If attribuut.TagString = "U_WTH250" And attribuut.textstring = "" Then attribuut.textstring = TextBox17 '250 METER
        If attribuut.TagString = "U_WTH165" And attribuut.textstring = "" Then attribuut.textstring = TextBox18 '165 METER
        If attribuut.TagString = "U_WTH125" And attribuut.textstring = "" Then attribuut.textstring = TextBox19 '125 METER
        If attribuut.TagString = "U_WTH105" And attribuut.textstring = "" Then attribuut.textstring = TextBox20 '105 METER
        If attribuut.TagString = "U_WTH90" And attribuut.textstring = "" Then attribuut.textstring = TextBox21 '90 METER
        If attribuut.TagString = "U_WTH75" And attribuut.textstring = "" Then attribuut.textstring = TextBox22 '75 METER
        If attribuut.TagString = "U_WTH63" And attribuut.textstring = "" Then attribuut.textstring = TextBox23 '63 METER
        If attribuut.TagString = "U_WTH50" And attribuut.textstring = "" Then attribuut.textstring = TextBox41 '50 METER
        If attribuut.TagString = "U_WTH40" And attribuut.textstring = "" Then attribuut.textstring = TextBox42 '40 METER
        If attribuut.TagString = "U_WTH120" And attribuut.textstring = "" Then attribuut.textstring = TextBox17 '120 METER
        If attribuut.TagString = "U_WTH90" And attribuut.textstring = "" Then attribuut.textstring = TextBox18 '90 METER
        If attribuut.TagString = "U_WTH60" And attribuut.textstring = "" Then attribuut.textstring = TextBox19 '60 METER
        If attribuut.TagString = "U_TOTAAL" And attribuut.textstring = "" Then attribuut.textstring = TextBox40 'TOTAAL AANTAL GROEPEN
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox70 'TOTAAL AANTAL GROEPEN
        
        'WTHZD "16*2,7
        If element.Name = "unitlogo_WTHZD1627" Or element.Name = "unitlogo_ZD1627PE14" _
           Or element.Name = "unitlogo_ZD1627PE14-16" Or element.Name = "unitlogo_ZD1627ALU" _
           Or element.Name = "unitlogo_ZD1627PE16_800" Or element.Name = "unitlogo_WTHZD1627werkleng" Then
            Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For P = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(P)
            
             If attribuut.TagString = "U_WTH105" And attribuut.textstring = "" Then attribuut.textstring = TextBox65 '105 METER
             If attribuut.TagString = "U_WTH90" And attribuut.textstring = "" Then attribuut.textstring = TextBox66 '90 METER
             If attribuut.TagString = "U_WTH75" And attribuut.textstring = "" Then attribuut.textstring = TextBox67 '75 METER
             If attribuut.TagString = "U_WTH63" And attribuut.textstring = "" Then attribuut.textstring = TextBox68 '63 METER
             If attribuut.TagString = "U_TOTAAL" And attribuut.textstring = "" Then attribuut.textstring = TextBox69 'TOTAAL AANTAL GROEPEN
             If attribuut.TagString = "U_TOTAALWERKLENG" And attribuut.textstring = "" Then attribuut.textstring = TextBox71 'TOTAAL AANTAL GROEPEN 500 meter
             'PE14
              If attribuut.TagString = "U_WTH1490" And attribuut.textstring = "" Then attribuut.textstring = TextBox52 '90 METER
              If attribuut.TagString = "U_WTH1460" And attribuut.textstring = "" Then attribuut.textstring = TextBox53 '60 METER
              If attribuut.TagString = "U_TOTAALPE14" And attribuut.textstring = "" Then attribuut.textstring = TextBox54 'TOTAAL PERT
             'PE16
             If attribuut.TagString = "U_WTH16120" And attribuut.textstring = "" Then attribuut.textstring = TextBox43 '120 METER
             If attribuut.TagString = "U_WTH1690" And attribuut.textstring = "" Then attribuut.textstring = TextBox44 '90 METER
             If attribuut.TagString = "U_WTH1660" And attribuut.textstring = "" Then attribuut.textstring = TextBox45 '60 METER
             If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox46 'TOTAAL PERT
             'ALUFLEX
              If attribuut.TagString = "ALU200" And attribuut.textstring = "" Then attribuut.textstring = TextBox47 '200 METER
              If attribuut.TagString = "ALU100" And attribuut.textstring = "" Then attribuut.textstring = TextBox48 '100 METER
              If attribuut.TagString = "ALU50" And attribuut.textstring = "" Then attribuut.textstring = TextBox49 '50 METER
              If attribuut.TagString = "ALUFLEX" Then attribuut.textstring = Label29  '60 METER
              If attribuut.TagString = "ALU_TOTAAL" And attribuut.textstring = "" Then attribuut.textstring = TextBox50 'TOTAAL AANTAL GROEPEN
             '800 METER
              If attribuut.TagString = "U_WTH16800" And attribuut.textstring = "" Then attribuut.textstring = TextBox61 '800 METER
              If attribuut.TagString = "U_WTH16120_800" And attribuut.textstring = "" Then attribuut.textstring = TextBox57 '120 METER
              If attribuut.TagString = "U_WTH1690_800" And attribuut.textstring = "" Then attribuut.textstring = TextBox58 '90 METER
              If attribuut.TagString = "U_WTH1660_800" And attribuut.textstring = "" Then attribuut.textstring = TextBox59 '60 METER
              If attribuut.TagString = "U_TOTAALPE16800" And attribuut.textstring = "" Then attribuut.textstring = TextBox60 'TOTAAL PERT
             Next P
             End If
        End If
         'WTHZD "16*2,7 + WTH-ZD 20*3,4
        If element.Name = "unitlogo_WTHZD2034_1627" Then
            Set symbool = element
            If symbool.HasAttributes Then
            attributen = symbool.GetAttributes
            For PE = LBound(attributen) To UBound(attributen)
            Set attribuut = attributen(PE)
            
             If attribuut.TagString = "U_WTH105_1627" And attribuut.textstring = "" Then attribuut.textstring = TextBox65 '105 METER
             If attribuut.TagString = "U_WTH90_1627" And attribuut.textstring = "" Then attribuut.textstring = TextBox66 '90 METER
             If attribuut.TagString = "U_WTH75_1627" And attribuut.textstring = "" Then attribuut.textstring = TextBox67 '75 METER
             If attribuut.TagString = "U_WTH63_1627" And attribuut.textstring = "" Then attribuut.textstring = TextBox68 '63 METER
             If attribuut.TagString = "U_TOTAAL_1627" And attribuut.textstring = "" Then attribuut.textstring = TextBox69 'TOTAAL AANTAL GROEPEN
             Next PE
             End If
        End If
        
        If CheckBox1.Value = False Then
        '--PERT 16*2
        If attribuut.TagString = "U_WTH16120" And attribuut.textstring = "" Then attribuut.textstring = TextBox43 '120 METER
        If attribuut.TagString = "U_WTH1690" And attribuut.textstring = "" Then attribuut.textstring = TextBox44 '90 METER
        If attribuut.TagString = "U_WTH1660" And attribuut.textstring = "" Then attribuut.textstring = TextBox45 '60 METER
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox46 'TOTAAL PERT
        Else '---+ 800 meter
        If attribuut.TagString = "U_WTH16800" And attribuut.textstring = "" Then attribuut.textstring = TextBox61 '800 METER
        If attribuut.TagString = "U_WTH16120" And attribuut.textstring = "" Then attribuut.textstring = TextBox57 '120 METER
        If attribuut.TagString = "U_WTH1690" And attribuut.textstring = "" Then attribuut.textstring = TextBox58 '90 METER
        If attribuut.TagString = "U_WTH1660" And attribuut.textstring = "" Then attribuut.textstring = TextBox59 '60 METER
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox60 'TOTAAL PERT
        End If
        
        
        '--PERT 14*2
        If attribuut.TagString = "U_WTH1490" And attribuut.textstring = "" Then attribuut.textstring = TextBox52 '90 METER
        If attribuut.TagString = "U_WTH1460" And attribuut.textstring = "" Then attribuut.textstring = TextBox53 '60 METER
        If attribuut.TagString = "U_TOTAALPE14" And attribuut.textstring = "" Then attribuut.textstring = TextBox54 'TOTAAL PERT
         
         'End If
         
         
         
         
       Next j
       
       End If
      End If
     End If
  Next element
 'End If
 
 For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "unitlogo_PERT14" Or element.Name = "unitlogo_PERT16" Or element.Name = "unitlogo_PERT16_800" _
      Or element.Name = "unitlogo_PE14PE16" Or element.Name = "unitlogo_PERT16_800wk" Or element.Name = "unitlogo_PE14PE16_800" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For m = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(m)
        If attribuut.TagString = "U_REGELUNITTYPE1" And attribuut.textstring = "" Then attribuut.textstring = TextBox1 'REGELUNIT1
        If attribuut.TagString = "U_REGELUNITTYPE2" And attribuut.textstring = "" Then attribuut.textstring = TextBox2 'REGELUNIT2
        If attribuut.TagString = "U_REGELUNITTYPE3" And attribuut.textstring = "" Then attribuut.textstring = TextBox3 'REGELUNIT3
        If attribuut.TagString = "U_REGELUNITTYPE4" And attribuut.textstring = "" Then attribuut.textstring = TextBox4 'REGELUNIT4
        If attribuut.TagString = "U_REGELUNITTYPE5" And attribuut.textstring = "" Then attribuut.textstring = TextBox5 'REGELUNIT5
        If attribuut.TagString = "U_REGELUNITTYPE6" And attribuut.textstring = "" Then attribuut.textstring = TextBox6 'REGELUNIT6
        If attribuut.TagString = "U_REGELUNITTYPE7" And attribuut.textstring = "" Then attribuut.textstring = TextBox7 'REGELUNIT7
        If attribuut.TagString = "U_REGELUNITTYPE8" And attribuut.textstring = "" Then attribuut.textstring = TextBox8 'REGELUNIT8
        If attribuut.TagString = "U_REGELUNITTYPE9" And attribuut.textstring = "" Then attribuut.textstring = TextBox24 'REGELUNIT9
        If attribuut.TagString = "U_REGELUNITTYPE10" And attribuut.textstring = "" Then attribuut.textstring = TextBox25 'REGELUNIT10

        If attribuut.TagString = "U_BEVESTIGINGSTYPE1" And attribuut.textstring = "" Then attribuut.textstring = TextBox9 'BEVESTIGING1
        If attribuut.TagString = "U_BEVESTIGINGSTYPE2" And attribuut.textstring = "" Then attribuut.textstring = TextBox10 'BEVESTIGING2
        If attribuut.TagString = "U_BEVESTIGINGSTYPE3" And attribuut.textstring = "" Then attribuut.textstring = TextBox11 'BEVESTIGING3
        If attribuut.TagString = "U_BEVESTIGINGSTYPE4" And attribuut.textstring = "" Then attribuut.textstring = TextBox12 'BEVESTIGING4
        If attribuut.TagString = "U_BEVESTIGINGSTYPE5" And attribuut.textstring = "" Then attribuut.textstring = TextBox13 'BEVESTIGING5
        If attribuut.TagString = "U_BEVESTIGINGSTYPE6" And attribuut.textstring = "" Then attribuut.textstring = TextBox14 'BEVESTIGING6
        If attribuut.TagString = "U_BEVESTIGINGSTYPE7" And attribuut.textstring = "" Then attribuut.textstring = TextBox15 'BEVESTIGING7
        If attribuut.TagString = "U_BEVESTIGINGSTYPE8" And attribuut.textstring = "" Then attribuut.textstring = TextBox16 'BEVESTIGING8
        If attribuut.TagString = "U_BEVESTIGINGSTYPE9" And attribuut.textstring = "" Then attribuut.textstring = TextBox32 'BEVESTIGING9
        If attribuut.TagString = "U_BEVESTIGINGSTYPE10" And attribuut.textstring = "" Then attribuut.textstring = TextBox33 'BEVESTIGING10
        
        If CheckBox1.Value = False Then
        '--PERT 16*2
        If attribuut.TagString = "U_WTH16120" And attribuut.textstring = "" Then attribuut.textstring = TextBox43 '120 METER
        If attribuut.TagString = "U_WTH1690" And attribuut.textstring = "" Then attribuut.textstring = TextBox44 '90 METER
        If attribuut.TagString = "U_WTH1660" And attribuut.textstring = "" Then attribuut.textstring = TextBox45 '60 METER
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox46 'TOTAAL PERT
        Else '---+ 800 meter
        If attribuut.TagString = "U_WTH16800" And attribuut.textstring = "" Then attribuut.textstring = TextBox61 '800 METER
        If attribuut.TagString = "U_WTH16120" And attribuut.textstring = "" Then attribuut.textstring = TextBox57 '120 METER
        If attribuut.TagString = "U_WTH1690" And attribuut.textstring = "" Then attribuut.textstring = TextBox58 '90 METER
        If attribuut.TagString = "U_WTH1660" And attribuut.textstring = "" Then attribuut.textstring = TextBox59 '60 METER
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" And element.Name = "unitlogo_PERT16_800" Then
            attribuut.textstring = TextBox60 'TOTAAL PERT
        End If
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" And element.Name = "unitlogo_PERT16_800wk" Then
            attribuut.textstring = TextBox70 'TOTAAL PERT
        End If
        
        End If
        
        
''        '--PE 16
''        If attribuut.TagString = "U_WTH16120" And attribuut.TextString = "" Then attribuut.TextString = TextBox43 '120 METER
''        If attribuut.TagString = "U_WTH1690" And attribuut.TextString = "" Then attribuut.TextString = TextBox44 '90 METER
''        If attribuut.TagString = "U_WTH1660" And attribuut.TextString = "" Then attribuut.TextString = TextBox45 '60 METER
''       ' If attribuut.TagString = "PERT" And attribuut.TextString = "" And Label51 = "PE-RT16" Then attribuut.TextString = "PE-RT 16*2 mm"
''        If attribuut.TagString = "U_TOTAALPE16" And attribuut.TextString = "" Then attribuut.TextString = TextBox46 'TOTAAL AANTAL GROEPEN
        
        '--PE 14
        If attribuut.TagString = "U_WTH1490" And attribuut.textstring = "" Then attribuut.textstring = TextBox52 '90 METER
        If attribuut.TagString = "U_WTH1460" And attribuut.textstring = "" Then attribuut.textstring = TextBox53 '60 METER
       ' If attribuut.TagString = "PERT" And attribuut.TextString = "" And Label52 = "PE-RT14" Then attribuut.TextString = "PE-RT 14*2 mm"
        If attribuut.TagString = "U_TOTAALPE14" And attribuut.textstring = "" Then attribuut.textstring = TextBox54 'TOTAAL AANTAL GROEPEN
        
        'End If
                
       
       Next m
       
       End If
      End If
     End If
  Next element
 
 For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "unitlogo_ZDALU" Or element.Name = "unitlogo_PE14ALU" Or element.Name = "unitlogo_PE16ALU" _
      Or element.Name = "unitlogo_PE16ALU_800" Or element.Name = "unitlogo_ALU" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For m = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(m)
        If attribuut.TagString = "U_REGELUNITTYPE1" And attribuut.textstring = "" Then attribuut.textstring = TextBox1 'REGELUNIT1
        If attribuut.TagString = "U_REGELUNITTYPE2" And attribuut.textstring = "" Then attribuut.textstring = TextBox2 'REGELUNIT2
        If attribuut.TagString = "U_REGELUNITTYPE3" And attribuut.textstring = "" Then attribuut.textstring = TextBox3 'REGELUNIT3
        If attribuut.TagString = "U_REGELUNITTYPE4" And attribuut.textstring = "" Then attribuut.textstring = TextBox4 'REGELUNIT4
        If attribuut.TagString = "U_REGELUNITTYPE5" And attribuut.textstring = "" Then attribuut.textstring = TextBox5 'REGELUNIT5
        If attribuut.TagString = "U_REGELUNITTYPE6" And attribuut.textstring = "" Then attribuut.textstring = TextBox6 'REGELUNIT6
        If attribuut.TagString = "U_REGELUNITTYPE7" And attribuut.textstring = "" Then attribuut.textstring = TextBox7 'REGELUNIT7
        If attribuut.TagString = "U_REGELUNITTYPE8" And attribuut.textstring = "" Then attribuut.textstring = TextBox8 'REGELUNIT8
        If attribuut.TagString = "U_REGELUNITTYPE9" And attribuut.textstring = "" Then attribuut.textstring = TextBox24 'REGELUNIT9
        If attribuut.TagString = "U_REGELUNITTYPE10" And attribuut.textstring = "" Then attribuut.textstring = TextBox25 'REGELUNIT10
 
        If attribuut.TagString = "U_BEVESTIGINGSTYPE1" And attribuut.textstring = "" Then attribuut.textstring = TextBox9 'BEVESTIGING1
        If attribuut.TagString = "U_BEVESTIGINGSTYPE2" And attribuut.textstring = "" Then attribuut.textstring = TextBox10 'BEVESTIGING2
        If attribuut.TagString = "U_BEVESTIGINGSTYPE3" And attribuut.textstring = "" Then attribuut.textstring = TextBox11 'BEVESTIGING3
        If attribuut.TagString = "U_BEVESTIGINGSTYPE4" And attribuut.textstring = "" Then attribuut.textstring = TextBox12 'BEVESTIGING4
        If attribuut.TagString = "U_BEVESTIGINGSTYPE5" And attribuut.textstring = "" Then attribuut.textstring = TextBox13 'BEVESTIGING5
        If attribuut.TagString = "U_BEVESTIGINGSTYPE6" And attribuut.textstring = "" Then attribuut.textstring = TextBox14 'BEVESTIGING6
        If attribuut.TagString = "U_BEVESTIGINGSTYPE7" And attribuut.textstring = "" Then attribuut.textstring = TextBox15 'BEVESTIGING7
        If attribuut.TagString = "U_BEVESTIGINGSTYPE8" And attribuut.textstring = "" Then attribuut.textstring = TextBox16 'BEVESTIGING8
        If attribuut.TagString = "U_BEVESTIGINGSTYPE9" And attribuut.textstring = "" Then attribuut.textstring = TextBox32 'BEVESTIGING9
        If attribuut.TagString = "U_BEVESTIGINGSTYPE10" And attribuut.textstring = "" Then attribuut.textstring = TextBox33 'BEVESTIGING10
        If attribuut.TagString = "ALU200" And attribuut.textstring = "" Then attribuut.textstring = TextBox47 '200 METER
        If attribuut.TagString = "ALU100" And attribuut.textstring = "" Then attribuut.textstring = TextBox48 '100 METER
        If attribuut.TagString = "ALU50" And attribuut.textstring = "" Then attribuut.textstring = TextBox49 '50 METER
        If attribuut.TagString = "ALUFLEX" Then attribuut.textstring = Label29  '60 METER
        If attribuut.TagString = "ALU_TOTAAL" And attribuut.textstring = "" Then attribuut.textstring = TextBox50 'TOTAAL AANTAL GROEPEN
        'End If
''           '--PE 16
''        If attribuut.TagString = "U_WTH16120" And attribuut.TextString = "" Then attribuut.TextString = TextBox43 '120 METER
''        If attribuut.TagString = "U_WTH1690" And attribuut.TextString = "" Then attribuut.TextString = TextBox44 '90 METER
''        If attribuut.TagString = "U_WTH1660" And attribuut.TextString = "" Then attribuut.TextString = TextBox45 '60 METER
''       ' If attribuut.TagString = "PERT" And attribuut.TextString = "" And Label51 = "PE-RT16" Then attribuut.TextString = "PE-RT 16*2 mm"
''        If attribuut.TagString = "U_TOTAALPE16" And attribuut.TextString = "" Then attribuut.TextString = TextBox46 'TOTAAL AANTAL GROEPEN
        
        If CheckBox1.Value = False Then
        '--PERT 16*2
        If attribuut.TagString = "U_WTH16120" And attribuut.textstring = "" Then attribuut.textstring = TextBox43 '120 METER
        If attribuut.TagString = "U_WTH1690" And attribuut.textstring = "" Then attribuut.textstring = TextBox44 '90 METER
        If attribuut.TagString = "U_WTH1660" And attribuut.textstring = "" Then attribuut.textstring = TextBox45 '60 METER
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox46 'TOTAAL PERT
        Else '---+ 800 meter
        If attribuut.TagString = "U_WTH16800" And attribuut.textstring = "" Then attribuut.textstring = TextBox61 '800 METER
        If attribuut.TagString = "U_WTH16120" And attribuut.textstring = "" Then attribuut.textstring = TextBox57 '120 METER
        If attribuut.TagString = "U_WTH1690" And attribuut.textstring = "" Then attribuut.textstring = TextBox58 '90 METER
        If attribuut.TagString = "U_WTH1660" And attribuut.textstring = "" Then attribuut.textstring = TextBox59 '60 METER
        If attribuut.TagString = "U_TOTAALPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox60 'TOTAAL PERT
        End If
        
        '--PE 14
        If attribuut.TagString = "U_WTH1490" And attribuut.textstring = "" Then attribuut.textstring = TextBox52 '90 METER
        If attribuut.TagString = "U_WTH1460" And attribuut.textstring = "" Then attribuut.textstring = TextBox53 '60 METER
       ' If attribuut.TagString = "PERT" And attribuut.TextString = "" And Label52 = "PE-RT14" Then attribuut.TextString = "PE-RT 14*2 mm"
        If attribuut.TagString = "U_TOTAALPE14" And attribuut.textstring = "" Then attribuut.textstring = TextBox54 'TOTAAL AANTAL GROEPEN
       
       Next m
       
       End If
      End If
     End If
  Next element
 
 
 For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "unitlogo_extragroepen" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For k = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(k)
        If attribuut.TagString = "U_REGELUNITTYPE11" And attribuut.textstring = "" Then attribuut.textstring = TextBox26 'REGELUNIT11
        If attribuut.TagString = "U_REGELUNITTYPE12" And attribuut.textstring = "" Then attribuut.textstring = TextBox27 'REGELUNIT12
        If attribuut.TagString = "U_REGELUNITTYPE13" And attribuut.textstring = "" Then attribuut.textstring = TextBox28 'REGELUNIT13
        If attribuut.TagString = "U_REGELUNITTYPE14" And attribuut.textstring = "" Then attribuut.textstring = TextBox29 'REGELUNIT14
        If attribuut.TagString = "U_REGELUNITTYPE15" And attribuut.textstring = "" Then attribuut.textstring = TextBox30 'REGELUNIT15
        If attribuut.TagString = "U_REGELUNITTYPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox31 'REGELUNIT16
               
        If attribuut.TagString = "U_BEVESTIGINGSTYPE11" And attribuut.textstring = "" Then attribuut.textstring = TextBox34 'BEVESTIGING11
        If attribuut.TagString = "U_BEVESTIGINGSTYPE12" And attribuut.textstring = "" Then attribuut.textstring = TextBox35 'BEVESTIGING12
        If attribuut.TagString = "U_BEVESTIGINGSTYPE13" And attribuut.textstring = "" Then attribuut.textstring = TextBox36 'BEVESTIGING13
        If attribuut.TagString = "U_BEVESTIGINGSTYPE14" And attribuut.textstring = "" Then attribuut.textstring = TextBox37 'BEVESTIGING14
        If attribuut.TagString = "U_BEVESTIGINGSTYPE15" And attribuut.textstring = "" Then attribuut.textstring = TextBox38 'BEVESTIGING15
        If attribuut.TagString = "U_BEVESTIGINGSTYPE16" And attribuut.textstring = "" Then attribuut.textstring = TextBox39 'BEVESTIGING16
               
        Next k
       
       End If
      End If
     End If
  Next element
 'End If
 
 For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "unitlogo_FLEX" Then
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For t = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(t)
        If attribuut.TagString = "FLEXMATTEN" And attribuut.textstring = "" Then attribuut.textstring = TextBox55 'AANTAL MATTEN
        If attribuut.TagString = "FLEXMEET_TOTAAL" And attribuut.textstring = "" Then attribuut.textstring = TextBox56 'AANTAL METER
        Next t
       End If
      End If
     End If
  Next element
 
 
 
 Unload Me
End Sub
Sub Schaal(scaal)
frmUnitlogo.Hide
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
Update
End Sub
Private Sub cmdAfsluiten_Click()
Unload Me
End Sub

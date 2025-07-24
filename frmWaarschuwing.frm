VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWaarschuwing 
   Caption         =   "WAARSCHUWING...."
   ClientHeight    =   4845
   ClientLeft      =   48
   ClientTop       =   540
   ClientWidth     =   5280
   OleObjectBlob   =   "frmWaarschuwing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWaarschuwing"
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
  
frmWaarschuwing.TextBox1 = frmGroeptekst.TextBox2
frmWaarschuwing.TextBox2 = frmGroeptekst.TextBox9

If frmGroeptekst.ToggleButton2.Value = True Then Call togglemodule.groeptekst_waarschuwing_toggle

Call ZOEK_DE_FOUT
End Sub
Sub ZOEK_DE_FOUT()
unittel = frmGroeptekst.TextBox9

If unittel > 0 And unittel < 10 Then
  unitonder10 = "0" & frmGroeptekst.TextBox9
Else
  unitonder10 = frmGroeptekst.TextBox9
End If

For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "groeptekstbloknew" Or element.Name = "GROEPTEKSTBLOKNEW" Then

      
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
           If UCase(attribuut.TagString) = "UNITNUMMER" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
           If UCase(attribuut.TagString) = "ROLLENGTE" And attribuut.textstring = "165 meter" Then zdf = attribuut.textstring
           If UCase(attribuut.TagString) = "GROEPTEKST" And attribuut.textstring <> "" Then gpt = attribuut.textstring
                
                  If bb = unitonder10 And zdf = "165 meter" Then
                  
                        
                        
                        
                     For Each element2 In ThisDrawing.ModelSpace
                        If element2.Layer = gpt Then
                        'BEREKENEN TOTALE LENGTE
                          If element2.EntityName = "AcDbLine" Then Lengte = Lengte + element2.Length
                          If element2.EntityName = "AcDbArc" Then Lengte = Lengte + element2.ArcLength
                        End If
                       Next element2
                       z = 0
                       Dim cirkel As Object
                       For Each cirkel In ThisDrawing.ModelSpace
                          If cirkel.Layer = gpt Then
                          If cirkel.EntityName = "AcDbCircle" Then z = z + 1
                          End If
                          Next cirkel
                          
                      If frmGroeptekst.ToggleButton1.Value = False Then wandhoogte = 2.5
                      If frmGroeptekst.ToggleButton1.Value = True Then wandhoogte = 2
                      If frmGroeptekst.OptionButton5 = True Then wandhoogte = frmGroeptekst.TextBox13
                          
                          If z <> 0 Then zlengte = (z * (100 * wandhoogte)) + 100
              
                      Lengte = Lengte + zlengte
                    
                     'LENGTE IN METERS
                      Lengte = Lengte / 100
                      Lengte = Round(Lengte, 1)
                      Lengte = Lengte ' + 3 ' 07-04-06
                      LLL = gpt & " = " & Lengte & " meter."
                      
                      If frmGroeptekst.ToggleButton2.Value = True Then ''''''''''''engeland
                             lp = Split(LLL, " ") ''''''''''''engeland
                             LLL = "group " & lp(1) & " = " & Lengte & " meter." ''''''''''''engeland
                      End If ''''''''''''engeland
                             
                             
                             
                      If zdf = "165 meter" Then ListBox1.AddItem (LLL)
                       
                       Lengte = 0
                       zdf = ""
                       bb = ""
                       gpt = ""
             
                   End If 'BB
           Next I
       
        End If
      End If
      End If
  Next element
End Sub

Private Sub cmdnietsveranderen_Click()
Unload Me
frmGroeptekst.TextBox2.BackColor = &HFFFFFF
End Sub
Private Sub cmdwijzigen_Click()
Unload Me
frmGroeptekst.TextBox2.BackColor = &HFFFFFF
b = Val(frmGroeptekst.TextBox2)
frmGroeptekst.TextBox2 = "0"
frmGroeptekst.TextBox3 = Val(frmGroeptekst.TextBox3) + b 'Val(frmGroeptekst.TextBox2)

unittel = frmGroeptekst.TextBox9

If unittel > 0 And unittel < 10 Then
  unitonder10 = "0" & frmGroeptekst.TextBox9
Else
  unitonder10 = frmGroeptekst.TextBox9
End If


For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
      If element.Name = "groeptekstbloknew" Or element.Name = "GROEPTEKSTBLOKNEW" Then
      
      Set symbool = element
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If UCase(attribuut.TagString) = "UNITNUMMER" And attribuut.textstring <> "" Then bb = attribuut.textstring 'REGELUNITNUMMER
            If bb = unitonder10 Then
              For k = LBound(attributen) To UBound(attributen)
              Set attribuut = attributen(k)
       
                 If UCase(attribuut.TagString) = "ROLLENGTE" And attribuut.textstring = "165 meter" Then attribuut.textstring = "125 meter"
                 
             Next k
             bb = ""
             End If
       Next I
       
        End If
      End If
      End If
  Next element

Update

End Sub



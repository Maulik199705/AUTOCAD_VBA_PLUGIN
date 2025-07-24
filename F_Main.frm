VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Main 
   Caption         =   "Leidingteken-programma WTH"
   ClientHeight    =   7680
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5688
   OleObjectBlob   =   "F_Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''
'''''    ' Display the delta of the new line
'''''    lineDelta = lineObj.Delta
'''''    MsgBox "The delta of the new Line is: " & vbCrLf & _
'''''                                            "DeltaX:" & lineDelta(0) & vbCrLf & _
'''''                                            "DeltaY:" & lineDelta(1) & vbCrLf & _
'''''                                            "DeltaZ:" & lineDelta(2)
'''''











Private Sub CommandButton33_Click()

    'NIEUW 31 MAART 2003
    
    ' --------------------------------------------------------------
    ' SELECTEREN ARC
    ' --------------------------------------------------------------
    
    F_Main.Hide
    
    Dim ReturnObj1 As AcadArc
    Dim basePnt1 As Variant
    
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity ReturnObj1, basePnt1, "Selecteer de aanvoer-polyline."
    If Err <> 0 Then
        Err.Clear
        'MsgBox "Verkeerd geselecteerd"
        Exit Sub
    End If
            
    If ReturnObj1.EntityName <> "AcDbArc" Then
        MsgBox "Geen Arc geselecteerd.", vbCritical, "Let op"
        Exit Sub
    End If
     
    'MsgBox basePnt1(0) & Chr(10) & Chr(13) & basePnt1(1) & Chr(10) & Chr(13) & basePnt1(2)
    
    
    ' --------------------------------------------------------------
    ' SELECTEREN NIEUWE PLAATSINGSPUNT
    ' --------------------------------------------------------------
    
    'LET OP: SNAPANG IN RADIALEN !
    'ThisDrawing.SetVariable "ORTHO", "ON"
    ThisDrawing.SetVariable "SNAPANG", ReturnObj1.StartAngle
    
    Dim p As Variant
    p = ThisDrawing.Utility.GetPoint(basePnt1, "Selecteer het plaatsingspunt")
    
    ThisDrawing.SetVariable "SNAPANG", 0
    
    ' --------------------------------------------------------------
    ' BEREKENEN LENGTE-VERSCHIL
    ' --------------------------------------------------------------
    
    Dim dLengte As Double
    dLengte = Lengte(basePnt1, p)
    
    ' --------------------------------------------------------------
    ' BEPALEN VERLENGEN OF VERKORTEN
    ' --------------------------------------------------------------
    
    Dim bVerlengen As Boolean
    If Lengte(p, ReturnObj1.EndPoint) > Lengte(p, basePnt1) Then
        bVerlengen = True
        MsgBox "verlengen"
    Else
        bVerlengen = False
        MsgBox "inkorten"
    End If
        
   
    ' --------------------------------------------------------------
    ' ALLE VOORGAANDE SELECTIESETS VERWIJDEREN
    ' --------------------------------------------------------------
    'geeft soms fatale error !
    'Call SelectiesetsVerwijderen
    
    ' --------------------------------------------------------------
    ' AANMAKEN SELECTIESET OM EINDPUNT ARC (ZOEK LIJN)
    ' --------------------------------------------------------------
   
    'Aanmaken selectieset
    
    Dim ReturnEindpunt As Variant
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    
    ReturnEindpunt = ReturnObj1.EndPoint

    p1(0) = ReturnEindpunt(0) - 0.1
    p1(1) = ReturnEindpunt(1) - 0.1
    p1(2) = 0
    p2(0) = ReturnEindpunt(0) + 0.1
    p2(1) = ReturnEindpunt(1) + 0.1
    p2(2) = 0
    
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")
    If Err Then
        ssetObj.Clear
        ssetObj.Delete
        End
    End If
    
    ssetObj.Select acSelectionSetCrossing, p1, p2
    
    
    ' --------------------------------------------------------------
    ' UITLEZEN SELECTIESET EN VINDEN LIJN
    ' --------------------------------------------------------------

    Dim element As Object
    Dim LijnObj1 As AcadLine
    Dim sVorigeHandle As String

    sVorigeHandle = ReturnObj1.handle

    For Each element In ssetObj
        element.Color = acGreen

        If element.handle <> sVorigeHandle Then
            If element.EntityName = "AcDbLine" Then
                    'element.Highlight True
                    LijnTeller = LijnTeller + 1
                    Set LijnObj1 = element
            End If
        End If
    Next element

    ssetObj.Clear
    ssetObj.Delete
    
    If LijnTeller = 0 Then
        MsgBox "Geen lijnen (leidingen) aan de arc verbonden.", vbExclamation, "let op"
        End
    End If
    
      
   
    
    
    ' --------------------------------------------------------------
    ' INKORTEN LIJN
    ' --------------------------------------------------------------
    
    'MsgBox LijnObj1.handle
    Dim dNieuweLengte As Double
    
    If bVerlengen = True Then
        dNieuweLengte = LijnObj1.Length + dLengte
    Else
        dNieuweLengte = LijnObj1.Length - dLengte
    End If
    

    If LijnTeller = 1 Then
        If Lengte(LijnObj1.EndPoint, p) > Lengte(LijnObj1.StartPoint, p) Then
            Call LijnlengteWijzigen(LijnObj1, dNieuweLengte, False)
        Else
            Call LijnlengteWijzigen(LijnObj1, dNieuweLengte, True)
        End If
    End If
    
    
    ' --------------------------------------------------------------
    '  VERPLAATSEN ARC
    ' --------------------------------------------------------------
    
    'BEPALEN OF BEGIN- OF EINDPUNT VAN LIJN GEVONDEN IS
    
    
    If Lengte(LijnObj1.EndPoint, ReturnObj1.EndPoint) Then
        'Beginpunt Lijn Gevonden
        ReturnObj1.Move ReturnObj1.EndPoint, LijnObj1.EndPoint
    Else
        'Eindpunt Lijn Gevonden
        ReturnObj1.Move ReturnObj1.EndPoint, LijnObj1.StartPoint
    End If
    
    
    

    
End Sub













Private Sub CommandButton37_Click()
    
    Me.Hide
    
    On Error Resume Next
    F_LengteMonitor.show
    If Err Then
        Err.Clear
    End If
    
End Sub






' OPROEPEN ABOUTBOX
Private Sub Image1_Click()
    F_About.show
End Sub
Private Sub Image11_Click()
    F_About.show
End Sub
Private Sub Label17_Click()
    F_About.show
End Sub
Private Sub Label26_Click()
     F_About.show
End Sub

Private Sub CommandButton34_Click()
    Me.Hide
    M_6PlaatsenSlingers.OpvragenHOH
End Sub








Private Sub TextBox1_Change()
    ' ebr 14 juni 2004
    
    If GetalIsDouble(TextBox1.Text) = False Then
        TextBox1.BackColor = vbRed
        MsgBox "Geen juist getal ingevuld (gebruik een punt maar geen komma).", vbCritical, "Let op"
        F_LengteMonitor.Label5.Caption = 0
    Else
        TextBox1.BackColor = vbWhite
        F_LengteMonitor.Label5.Caption = TextBox1.Text
    End If
    
End Sub

'returnPnt = ThisDrawing.Utility.GetCorner(basePnt, "Enter Other corner: ")
'ThisDrawing.Utility.AngleFromXAxis
'THISDRAWING.HandleToObject
'ThisDrawing.Utility.TranslateCoordinates



Private Sub UserForm_Initialize()
    
    Call M_0SaveGetSettings.GetSettings
    Call LeesCfgFile

End Sub


Private Sub UserForm_Activate()
    'MsgBox "activate"
End Sub


Private Sub CommandButton28_Click()
    'TEST WERKING BEPALEN VAN: TEKEN_LIJN_OVER_POLYLINE EN BEPALEN SNIJPUNT_POLYLINE

    Dim plineObj As AcadLWPolyline
    Dim LijnObj As AcadLine
    
    Set plineObj = ThisDrawing.ModelSpace.Item(0)
    Set LijnObj = ThisDrawing.ModelSpace.Item(1)

    Call M_4BepalenLengteNEW.TekenLijnenOverPolyline(plineObj)
    Call M_4BepalenLengteNEW.BepalenSnijpunt(plineObj, LijnObj)
    Call M_4BepalenLengteNEW.VerwijderenLosseLines
    
End Sub



Private Sub CommandButton24_Click()
    M_TekenenAftaklijnen.StartTekenenAftaklijnen
End Sub

Private Sub CommandButton26_Click()
    ThisDrawing.SetVariable "LIMCHECK", 0      'limits off
    Call M_0SaveGetSettings.SaveSettings
    Call M_3PlaatsenArcsNew.PlaatsenArcsNew
    ThisDrawing.SetVariable "LIMCHECK", 1      'limits on
    End
End Sub



Private Sub CommandButton30_Click()

    For Each element In ThisDrawing.ModelSpace
         
        If element.EntityName = "AcDbLine" Then
            Call Punt1(element.StartPoint, 1)
            
        End If
                  
        If element.EntityName = "AcDbArc" Then
            'Call Punt1(Element.StartPoint, 2)
        End If
        
    Next element
End Sub





'**********************************************************************************************
'ONDERSTAANDE GEDEELTE: VOOR HET SELECTREREN VAN ROLTYPEN EN ROLLENGTEN
'**********************************************************************************************

Private Sub ComboBox3_Change()
    Call M_0CfgFiles.ChangeRollengten
End Sub


''''UITGEZET OP 6 MAART 2003:
''''
'''''''''' VERWIJDERD 24 FEB 2003
''''Private Sub CommandButton1_Click()
''''   Call PlaatsenBochten
''''End Sub
''''
''''Sub PlaatsenBochten()
''''
''''    ' -----------------------------
''''    ' Teken een hatch (legpatroon), deze vervolgens exploderen (eerste overige elementen verwijderen)
''''    ' daarna programma runnen.
''''    '----------------------------------------
''''
''''    MsgBox "FilterenLijnen IS VERWIJDERD"
''''    'Call FilterenLijnen
''''
''''
''''    'Laat handles zien: 'F_Main.Show
''''    'F_Main.Show
''''
''''    Dim HoHafstand As Double
''''    HoHafstand = F_Main.ComboBox2.Text
''''    Call VerbindLijnenMetArc(HoHafstand / 2)    ' principe zie: Call VerbindLijnen
''''End Sub



Private Sub CommandButton10_Click()
    Call M_0SaveGetSettings.SaveSettings
    F_Main.Hide
        
    
    F_LengteMonitor.Image1.Visible = True
    F_LengteMonitor.Image2.Visible = False
    
    'MsgBox "dit weggehaald op 27 mei 2004"
    'Unload Me
    
    'End ' dit weggehaald op 27 jan 2004
End Sub



Private Sub CommandButton12_Click()
    MsgBox "6 MAART 2003: KAN WEG ???", vbInformation

    Dim objPline As Object
    Set objPline = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)

    MsgBox BepaalLijnAanEindePolyline(objPline)
    
End Sub



Private Sub CommandButton2_Click()
    MsgBox "6 MAART 2003: KAN WEG ???", vbInformation
    
    Dim sLogFile As String
    sLogFile = "C:\Temp\WTH-logfile.txt"
    If Dir(sLogFile) <> "" Then Kill (sLogFile)         'verwijderen voorgaande logfile
    
    Call BerekenRollengte
End Sub




Private Sub CommandButton5_Click()

    Dim AantalElementVoorHethatchen As Double
    AantalElementVoorHethatchen = ThisDrawing.ModelSpace.Count
    
    Call M_0SaveGetSettings.SaveSettings
    Call PlaatsenHatch
    Call AlleLijnenInkorten(AantalElementVoorHethatchen)
    
    ThisDrawing.Regen (True)
    ThisDrawing.Application.Update
    
    '24 feb 2003 verwijderd en nieuwe routine gemaakt:
    'Call PlaatsenBochten
    
    Call M_3PlaatsenArcsNew.PlaatsenArcsNew
End Sub

Private Sub CommandButton6_Click()
    Call M_0SaveGetSettings.SaveSettings
    Call BepaalObstakels(True)
End Sub


Private Sub CommandButton7_Click()
    MsgBox "DE OUDE ROUTINE IS VERWIJDERD (ZIE TXT-BACKUP)"
    'Call BerekenKopStaart
End Sub


Private Sub CommandButton9_Click()
    
    'SHOW LOGFILE
    Dim sAntw As String
    antw = MsgBox("Openen WTH-logfile ?", vbCritical + vbYesNo, "WTH-logfile.txt")
    
    If antw = vbNo Then Exit Sub
    Shell "notepad.exe C:\Temp\WTH-logfile.txt", vbMaximizedFocus
End Sub

Private Sub CommandButton35_Click()
    'SHOW ROLLENGTEN.CFG
    Dim sAntw As String
    antw = MsgBox("Configuratie-bestand met rollengten wijzigen ?", vbCritical + vbYesNo, "Rollengten.cfg")
    
    If antw = vbNo Then Exit Sub

    sCfgFile = FindZoekpad("Rollengten.cfg")
    Shell "notepad.exe " & sCfgFile, vbMaximizedFocus
End Sub

Private Sub CommandButton14_Click()
   Call OpschuivenLegpatroon
End Sub

Sub OpschuivenLegpatroon()
    '* nieuwe sub 21 mei 2003
    Call M_0SaveGetSettings.SaveSettings
    Call M_6TrimmenLijnen.TrimmenLijnen
End Sub

Private Sub CommandButton31_Click()
    Call M_0SaveGetSettings.SaveSettings
    Call OffsetPoly
End Sub


Private Sub CommandButton21_Click()
    Call M_0SaveGetSettings.SaveSettings
    Call M_5BerekenLengteInLayers.BerekenLengteInColors
End Sub



Private Sub CommandButton29_Click()

    'MsgBox "*** 5 mei 2004 deze regel ertussen gezet"
    Call M_0CfgFiles.ChangeRollengten
        
    Call M_0SaveGetSettings.SaveSettings
    
    ' --------------------------------------------------------------
    ' CONTROLEREN RESERVE LENGTE (SNIT)
    ' --------------------------------------------------------------
    On Error Resume Next
    Dim dSnit As Double
    dSnit = F_Main.TextBox1.Text
    If Err Then
        Err.Clear
        MsgBox "De ingevoerde reservelengte (snit) is niet juist.", vbCritical, "Let op"
        Exit Sub
    End If
    ' --------------------------------------------------------------
    
    Load F_InvoerRollengte
    F_InvoerRollengte.Label3.Caption = F_Main.ComboBox3.Text
   
    F_InvoerRollengte.show
    
End Sub





Private Sub CommandButton36_Click()
    Call M_7WijzigenLengte.WijzigenLengte
    Call F_LengteMonitor.MeetLengte
End Sub



Private Sub UserForm_Deactivate()
    'automatisch meten met lengte monitor tijdelijk uitzetten
        'MsgBox "de-activate"
    
        'F_LengteMonitor.Image1.Visible = True
        'F_LengteMonitor.Image2.Visible = False
    
End Sub













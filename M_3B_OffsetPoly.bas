Attribute VB_Name = "M_3B_OffsetPoly"
Sub OffsetPoly()
    
    '-----------------------------------------------------------------------
    ' DEZE ROUTINE OFFSET DE GESELECTEERDE AANVOER-POLYLINE MET AFSTAND HOH
    '-----------------------------------------------------------------------
    
    'GEHEEL NIEUW OP 21 MAART 2003 EBR
    'TEN 15:09 TO 16:10= TOTAAL 1 UUR


    F_Main.Hide
    
    '-----------------------------------------------------------------------
    ' CONTROLE HOH
    '-----------------------------------------------------------------------
    
    Dim dHOH As Double
    
    On Error Resume Next
    dHOH = F_Main.ComboBox2.Text
    
    If Err Then
        Err.Clear
        MsgBox "De ingevoerde HOH-afstand is niet juist.", vbCritical, "Let op"
        Exit Sub
        F_Main.show
    End If
    
    '-----------------------------------------------------------------------
    ' LAAG "LEGPLANOMTREK" UITZETTEN.
    '-----------------------------------------------------------------------
        
    ThisDrawing.Layers.Item("Legplanomtrek").LayerOn = False
    If Err Then
        Err.Clear
        'MsgBox "Laag legplanomtrek kan niet worden uitgezet", vbCritical, "WTH-offset"
    End If
    
    
    '-----------------------------------------------------------------------
    ' OSNAP UITZETTEN EN ONTHOUDEN (LUKT NIET)
    '-----------------------------------------------------------------------
    
'    Dim iOsmode 'As Integer
'    iOsmode = ThisDrawing.GetVariable("OSMODE")
'    MsgBox iOsmode
'    ThisDrawing.SetVariable "OSMODE", 0
    
    '-----------------------------------------------------------------------
    ' SELECTEER DE AANVOER-POLYLINE
    '-----------------------------------------------------------------------

    
    Dim p1 As Variant
    Dim RetObj As AcadEntity
    
    Dim iFoutSelectieTeller As Integer
    
              
Opnieuw1:
    
    If iFoutSelectieTeller > 1 Then End
    
    ThisDrawing.Utility.GetEntity RetObj, p1, "Selecteer de aanvoer-polyline."
    'RetObj.Highlight (True)
    
    If Err Then
        If iFoutSelectieTeller < 1 Then MsgBox "Verkeerd geselecteerd", vbCritical, "Let op"
        Err.Clear
        'RetObj.Highlight (False)
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw1
    End If
    
    
    If RetObj.EntityName = "AcDbPolyline" Or RetObj.EntityName = "AcDbLine" Then
    Else
        If iFoutSelectieTeller < 1 Then MsgBox "Geen line of polyline geselecteerd", vbCritical
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw1
        'End
    End If
   
    '-----------------------------------------------------------------------
    ' GESELECTEERDE POLYLINE IN LAAG LEGPLAN ZETTEN
    '-----------------------------------------------------------------------
    
    If RetObj.EntityName = "AcDbPolyline" Then
        RetObj.Highlight (True)
        Call AanmakenLaag("Legplan", 2, False)
        RetObj.Layer = "Legplan"
    End If
    
    
    '-----------------------------------------------------------------------
    ' OSNAP WEER TERUGZETTEN
    '-----------------------------------------------------------------------
'
'    ThisDrawing.SetVariable "OSMODE", iOsmode
    
    '-----------------------------------------------------------------------
    ' LAAG "LEGPLANOMTREK" AANZETTEN.
    '-----------------------------------------------------------------------
        
    ThisDrawing.Layers.Item("Legplanomtrek").LayerOn = True
    If Err Then
        Err.Clear
    End If

    '-----------------------------------------------------------------------
    ' GEEF OFFSET RICHTING (EN DAARMEE AANTAL AAN)
    '-----------------------------------------------------------------------

opnieuw:

    Dim sAantal As String
    Dim Aantal As Integer
    Dim p2 As Variant
    Dim sMelding As String
    Dim sAntw As String
    
    'AANTAL OPGEVEN EN RICHTING AANWIJZEN
    If F_Main.CheckBox10.Value = True Then
        sMelding = "Geef offset-richting aan:"
        
        sAantal = InputBox("Geeft aantal groepen op.", "Aantal groepen", "2")
        If sAantal = "" Then Exit Sub
        
        If sAantal = Val(sAantal) Then
            Aantal = Val(sAantal)
            Aantal = Aantal * 2
            
            ' CONTROLE INDIEN > 9 GROEPEN
            If Aantal > 18 Then
                sAntw = MsgBox("Meer dan 10 groepen plaatsen ?", vbYesNo, "Offset leidingen")
                If sAntw = vbNo Then GoTo opnieuw
            End If
            
            
            
        Else
            MsgBox " verkeerde waarde Invoer Aantal:" & sAantal, vbInformation
        End If
    Else
        ' EINDPUNT EN RICHTING AANWIJZEN:
        sMelding = "Geef aan tot waar de offset-polylines (leidingen) moeten komen:"
    End If
    
    p2 = ThisDrawing.Utility.GetPoint(p1, sMelding)
    If Err Then
        Err.Clear
        End
    End If
    
        
        
        

   
    
   
    '-----------------------------------------------------------------------
    ' BEPAAL DE AANTAL OFFSETS
    '-----------------------------------------------------------------------
    
    Dim dLengte As Double
    
    If F_Main.CheckBox10.Value = False Then
        dLengte = Lengte(p1, p2)
        Aantal = dLengte / dHOH
    End If
    
    ''Aantal = Aantal / 2
    'MsgBox "Aantal offset is: " & Aantal
    
    
    '6 aug 2003
    'indien gekozen voor duo, dan c.a. 2x zo veel als geselecteerde punt (behalve als aantal opgeven wordt.)
    If F_Main.CheckBox9.Value = True And F_Main.CheckBox10.Value = False Then
        Aantal = Aantal * 4
        If Aantal Mod (2) <> 0 Then Aantal = Aantal + 1
    End If
    
    
    
    '-----------------------------------------------------------------------
    ' MAAK EERST EEN TEST OFFSET (0.5 hoh), OM TE KIJKEN OF RICHTING GOED IS
    '-----------------------------------------------------------------------
    
'    LET OP: + WAARDE BIJ OFFSET, BETEKENT DAT DE KOPIE GROTER WORDT
'    The distance to offset the object.
'    The offset can be a positive or negative number, but it cannot equal zero.
'    If the offset is negative, this is interpreted as being an offset to make a "smaller"
'    curve (that is, for an arc it would offset to a radius that is
'    "Distance less" than the starting curve's radius).
'    If "smaller" has no meaning, then it would offset
'    in the direction of smaller X, Y, and Z WCS coordinates.

    
    OffsetObj = RetObj.Offset(dHOH / 2)
    'OffsetObj(0).Color = acMagenta
    
    'LET OP: UPDATE MOET ANDERS ZAL SELECTIESET FENCE DEZE NIET VINDEN !!
    OffsetObj(0).Update
        
    'MsgBox OffsetObj(0).handle
    
    '--------------------------------------------------------------------------
    ' SELECTIESET FENCE AANMAKEN (OP HULP-OFFSET-POLYLINE) OM TE CONTROLEREN
    ' OF DE OFFSET-RICHTING GOED WAS.
    '--------------------------------------------------------------------------
    
    Call SelectiesetsVerwijderen
    
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")
        
    
    Dim mode As Integer
    Dim pointsArray(0 To 5) As Double
    mode = acSelectionSetFence
    
    pointsArray(0) = p1(0):     pointsArray(1) = p1(1):     pointsArray(2) = 0
    pointsArray(3) = p2(0):     pointsArray(4) = p2(1):     pointsArray(5) = 0
    
        
    ssetObj.SelectByPolygon mode, pointsArray
    'LET OP DUS GEEN ssetObj.Select mode, pointsArray

    
    Dim element As Object
    Dim bJuisteOffsetRichting As Boolean
    
    For Each element In ssetObj
        'MsgBox OffsetObj(0).handle
        If element.handle = OffsetObj(0).handle Then bJuisteOffsetRichting = True
    Next element
    
    
    ssetObj.Clear
    ssetObj.Delete
    
    'MsgBox "bJuisteOffsetRichting = " & bJuisteOffsetRichting
        
      
    '-----------------------------------------------------------------------
    ' VOORGAANDE HULP-OFFSET VERWIJDEREN
    '-----------------------------------------------------------------------

    OffsetObj(0).Erase
    
    '-----------------------------------------------------------------------
    ' ALLE OFFSET-POLYLINES PLAATSEN IN DE JUISTE RICHTING
    '-----------------------------------------------------------------------

    If F_Main.CheckBox5.Value = True And RetObj.EntityName <> "AcDbPolyline" Then MsgBox "De beginpunten van de lijnen worden alleen gelijk gemaakt voor (aanvoer/ retour) polylines.", vbInformation, "WTH-Offset"
        
    Dim t As Integer
    'Dim S As Integer    'sign, plus of min
    Dim LaatsteOffsetObj As AcadEntity
    
    
'''     ORIGINEEL VOOR 7 APRIL 2003
'''     For t = 1 To Aantal
'''        If bJuisteOffsetRichting = True Then
'''            OffsetObj = RetObj.Offset(t * dHOH)
'''        Else
'''            OffsetObj = RetObj.Offset(-t * dHOH)
'''        End If
    
'''     NIEUW OP 7 APRIL 2003


    
     Dim c As Double
     For t = 1 To Aantal
        
        ''' NIEUW INBOUW VAN DUOLEIDING-MOGELIJKHEID 6 AUG 2003
        'offset instellen d.m.v. counter (c)
        If F_Main.CheckBox9.Value = False Then
            'geen duo-leidingen
            c = c + 1
        Else
            'wel duo-leidingen
            If t > 1 Then
                    'duo-leidingen, offset is altijd 1 cm en tussenliggend 0.5 HOH
                    If t Mod (2) = 0 Then
                        'c = c + (1 / dHOH)                      'offset van 1 cm
                        c = c + (2 / dHOH)                      'offset van 2 cm   (gewijzigd 10 feb 2004)
                    Else
                        '*** origineel 14 feb 2004: 'c = c + ((1 / dHOH) * (dHOH / 2))       'offset van 1cm x 0.5 HOH
                         c = c + ((1 / dHOH) * dHOH)        'offset van 1cm x 0.5 HOH
                    End If
            Else
                    '*** origineel 14 feb 2004:
                    'c = c + 0.5    'dan is eerste offsetlijn 0.5 HOH
                    c = c + 1       'dan is eerste offsetlijn 1 HOH
            End If
        End If
     

        If bJuisteOffsetRichting = True Then
                If F_Main.CheckBox7 Then
                    'start met halve hoh-afstand
                    '***origineel 14 feb 2004:
                    OffsetObj = RetObj.Offset((c * dHOH) - (0.5 * dHOH))
                     'OffsetObj = RetObj.Offset((c * dHOH) - (dHOH))
                Else
                    'start met hele hoh-afstand
                    OffsetObj = RetObj.Offset(c * dHOH)
                    
                End If
                
        Else
                If F_Main.CheckBox7 Then
                    'start met halve hoh-afstand
                    '*** origineel 14 feb 2004
                    OffsetObj = RetObj.Offset((-c * dHOH) + (0.5 * dHOH))
                    'OffsetObj = RetObj.Offset((-c * dHOH) + dHOH)
                Else
                    'start met hele hoh-afstand
                    OffsetObj = RetObj.Offset(-c * dHOH)
                End If
        End If


        'MsgBox "wachten"
                        
        'OffsetObj.Layer = RetObj.Layer
        
        
        'BEGINPUNTEN VAN LWT-POLYLINE GELIJK MAKEN AAN BEGINPUNT VAN GESELECTEERDE POLYLINE.
        If F_Main.CheckBox5.Value = True Then
            Set LaatsteOffsetObj = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)
            
            
            'Dim intVCnt As Integer
            'Dim varVert As Variant
            Dim coord As Variant
            
            coord = RetObj.Coordinate(0)
            coord(0) = coord(0)
            LaatsteOffsetObj.Coordinate(0) = coord
            LaatsteOffsetObj.Update
        End If
        
        
    Next t
    
    'INDIEN GEKOZEN IS OM DE EERSTE POLYLINE MET HALVE HOH TE OFFSETTEN (Start halve hoh)
    'DAN AUTOMATISCH EERSTE GESELECTEERDE POLYLINE VERWIJDEREN (ERASE)
    '+ automatisch vragen of legpatroon opgeschoven moet worden.
    
    Dim antw As String
    If F_Main.CheckBox7 Then
        RetObj.Delete
        
        antw = MsgBox("Legpatroon opschuiven ?", vbQuestion + vbYesNo, "Doorgaan met opschuiven")
        If antw = vbYes Then Call F_Main.OpschuivenLegpatroon
    End If
    
    '* nieuw 21 mei 2003
    'MsgBox "tot hier 5 mei 2004: end verwijderd", vbExclamation
    
    'End
End Sub



   'DUO
'''                    If t Mod 2 = 0 Then
'''                        OffsetObj = RetObj.Offset(t * dHOH / 2)
'''                    Else
'''                        OffsetObj = RetObj.Offset(t * dHOH / 2 - 2)
'''                    End If
'''

Attribute VB_Name = "M_3PlaatsenArcsNew"
Public Veld(0 To 500) As String

Public Sub PlaatsenArcsNew()

    F_Main.Hide
    On Error Resume Next
    
    Dim KleurNr As Integer
    Dim KleurOrgNr As Integer
    
    KleurNr = 1
    
    '------------------------------------------------------------------
    ' EERST REGENEREREN ANDERS SOMS NIET GOED
    '------------------------------------------------------------------
    '*verplaatst 21 mei 2003
    'ThisDrawing.Regen acActiveViewport
    

    '------------------------------------------------------------------
    ' AANGEVEN BEGIN EN EINDPUNT VAN FENCE SELECTIE-LIJN
    '------------------------------------------------------------------
    
 
    Dim ReturnObj1 As AcadObject
    Dim returnObj2 As AcadObject
    Dim bStartpointGevonden As Boolean

    Dim p1 As Variant
    Dim p2 As Variant
    
    
       
    Dim iFoutSelectieTeller As Integer
    
Opnieuw1:

    '*** voor 27 jan 2004 was dit: If iFoutSelectieTeller > 1 Then End
    If iFoutSelectieTeller > 1 Then End
 
    ThisDrawing.Utility.GetEntity ReturnObj1, p1, "Selecteer beginpunt van de eerste leiding"
    
    
    
    If Err Then
        Err.Clear
        If iFoutSelectieTeller < 1 Then MsgBox "Verkeerde selectie", vbCritical, "Let op"
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw1
        
    End If
    If ReturnObj1.EntityName <> "AcDbLine" Then
        MsgBox "Geen lijn geselecteerd.", vbCritical, "Let op"  ': End
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw1
    End If
        
    KleurOrgNr = ReturnObj1.Color
    ReturnObj1.Color = KleurNr
    
    
    iFoutSelectieTeller = 0
    
Opnieuw2:

    If iFoutSelectieTeller > 1 Then End
    
    ThisDrawing.Utility.GetEntity returnObj2, p2, "Selecteer de tweede leiding."
    
    
    
    If Err Then
        Err.Clear
        If iFoutSelectieTeller < 1 Then MsgBox "Verkeerde selectie", vbCritical, "Let op"
        ReturnObj1.Color = KleurOrgNr
        
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw2
        
    End If
    If returnObj2.EntityName <> "AcDbLine" Then
        MsgBox "Geen lijn geselecteerd.", vbCritical, "Let op"
        ReturnObj1.Color = KleurOrgNr
        
        iFoutSelectieTeller = iFoutSelectieTeller + 1
        GoTo Opnieuw2
    End If
    'returnObj2.Color = KleurNr
    
    If ReturnObj1.handle = returnObj2.handle Then
        MsgBox "Tweemaal dezelfde lijn geselecteerd.", vbExclamation, "Let op"
        End 'Exit Sub
    End If
    
    'ZOOMEN ANDERS VERKEERDE SELECTIESET:
    ZoomExtents
    '* NIEUW OP 21 MEI 2003 !!!
    ThisDrawing.Regen acActiveViewport
    
    'BEPALEN OF BEGIN OF EINDPUNT VAN DE LIJN GESELECTEERD IS.
    
    If Lengte(ReturnObj1.EndPoint, p1) > Lengte(ReturnObj1.StartPoint, p1) Then bStartpointGevonden = True
    'MsgBox "bStartpointGevonden=" & bStartpointGevonden

    '* 20 mei 2003 er bij geschreven
    If bStartpointGevonden = True Then bBijStartpointBeginnen = True

'''''    '* 20 mei 2003 uitgezet want is niet van toepassing (berekening stond er eigenlijk al boven)
'''''    '------------------------------------------------------------------
'''''    ' BEPALEN AAN WELKE ZIJDE BEGINNEN (BEGINLEIDING, PUNT P1)
'''''    '------------------------------------------------------------------
''''''    Dim bBijStartpointBeginnen As Boolean
''''''    If Lengte(P1, ReturnObj1.EndPoint) > Lengte(P1, ReturnObj1.StartPoint) Then
''''''        'MsgBox "Bij startpoint van de eerste lijn beginnen"
''''''        bBijStartpointBeginnen = True
''''''    Else
''''''        'MsgBox "Bij endpoint van de eerste lijn beginnen"
''''''        bBijStartpointBeginnen = False
''''''    End If
    

    '------------------------------------------------------------------
    ' AANMAKEN SELECTIESET MBV FENCE
    '------------------------------------------------------------------
 
    Dim ssetObj As AcadSelectionSet
    Dim mode As Integer
    Dim pointsArray(0 To 5) As Double
    
    'MsgBox "hier fout"
    
    For Each ssetObj In ThisDrawing.SelectionSets
        ssetObj.Clear
        ssetObj.Delete
    Next ssetObj
    
    
    Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")
    mode = acSelectionSetFence
    
    pointsArray(0) = p1(0):     pointsArray(1) = p1(1):     pointsArray(2) = 0
    pointsArray(3) = p2(0):     pointsArray(4) = p2(1):     pointsArray(5) = 0
    
    
    ssetObj.SelectByPolygon mode, pointsArray
    'MsgBox "ssetObj.count=" & ssetObj.Count
  
    
    '---------------------------------------------------------------------------------------------------------
    ' * nieuw 21 mei 2003 (BIJ ONEVEN AANTAL GING AFRONDEN NIET ALTIJD GOED)
    '
    ' CONTROLEREN OF DE TWEE GESLECTEERDE LIJNEN OOK IN DE SELECTIE ZITTEN
    ' DEZE ANDERS TOEVOEGEN IN SELECTIESET
    ' INDIEN ONEVEN-AANTAL ELEMENTEN IN SELECTIESET, DAN LAATST GESELECTEERDE LIJN VERWIJDEREN
    '---------------------------------------------------------------------------------------------------------
            
    
    Dim element As Object
    Dim bReturnObj1Gevonden As Boolean
    Dim bReturnObj2Gevonden As Boolean
    
    'CONTROLEREN OF GESELECTEERDE LIJN IN SELECTIESET AANWEZIG IS, ANDERS TOEVOEGEN
    For Each element In ssetObj
        'element.Color = acGreen
        If element.handle = ReturnObj1.handle Then bReturnObj1Gevonden = True
        If element.handle = returnObj2.handle Then bReturnObj2Gevonden = True
    Next element
    
    Dim addObjects(0) As AcadEntity
    Set addObjects(0) = ReturnObj1
    If bReturnObj1Gevonden = False Then ssetObj.AddItems addObjects
    
    Set addObjects(0) = returnObj2
    If bReturnObj2Gevonden = False Then ssetObj.AddItems addObjects
    
    'MsgBox ssetObj.Count
    
    'INDIEN ELEMENTEN NIET IN LAAG LEGPLAN OF GROEP_ DAN VERWIJDEREN UIT SELECTIESET
    Dim removeObjects(0) As AcadEntity
    
    If F_Main.CheckBox8.Value = True Then
        For Each element In ssetObj
            'element.Color = acGreen
            If element.Layer = "Legplan" Or Left$(element.Layer, 6) = "groep_" Then
            Else
                Set removeObjects(0) = element
                ssetObj.RemoveItems removeObjects
            End If
        Next element
    End If
    'MsgBox ssetObj.Count
           
      
                        
    
    'INDIEN EEN ONEVEN AANTAL LIJNEN IN SELECTIESET,DAN LAATST GESELECTEERDE LIJN VERWIJDEREN
    'Dim removeObjects(0) As AcadEntity
    If ssetObj.Count Mod 2 <> 0 Then
        Set removeObjects(0) = returnObj2
        ssetObj.RemoveItems removeObjects
    End If
    
    'MsgBox ssetObj.Count

'    For Each element In ssetObj
'        element.Color = acMagenta
'    Next element
    
    
    '---------------------------------------------------------------------------------------------------------
    ' UITLEZEN WELKE LEIDINGEN IN SELECTIESET ZITTEN, Y-BEGINPUNT LIJN + GEVONDEN HANDLES IN ARRAY ZETTEN
    ' ofwel in array: beginY/handle (zodat dit gesorteerd kan worden)
    '---------------------------------------------------------------------------------------------------------
        
   
    
    'Dim element As Object          'HIERBOVEN GEDECLAREERD
    Dim ArrHandles() As String

    ReDim Preserve ArrHandles(0)
    Dim bWelAfronden As Boolean
    Dim iLijnteller As Integer

    
    
     For Each element In ssetObj

            bWelAfronden = False
            If element.EntityName = "AcDbLine" Then


                    If F_Main.CheckBox8.Value = True Then
                        If element.Layer = "Legplan" Or Left$(element.Layer, 6) = "groep_" Then bWelAfronden = True
                    Else
                        bWelAfronden = True
                    End If

                    If bWelAfronden = True Then
                        iLijnteller = iLijnteller + 1

                        element.Color = KleurNr
                        element.Update

                        'MsgBox Element.handle
                        '((Element.Highlight True)) ZIE ARRAY
                        'VULLEN ARRAY
                        ReDim Preserve ArrHandles(UBound(ArrHandles) + 1)
                        ArrHandles(UBound(ArrHandles)) = element.handle
                    End If


            End If
    Next element

   
    ssetObj.Clear
    ssetObj.Delete
    
    '--------------------------------------------------------------------------------------
    ' CONTROLEREN HOEVEEL LIJNEN ER GEVONDEN ZIJN.
    '--------------------------------------------------------------------------------------
   
    If iLijnteller = 0 Then
        MsgBox "Geen lijnen gefilterd. (verkeerde lagen > vink optie uit)", vbExclamation, "Geen lijnen gevonden."
        Exit Sub
    End If
    
    
    
'''''   ONDERSTAANDE ROUTINE NIET MEER NODIG DOOR AANPASSINGEN 21 MEI 2003
'''''    '--------------------------------------------------------------------------------------
'''''    ' UITLEZEN ARRAY MET DE HANDLES EN CONTROLEREN OF HANDLE VAN DE TWEE GESELECTEERDE
'''''    ' LIJNEN IN ARRAY BESTAAT, ANDERS TOEVOEGEN.
'''''    '--------------------------------------------------------------------------------------
'''''
'''''
'''''    Dim t As Integer
'''''    Dim bHandleEersteLijnGevonden As Boolean
'''''    Dim bHandleTweedeLijnGevonden As Boolean
'''''
'''''
'''''    For t = LBound(ArrHandles) + 1 To UBound(ArrHandles)            ' + 1 anders wordt eerste lege veld afgebeeld
'''''        'MsgBox ArrHandles(t), vbExclamation
'''''
'''''        If ArrHandles(t) = ReturnObj1.handle Then bHandleEersteLijnGevonden = True
'''''        If ArrHandles(t) = returnObj2.handle Then bHandleTweedeLijnGevonden = True
'''''    Next t
'''''
'''''    If bHandleEersteLijnGevonden = False Then
'''''        'TOEVOEGEN HANDLE VAN EERSTE GESELECTEERDE LIJN
'''''        ReDim Preserve ArrHandles(UBound(ArrHandles) + 1)
'''''        ArrHandles(UBound(ArrHandles)) = ReturnObj1.handle
'''''    End If
'''''
'''''    If bHandleTweedeLijnGevonden = False Then
'''''         'TOEVOEGEN HANDLE VAN TWEEDE GESELECTEERDE LIJN
'''''        ReDim Preserve ArrHandles(UBound(ArrHandles) + 1)
'''''        ArrHandles(UBound(ArrHandles)) = returnObj2.handle
'''''    End If
    
'''''    '--------------------------------------------------------------------------------------
'''''    ' TELLEN VAN HET AANTAL GESLECTEERDE LEIDINGEN
'''''    '--------------------------------------------------------------------------------------
'''''
'''''    'MsgBox "Er zijn " & UBound(ArrHandles) & " leidingen geselecteerd", vbInformation, "Leiding selectie"
'''''
'''''    Dim bEvenAantal As Integer
'''''    If UBound(ArrHandles) Mod 2 = 0 Then
'''''        bEvenAantal = True
'''''        MsgBox "Even aantal", vbExclamation
'''''    Else
'''''         bEvenAantal = False
'''''
'''''        '* 21 mei gewijzigd in true
'''''        'bEvenAantal = True
'''''
'''''        '19 mei 2003 melding uitgezet
'''''        MsgBox "Oneven aantal leidingen !", vbExclamation, "Let op"
'''''    End If
    
    
    bEvenAantal = True
    
    '--------------------------------------------------------------------------------------
    ' TEST UITLEZEN HANDLES
    '--------------------------------------------------------------------------------------
'    For t = LBound(ArrHandles) + 1 To UBound(ArrHandles)            ' + 1 anders wordt eerste lege veld afgebeeld
'        MsgBox LeesHandle(ArrHandles(t))
'    Next t

    '--------------------------------------------------------------------------------------
    ' ROTEREN GESELCTEERDE LIJNEN
    '(UITLEZEN ARRAY MET DE HANDLES EN OP 0 GRADEN ROTEREN + HIGHLIGHTEN)
    '--------------------------------------------------------------------------------------
    
    Dim LijnObj As AcadLine
    Dim RotHoek As Double
    
    RotHoek = ReturnObj1.Angle
    Dim Pbegin As Variant
    Dim Peind As Variant
    Dim Ptemp As Variant
    
    Dim Y() As String
    ReDim Preserve Y(0)
    
    'Dim sHandle As String
    
    For t = LBound(ArrHandles) + 1 To UBound(ArrHandles)            ' + 1 anders wordt eerste lege veld afgebeeld
        
        Set LijnObj = ThisDrawing.HandleToObject(ArrHandles(t))
        LijnObj.Color = KleurOrgNr
        
        'RECHT ROTEREN LIJNEN (OP NUL GRADEN):
        LijnObj.Rotate ReturnObj1.EndPoint, -RotHoek
        'LijnObj.Highlight True
        
        
        'BEPALEN OF LIJNEN RECHT LIGGEN:
        Pbegin = LijnObj.StartPoint
        Peind = LijnObj.EndPoint
        'MsgBox Pbegin(1) & Chr(10) & Chr(13) & Peind(1), vbExclamation
        If Round(Pbegin(1), 1) <> Round(Peind(1), 1) Then MsgBox "De leiding met handle " & LijnObj.handle & " heeft een andere hoek dan de geselecteerde leiding.", vbCritical, "Let op": End
            
            
        'Y-COORDINAAT INVULLEN IN ARRAY (Y-BEGINPUNT = Y-EINDPUNT, WANT LIJNEN OP 0 GRADEN !)
        ReDim Preserve Y(UBound(Y) + 1)
        Y(UBound(Y)) = Pbegin(1)
         
         
        'ZELFDE RICHTING ROTEREN LIJNEN, BEGINPUNT LINKS VAN EINDPUNT ANDERS DRAAIEN
        Pbegin = LijnObj.StartPoint
        Peind = LijnObj.EndPoint
        If Pbegin(0) > Peind(0) Then
        'beginpunt van lijn ligt rechts van eindpunt, daarom omdraaien.
            Ptemp = LijnObj.EndPoint
            LijnObj.EndPoint = LijnObj.StartPoint
            LijnObj.StartPoint = Ptemp
        End If
        
        'handle in array overschrijven met y-coordinaat & "/" &  handle
        ArrHandles(t) = ycoordENhandle(LijnObj)
        
        
        '*** VULLEN ARRAY VELD (DEZE LATER SORTEREN !)
        Veld(t) = ycoordENhandle(LijnObj)
        
        
    Next t
    
    'sHandle = LeesHandle(ArrHandles(t))
   
    
    '**************************************************************************************
    ' (1) OP VOLGORDE SORTEREN VAN DE LIJNEN (OP Y-COORDINAAT) + ARCS PLAATSEN
    '**************************************************************************************

    'MsgBox "einde array", vbInformation
    
    
    'SORTEREN GEGEVENS:
    Call BubbleSort
    
'''    Debug.Print "*******************"
'''    Debug.Print Time
'''    Debug.Print "*******************"
'''    For x = 0 To UBound(Veld) '- 1
'''        Debug.Print Veld(x)
'''    Next x
'''
'''    MsgBox "tot hier 30 mei 2003"
'''    End
    
'''    'GESORTEERDE GEGEVENS UITLEZEN UIT ARRAY:
'''    Dim x As Integer
'''    For x = 0 To UBound(Veld)
'''        'MsgBox Veld(x)
'''        If Veld(x) <> "" Then
'''            MsgBox Veld(x), vbInformation, "GESORTEERD"
'''            Debug.Print Veld(x)
'''        End If
'''    Next x
            
    
    '--------------------------------------------------------------------------------------
    ' EXTRA ZOOMEXTENTS 30 MAART 2003 (WANT SOMS FOUTEN BIJ HET TERUGPLAATSEN/ ROTEREN)
    '--------------------------------------------------------------------------------------
   
   ZoomExtents
   
   
    '**************************************************************************************
    ' (2) OP VOLGORDE SORTEREN VAN DE LIJNEN (OP Y-COORDINAAT) + ARCS PLAATSEN
    '**************************************************************************************
    
    
    'GESORTEERDE ARRAY !
    Dim LijnObj2 As AcadLine
    
    Dim Plijn1Beginpunt As Variant
    Dim Plijn2Beginpunt As Variant
    Dim Plijn1Eindpunt As Variant
    Dim Plijn2Eindpunt As Variant
    Dim Pnew(0 To 2) As Double       'nieuwe start of eindpunt van de lijn (inkorten na plaatsen arc)
    
    
    Dim Centerp(0 To 2) As Double
    Dim PuntObj As AcadPoint
    Dim Radius As Double
    Dim ArcObj As AcadArc
    
    Dim sHandle1 As String
    Dim sHandle2 As String
    Dim pos As Integer
    
    Dim ModIsWaar As Boolean
    
    
    'UITLEZEN GESORTEERD ARRAY:
    For x = 0 To UBound(Veld) '- 1
        'MsgBox Veld(x)
        If Veld(x) <> "" Then
            'Debug.Print Veld(X)
            
            pos = InStr(1, Veld(x), "/", vbTextCompare)
            sHandle1 = Right(Veld(x), Len(Veld(x)) - pos)
            
            pos = InStr(1, Veld(x + 1), "/", vbTextCompare)
            sHandle2 = Right(Veld(x + 1), Len(Veld(x + 1)) - pos)
            
            'afbeelden y-coordinaat van geroteerde lijnen + handle
            'Debug.Print sHandle
            'MsgBox "sHandle=" & sHandle
            
        'End If
        'Next x
    
    



        'EERSTE LIJN UIT LIJST (ARRAY)
         Set LijnObj = ThisDrawing.HandleToObject(sHandle1)
         Plijn1Beginpunt = LijnObj.StartPoint
         Plijn1Eindpunt = LijnObj.EndPoint


        'BIJ LAATSTE LIJN, DEZE TERUGDRAAIEN EN EINDE
        'OF DIT WEG EN FOR T =.... -1 !!!! ZIE HIERBOVEN.
'        If t = UBound(ArrHandles) Then
'            LijnObj.Rotate returnObj1.EndPoint, RotHoek
'            LijnObj.Highlight True
'            'MsgBox "Laatste lijn: " & ArrHandles(t)
'            End
'        End If


          'TWEEDE LIJN UIT LIJST (ARRAY)
         'If t < UBound(ArrHandles) - 1 Then
            Set LijnObj2 = ThisDrawing.HandleToObject(sHandle2)
            
            
            '*** nieuw 30 mei 2003:
            If Err Then
                '* 30 mei 2003, bij overgang van negatieve naar positieve y-coordinaat
                'van de gefilterde lijnen ontstaat error want in array zitten 500
                'elementen waaronder veel lege velden: deze worden
                'door de filter (bubblesort) routine tussen de positieve en nagatieve y-waarden ingeplaatst !
                Err.Clear
            End If
            
            Plijn2Beginpunt = LijnObj2.StartPoint
            Plijn2Eindpunt = LijnObj2.EndPoint
         'End If


        '----------------------------------------------------------------------------------
        'PLAATSEN ARC
        '----------------------------------------------------------------------------------
        
        'AANGEVEN RICHTING (BEGINPUNT OF EINDPUNT VAN EERSTE LEIDING)
        If x Mod 2 <> 0 Then
            If bBijStartpointBeginnen = True Then
                ModIsWaar = False
            Else
                ModIsWaar = True
            End If
        Else
            If bBijStartpointBeginnen = True Then
                ModIsWaar = True
            Else
                ModIsWaar = False
            End If
        End If
        
        'INDIEN ONEVEN AANTAL LEIDINGEN, DAN 'RICHTING' OMKEREN:
        If bEvenAantal = False Then ModIsWaar = Not (ModIsWaar)
        
        
        If ModIsWaar = True Then
            'ARCS AAN DE BEGINZIJDE (LINKS):

            'X-PUNT ARCCENTER (IS GELIJK AAN BEGINPUNT VAN KORTSTE LIJN):
            If Plijn1Beginpunt(0) > Plijn2Beginpunt(0) Then
                Centerp(0) = Plijn1Beginpunt(0)

                Pnew(0) = Plijn1Beginpunt(0)            '2 ipv 1 !!!
                Pnew(1) = Plijn2Beginpunt(1)
                LijnObj2.StartPoint = Pnew
                LijnObj2.Color = KleurOrgNr
                'LijnObj2.Color = acMagenta
            Else
                Centerp(0) = Plijn2Beginpunt(0)

                Pnew(0) = Plijn2Beginpunt(0)
                Pnew(1) = Plijn1Beginpunt(1)
                LijnObj.StartPoint = Pnew
                'LijnObj.Color = acGreen
                LijnObj.Color = KleurOrgNr
            End If

            'Y-PUNT ARCCENTER:
            Centerp(1) = (Plijn1Beginpunt(1) + Plijn2Beginpunt(1)) / 2
            Radius = Abs(Plijn1Beginpunt(1) - Plijn2Beginpunt(1))       'dit is de hoh
            Radius = Radius / 2
            If Radius <> 0 Then Set ArcObj = ThisDrawing.ModelSpace.AddArc(Centerp, Radius, 1.5707963267949, 4.71238898038469)

        Else
            'ARCS AAN DE EINDZIJDE (RECHTS):

            'X-PUNT ARCCENTER (IS GELIJK AAN EINDPUNT VAN KORTSTE LIJN):
            If Plijn1Eindpunt(0) < Plijn2Eindpunt(0) Then
                Centerp(0) = Plijn1Eindpunt(0)

                Pnew(0) = Plijn1Eindpunt(0)
                Pnew(1) = Plijn2Eindpunt(1)
                LijnObj2.EndPoint = Pnew
                'LijnObj2.Color = acMagenta
                LijnObj2.Color = KleurOrgNr
            Else
                Centerp(0) = Plijn2Eindpunt(0)

                Pnew(0) = Plijn2Eindpunt(0)
                Pnew(1) = Plijn1Eindpunt(1)
                LijnObj.EndPoint = Pnew
                LijnObj.Color = KleurOrgNr
            End If

            'Y-PUNT ARCCENTER:
            Centerp(1) = (Plijn1Eindpunt(1) + Plijn2Eindpunt(1)) / 2
            Radius = Abs(Plijn1Eindpunt(1) - Plijn2Eindpunt(1))       'dit is de hoh
            Radius = Radius / 2
            If Radius <> 0 Then Set ArcObj = ThisDrawing.ModelSpace.AddArc(Centerp, Radius, 4.71238898038469, 1.5707963267949)

         End If


        'TERUG DRAAIEN ARC EN LINES
        If Radius <> 0 Then
            ArcObj.Rotate ReturnObj1.EndPoint, RotHoek
            ArcObj.Highlight True
        End If

        ArcObj.Color = KleurOrgNr
        
        'nieuw 1 april 2003:
        'arcobj.Layer = ReturnObj1.Layer
        
        ArcObj.Highlight True

        LijnObj.Rotate ReturnObj1.EndPoint, RotHoek
        LijnObj.Highlight True
        
        'test 30 maart 2003:
        'Punt (ReturnObj1)
        'LijnObj.Update
        'MsgBox "wacht"


        End If
    Next x


 'WEER TERUGZOOMEN:
 ZoomPrevious
 '* 21 mei 2003 extra zoomprevious want ging niet terug naar begin
 ZoomPrevious
    
End Sub

Function ycoordENhandle(Object As AcadObject) As String

    'object = line

    Dim Beginpunt As Variant
    Dim Xbegincoord As Double

    Beginpunt = Object.StartPoint
    Xbegincoord = Beginpunt(1)
    Xbegincoord = Round(Xbegincoord, 0)
    

    ycoordENhandle = Xbegincoord & "/" & Object.handle
End Function

Function LeesHandle(zoekstring As String) As String
    Dim pos As Integer
    pos = InStr(1, zoekstring, "/")
    LeesHandle = Right(zoekstring, Len(zoekstring) - pos)
End Function

  
''''------------------------------------------------------------------------------------------------
''''PRINCIPE ARRAY: (SCHRIJVEN EN LEZEN)
''''------------------------------------------------------------------------------------------------

'''Public ArrLokatie() As String
'''
''''Dit boven de aanroep van het array plaatsen:
'''ReDim Preserve ArrLokatie(0)
'''
''''vullen array
'''ReDim Preserve ArrLokatie(UBound(ArrLokatie) + 1)
'''
'''ArrLokatie(UBound(ArrLokatie)) = sLokatie
'''
'''Uitlezen
'''Dim t As Integer
'''For t = LBound(ArrLokatie) + 1 To UBound(ArrLokatie)  ' + 1 anders wordt eerste lege veld afgebeeld
'''    MsgBox ArrLokatie(t)
'''Next t
''''------------------------------------------------------------------------------------------------




'------------------------------------------------------------------------------------------------
'SORTEREN VAN ARRAY-INHOUD MBV BUBBLESORT
'------------------------------------------------------------------------------------------------



'Sub TestSorteren()
'
'    Dim Veld(0 To 500)
'    Veld(0) = "AAA3"
'    Veld(1) = "AA1"
'    Veld(2) = "AA2"
'
'    Dim X%
'
'    Call BubbleSort(Veld)
'
'    For X = 0 To UBound(Veld)
'      If Veld(X) <> "" Then MsgBox Veld(X)
'    Next X
'End Sub
 
 
Private Sub BubbleSort()

''''    TEST: BEKIJKEN INHOUD VAN ARRAY "VELD"
''''    Dim XX As Integer
''''    For XX = 0 To UBound(Veld)
''''        If Veld(XX) <> "" Then MsgBox Veld(XX)
''''    Next XX
''''
    
    
  Dim LB As Integer     'of dim LB&
  Dim UB As Integer
  Dim TEMP$, pos&, x&
 
    LB = LBound(Veld)
    UB = UBound(Veld)
 
    While UB > LB
    
    
      pos = LB
 
      For x = LB To UB - 1
      
      '*** onderstaande regel is nieuw op 30 mei 2003 omdat filteren niet goed ging
      'bij de overgang van positieve en negatieve y-waarden. (want lege velden worden hier
      'blijkbaar tussen gesorteerd
      
      If Veld(x) <> "" And Veld(x + 1) <> "" Then
        
        '19 MEI 2003, DEZE REGEL UITGEZET. ARC'S PLAATSEN BIJ GROTE LEGPLANNEN GAAT NU WEL GOED !!
        'If Val(Veld(X)) < 0 Then MsgBox "Let op, leidingen met negatieve Y-waarde gevonden. Plaats de tekening in het eerste kwadrant (x>0 en Y>0 !).", vbCritical, "Afrondingen kunnen niet worden gemaakt": End
        
                
        
        'If Veld(x) > Veld(x + 1) Then                  'sorteren volgens ascii
        If Val(Veld(x)) > Val(Veld(x + 1)) Then         'sorteren volgens nummer
        
          TEMP = Veld(x + 1)
          Veld(x + 1) = Veld(x)
          Veld(x) = TEMP
          pos = x
        End If
        
        End If
      Next x
      
      'MsgBox Veld(x) & "    " & UBound(Veld)
 
      UB = pos
    
    Wend
    
'''    GESORTEERD ARRAY:
'''    For x = 0 To UBound(Veld)
'''        If Veld(x) <> "" Then MsgBox Veld(x), vbInformation
'''    Next x
    
End Sub


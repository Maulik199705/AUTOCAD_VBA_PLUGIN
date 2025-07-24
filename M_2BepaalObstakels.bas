Attribute VB_Name = "M_2BepaalObstakels"
Public Sub BepaalObstakels(Stoppen As Boolean)

F_Main.ListBox1.Clear

Dim element As Object
Dim RotHoek As Double
Dim Pdraai(0 To 2) As Double
Dim LijnInJuistelaagGevonden As Boolean

Pdraai(0) = 0
Pdraai(1) = 0
Pdraai(2) = 0

'--------------------------------------------------------------------------
'BEPALEN ROTATIEHOEK
'---------------------------------------------------------------------------

For Each element In ThisDrawing.ModelSpace
    If element.EntityName = "AcDbLine" And element.Layer = "Legplan" Then
        RotHoek = element.Angle
        'MsgBox RotHoek
        LijnInJuistelaagGevonden = True
        Exit For
    End If
Next element

If LijnInJuistelaagGevonden = False Then
    MsgBox "Er zijn geen lijnen in laag 'Legplan' gevonden.", vbExclamation, "Obstakels kunnen niet worden bepaald"
    End
End If

'--------------------------------------------------------------------------
'ALLE LIJNEN RECHT PLAATSEN (HOEK 0 GRADEN) + OPSLAAN HANDLE IN LISTBOX
'---------------------------------------------------------------------------

For Each element In ThisDrawing.ModelSpace
    If element.EntityName = "AcDbLine" And element.Layer = "Legplan" Then
        F_Main.ListBox1.AddItem element.handle
        element.Rotate Pdraai, -RotHoek
    End If
    
Next element


'--------------------------------------------------------------------------
'BEPALEN OF ER LIJNEN OP DEZELFDE X-COORDINAAT LIGGEN (DUS NAAST ELKAAR)
'---------------------------------------------------------------------------
Dim lijn1 As Object
Dim lijn2 As Object
Dim lijn3 As Object

Dim StartpuntLijn1 As Variant
Dim StartpuntLijn2 As Variant
Dim StartpuntLijn3 As Variant
            
Dim ListTel As Integer
For ListTel = 0 To F_Main.ListBox1.ListCount - 3
                 
                'MsgBox ListBox1.List(2)
            
                Set lijn1 = ThisDrawing.HandleToObject(F_Main.ListBox1.List(ListTel))
                Set lijn2 = ThisDrawing.HandleToObject(F_Main.ListBox1.List(ListTel + 1))
                Set lijn3 = ThisDrawing.HandleToObject(F_Main.ListBox1.List(ListTel + 2))
                
                StartpuntLijn1 = lijn1.StartPoint
                StartpuntLijn2 = lijn2.StartPoint
                StartpuntLijn3 = lijn3.StartPoint
                
                'nieuw op 14 jan 2003:
                'lijn1.Color = acBlue
                
                ' *** 27 mei 2003 kleur aangepast in geel
                lijn1.Color = acYellow
                
                
                  
                'CONTROLEREN OF LIJNEN NAAST ELKAAR LIGGEN OP ZELFDE Y-COORDINAAT
                '(GEBRUIK VAN FORMAT OM ENIGE AFWIJKING IN Y-RICHTING TOCH GOED OP TE VANGEN):
                
                If Round(StartpuntLijn1(1), 0) = Round(StartpuntLijn2(1), 0) = Round(StartpuntLijn3(1), 0) Then
                  
                    'lijn1.Color = acMagenta
                   'lijn2.Color = acYellow
                    'lijn3.Color = acGreen
                   'ListTel = ListTel + 2
                ElseIf Round(StartpuntLijn1(1), 0) = Round(StartpuntLijn2(1), 0) Then
 
                    '*** ORIGINEEL VOOR 27 MEI 2003
'                    lijn1.Color = acYellow
'                    lijn2.Color = acGreen
'                    lijn3.Color = acRed  'DEZE UITVINKEN !?
                    
                    lijn1.Color = 21
                    lijn2.Color = 171
                    lijn3.Color = acRed  'DEZE UITVINKEN !?
                    
                    ListTel = ListTel + 1
                End If
       
Next ListTel



'--------------------------------------------------------------------------
'TERUGDRAAIEN LIJNEN MET DE ROTATIEHOEK
'---------------------------------------------------------------------------


For Each element In ThisDrawing.ModelSpace
    If element.EntityName = "AcDbLine" And element.Layer = "Legplan" Then
        element.Rotate Pdraai, RotHoek
    End If

Next element

If Stoppen = True Then End


End Sub

' ORGINEEL PRINCIPE

'''''Public Sub BepaalObstakels()
'''''
'''''Dim lijn1 As Object
'''''Dim lijn2 As Object
'''''Dim lijn3 As Object
'''''
'''''Dim StartpuntLijn1 As Variant
'''''Dim StartpuntLijn2 As Variant
'''''Dim StartpuntLijn3 As Variant
'''''
'''''Dim teller As Integer
'''''For teller = 0 To ThisDrawing.ModelSpace.Count - 3          '2 lijnen dan -2
'''''
'''''
'''''
'''''       If ThisDrawing.ModelSpace.Item(teller).Layer = "Legplan" Then Set lijn1 = ThisDrawing.ModelSpace.Item(teller)
'''''       If ThisDrawing.ModelSpace.Item(teller + 1).Layer = "Legplan" Then Set lijn2 = ThisDrawing.ModelSpace.Item(teller + 1)
'''''       If ThisDrawing.ModelSpace.Item(teller + 2).Layer = "Legplan" Then Set lijn3 = ThisDrawing.ModelSpace.Item(teller + 2)
'''''
'''''       StartpuntLijn1 = lijn1.StartPoint
'''''       StartpuntLijn2 = lijn2.StartPoint
'''''       StartpuntLijn3 = lijn3.StartPoint
'''''
'''''       'CONTROLEREN OF LIJNEN NAAST ELKAAR LIGGEN OP ZELFDE Y-COORDINAAT
'''''
'''''       If StartpuntLijn1(1) = StartpuntLijn2(1) = StartpuntLijn3(1) Then
'''''            lijn1.Color = acRed
'''''            lijn2.Color = acBlue
'''''            lijn3.Color = acGreen
'''''            teller = teller + 2
'''''       ElseIf StartpuntLijn1(1) = StartpuntLijn2(1) Then
'''''            lijn1.Color = acBlue
'''''            lijn2.Color = acGreen
'''''            'lijn3.Color = acRed
'''''            teller = teller + 1
'''''       End If
'''''
'''''Next teller
'''''
'''''
'''''End Sub

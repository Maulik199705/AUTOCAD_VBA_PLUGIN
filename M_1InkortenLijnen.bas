Attribute VB_Name = "M_1InkortenLijnen"


Public Sub AlleLijnenInkorten(ItemNrVanEersteElement)

'''GEWIJZIGD OP 24 FEB 2003

    'inkorten alle lijnen in laag "Legplan"

'    Dim Element As Object
'    For Each Element In ThisDrawing.ModelSpace
'        If Element.EntityName = "AcDbLine" Then
'            If Element.Layer = "Legplan" Then
'                Call M_InkortenLijnen.InkortenLijnen(Element, F_Main.ComboBox2)
'            End If
'        End If
'    Next Element

    Dim element As Object
    Dim t As Double
    
    For t = ThisDrawing.ModelSpace.Count - 1 To 0 Step -1
    
        If t > ItemNrVanEersteElement - 1 Then
    
            Set element = ThisDrawing.ModelSpace.Item(t)
            If element.EntityName = "AcDbLine" Then
                If element.Layer = "Legplan" Then
                    Call M_1InkortenLijnen.InkortenLijnen(element, F_Main.ComboBox2)
                End If
            End If
        End If
    Next t
    
End Sub
Sub InkortenLijnen(LijnObj As AcadLine, HoH As Double)

    'Principe:
    'Hoek opvragen van lijn, lijn op 0 graden draaien en
    'begin en eindpunt (x-coordinaten) inkorten met halve hart-op-hart afstand
    'daarna lijn weer goed terugdraaien om het oorspronkelijke draaipunt Pdraai.
    
    If LijnObj.Length < (2 * HoH) Then Exit Sub

    Dim Pbegin As Variant
    Dim Peind As Variant
    Dim HOEK As Double
    Dim Pdraai As Variant
    
    Pdraai = LijnObj.StartPoint
    HOEK = LijnObj.Angle
    'MsgBox Hoek
    
    LijnObj.Rotate Pdraai, -HOEK
    
    Pbegin = LijnObj.StartPoint
    Peind = LijnObj.EndPoint
        
'    If Pbegin(0) < Peind(0) Then
'        Pbegin(0) = Pbegin(0) + (HOH / 2)
'        Peind(0) = Peind(0) - (HOH / 2)
'    End If
'    If Pbegin(0) > Peind(0) Then
'        MsgBox "Anders"
'        Pbegin(0) = Pbegin(0) - (HOH / 2)
'        Peind(0) = Peind(0) + (HOH / 2)
'    End If

    Pbegin(0) = Pbegin(0) + (HoH)
    Peind(0) = Peind(0) - (HoH)
    
     
    LijnObj.StartPoint = Pbegin
    LijnObj.EndPoint = Peind
    
    LijnObj.Rotate Pdraai, HOEK

End Sub

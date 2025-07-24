Attribute VB_Name = "M_5BerekenLengteInLayers"
Public Sub BerekenLengteInColors()

        F_Main.Hide

        Dim RetourObj As Object
        On Error Resume Next
        ThisDrawing.Utility.GetEntity RetourObj, basePnt, "Selecteer van welke kleur en laag de lengte berekend moet worden"
       
        If Err <> 0 Then
            Err.Clear
            'MsgBox "Verkeerd geselecteerd"
            Exit Sub
        End If
        
        RetourObj.Highlight (True)
        
        Dim element As Object
        Dim Lengte As Double
        Dim VerkeerdeObject As Boolean
        
        For Each element In ThisDrawing.ModelSpace
            If element.Color = RetourObj.Color Then
            If element.Layer = RetourObj.Layer Then
            
                    Select Case element.EntityName
                    Case "AcDbLine"
                        Lengte = Lengte + element.Length
                        element.Highlight True
                    Case "AcDbArc"
                        Lengte = Lengte + element.ArcLength
                        element.Highlight True
                    Case "AcDbPolyline"
                        Lengte = Lengte + LengtePolyline2(element)
                        element.Highlight True
                    Case Else
                        VerkeerdeObject = True
                        
                        'nieuw op 26 maart 2003
                        'circle plaatsen om verkeerde element, boundingbox om midden te bepalen
                        Dim minExt As Variant
                        Dim maxExt As Variant
                        element.GetBoundingBox minExt, maxExt
                        Dim Pcenter(0 To 2) As Double
                        Pcenter(0) = (minExt(0) + maxExt(0)) / 2
                        Pcenter(1) = (minExt(1) + maxExt(1)) / 2
                        Dim circleObj As AcadCircle
                        Set circleObj = ThisDrawing.ModelSpace.AddCircle(Pcenter, 60)
                        circleObj.Color = acRed
                        circleObj.Lineweight = acLnWt050
                        ThisDrawing.Preferences.LineWeightDisplay = True
                        
                    End Select
            End If
            End If
        Next element
        
        If VerkeerdeObject = True Then MsgBox "Element in deze laag is geen lijn of arc of polyline !", vbCritical
                
        Lengte = Lengte / 100
        Lengte = Format(Lengte, "0.0")
        
        Dim msg As String
        msg = "De gemeten lengte: " & Lengte & " m" & Chr(10) & Chr(13) & Chr(10) & Chr(13)
        msg = msg & "- reserve lengte: " & F_Main.TextBox1.Text & " m" & Chr(10) & Chr(13)
        msg = msg & "- rollengte: " & Lengte & " + " & F_Main.TextBox1.Text & " = " & F_Main.TextBox1.Text + Lengte & " m" & Chr(10) & Chr(13) & Chr(10) & Chr(13)
        
        MsgBox msg, vbInformation, "Lengte " & "[Laag: " & RetourObj.Layer & " / Kleur: " & RetourObj.Color & "]"
        
        
        '---------------------------------------------------------------------------------
        'JUIST INVULLEN MODEMACRO
        '---------------------------------------------------------------------------------
        
        Dim sModemacro As String
        Dim pos As Integer
        Dim sRechterdeel As String
        Dim sLinkerdeel As String
        
        sModemacro = ThisDrawing.GetVariable("modemacro")
        pos = InStr(1, sModemacro, "   ")
        
        If pos = 0 Then
                pos = InStr(1, sModemacro, "Rollengte=")
                If pos <> 0 Then
                    sModemacro = sModemacro & "   " & "Gemeten lengte= " & Lengte & " m."
                Else
                    sModemacro = "   " & "gemeten lengte= " & Lengte & " m."
                End If
        Else
                sModemacro = Left(sModemacro, pos) & "   " & "Gemeten lengte= " & Lengte & " m."
                sLinkerdeel = sLinkerdeel & "   "
        End If
                
        ThisDrawing.SetVariable "modemacro", sModemacro
   
        
        
        
End Sub











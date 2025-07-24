VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''''Sub fopen()
'''''Dim pe2(0 To 2) As Double
'''''Dim element
'''''Dim insp
'''''
'''''For Each element In ThisDrawing.ModelSpace
'''''      If element.ObjectName = "AcDbBlockReference" Then
'''''         If element.Name = "ba3" Or element.Name = "ba2" Or element.Name = "ba1" _
'''''         Or element.Name = "ba0" Or element.Name = "ba0+" Then
'''''          insp = element.InsertionPoint
'''''
'''''          pe2(0) = insp(0)
'''''          pe2(1) = insp(1)
'''''          pe2(2) = 0
'''''          If pe2(0) <> 0 Or pe2(1) <> 0 Then
'''''          MsgBox "Je kader staat niet goed...!!!", vbCritical
'''''          element.Erase
'''''          End If
'''''          'MsgBox pe2(0) & " - " & pe2(1) & " - " & pe2(2) & " - "
'''''
'''''          End If
'''''      Update
'''''      End If
'''''     Next element
'''''End Sub

'''''Private Sub AcadDocument_BeginPlot(ByVal DrawingName As String)
'''''Dim pe2(0 To 2) As Double
'''''Dim element
'''''Dim insp
'''''
'''''For Each element In ThisDrawing.ModelSpace
'''''      If element.ObjectName = "AcDbBlockReference" Then
'''''         If element.Name = "ba3" Or element.Name = "ba2" Or element.Name = "ba1" _
'''''         Or element.Name = "ba0" Or element.Name = "ba0+" Then
'''''          insp = element.InsertionPoint
'''''
'''''          pe2(0) = insp(0)
'''''          pe2(1) = insp(1)
'''''          pe2(2) = 0
'''''          If pe2(0) <> 0 Or pe2(1) <> 0 Then
'''''          MsgBox "Je kader staat niet goed...!!!", vbCritical
'''''          element.Erase
'''''          End If
'''''          'MsgBox pe2(0) & " - " & pe2(1) & " - " & pe2(2) & " - "
'''''
'''''          End If
'''''      Update
'''''      End If
'''''     Next element
'''''
'''''End Sub

'''''Private Sub AcadDocument_BeginClose()
'''''For Each element In ThisDrawing.ModelSpace
'''''        If element.ObjectName = "AcDbBlockReference" Then
'''''            If UCase(element.Name) = "GROEPTEKSTBLOK" Then
'''''                Set symbool = element
'''''                If symbool.HasAttributes Then
'''''                    attributen = symbool.GetAttributes
'''''                    For I = LBound(attributen) To UBound(attributen)
'''''                    Set attribuut = attributen(I)
'''''                      If attribuut.TagString = "FLEXFIX" And attribuut.textstring = "ja" Then
'''''                      ThisDrawing.SendCommand "ffexport" & vbCr
'''''                      Exit Sub
'''''                      End If
'''''
'''''                    Next I
'''''                End If
'''''            End If
'''''        End If
'''''    Next element
'''''End Sub
Sub papierformaatA4()
Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then a = a + 1
            If UCase(element.Name) = "BA2" Or UCase(element.Name) = "BA3" Or UCase(element.Name) = "BA1" _
            Or UCase(element.Name) = "BA0" Or UCase(element.Name) = "BA0+" Then element.Erase
          End If
Next element

If a > 1 Then

For Each element1 In ThisDrawing.ModelSpace
      If element1.ObjectName = "AcDbBlockReference" Then
       If element1.Name = "Kaderlogo" Or element1.Name = "KaderlogoEngels" Or element1.Name = "logotgh" Then
       Set symbool = element1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" And attribuut.textstring = "" Then
          ' insp = element1.InsertionPoint
           element1.Erase
         End If
        Next I
       End If
     End If
   End If
  Next element1

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If element2.Name = "Kaderlogo" Or element2.Name = "KaderlogoEngels" Or element2.Name = "logotgh" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = "A4"  'FORMAATPAPIER
        Next j

      End If
      End If
      End If
  Next element2
Update
End If

'ThisDrawing.PurgeAll

End Sub
Sub papierformaatA3()
Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then a = a + 1
            If UCase(element.Name) = "BA2" Or UCase(element.Name) = "BA1" Or UCase(element.Name) = "BA4" _
            Or UCase(element.Name) = "BA0" Or UCase(element.Name) = "BA0+" Then element.Erase
          End If
Next element

If a > 1 Then

For Each element1 In ThisDrawing.ModelSpace
      If element1.ObjectName = "AcDbBlockReference" Then
       If element1.Name = "Kaderlogo" Or element1.Name = "KaderlogoEngels" Or element1.Name = "logotgh" Then
       Set symbool = element1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" And attribuut.textstring = "" Then
          ' insp = element1.InsertionPoint
           element1.Erase
         End If
        Next I
       End If
     End If
   End If
  Next element1

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If element2.Name = "Kaderlogo" Or element2.Name = "KaderlogoEngels" Or element2.Name = "logotgh" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = "A3"  'FORMAATPAPIER
        Next j

      End If
      End If
      End If
  Next element2
Update
End If

'ThisDrawing.PurgeAll

End Sub
Sub papierformaatA2()

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then a = a + 1
            If UCase(element.Name) = "BA1" Or UCase(element.Name) = "BA3" Or UCase(element.Name) = "BA4" _
            Or UCase(element.Name) = "BA0" Or UCase(element.Name) = "BA0+" Then element.Erase
          End If
Next element

If a > 1 Then

For Each element1 In ThisDrawing.ModelSpace
      If element1.ObjectName = "AcDbBlockReference" Then
       If element1.Name = "Kaderlogo" Or element1.Name = "KaderlogoEngels" Or element1.Name = "logotgh" Then
       Set symbool = element1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" And attribuut.textstring = "" Then
          ' insp = element1.InsertionPoint
           element1.Erase
         End If
        Next I
       End If
     End If
   End If
  Next element1

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If element2.Name = "Kaderlogo" Or element2.Name = "KaderlogoEngels" Or element2.Name = "logotgh" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = "A2"  'FORMAATPAPIER
        Next j

      End If
      End If
      End If
  Next element2
Update
End If

'ThisDrawing.PurgeAll

End Sub
Sub papierformaatA1()

Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then a = a + 1
            If UCase(element.Name) = "BA2" Or UCase(element.Name) = "BA3" Or UCase(element.Name) = "BA4" _
            Or UCase(element.Name) = "BA0" Or UCase(element.Name) = "BA0+" Then element.Erase
          End If
Next element


If a > 1 Then

For Each element1 In ThisDrawing.ModelSpace
      If element1.ObjectName = "AcDbBlockReference" Then
       If element1.Name = "Kaderlogo" Or element1.Name = "KaderlogoEngels" Or element1.Name = "logotgh" Then
       Set symbool = element1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" And attribuut.textstring = "" Then
          ' insp = element1.InsertionPoint
           element1.Erase
         End If
        Next I
       End If
     End If
   End If
  Next element1

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If element2.Name = "Kaderlogo" Or element2.Name = "KaderlogoEngels" Or element2.Name = "logotgh" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = "A1"  'FORMAATPAPIER
        Next j

      End If
      End If
      End If
  Next element2
Update
End If

'ThisDrawing.PurgeAll




' -------!!!!!!!!!!!!!!!!!!!! DIT STUK EFFE BEWAREN VOOR ALS HET BOVENSTAANDE FOUT GAAT !!!!!!!!!!!!!!!!!!!!-----
'For Each element2 In ThisDrawing.ModelSpace
'      If element2.ObjectName = "AcDbBlockReference" Then
'      If UCase(element2.Name) = "KADERLOGO" Or UCase(element2.Name) = "LOGOTGH" Then
'      Set symbool = element2
'        If symbool.HasAttributes Then
'        attributen = symbool.GetAttributes
'        For I = LBound(attributen) To UBound(attributen)
'        Set attribuut = attributen(I)
'        If attribuut.TagString = "FORMAAT" Then attribuut.TextString = "A1"  'FORMAATPAPIER
'        Next I
'
'      End If
'      End If
'      End If
'  Next element2
'Update
End Sub
Sub papierformaatA0()
Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then a = a + 1
            If UCase(element.Name) = "BA2" Or UCase(element.Name) = "BA3" Or UCase(element.Name) = "BA4" _
            Or UCase(element.Name) = "BA1" Or UCase(element.Name) = "BA0+" Then element.Erase
          End If
Next element

If a > 1 Then

For Each element1 In ThisDrawing.ModelSpace
      If element1.ObjectName = "AcDbBlockReference" Then
       If element1.Name = "Kaderlogo" Or element1.Name = "KaderlogoEngels" Or element1.Name = "logotgh" Then
       Set symbool = element1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" And attribuut.textstring = "" Then
          ' insp = element1.InsertionPoint
           element1.Erase
         End If
        Next I
       End If
     End If
   End If
  Next element1

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If element2.Name = "Kaderlogo" Or element2.Name = "KaderlogoEngels" Or element2.Name = "logotgh" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = "A0"  'FORMAATPAPIER
        Next j

      End If
      End If
      End If
  Next element2
Update
End If


'ThisDrawing.PurgeAll

End Sub
Sub papierformaatA0plus()
Dim element As Object
For Each element In ThisDrawing.ModelSpace
      If element.ObjectName = "AcDbBlockReference" Then
         If element.Name = "Kaderlogo" Or element.Name = "KaderlogoEngels" Or element.Name = "logotgh" Then a = a + 1
            If UCase(element.Name) = "BA2" Or UCase(element.Name) = "BA3" Or UCase(element.Name) = "BA4" _
            Or UCase(element.Name) = "BA0" Or UCase(element.Name) = "BA1" Then element.Erase
          End If
Next element

If a > 1 Then

For Each element1 In ThisDrawing.ModelSpace
      If element1.ObjectName = "AcDbBlockReference" Then
       If element1.Name = "Kaderlogo" Or element1.Name = "KaderlogoEngels" Or element1.Name = "logotgh" Then
       Set symbool = element1
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For I = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(I)
        If attribuut.TagString = "OPDRACHTGEVER" And attribuut.textstring = "" Then
          ' insp = element1.InsertionPoint
           element1.Erase
         End If
        Next I
       End If
     End If
   End If
  Next element1

For Each element2 In ThisDrawing.ModelSpace
      If element2.ObjectName = "AcDbBlockReference" Then
      If element2.Name = "Kaderlogo" Or element2.Name = "KaderlogoEngels" Or element2.Name = "logotgh" Then
      Set symbool = element2
        If symbool.HasAttributes Then
        attributen = symbool.GetAttributes
        For j = LBound(attributen) To UBound(attributen)
        Set attribuut = attributen(j)
        If attribuut.TagString = "FORMAAT" Then attribuut.textstring = "A0+"  'FORMAATPAPIER
        Next j

      End If
      End If
      End If
  Next element2
Update
End If

'ThisDrawing.PurgeAll


End Sub
Sub mflex(groepsnummer)

   On Error Resume Next
   Dim RetObj As Object
   Dim Pbase As Variant
   Dim element As Object
   Dim Lengte As Double
   Dim Lengte1 As Double
   Dim lengte20 As Double
   Dim Lengte3 As Double
   
    Lengte = 0
    Lengte1 = 0
    lengte20 = 0
    Lengte3 = 0

RetObj = groepsnummer
    'ThisDrawing.Utility.GetEntity RetObj, Pbase, "Selecteer het MAT-PATROON."
'    If Err <> 0 Then
'        Err.Clear
'        message = "Mis"
'        ThisDrawing.Utility.Prompt message
'        ThisDrawing.SendCommand Chr(27)
'
'        Exit Sub
'        'End
'    End If


'MsgBox RetObj.Layer
   ' aktievelaag = retobj.layer   'ThisDrawing.Utility.GetEntity RetObj, Pbase, "Selecteer een lijn of cirkelboog"
   ' If Err Then
   '     Err.Clear
   '     'OPNIEUW OF STOPPEN
   '     GoTo Opnieuw
   '     'Exit sub
   checkpoint1 = Right(RetObj.Layer, 4)
   If checkpoint1 = "voer" Then
      checkpoint2 = Len(RetObj.Layer)
      checkpoint3 = checkpoint2 - 16
      checkpoint4 = Left(RetObj.Layer, checkpoint3)
      Set newLayer = ThisDrawing.Layers.Add(checkpoint4)
      ThisDrawing.ActiveLayer = newLayer
      zetlaagterug = RetObj.Layer
      'MsgBox checkpoint4
   End If
      If checkpoint1 = "tour" Then
      checkpoint2 = Len(RetObj.Layer)
      checkpoint3 = checkpoint2 - 15
      checkpoint4 = Left(RetObj.Layer, checkpoint3)
      Set newLayer = ThisDrawing.Layers.Add(checkpoint4)
      ThisDrawing.ActiveLayer = newLayer
      zetlaagterug = RetObj.Layer
      'MsgBox checkpoint4
   End If
   zetlaagterug = RetObj.Layer
   
   'retobj.layer = ThisDrawing.ActiveLayer.Name
     'If checkpoint1 <> "voer" And checkpoint1 <> "tour" Then
        For Each element In ThisDrawing.ModelSpace
            If element.Layer = RetObj.Layer Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte = Lengte + element.length
                If element.EntityName = "AcDbArc" Then Lengte = Lengte + element.ArcLength
                If element.EntityName = "AcDbPolyline" Then Lengte = Lengte + element.length
                
            End If
        Next element
     'End If
    
    Lengte = Lengte / 100
    Lengte = Round(Lengte, 1)
    
    If checkpoint1 = "xfix" Then
        samen100 = "_aanvoer"
        Else
        samen100 = "_Flexfix_aanvoer"
    End If
    
    
    checkpointh = Right(RetObj.Layer, 1)
    If checkpointh = "h" Then
         checkpoint20 = Len(RetObj.Layer)
         checkpoint20 = checkpoint20 - 1
         checkpointh2 = Left(RetObj.Layer, checkpoint20)
            If checkpoint1 = "fixh" Then
            samen100 = "_aanvoerh"
            Else
            samen100 = "_Flexfix_aanvoerh"
            End If
         samenvoeging1 = checkpointh2 & samen100
    End If
    
    If checkpointh <> "h" Then samenvoeging1 = RetObj.Layer & samen100
    
    
    
    
    'If checkpoint1 = "voer" Then
         For Each element In ThisDrawing.ModelSpace
            If element.Layer = samenvoeging1 Then '(RetObj.Layer & samen100) Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then Lengte1 = Lengte1 + element.length
                If element.EntityName = "AcDbArc" Then Lengte1 = Lengte1 + element.ArcLength
                If element.EntityName = "AcDbPolyline" Then Lengte1 = Lengte1 + element.length
                
            End If
     Next element
     'End If
    Lengte1 = Lengte1 / 100
    Lengte1 = Round(Lengte1, 1)
    
    If checkpoint1 = "xfix" Then
        samen100 = "_retour"
        Else
        samen100 = "_Flexfix_retour"
    End If
    
    
     If checkpointh = "h" Then
            If checkpoint1 = "fixh" Then
            samen100 = "_retourh"
            Else
            samen100 = "_Flexfix_retourh"
            End If
         samenvoeging2 = checkpointh2 & samen100
     End If
    
     If checkpointh <> "h" Then samenvoeging2 = RetObj.Layer & samen100
     
    
    'If checkpoint1 = "tour" Then
         For Each element In ThisDrawing.ModelSpace
            If element.Layer = samenvoeging2 Then '(RetObj.Layer & samen100) Then
                'BEREKENEN TOTALE LENGTE
                If element.EntityName = "AcDbLine" Then lengte20 = lengte20 + element.length
                If element.EntityName = "AcDbArc" Then lengte20 = lengte20 + element.ArcLength
                If element.EntityName = "AcDbPolyline" Then lengte20 = lengte20 + element.length
                
            End If
     Next element
     'End If
    lengte20 = lengte20 / 100
    lengte20 = Round(lengte20, 1)
    Lengte3 = Lengte + Lengte1 + lengte20
    
    zzz1 = (Chr(13) & Chr(10)) & "De lengte van het mat-patroon: " & Lengte & " meter(s)" & " [" & zetlaagterug & "]" & (Chr(13) & Chr(10)) _
    & "Aanvoerlengte van het mat-patroon: " & Lengte1 & " meter(s)" & (Chr(13) & Chr(10)) _
    & "Retourlengte van het mat-patroon: " & lengte20 & " meter(s)" & (Chr(13) & Chr(10)) _
    & "Totale lengte: " & Lengte3 & " meter(s)"
    
    'MsgBox Lengte & " meter(s)" & " | De lengte van het mat-patroon." & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) _
    '& lengte1 & " meter(s)" & " | Aanvoerlengte van het mat-patroon. " & (Chr(13) & Chr(10)) _
    '& lengte20 & " meter(s)" & " | Retourlengte van het mat-patroon. " & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10)) _
    '& "Totale lengte = " & lengte3, vbInformation, "Laag: " & aktievelaag   'RetObj.Layer
    Lengte = 0
    Lengte1 = 0
    lengte20 = 0
    Lengte3 = 0

    ThisDrawing.Utility.Prompt zzz1
    ThisDrawing.SendCommand Chr(27)

      Set newLayer = ThisDrawing.Layers.Add(zetlaagterug)
      ThisDrawing.ActiveLayer = newLayer

  End Sub



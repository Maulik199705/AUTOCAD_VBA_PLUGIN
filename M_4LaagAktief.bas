Attribute VB_Name = "M_4LaagAktief"
Public Sub LaagAktief()

    '*** 19 mei 2003 routine geheel aangepast: nu > 99 groep_lagen mogelijk

    ' ------------------------------------------------------------------------------
    ' DEZE ROUTINE MAAK EEN NIEUWE LAAG AAN MET EEN JUISTE GROEPSNUMMER
    ' ------------------------------------------------------------------------------
    

    ' ------------------------------------------------------------------------------
    ' BEPALEN LAATSTE GROEP-NUMMER (UIT LAYERS)
    ' ------------------------------------------------------------------------------
    
    SchrijfLogFile ("BEPALEN LAATSTE GROEP-NUMMER (UIT LAYER)EN AANMAKEN NIEUWE LAYER")

    Dim LaagObj As AcadLayer
    Dim sLaagNaam As String
    Dim sLaagNummer As String
    Dim iLaagNummer As String
    Dim iHoogsteNummer As Integer
    'Dim sNwLaagnaam As String
    'Dim bGroepLaagBestaat As Boolean
    Dim sHoogsteGroepNaam As String

    For Each LaagObj In ThisDrawing.Layers
        sLaagNaam = UCase(LaagObj.Name)
        If Left$(sLaagNaam, 6) = "GROEP_" Then
        
            bGroepLaagBestaat = True

            sLaagNummer = Right(sLaagNaam, Len(sLaagNaam) - 6)
            If Val(sLaagNummer) <> sLaagNummer Then
                MsgBox "Laag " & sLaagNummer & " bevat een verkeerd nummer.", vbCritical, "Let op"
                End
            Else
                iLaagNummer = sLaagNummer
                If iLaagNummer > iHoogsteNummer Then iHoogsteNummer = iLaagNummer: sHoogsteGroepNaam = LaagObj.Name
            End If

        End If
    Next LaagObj
        
    
    
    ' ------------------------------------------------------------------------------
    ' ALS ER NOG GEEN GROEPSLAAG IS, DAN DEZE AANMAKEN EN EXIT SUB
    ' ------------------------------------------------------------------------------
      
    
    If sHoogsteGroepNaam = "" Then
        Call AanmakenLaag("groep_01", acCyan, True)
        Exit Sub
    End If
    
        
    
    ' ------------------------------------------------------------------------------
    ' BEPALEN OF ER ELEMENTEN IN DE HOOGSTE (LAATSTE)GROEPSLAAG STAAN
    ' ZO NIET, DAN GEEN NIEUWE LAAG AANMAKEN MAAR LAAG WEL AKTIEF MAKEN !
    ' ------------------------------------------------------------------------------
    
    'MsgBox sHoogsteGroepNaam
    
    Dim element As AcadEntity
    Dim bElementenGevonden As Boolean
    Dim LaatsteGroepLaagnaam As String
    
    For Each element In ThisDrawing.ModelSpace
        If element.Layer = sHoogsteGroepNaam Then
            bElementenGevonden = True
        End If
    Next element
        
    If bElementenGevonden <> True Then
        If ThisDrawing.Layers.Item(sHoogsteGroepNaam).Freeze = True Then ThisDrawing.Layers.Item(sHoogsteGroepNaam).Freeze = False
        ThisDrawing.Layers.Item(sHoogsteGroepNaam).LayerOn = True
        ThisDrawing.ActiveLayer = ThisDrawing.Layers.Item(sHoogsteGroepNaam)
        Exit Sub
    End If
    
    
    
    ' ------------------------------------------------------------------------------
    ' NIEUWE GROEPS-LAAG AANMAKEN (MET HOGER NUMMER)
    ' ------------------------------------------------------------------------------
    
    'MsgBox sHoogsteGroepNaam
    
    Dim pos As Integer
    Dim sGroepsnummer As String
    Dim iGroepsnummer As Integer
    
    sGroepsnummer = Right$(sHoogsteGroepNaam, Len(sHoogsteGroepNaam) - 6)
    sGroepsnummer = Trim(sGroepsnummer)
    
    If sGroepsnummer <> Val(sGroepsnummer) Then MsgBox "Verkeerd groepnummer in laag " & sHoogsteGroepNaam, vbCriticalm, "Let op": End
    
    iGroepsnummer = Val(sGroepsnummer)
    iGroepsnummer = iGroepsnummer + 1
    
    
    Dim sNieuweLaagNaam As String
    
    If iGroepsnummer < 10 Then
        sNieuweLaagNaam = "groep_0" & iGroepsnummer
    Else
        sNieuweLaagNaam = "groep_" & iGroepsnummer
    End If
    
    
    'MsgBox sNieuweLaagNaam
    
    ' ------------------------------------------------------------------------------
    ' LAAG KLEUR BEPALEN
    ' ------------------------------------------------------------------------------
    
    Dim KleurNo As Integer
    
    On Error Resume Next
    KleurNo = ThisDrawing.Layers.Item(sHoogsteGroepNaam).Color
    If Err Then Err.Clear
    'ER BESTAAT NOG GEEN LAAG GROEP-....

    If KleurNo = acGreen Then
        KleurNo = acCyan
    Else
        KleurNo = acGreen
    End If

    
    ' ------------------------------------------------------------------------------
    ' LAAG AANMAKEN EN AKTIEF ZETTEN
    ' ------------------------------------------------------------------------------
    
    Call AanmakenLaag(sNieuweLaagNaam, KleurNo, True)
   


    

  

End Sub

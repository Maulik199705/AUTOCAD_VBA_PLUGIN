Attribute VB_Name = "M_0CfgFiles"
    Dim sCfgFile As String
    
''''
''''Sub ControleCfgFile()
''''
'''''------------------------------------------------------
'''''CONTROLEREN OF CFG-FILE AANWEZIG IS.
'''''------------------------------------------------------
''''
''''    sCfgFile = "C:\TEMP\Rollengten.cfg"
''''
''''    'CONTROLE OF FILE BESTAAT:
''''    Dim ControleDirectory As String
''''    Dim DirBesaat As String
''''    DirBesaat = Dir(ControleDirectory)
''''    If DirBesaat = "" Then
''''        MsgBox "Configuratie bestand '" & sCfgFile & "' niet gevonden", vbCritical, "Leidingleg-programma"
''''        End
''''    End If
''''End Sub

Sub LeesCfgFile()
    
'------------------------------------------------------
'LEZEN CONFIGURATIE-FILE MET ROLLENGTEN.
'------------------------------------------------------
    
''''    Call ControleCfgFile

    
    
    sCfgFile = FindZoekpad("Rollengten.cfg")
    
    'LEZEN CONFIGURATIE-FILE:
    Dim sLine As String
    Dim bSectieGevonden As Boolean
    
    Open sCfgFile For Input As #1
    Do While Not EOF(1)
            Line Input #1, sLine
            sLine = Trim(sLine)
            If Left$(sLine, 1) <> "'" And sLine <> "" Then
                 
                'UITLEZEN SECTIE-NAAM [HOH] EN VULLEN COMBOBOX
                If UCase(sLine) = "[HOH]" Then bSectieGevonden = True
                If Left$(sLine, 1) = "[" And UCase(sLine) <> "[HOH]" Then bSectieGevonden = False
                
                If bSectieGevonden = True And Left$(sLine, 1) <> "[" Then F_Main.ComboBox2.AddItem sLine
            
                'UITLEZEN SECTIE-NAMEN MET DE ROLTYPEN, NIEUWE SECTIES KUNNEN WORDEN TOEGEVOEGD
                If Left$(sLine, 1) = "[" And Right$(sLine, 1) = "]" Then
                    If UCase(sLine) <> "[HOH]" Then
                            sLine = Right$(sLine, Len(sLine) - 1)
                            sLine = Left$(sLine, Len(sLine) - 1)
                            F_Main.ComboBox3.AddItem sLine
                    End If
                End If
               
                
            End If
    Loop
    Close #1

End Sub


Sub ChangeRollengten()

    
    F_InvoerRollengte.ComboBox1.Clear
    
    'Call ControleCfgFile
    
    sCfgFile = FindZoekpad("Rollengten.cfg")
    
    'LEZEN CONFIGURATIE-FILE:
    Dim sLine As String
    Open sCfgFile For Input As #1
    Dim bSectieGevonden As Boolean
    
    Do While Not EOF(1)
            Line Input #1, sLine
            sLine = Trim(sLine)
            If Left$(sLine, 1) <> "'" And sLine <> "" Then
            
                'CONTROLE OF JUISTE SECTIE GELEZEN WORDT
                If Left$(sLine, 1) = "[" Then
                    If UCase(sLine) = "[" & UCase(F_Main.ComboBox3.Text) & "]" Then
                        bSectieGevonden = True
                    Else
                        bSectieGevonden = False
                    End If
                End If
                
                'INDIEN JUISTE SECTIE,DAN COMBO ROLLENGTEN INVULLEN
                If bSectieGevonden = True And Left$(sLine, 1) <> "[" Then F_InvoerRollengte.ComboBox1.AddItem sLine
                'F_Main.ListBox5.AddItem sLine
                
            End If
    Loop
    Close #1
End Sub



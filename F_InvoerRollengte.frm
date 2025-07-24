VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_InvoerRollengte 
   Caption         =   "Invoer rollengte"
   ClientHeight    =   1755
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3048
   OleObjectBlob   =   "F_InvoerRollengte.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_InvoerRollengte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    ComboBox1.Text = GetSetting("Leidinglegprogramma", "Startup", "InvoerRolLengte", "")
    ComboBox1.SelStart = 0
    ComboBox1.SelLength = Len(ComboBox1.Text)
End Sub

Private Sub CommandButton1_Click()
    'EINDE PROGRAMMA
    
    ' SaveSetting "Leidinglegprogramma", "Startup", "InvoerRolLengte", ComboBox1.Text
    ' F_InvoerRollengte.Hide
    Unload Me
End Sub


Private Sub ComboBox1_Change()

    If ComboBox1.Text = "" Then
        CommandButton2.Enabled = False
    Else
        CommandButton2.Enabled = True
    End If
    
    Dim a As String
    Dim b As String
    a = ComboBox1.Text
    b = Val(ComboBox1.Text)
    
    If a <> b Then
        CommandButton2.Enabled = False
        MsgBox "Geen juiste waarde ingevuld.", vbCritical, "Let op"
    End If
    
    SaveSetting "Leidinglegprogramma", "Startup", "InvoerRolLengte", ComboBox1.Text
    
End Sub

Private Sub ComboBox1_KeyDown(ByVal Keycode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'druk op enter of spatie

    If Keycode = 13 Or Keycode = 32 Then
        If ComboBox1.Text <> "" Then
            F_InvoerRollengte.Hide
            Call StartLengteProgramma(ComboBox1.Text)
        Else
            MsgBox "Geen Rollengte ingevuld"
        End If
    End If
End Sub


Private Sub CommandButton2_Click()
    
    Call StartLengteProgramma(ComboBox1.Text)
    SaveSetting "Leidinglegprogramma", "Startup", "InvoerRolLengte", ComboBox1.Text
End Sub

Sub StartLengteProgramma(Lengte As Double)
    Unload F_InvoerRollengte
    
    'MsgBox "Opgegeven rollengte: " & Lengte & " meter"
    
    '----------------------------------------------------------------------------------------
    'START BEPALEN ROLLENGTE
    '----------------------------------------------------------------------------------------
    
    ' invoer in meters, lijnen getekend in cm
    
    Lengte = Lengte * 100
    Call M_4BerekenKopStaartNW.BerekenKopStaart1(Lengte)

End Sub





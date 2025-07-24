VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_LengteMonitor 
   Caption         =   "Lengte monitor WTH"
   ClientHeight    =   1755
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3840
   OleObjectBlob   =   "F_LengteMonitor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_LengteMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'showmodal is in properties op TRUE gezet !

Dim bMetenOverslaan As Boolean



Private Sub CommandButton2_Click()
    '---------------------------------------------------
    'MEET LENGTE ('zie ook knop op: F_main)
    '---------------------------------------------------
    
    Call M_0SaveGetSettings.SaveSettings
    Call M_5BerekenLengteInLayers.BerekenLengteInColors
End Sub


Private Sub CommandButton3_Click()
    '---------------------------------------------------
    'STRETCH LEIDING ('zie ook knop op: F_main)
    '---------------------------------------------------
    Call M_7WijzigenLengte.WijzigenLengte
    Call MeetLengte
End Sub




Private Sub UserForm_Initialize()
    Label5.Caption = F_Main.TextBox1.Text
    bMetenOverslaan = False
    Call MeetLengte
End Sub

Private Sub Image1_Click()
    Call Afbeelding
End Sub

Private Sub Image2_Click()
    Call Afbeelding
    
End Sub
Sub Afbeelding()
        
    Image1.Visible = Not (Image1.Visible)
    Image2.Visible = Not (Image2.Visible)
    
    If Image1.Visible = True Then
        TextBox1.BackColor = vbGreen
        Call MeetLengte
    Else
        TextBox1.BackColor = vbRed
    End If
    
End Sub

Public Sub MeetLengte()
   
    'NIET METEN INDIEN AFBEELDING NIET-METEN ZICHTBAAR IS
    If Image1.Visible = False Then Exit Sub
    If F_LengteMonitor.Visible = False Then Exit Sub
    'MsgBox "meet lengte"

    Dim sAktieveLaag As String
    sAktieveLaag = ThisDrawing.ActiveLayer.Name
    
    Label1.Caption = "Lengte laag: " & sAktieveLaag
    TextBox1.BackColor = vbGreen
    Label2.Visible = False
    Label3.Visible = False
    
     
     Dim ssetObj As Object
    'Set ssetObj = ThisDrawing.SelectionSets.Add("SSET")
    
    Set ssetObj = CreateSelectionSet
  
    Dim gpCode(0) As Integer
    Dim dataValue(0) As Variant
    
    'groepscode 8 > laagnaam
    'groepscode 0 > entityname bijv Circle
    
    gpCode(0) = 8
    dataValue(0) = sAktieveLaag
    
    Dim groupCode As Variant, dataCode As Variant
    groupCode = gpCode
    dataCode = dataValue
    
    ssetObj.Select acSelectionSetAll, , , groupCode, dataCode
    
    Dim element As Object
    Dim dLengte As Double
    Dim dReserveLengte As Double
       
    
    
    Dim bFout As Boolean
    
    For Each element In ssetObj
        Select Case element.EntityName
        Case "AcDbLine"
            dLengte = dLengte + element.Length
        Case "AcDbArc"
            dLengte = dLengte + element.ArcLength
        Case "AcDbPolyline"
            TextBox1.BackColor = vbRed
            element.Highlight True
            
            If bFout <> True Then
                Label2.Visible = True
                Label3.Visible = True
            End If
            
        Case Else
        
        End Select
        
    Next element
    
        
    dLengte = dLengte / 100
    dLengte = Round(dLengte, 1)
    
    dReserveLengte = Label5.Caption
    
    
    TextBox1.Text = dLengte
    Label6.Caption = Round(dReserveLengte + dLengte, 1)
    ' ThisDrawing.SetVariable "MODEMACRO", Str(dLengte) & " m"
    
    ssetObj.Clear
    ssetObj.Delete
    

End Sub

Public Function CreateSelectionSet(Optional ssName As String = "ss") As AcadSelectionSet

    Dim ss As AcadSelectionSet
    
    On Error Resume Next
    Set ss = ThisDrawing.SelectionSets(ssName)
    If Err Then Set ss = ThisDrawing.SelectionSets.Add(ssName)
    ss.Clear
    Set CreateSelectionSet = ss

End Function


Private Sub UserForm_Terminate()
   bMetenOverslaan = True
End Sub


Private Sub CommandButton1_Click()

'----------------------------------------------------------------------
'GESELECTEERDE LAAG AKTIEF MAKEN
'----------------------------------------------------------------------

   
    Dim ReturnObj1 As AcadEntity
    Dim basePnt1 As Variant
    
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity ReturnObj1, basePnt1, "Selecteer een object."
    If Err <> 0 Then
        Err.Clear
        Exit Sub
    End If
                    
    If ThisDrawing.Layers.Item(ReturnObj1.Layer).Freeze = True Then ThisDrawing.Layers.Item(ReturnObj1.Layer).Freeze = False
    
    
    ThisDrawing.ActiveLayer = ThisDrawing.Layers.Item(ReturnObj1.Layer)
    
    Label1.Caption = "Lengte laag: " & ThisDrawing.ActiveLayer.Name
    
    'meetfunctie aanzetten
    Image1.Visible = True
    Image2.Visible = False
    
    Call MeetLengte
    
End Sub















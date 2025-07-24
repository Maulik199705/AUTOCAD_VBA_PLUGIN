VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub AcadDocument_EndCommand(ByVal CommandName As String)
    'If UserForm1.Visible = True Then
        If F_LengteMonitor.Image1.Visible = True Then Call F_LengteMonitor.MeetLengte
    'End If
End Sub

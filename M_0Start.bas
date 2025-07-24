Attribute VB_Name = "M_0Start"


Public Sub Start()
    F_LengteMonitor.Image1.Visible = False
    F_LengteMonitor.Image2.Visible = True
    F_Main.show
End Sub

Public Sub Start1()
    Call M_6PlaatsenSlingers.OpvragenHOH
End Sub

Public Sub Start2()
    Call F_Main.OpschuivenLegpatroon
End Sub

Public Sub Start3()
    Call M_0SaveGetSettings.SaveSettings
    Call M_5BerekenLengteInLayers.BerekenLengteInColors
End Sub



 




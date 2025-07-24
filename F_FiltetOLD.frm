VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_FiltetOLD 
   Caption         =   "UserForm1"
   ClientHeight    =   1485
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2628
   OleObjectBlob   =   "F_FiltetOLD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_FiltetOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandButton3_Click()
    Dim LijnObj1 As AcadLine
    Dim LijnObj2 As AcadLine
    
    Dim Psnij As Variant
    
    Set LijnObj1 = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 1)
    Set LijnObj2 = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.Count - 2)
    
    Psnij = LijnObj1.IntersectWith(LijnObj2, acExtendNone)
    MsgBox Psnij(0) & Chr(10) & Chr(13) & Psnij(1)

    Dim HoekVerchil As Double
    MsgBox "Hoek1=" & ThisDrawing.Utility.AngleToString(LijnObj1.Angle, acDegrees, 6) '= hoek in graden
    MsgBox "Hoek1=" & ThisDrawing.Utility.AngleToString(LijnObj2.Angle, acDegrees, 6) '= hoek in graden
    HoekVerchil = (LijnObj2.Angle - LijnObj1.Angle) / 2
    
    'MsgBox "Hoek=" & Hoek
    MsgBox "Hoekverschil" & ThisDrawing.Utility.AngleToString(HoekVerchil, acDegrees, 6) '= hoek in graden
    
    'If HoekVerchil < 0 Then HoekVerchil = HoekVerchil + 3.1415
    
'    Dim HoekInGraden As Double
'    HoekInGraden = ThisDrawing.Utility.AngleToString(6.283, acDegrees, 6)  '= hoek in graden
'    MsgBox HoekInGraden
    
    Dim Aanliggende As Double
    Dim SchuineZijde As Double
    Dim Radius As Double
    
    Radius = 10
    Aanliggende = Radius / Tan(HoekVerchil)
    MsgBox "Aanliggende = " & Aanliggende
    
    SchuineZijde = Radius / Sin(HoekVerchil)
    MsgBox "SchuineZijde = " & SchuineZijde
    
    Dim CenterPunt As Variant
    Dim HOEK As Variant
    HOEK = (LijnObj2.Angle + LijnObj1.Angle) / 2
    MsgBox "Hoek=" & ThisDrawing.Utility.AngleToString(HOEK, acDegrees, 6) '= hoek in graden
    CenterPunt = ThisDrawing.Utility.PolarPoint(Psnij, HOEK, SchuineZijde)
    MsgBox "x=" & CenterPunt(0) & Chr(10) & Chr(13) & "y=" & CenterPunt(1)
    
    
    Dim CirkelObj As AcadCircle
    Set CirkelObj = ThisDrawing.ModelSpace.AddCircle(CenterPunt, Radius)
    CirkelObj.Update
    
    
    
End Sub





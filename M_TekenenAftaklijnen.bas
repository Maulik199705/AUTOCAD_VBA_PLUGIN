Attribute VB_Name = "M_TekenenAftaklijnen"
Const NegentigGraden = 1.5707963267949

Sub StartTekenenAftaklijnen()
    'DEZE SUB ALLEEN NODIG VOOR TESTDOELEINDEN VAN DE ONDERSTAANDE SUBROUTINER.
    Dim HoH As Double
    Dim Pbegin(0 To 2) As Double
    
    HoH = 10
    Pbegin(0) = 0
    Pbegin(1) = 0
    
    HOEK = 1
    
    Call TekenenAftaklijnen(HoH, Pbegin, HOEK)
End Sub

Sub TekenenAftaklijnen(HoH, Pbegin, HOEK)
    
   
    
    Dim LijnObj As AcadLine
    Dim Pstart(0 To 2) As Double
    Dim Peind(0 To 2) As Double
    
    Dim ArcObj As AcadArc
    Dim Pcenter(0 To 2) As Double
    Dim Radius As Double
    Dim Starthoek As Double
    Dim EindHoek As Double
    
    Dim LaagNaam As String
    LaagNaam = "Legplan"
    
    '-----------------------------------------------------------
    
'    'Horizontale lijnstuk
'    Pstart(0) = Pbegin(0) + 2 * HOH
'    Pstart(1) = Pbegin(1)
'
'    Peind(0) = Pbegin(0) + 50
'    Peind(1) = Pbegin(1)
'
'    Set LijnObj = ThisDrawing.ModelSpace.AddLine(Pstart, Peind)
'    LijnObj.Rotate Pbegin, Hoek
'    LijnObj.Layer = "Legplan"
    
    
    '-----------------------------------------------------------
    'Eerste Arc
    
    
    Pcenter(0) = Pbegin(0) + 2 * HoH
    Pcenter(1) = Pbegin(1) + (HoH / 2)
    Radius = HoH / 2
    Starthoek = 2 * NegentigGraden
    EindHoek = -NegentigGraden
    
    Set ArcObj = ThisDrawing.ModelSpace.AddArc(Pcenter, Radius, Starthoek, EindHoek)
    ArcObj.Rotate Pbegin, HOEK
    ArcObj.Layer = "Legplan"
   
    
    '-----------------------------------------------------------
    
    'Verticale lijnstuk
    Pstart(0) = Pbegin(0) + (1.5 * HoH)
    Pstart(1) = Pbegin(1) + (HoH / 2)
    
    Peind(0) = Pstart(0)
    Peind(1) = Pbegin(1) + (1.5 * HoH)
    Set LijnObj = ThisDrawing.ModelSpace.AddLine(Pstart, Peind)
    LijnObj.Rotate Pbegin, HOEK
    LijnObj.Layer = "Legplan"
    
    '------------------------------------------------------------------
    
    'Tweede Arc
    
    Pcenter(0) = Pbegin(0) + (HoH)
    Pcenter(1) = Pbegin(1) + (HoH * 1.5)
    Radius = HoH / 2
    Starthoek = 0
    EindHoek = NegentigGraden
    
    Set ArcObj = ThisDrawing.ModelSpace.AddArc(Pcenter, Radius, Starthoek, EindHoek)
    ArcObj.Rotate Pbegin, HOEK
    ArcObj.Layer = "Legplan"
  
    
    
    '------------------------------------------------------------------
    
    'Derde Arc
    
    Pcenter(0) = Pbegin(0) + (3 * HoH)
    Pcenter(1) = Pbegin(1) + (1.5 * HoH)
    Radius = HoH / 2
    Starthoek = NegentigGraden
    EindHoek = -NegentigGraden
    
    Set ArcObj = ThisDrawing.ModelSpace.AddArc(Pcenter, Radius, Starthoek, EindHoek)
    ArcObj.Rotate Pbegin, HOEK
    ArcObj.Layer = "Legplan"
    
    '-----------------------------------------------------------
    
    End
    

End Sub

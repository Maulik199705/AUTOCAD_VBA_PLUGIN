VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub plotten()
  
  Call frmplotten.show(vbModal)

End Sub

Sub acadnavi()
  
  Call frmAcadNavi.show(vbModal)

End Sub

Sub regelstaat()
  
  Call frmregelstaat.show(vbModal)

End Sub
Sub replacet()
  ' gegevens exporteren voor Navision
  
    Dim drawing As AcadDocument
    For Each drawing In ThisDrawing.Application.Documents
        drawing.Activate
        ThisDrawing.SendCommand "pdf" & vbCr
'''        ThisDrawing.SendCommand "-VBARUN" & vbCr & "REGELXLS" & vbCr
'''        SendKeys "{ENTER}", True
    Next drawing
End Sub
Sub pfvba()
  ' plotfile aanmaken voor alle openstaande tekeningen
  
    Dim drawing As AcadDocument
    For Each drawing In ThisDrawing.Application.Documents
        drawing.Activate
        ThisDrawing.SendCommand "pf" & vbCr
        ThisDrawing.Close
        
    Next drawing
End Sub
Sub dp()
  ' dplot aanmaken voor alle openstaande tekeningen
  
    Dim drawing As AcadDocument
    For Each drawing In ThisDrawing.Application.Documents
        drawing.Activate
        ThisDrawing.SendCommand "a4laser" & vbCr
        ThisDrawing.SendCommand "pf" & vbCr
        ThisDrawing.SendCommand "inpak" & vbCr
        ThisDrawing.Close
    Next drawing
End Sub
Sub inpakvba()
  ' inpakken alle openstaande tekeningen
  
    Dim drawing As AcadDocument
    For Each drawing In ThisDrawing.Application.Documents
        drawing.Activate
        ThisDrawing.SendCommand "inpak" & vbCr
        ThisDrawing.Close
        
    Next drawing
End Sub
Sub direkt9400()
  ' direkt naar de oce9400 plotten
  
    Dim drawing As AcadDocument
    For Each drawing In ThisDrawing.Application.Documents
        drawing.Activate
        ThisDrawing.SendCommand "d9400" & vbCr
        ThisDrawing.Close
        
    Next drawing
End Sub
''''Sub pfs()
''''  ' gtekening plotfiles aanmaken
''''
''''    Dim drawing As AcadDocument
''''    For Each drawing In ThisDrawing.Application.Documents
''''        drawing.Activate
''''        ThisDrawing.SendCommand "pfspecial" & vbCr
''''          ThisDrawing.Close
''''    Next drawing
''''End Sub

Sub blokplotfile()
  ' blokfiles plotten bloks opzoeken, filteren op naam en dan een plotfilemaken
  

''    For Each drawing In ThisDrawing.Application.Documents
''        drawing.Activate
  ThisDrawing.SendCommand "filedia" & vbCr & "0" & vbCr
  ThisDrawing.SendCommand "cmddia" & vbCr & "0" & vbCr
  
        
        Dim teller
        Dim TELNAAM As String
        
        Dim pe1(0 To 2) As Double
        Dim insp
        Dim pe3 As String
        Dim pe4 As String
        teller = 0
        
        Dim element As Object
        For Each element In ThisDrawing.ModelSpace
              If element.ObjectName = "AcDbBlockReference" Then
                  If UCase(element.Name) = "KADER" Then
                  insp = element.InsertionPoint
                  
                  pe1(0) = insp(0) + 42000
                  pe1(1) = insp(1) + 28700
                  pe1(2) = 0
                       
                  pe3 = pe1(0) & "," & pe1(1)
                  pe4 = insp(0) & "," & insp(1)
                  
                  'MsgBox pe3
                  teller = teller + 1
                  TELNAAM = "C:\PLOT\" & teller & ".PLT"
                  'ThisDrawing.SendCommand "text" & vbCr & pe4 & vbCr & "100" & vbCr & "90" & vbCr & dwgnm & vbCr
                  'PLOTFILE AANMAKEN VOOR DE 5150
                  ThisDrawing.SendCommand "-plot" & vbCr & "Y" & vbCr & "Model" & vbCr & "5150.pc3" & vbCr & "A3" & vbCr & "M" & vbCr & "L" & vbCr & "N" & vbCr & "W" & vbCr & pe4 & vbCr & pe3 & vbCr & "1=100" & vbCr & "0,0" & vbCr & "Y" & vbCr & "monochrome.ctb" & vbCr & "Y" & vbCr & "W" & vbCr & "Y" & vbCr & TELNAAM & vbCr & "Y" & vbCr & "Y" & vbCr
                  End If

                  
                
              End If
         Next element
        
        
        
        
''''        ThisDrawing.Close
        
''''    Next drawing
End Sub

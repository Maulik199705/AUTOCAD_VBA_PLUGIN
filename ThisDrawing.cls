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
Private Sub AcadDocument_BeginClose()
Dim element
Dim SYMBOOL
Dim ATTRIBUTEN: Dim ATTRIBUUT: Dim zz: Dim rol
Dim i
For Each element In ThisDrawing.ModelSpace
        If element.ObjectName = "AcDbBlockReference" Then
            If UCase(element.Name) = "GROEPTEKSTBLOKNEW" Then
                Set SYMBOOL = element
                If SYMBOOL.HasAttributes Then
                    ATTRIBUTEN = SYMBOOL.GetAttributes
                    For i = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
                    Set ATTRIBUUT = ATTRIBUTEN(i)
                      If ATTRIBUUT.TagString = "FLEXFIX" And UCase(ATTRIBUUT.TextString) = "JA" Then
                       frmflexfixexport.show
                       Exit Sub
                      End If
                    Next i
                End If
            End If
        End If
        Next element

End Sub
Sub revi()
Dim ELEMENT700
Dim INSP700
Dim PE700(0 To 2) As Double
Dim NAAMBLOK700
NAAMBLOK700 = "C:\ACAD2002\DWG\REVISIEBLOK.DWG"
Dim OP
Dim OP1
Dim ATTRIBUUT: Dim BLOCKREFOBJ: Dim SYMBOOL: Dim ATTRIBUTEN: Dim x

 For Each ELEMENT700 In ThisDrawing.ModelSpace
             If ELEMENT700.ObjectName = "AcDbBlockReference" Then
                If UCase(ELEMENT700.Name) = "KADERLOGO" Then
'''                  Set SYMBOOL = ELEMENT700
'''                  If SYMBOOL.HasAttributes Then
'''                  ATTRIBUTEN = SYMBOOL.GetAttributes
'''                    For X = LBound(ATTRIBUTEN) To UBound(ATTRIBUTEN)
'''                    Set ATTRIBUUT = ATTRIBUTEN(X)
'''                        If ATTRIBUUT.TagString = "SCHAAL" Then
'''                        OP = ATTRIBUUT.TextString
'''                        If OP = "1:50" Then OP1 = 1
'''                        If OP = "1:100" Then OP1 = 0.5
'''                        If OP = "1:200" Then OP1 = 0.25
                        INSP700 = ELEMENT700.InsertionPoint
                        PE700(0) = INSP700(0) - 882
                        PE700(1) = INSP700(1) + 329
                        PE700(2) = 0
                        Set BLOCKREFOBJ = ThisDrawing.ModelSpace.InsertBlock(PE700, NAAMBLOK700, 1, 1, 1, 0)
                        End If
''                     Next X
''                   End If
''                 End If
              End If
              Update
Next ELEMENT700
End Sub
Sub vervangen()
  
  Call frmvervangen.show(vbModal)

End Sub
Sub translate()
  
  Call frmRenLayer.show(vbModal)

End Sub
Sub blokplot()
  ' blokfiles plotten bogro tekeningen, bloks opzoeken, filteren op naam en dan een A4 printje maken
  
    Dim dwgnm
    Dim drawing As AcadDocument
''    For Each drawing In ThisDrawing.Application.Documents
''        drawing.Activate
        
        Dim controle1
        Dim controle2
        Dim controle3
        Dim controle4
        
        
        Dim pe1(0 To 2) As Double
        Dim PE2(0 To 2) As Double
        Dim pe10(0 To 2) As Double
        Dim pe11(0 To 2) As Double
        Dim pe12(0 To 2) As Double
        Dim pe6(0 To 2) As Double
        Dim insp
        Dim pe3 As String
        Dim pe4 As String
        Dim pe30 As String
        Dim pe40 As String
        
        Dim element As Object
        For Each element In ThisDrawing.ModelSpace
              If element.ObjectName = "AcDbBlockReference" Then
                  controle1 = InStr(1, element.Name, "h0.kl.1", vbTextCompare)
                  controle2 = InStr(1, element.Name, "h1.kl.1", vbTextCompare)
                  controle3 = InStr(1, element.Name, "h2.kl.1", vbTextCompare)
                  'a''controle4 = InStr(1, element.Name, "h3.kl.1", vbTextCompare)
                   
                  
                  If controle1 <> 0 Or controle2 <> 0 Or controle3 <> 0 Or controle4 <> 0 Then
                    insp = element.InsertionPoint
                  pe1(0) = insp(0) + 7000
                  pe1(1) = insp(1) + 10000
                  pe1(2) = 0
                  
                  PE2(0) = insp(0) - 7000
                  PE2(1) = insp(1) - 10000
                  PE2(2) = 0
                  
                  pe6(0) = insp(0)
                  pe6(1) = insp(1)
                  pe6(2) = 0
 
                  
                  pe3 = pe1(0) & "," & pe1(1)
                  pe4 = PE2(0) & "," & PE2(1)
                  'MsgBox pe3
                  
        
                  'ThisDrawing.SendCommand "text" & vbCr & pe4 & vbCr & "100" & vbCr & "90" & vbCr & dwgnm & vbCr
                  'a''ThisDrawing.SendCommand "-plot" & vbCr & "Y" & vbCr & "Model" & vbCr & "\\FSERVER2\P39 L2100TN" & vbCr & "A4" & vbCr & "M" & vbCr & "P" & vbCr & "N" & vbCr & "W" & vbCr & pe3 & vbCr & pe4 & vbCr & "F" & vbCr & "0,0" & vbCr & "Y" & vbCr & "wth.ctb" & vbCr & "Y" & vbCr & "N" & vbCr & "N" & vbCr & "Y" & vbCr & "Y" & vbCr
                 
                  pe10(0) = PE2(0) + 300000
                  pe10(1) = PE2(1)
                  pe10(2) = 0
                    
                  pe11(0) = pe10(0) + 308000
                  pe11(1) = pe10(1) + 140000
                  pe11(2) = 0
                  
'''                  pe11(0) = pe2(0) + 8000
'''                  pe11(1) = pe2(1) + 140000
'''                  pe11(2) = 0
                  
                  element.Move PE2, pe10
                  pe30 = pe10(0) & "," & pe10(1)
                  pe40 = pe11(0) & "," & pe11(1)
                  
                  ThisDrawing.SendCommand "copyclip" & vbCr & "W" & vbCr & pe30 & vbCr & pe40 & vbCr & vbCr
                  dwgnm = ThisDrawing.GetVariable("dwgname")
                  Dim textObj As AcadText
                  ''Set textObj = ThisDrawing.ModelSpace.AddText(dwgnm, pe10, 100)
                  
                  
        ''          Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(insp, naamblok, 1, 1, 1, 0)
                  End If
                
                End If
             Next element
          
          Dim dwgnm2
          Dim check2
          dwgnm2 = ThisDrawing.GetVariable("dwgname")
          check2 = Split(dwgnm2, " ")
          Dim check3
          check3 = check2(2)
          
          ThisDrawing.Application.Documents.Add ("Autocad.dwt")
          
          ThisDrawing.SendCommand "pasteclip" & vbCr & "0,0" & vbCr
    
        
        'Zoom all
        
''    Next drawing
Call verder(check3)
End Sub
Sub verder(check3)
        
        Dim controle100
        Dim controle200
        Dim controle300
        Dim controle400
        Dim dw
        dw = "c:\acad2002\"
        Dim dw1
        dw1 = dw & check3 & ".dwg"
        
        Dim pe100(0 To 2) As Double
        Dim pe200(0 To 2) As Double
        Dim insp100
        Dim insp200
                
        Dim element As Object
        For Each element In ThisDrawing.ModelSpace
              If element.ObjectName = "AcDbBlockReference" Then
                  ''controle1 = InStr(1, element.Name, "h0.kl.1", vbTextCompare)
                  controle200 = InStr(1, element.Name, "h1.kl.1", vbTextCompare)
                  controle300 = InStr(1, element.Name, "h2.kl.1", vbTextCompare)
                  'a''controle4 = InStr(1, element.Name, "h3.kl.1", vbTextCompare)
                   
                  
                  If controle200 <> 0 Then
                   insp100 = element.InsertionPoint
                  
                  pe100(0) = insp100(0) + 10000
                  pe100(1) = insp100(1) - 39500
                  pe100(2) = 0
                  
                  element.Move insp100, pe100
                  End If
                  
                  
                  If controle300 <> 0 Then
                   insp200 = element.InsertionPoint
                  
                  pe200(0) = insp200(0) + 20000
                  pe200(1) = insp200(1) - 79000
                  pe200(2) = 0
                  
                  element.Move insp200, pe200
                  End If
                
                
                End If
        Next element
        
        
        
          ZoomAll
          ThisDrawing.SendCommand "scale" & vbCr & "all" & vbCr & vbCr & "0,0" & vbCr & "0.1" & vbCr
          
          
         ' ThisDrawing.SaveAs (dw1)
         ' ThisDrawing.Application.Documents.Close
End Sub
Sub delhatch()
  ' plotfile aanmaken voor alle openstaande tekeningen
  

  
    Dim drawing As AcadDocument
    For Each drawing In ThisDrawing.Application.Documents
        drawing.Activate
        
'''          ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr
'''            ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr
'''            ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr

        ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "AA$1A---VERBLIJFSGEBIED" & vbCr & vbCr
'''''        Dim layer
'''''        Dim entHandle As String
'''''        Dim entry As AcadEntity
'''''        For Each entry In ThisDrawing.ModelSpace
'''''            entHandle = entry.ObjectName
'''''            entry.Highlight (True)
'''''            entry.Highlight (False)
'''''            If entHandle = "AcDbHatch" Then entry.Delete
'''''            If entHandle = "AcDbSolid" Then entry.Delete
'''''            Update
'''''        Next
        
        
        
        ThisDrawing.Save
        ThisDrawing.Close
        
    Next drawing
End Sub
Sub ops()
'''    Dim drawing As AcadDocument
'''    For Each drawing In ThisDrawing.Application.Documents
'''    drawing.Activate
On Error Resume Next
ThisDrawing.SendCommand "audit" & vbCr & "y" & vbCr
        
    Dim objBlock As AcadBlock
    Dim objEntity As AcadEntity
    Dim objLayer As AcadLayer
'      Dim mypos
'
'      For Each objBlock In ThisDrawing.Blocks
'        For Each objEntity In objBlock
'          mypos = InStr(1, objEntity, "STRAMIEN")
'          objEntity.layer = "STRAMIEN"
'          objEntity.color = acByLayer
'          objEntity.Lineweight = acLnWtByLwDefault
'        Next objEntity
'      Next objBlock
'      ThisDrawing.SendCommand "-layer" & vbCr & "Freeze" & vbCr & "STRAMIEN" & vbCr & vbCr

      
         
Dim olddata As String
olddata = ThisDrawing.GetVariable("clayer")
ThisDrawing.SendCommand "-layer" & vbCr & "M" & vbCr & "BOUWKUNDIG" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "M" & vbCr & "3" & vbCr & vbCr
ThisDrawing.SendCommand "-layer" & vbCr & "Set" & vbCr & olddata & vbCr & vbCr
Update




  
      'alles naar layer bouwkundig
      Dim object3
        For Each object3 In ThisDrawing.ModelSpace
           object3.layer = "bouwkundig"
        Next
       
       
'set all object colors to bylayer
      For Each objBlock In ThisDrawing.Blocks
        For Each objEntity In objBlock
          objEntity.layer = "bouwkundig"
          objEntity.color = acByLayer
          objEntity.Lineweight = acLnWtByLwDefault
        Next objEntity
      Next objBlock
      
      ThisDrawing.SendCommand "burst" & vbCr & "all" & vbCr & vbCr
      
      
Dim layer
Dim entHandle As String
    Dim entry As AcadEntity
    For Each entry In ThisDrawing.ModelSpace
        entHandle = entry.ObjectName
        entry.Highlight (True)
        'MsgBox "The handle of this object is " & entHandle, vbInformation, "Handle Example"
        entry.Highlight (False)
        If entHandle = "AcDbDimension" Then entry.Delete
        If entHandle = "AcDbHatch" Then entry.Delete
        If entHandle = "AcDbSolid" Then entry.Delete
        If entHandle = "AcDbRotatedDimension" Then entry.Delete
        If entHandle = "AcDbAlignedDimension" Then entry.Delete
        If entHandle = "AcDbAngularDimension" Then entry.Delete
        
        If entHandle = "AcDbText" Then
            entry.layer = "3"
            entry.color = acByLayer
            If InStr(1, entry.TextString, "merk", vbBinaryCompare) Then entry.Delete
            If InStr(1, entry.TextString, "MERK", vbBinaryCompare) Then entry.Delete
        End If
        If entHandle = "AcDbMText" Then
    
            If InStr(1, entry.TextString, "merk", vbBinaryCompare) Then entry.Delete
            If InStr(1, entry.TextString, "MERK", vbBinaryCompare) Then entry.Delete
            If UCase(entry.TextString) = "K." Then entry.Delete
            If UCase(entry.TextString) = "MV" Then entry.Delete
            If UCase(entry.TextString) = "HWA" Then entry.Delete
            If UCase(entry.TextString) = "H.W.A." Then entry.Delete
            If UCase(entry.TextString) = "WINDVERBAND" Then entry.Delete
            entry.layer = "3"
            entry.color = acByLayer
        End If
        If entHandle = "AcDbImage" Then entry.Delete
        Update
    Next



ThisDrawing.PurgeAll
ThisDrawing.SendCommand "wthlayer" & vbCr
ThisDrawing.Save

'''    Next drawing
End Sub


Attribute VB_Name = "Module2"
'vul de waardes in excel in
Sub excel_schrijven1(text5)


Dim objworksheet As Object '' ---


    Dim MyXl As Object    ' Variable to hold reference
                                ' to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean    ' Flag for final release.

' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    Set MyXl = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear    ' Clear Err object in case error occurred.

    'Dim pak
    'pak = frmCalculatie.TextBox21
    Set MyXl = GetObject("G:\EXCEL\Calculatiesheet!\inregelen\Inregelstaat t.b.v. tekenkamer\digital.xlt")
    'c:\acad2002\digital.xlt")

    MyXl.Application.Visible = False
    MyXl.Parent.Windows(1).Visible = True
    Dim c
    c = frmregelxls.TextBox20 & frmregelxls.TextBox17 & "-RU" & text5 & ".xls"
    MyXl.SaveAs (c)
         
         Set MyXl = Nothing    ' Release reference to the
         'MyXl.Close                             ' application and spreadsheet.
        Call Module2.excel_schrijven2(c)
         
End Sub
Sub excel_schrijven2(c)
Dim objworksheet As Object '' ---
'MsgBox c

    Dim MyXl As Object    ' Variable to hold reference
                                ' to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean    ' Flag for final release.

' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    Set MyXl = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear    ' Clear Err object in case error occurred.

    Set MyXl = GetObject(c) '"g:\tekeningen\calculatie\Calculatie Autocad_test.xls")

    MyXl.Application.Visible = False
    MyXl.Parent.Windows(1).Visible = False
    
     MyXl.Activate
     
    'project gegevens
     Set objworksheet = MyXl.Sheets("Project gegevens")   'Projectgegevens")
     objworksheet.Activate
     With MyXl.ActiveSheet
     
          Dim teller3
          Dim k
          Dim textstring5
          Dim textstring6
          teller3 = frmregelxls.ListBox4.ListCount
                 
          For k = 0 To teller3 - 1
          
           'Define the text object
            textstring5 = frmregelxls.ListBox4.List(k)
            textstring6 = Split(textstring5, ("#"))
            L = k + 1
           
           
            .cells(L, 1).Value = textstring6(0)
            .cells(L, 2).Value = textstring6(1)
            
          Next k
       End With
       
 
     ' groepen en lengtes
     Set objworksheet = MyXl.Sheets("Calculatie")   'Projectgegevens")
     objworksheet.Activate
     With MyXl.ActiveSheet
          Dim teller4
          Dim kk
          Dim textstring7
          Dim textstring8
          teller4 = frmregelxls.ListBox2.ListCount
                 
          For kk = 0 To teller4 - 1
            'Define the text object
            textstring7 = frmregelxls.ListBox2.List(kk)
            textstring8 = Split(textstring7, ("#"))
            textstring10 = Split(textstring8(0), ".")
            nn = kk + 8 ' 7  'B7 in excel
            
            .cells(nn, 2).Value = textstring8(0)
            .cells(nn, 3).Value = textstring8(1)
            .cells(nn, 4).Value = textstring8(2)
            .cells(1, 3).Value = textstring8(4)
            .cells(4, 3).Value = textstring8(3)
           Next kk

     End With

   MyXl.WindowState = -4140    ' xlminimized


   Set MyXl = Nothing    ' Release reference to the application and spreadsheet.
                  
Call klaar(c)
                  
End Sub
Sub klaar(c)

 Dim objworksheet As Object '' ---
 Dim MyXl As Object    ' Variable to hold reference

    Dim ExcelWasNotRunning As Boolean    ' Flag for final release.
    On Error Resume Next    ' Defer error trapping.

    Set MyXl = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear    ' Clear Err object in case error occurred.

    Set MyXl = GetObject(c) '"g:\tekeningen\calculatie\Calculatie Autocad_test.xls")

    MyXl.Application.Visible = True
    MyXl.Parent.Windows(1).Visible = True
    
'    Set objworksheet = MyXl.Sheets("Output")   'Projectgegevens")
'    objworksheet.Activate
'
        Set objworksheet = MyXl.Sheets("CloseXcel")
        objworksheet.Activate

        
    
    Set MyXl = Nothing
    
   If frmregelxls.CheckBox4.Value = False Then Kill (c)
    
    'MyXl.Close (c)
End Sub


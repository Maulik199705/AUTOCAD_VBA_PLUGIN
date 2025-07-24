Attribute VB_Name = "M_0Algemeen_FindFile"
#If VBA7 Then
    Public Declare PtrSafe Function SearchPath Lib "kernel32" Alias "SearchPathA" _
        (ByVal lpPath As String, _
         ByVal lpFileName As String, _
         ByVal lpExtension As String, _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String, _
         ByVal lpFilePart As String) As Long
#Else
    Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" _
        (ByVal lpPath As String, _
         ByVal lpFileName As String, _
         ByVal lpExtension As String, _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String, _
         ByVal lpFilePart As String) As Long
#End If

Public Const MAX_PATH As Long = 260

' Returns the full path to the file if found in the search path, otherwise "ERROR"
Public Function FindFile(sFileName As String, Optional sPath As String = vbNullString) As String
    Dim retVal As Long
    Dim lpBuffer As String * MAX_PATH

    retVal = SearchPath(sPath, sFileName, vbNullString, MAX_PATH, lpBuffer, vbNullString)
    If retVal > 0 Then
        FindFile = Left(lpBuffer, retVal)
    Else
        FindFile = "ERROR"
    End If
End Function

' Looks for a file in SupportPath, Application Path, then Project Folder
Public Function FindZoekpad(bestandnaam As String) As String
    Dim sBestandnaam As String
    Dim sZoekpad As String

    sBestandnaam = bestandnaam
    sZoekpad = FindFile(sBestandnaam, ThisDrawing.Application.Preferences.Files.SupportPath)

    If sZoekpad = "ERROR" Then
        If Dir(ThisDrawing.Application.Path & "\" & sBestandnaam) <> "" Then
            FindZoekpad = ThisDrawing.Application.Path & "\" & sBestandnaam
        ElseIf Dir(CurDir & "\" & sBestandnaam) <> "" Then
            FindZoekpad = CurDir & "\" & sBestandnaam
        Else
            FindZoekpad = "ERROR"
        End If
    Else
        FindZoekpad = sZoekpad
    End If
End Function



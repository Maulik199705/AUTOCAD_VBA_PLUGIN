VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmpdf 
   Caption         =   "PDF-scale"
   ClientHeight    =   3060
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   4704
   OleObjectBlob   =   "frmpdf.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmpdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
frmpdf.Hide
Dim a As String
Dim b As String
a = frmpdf.TextBox1
b = frmpdf.TextBox2
'MsgBox a
'MsgBox b
ThisDrawing.SendCommand "scale" & vbCr & "All" & vbCr & vbCr & "0.0" & vbCr & "R" & vbCr & a & vbCr & b & vbCr

ZoomAll
Unload Me
End Sub

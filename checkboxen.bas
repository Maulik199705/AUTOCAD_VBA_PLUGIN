Attribute VB_Name = "checkboxen"
Option Explicit
Sub checkbox()

'hoeveel units van hetzelfde type heb ik

Dim b
b = frmAcadNavi.TextBox1
If b <> "-" Then
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox2 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox2 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox3 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox3 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox4 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox4 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox5 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox5 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox6 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox6 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox7 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox7 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox8 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox8 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox1 <> "-" And frmAcadNavi.TextBox1 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox21 = Val(frmAcadNavi.TextBox21) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b


'------------------------------------ textbox2

b = frmAcadNavi.TextBox2
If b <> "-" Then
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox3 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox3 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox4 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox4 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox5 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox5 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox6 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox6 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox7 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox7 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox8 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox8 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox2 <> "-" And frmAcadNavi.TextBox2 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox22 = Val(frmAcadNavi.TextBox22) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

     
'------------------------------------ textbox3

b = frmAcadNavi.TextBox3
If b <> "-" Then
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox4 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox4 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox5 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox5 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox6 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox6 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox7 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox7 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox8 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox8 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox3 <> "-" And frmAcadNavi.TextBox3 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox23 = Val(frmAcadNavi.TextBox23) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox4
b = frmAcadNavi.TextBox4
If b <> "-" Then
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox5 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox5 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox6 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox6 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox7 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox7 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox8 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox8 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox4 <> "-" And frmAcadNavi.TextBox4 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox24 = Val(frmAcadNavi.TextBox24) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b


'------------------------------------ textbox5
b = frmAcadNavi.TextBox5
If b <> "-" Then
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox6 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox6 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox7 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox7 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox8 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox8 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox5 <> "-" And frmAcadNavi.TextBox5 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox25 = Val(frmAcadNavi.TextBox25) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox6
b = frmAcadNavi.TextBox6
If b <> "-" Then
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox7 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox7 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox8 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox8 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox6 <> "-" And frmAcadNavi.TextBox6 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox26 = Val(frmAcadNavi.TextBox26) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox7
b = frmAcadNavi.TextBox7
If b <> "-" Then
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox8 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox8 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox7 <> "-" And frmAcadNavi.TextBox7 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox27 = Val(frmAcadNavi.TextBox27) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox8
b = frmAcadNavi.TextBox8
If b <> "-" Then
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox9 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox9 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox8 <> "-" And frmAcadNavi.TextBox8 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox28 = Val(frmAcadNavi.TextBox28) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox9
b = frmAcadNavi.TextBox9
If b <> "-" Then
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox10 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox10 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox9 <> "-" And frmAcadNavi.TextBox9 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox29 = Val(frmAcadNavi.TextBox29) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b


'------------------------------------ textbox10
b = frmAcadNavi.TextBox10
If b <> "-" Then
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox11 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox11 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox10 <> "-" And frmAcadNavi.TextBox10 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox30 = Val(frmAcadNavi.TextBox30) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox11
b = frmAcadNavi.TextBox11
If b <> "-" Then
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox12 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox12 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox11 <> "-" And frmAcadNavi.TextBox11 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox31 = Val(frmAcadNavi.TextBox31) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox12
b = frmAcadNavi.TextBox12
If b <> "-" Then
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox13 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox13 = "-"
End If
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox12 <> "-" And frmAcadNavi.TextBox12 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox32 = Val(frmAcadNavi.TextBox32) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox13
b = frmAcadNavi.TextBox13
If b <> "-" Then
If frmAcadNavi.TextBox13 <> "-" And frmAcadNavi.TextBox13 = frmAcadNavi.TextBox14 Then
      frmAcadNavi.TextBox33 = Val(frmAcadNavi.TextBox33) + 1
      frmAcadNavi.TextBox14 = "-"
End If
If frmAcadNavi.TextBox13 <> "-" And frmAcadNavi.TextBox13 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox33 = Val(frmAcadNavi.TextBox33) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox13 <> "-" And frmAcadNavi.TextBox13 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox33 = Val(frmAcadNavi.TextBox33) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox13 <> "-" And frmAcadNavi.TextBox13 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox33 = Val(frmAcadNavi.TextBox33) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox13 <> "-" And frmAcadNavi.TextBox13 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox33 = Val(frmAcadNavi.TextBox33) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox13 <> "-" And frmAcadNavi.TextBox13 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox33 = Val(frmAcadNavi.TextBox33) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox13 <> "-" And frmAcadNavi.TextBox13 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox33 = Val(frmAcadNavi.TextBox33) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox14
b = frmAcadNavi.TextBox14
If b <> "-" Then
If frmAcadNavi.TextBox14 <> "-" And frmAcadNavi.TextBox14 = frmAcadNavi.TextBox15 Then
      frmAcadNavi.TextBox34 = Val(frmAcadNavi.TextBox34) + 1
      frmAcadNavi.TextBox15 = "-"
End If
If frmAcadNavi.TextBox14 <> "-" And frmAcadNavi.TextBox14 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox34 = Val(frmAcadNavi.TextBox34) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox14 <> "-" And frmAcadNavi.TextBox14 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox34 = Val(frmAcadNavi.TextBox34) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox14 <> "-" And frmAcadNavi.TextBox14 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox34 = Val(frmAcadNavi.TextBox34) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox14 <> "-" And frmAcadNavi.TextBox14 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox34 = Val(frmAcadNavi.TextBox34) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox14 <> "-" And frmAcadNavi.TextBox14 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox34 = Val(frmAcadNavi.TextBox34) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox15
b = frmAcadNavi.TextBox15
If b <> "-" Then
If frmAcadNavi.TextBox15 <> "-" And frmAcadNavi.TextBox15 = frmAcadNavi.TextBox16 Then
      frmAcadNavi.TextBox35 = Val(frmAcadNavi.TextBox35) + 1
      frmAcadNavi.TextBox16 = "-"
End If
If frmAcadNavi.TextBox15 <> "-" And frmAcadNavi.TextBox15 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox35 = Val(frmAcadNavi.TextBox35) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox15 <> "-" And frmAcadNavi.TextBox15 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox35 = Val(frmAcadNavi.TextBox35) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox15 <> "-" And frmAcadNavi.TextBox15 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox35 = Val(frmAcadNavi.TextBox35) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox15 <> "-" And frmAcadNavi.TextBox15 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox35 = Val(frmAcadNavi.TextBox35) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox16
b = frmAcadNavi.TextBox16
If b <> "-" Then
If frmAcadNavi.TextBox16 <> "-" And frmAcadNavi.TextBox16 = frmAcadNavi.TextBox17 Then
      frmAcadNavi.TextBox36 = Val(frmAcadNavi.TextBox36) + 1
      frmAcadNavi.TextBox17 = "-"
End If
If frmAcadNavi.TextBox16 <> "-" And frmAcadNavi.TextBox16 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox36 = Val(frmAcadNavi.TextBox36) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox16 <> "-" And frmAcadNavi.TextBox16 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox36 = Val(frmAcadNavi.TextBox36) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox16 <> "-" And frmAcadNavi.TextBox16 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox36 = Val(frmAcadNavi.TextBox36) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox17
b = frmAcadNavi.TextBox17
If b <> "-" Then

If frmAcadNavi.TextBox17 <> "-" And frmAcadNavi.TextBox17 = frmAcadNavi.TextBox18 Then
      frmAcadNavi.TextBox37 = Val(frmAcadNavi.TextBox37) + 1
      frmAcadNavi.TextBox18 = "-"
End If
If frmAcadNavi.TextBox17 <> "-" And frmAcadNavi.TextBox17 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox37 = Val(frmAcadNavi.TextBox37) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox17 <> "-" And frmAcadNavi.TextBox17 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox37 = Val(frmAcadNavi.TextBox37) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox18
b = frmAcadNavi.TextBox18
If b <> "-" Then
If frmAcadNavi.TextBox18 <> "-" And frmAcadNavi.TextBox18 = frmAcadNavi.TextBox19 Then
      frmAcadNavi.TextBox38 = Val(frmAcadNavi.TextBox38) + 1
      frmAcadNavi.TextBox19 = "-"
End If
If frmAcadNavi.TextBox18 <> "-" And frmAcadNavi.TextBox18 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox38 = Val(frmAcadNavi.TextBox38) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b

'------------------------------------ textbox19
b = frmAcadNavi.TextBox19
If b <> "-" Then
If frmAcadNavi.TextBox19 <> "-" And frmAcadNavi.TextBox19 = frmAcadNavi.TextBox20 Then
      frmAcadNavi.TextBox39 = Val(frmAcadNavi.TextBox39) + 1
      frmAcadNavi.TextBox20 = "-"
End If
End If 'eind b


End Sub
     
     

Attribute VB_Name = "M_0SaveGetSettings"
Sub SaveSettings()
    SaveSetting "Leidinglegprogramma", "Startup", "HOHAfstand", F_Main.ComboBox2.Value
    SaveSetting "Leidinglegprogramma", "Startup", "LeidingType", F_Main.ComboBox3.Value
    SaveSetting "Leidinglegprogramma", "Startup", "Snit", F_Main.TextBox1.Text
    SaveSetting "Leidinglegprogramma", "Startup", "OpgevenAantalLeid", F_Main.CheckBox10.Value
    
    

End Sub


Sub GetSettings()
    F_Main.ComboBox2.Text = GetSetting("Leidinglegprogramma", "Startup", "HOHAfstand", "")
    F_Main.ComboBox3.Text = GetSetting("Leidinglegprogramma", "Startup", "LeidingType", "")
    F_Main.TextBox1.Text = GetSetting("Leidinglegprogramma", "Startup", "Snit", "0")
    F_Main.CheckBox10.Value = GetSetting("Leidinglegprogramma", "Startup", "OpgevenAantalLeid", False)
       
    
    If Trim(F_Main.TextBox1.Text) = "" Then F_Main.TextBox1.Text = 0
End Sub




        


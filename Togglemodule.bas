Attribute VB_Name = "togglemodule"
Sub groeptekst_dutch()
frmGroeptekst.Caption = "GROEPTEKST PLAATSEN  [proefversie A]"
frmGroeptekst.Frame1.Caption = "Rollengte's"
frmGroeptekst.Frame2.Caption = "Groeptekst"
frmGroeptekst.Frame3.Caption = "Unitlogo"
frmGroeptekst.Frame4.Caption = "Wandhoogte"
frmGroeptekst.Frame5.Caption = "Groep en unitlogo"
frmGroeptekst.Frame6.Caption = "Groep-range verwijderen"
frmGroeptekst.Label8.Caption = "Unit"
frmGroeptekst.Label9.Caption = "Groepsnummer"
frmGroeptekst.Label12.Caption = "Type unit"
frmGroeptekst.Label13.Caption = "Type regeling"
frmGroeptekst.Label14.Caption = "Bevestiging"
frmGroeptekst.Label26.Caption = "Type buis"
frmGroeptekst.Label23.Caption = "Unitnummer"
frmGroeptekst.Label29.Caption = "Extra groepen"
frmGroeptekst.Label34.Caption = "Unitnummer"
frmGroeptekst.Label35.Caption = "van groepsnummer"
frmGroeptekst.Label36.Caption = "t/m groepsnummer"
frmGroeptekst.Label37.Caption = "Totaal Flexfix"
frmGroeptekst.Label38.Caption = "Aantal matten"
frmGroeptekst.CheckBox2.Caption = "Werkelijke rollengte"
frmGroeptekst.cmdAfsluiten.Caption = "Afsluiten"
frmGroeptekst.CmdBloklogo.Caption = "Unitlogo"
frmGroeptekst.CmdErase.Caption = "Verwijder groeptekst"
frmGroeptekst.Cmdmeten.Caption = "Meet de groep"
frmGroeptekst.cmdTelrollen.Caption = "Aantal rollen tellen"
frmGroeptekst.CommandButton4.Caption = "Verwijder lege groep- & wandlayers"
frmGroeptekst.CommandButton5.Caption = "Verwijder groeptekst"
frmGroeptekst.CommandButton6.Caption = "Unitlogo bijwerken"
frmGroeptekst.CheckBox1.Caption = "Variabele rollengte"
frmGroeptekst.CheckBox7.Caption = "Afwijkende A&R"
frmGroeptekst.Label40.Caption = "Restlengte"
frmGroeptekst.TextBox13.ControlTipText = "EEN PUNT GEBRUIKEN (GEEN KOMMA)"
frmGroeptekst.TextBox24.ControlTipText = "Vul hier de totale restlengte in. (restlengte = A en  R)"
frmGroeptekst.TextBox28.ControlTipText = "Restlengte (= aanvoer EN retour)"
frmGroeptekst.OptionButton5.Caption = "Variabel"
frmGroeptekst.OptionButton7.Caption = "Ringleiding"

End Sub
Sub groeptekst_english()
frmGroeptekst.Caption = "PLACE GROUPTEXT [proefversie A]"
frmGroeptekst.Frame1.Caption = "Roll length"
frmGroeptekst.Frame2.Caption = "Grouptext"
frmGroeptekst.Frame3.Caption = "Manifold logo"
frmGroeptekst.Frame4.Caption = "Wall height"
frmGroeptekst.Frame5.Caption = "Group/Manifold"
frmGroeptekst.Frame6.Caption = "Delete group-range"
frmGroeptekst.Label8.Caption = "Manifold"
frmGroeptekst.Label9.Caption = "Groupnumber"
frmGroeptekst.Label12.Caption = "Manifold type"
frmGroeptekst.Label13.Caption = "Control type"
frmGroeptekst.Label14.Caption = "Fixing material"
frmGroeptekst.Label26.Caption = "Tube type"
frmGroeptekst.Label23.Caption = "Manifold"
frmGroeptekst.Label29.Caption = "Extra groups"
frmGroeptekst.Label34.Caption = "Manifold"
frmGroeptekst.Label35.Caption = "from groupnumber"
frmGroeptekst.Label36.Caption = "including groupnumber"
frmGroeptekst.Label37.Caption = "Total Flexfix"
frmGroeptekst.Label38.Caption = "Amount mats"
frmGroeptekst.CheckBox2.Caption = "True roll length"
frmGroeptekst.cmdAfsluiten.Caption = "Close"
frmGroeptekst.CmdBloklogo.Caption = "Manifold logo"
frmGroeptekst.CmdErase.Caption = "Delete grouptext"
frmGroeptekst.Cmdmeten.Caption = "Measure the group"
frmGroeptekst.cmdTelrollen.Caption = "Count the groups"
frmGroeptekst.CommandButton4.Caption = "Delete empty grouplayers"
frmGroeptekst.CommandButton5.Caption = "Delete grouptext"
frmGroeptekst.CommandButton6.Caption = "Update Manifold logo"
frmGroeptekst.CheckBox1.Caption = "Variable roll length"
frmGroeptekst.CheckBox7.Caption = "Different flow & return"
frmGroeptekst.Label40.Caption = "Rest length"
frmGroeptekst.TextBox13.ControlTipText = "USE A COMMA TO SEPARATE FIGURES (EXAMPLE: 2,4 and not 2.4)"
frmGroeptekst.TextBox24.ControlTipText = "Fill in the rest length (flow and return)"
frmGroeptekst.TextBox28.ControlTipText = "Rest length. (flow and return)"
frmGroeptekst.OptionButton5.Caption = "Variable"
frmGroeptekst.OptionButton7.Caption = "Loop"
End Sub
Sub groeptekst_combolijst5()
frmGroeptekst.ComboBox1.Clear
frmGroeptekst.ComboBox1.AddItem "Sack tie"
frmGroeptekst.ComboBox1.AddItem "Hammerclips"
frmGroeptekst.ComboBox1.AddItem "Nails"
frmGroeptekst.ComboBox1.AddItem "IFD Cardboard"
frmGroeptekst.ComboBox1.AddItem "IFD Polystyrene"
frmGroeptekst.ComboBox1.AddItem "Isoclips"
frmGroeptekst.ComboBox1.AddItem "Wedge"
frmGroeptekst.ComboBox1.AddItem "Tiewrap"
frmGroeptekst.ComboBox1.AddItem "Varisoclips"
End Sub
Sub groeptekst_waarschuwing_toggle()
frmWaarschuwing.Label1.Caption = "roll(s) on manifold"
frmWaarschuwing.Label2.Caption = " is/are 165 meter"
frmWaarschuwing.cmdnietsveranderen.Caption = " Correct, no changes"
frmWaarschuwing.cmdwijzigen.Caption = " Change the roll length to 125 meter"
frmWaarschuwing.Caption = "Caution"
End Sub

Sub kaderlogo_english()
frmKaderlogo.Caption = "Framework logo"
frmKaderlogo.Frame1.Caption = "Client & Project information"
frmKaderlogo.Frame2.Caption = "Alterations"
frmKaderlogo.cmdAfsluiten.Caption = "Close"
frmKaderlogo.cmdAfsluiten.Accelerator = "C"
frmKaderlogo.CmdUpdate.ControlTipText = "Fill in or change the framework logo"
frmKaderlogo.cmdAfsluiten.ControlTipText = "Close window"
frmKaderlogo.ToggleButton1.ControlTipText = "Show/hide the alterations"
frmKaderlogo.Label42.Caption = "Client": frmKaderlogo.Label41.Caption = "Sheet": frmKaderlogo.Label48.Caption = "Alt. a"
frmKaderlogo.Label43.Caption = "Location": frmKaderlogo.Label40.Caption = "Size": frmKaderlogo.Label49.Caption = "Alt. b"
frmKaderlogo.Label45.Caption = "Description": frmKaderlogo.Label39.Caption = "Name": frmKaderlogo.Label50.Caption = "Alt.c"
frmKaderlogo.Label44.Caption = "Instal.Adress": frmKaderlogo.Label38.Caption = "Scale": frmKaderlogo.Label51.Caption = "Alt. d"
frmKaderlogo.Label46.Caption = "Instal.Location ": frmKaderlogo.Label37.Caption = "Date": frmKaderlogo.Label52.Caption = "Alt. e"
frmKaderlogo.Label47.Caption = "Project": frmKaderlogo.Label55.Caption = "Revision": frmKaderlogo.Label53.Caption = "Alt. f"
frmKaderlogo.Label54.Caption = "Alt. g"
frmKaderlogo.TextBox1.ControlTipText = "The name of the client": frmKaderlogo.TextBox2.ControlTipText = "Place of the client"
frmKaderlogo.TextBox3.ControlTipText = "The name of the project": frmKaderlogo.TextBox4.ControlTipText = "Installation adress"
frmKaderlogo.TextBox5.ControlTipText = "Installation location": frmKaderlogo.TextBox6.ControlTipText = "Projectnumber"
frmKaderlogo.TextBox8.ControlTipText = "Size of the drawing": frmKaderlogo.TextBox9.ControlTipText = "Name of the draftsman"
frmKaderlogo.TextBox10.ControlTipText = "Scale": frmKaderlogo.TextBox11.ControlTipText = "Sign date"
frmKaderlogo.TextBox12.ControlTipText = "Alteration 1": frmKaderlogo.TextBox13.ControlTipText = "Alteration 2"
frmKaderlogo.TextBox14.ControlTipText = "Alteration 3": frmKaderlogo.TextBox15.ControlTipText = "Alteration 4"
frmKaderlogo.TextBox16.ControlTipText = "Alteration 5": frmKaderlogo.TextBox17.ControlTipText = "Alteration 6"
frmKaderlogo.TextBox18.ControlTipText = "Alteration 7"
frmKaderlogo.ComboBox1.ControlTipText = "Sheet number"
End Sub
Sub kaderlogo_dutch()
frmKaderlogo.Caption = "Kaderlogo"
frmKaderlogo.Frame1.Caption = "Invulgegevens klant"
frmKaderlogo.Frame2.Caption = "Wijzigingen"
frmKaderlogo.cmdAfsluiten.Caption = "Afsluiten"
frmKaderlogo.cmdAfsluiten.Accelerator = "A"
frmKaderlogo.CmdUpdate.ControlTipText = "Kaderlogo invullen of wijzigen"
frmKaderlogo.cmdAfsluiten.ControlTipText = "Venster sluiten"
frmKaderlogo.ToggleButton1.ControlTipText = "Hiermee komen de wijzigingsdatums te voorschijn"
frmKaderlogo.Label42.Caption = "Opdrachtgever": frmKaderlogo.Label41.Caption = "Blad": frmKaderlogo.Label48.Caption = "Wijziging a"
frmKaderlogo.Label43.Caption = "Plaats": frmKaderlogo.Label40.Caption = "Formaat": frmKaderlogo.Label49.Caption = "Wijziging b"
frmKaderlogo.Label45.Caption = "Projectnaam": frmKaderlogo.Label39.Caption = "Tekenaar": frmKaderlogo.Label50.Caption = "Wijziging c"
frmKaderlogo.Label44.Caption = "Montageadres": frmKaderlogo.Label38.Caption = "Schaal": frmKaderlogo.Label51.Caption = "Wijziging d"
frmKaderlogo.Label46.Caption = "Montageplaats": frmKaderlogo.Label37.Caption = "Datum": frmKaderlogo.Label52.Caption = "Wijziging e"
frmKaderlogo.Label47.Caption = "Projectnummer": frmKaderlogo.Label55.Caption = "Revisie": frmKaderlogo.Label53.Caption = "Wijziging f"
frmKaderlogo.Label54.Caption = "Wijziging g"
frmKaderlogo.TextBox1.ControlTipText = "Vul naam van de opdrachtgever in": frmKaderlogo.TextBox2.ControlTipText = "Plaats v/d opdr.gever"
frmKaderlogo.TextBox3.ControlTipText = "Vul projectnaam in": frmKaderlogo.TextBox4.ControlTipText = "Vul montage adres in"
frmKaderlogo.TextBox5.ControlTipText = "Vul montageplaats in": frmKaderlogo.TextBox6.ControlTipText = "Projectnummer"
frmKaderlogo.TextBox8.ControlTipText = "Formaat v/d tekening": frmKaderlogo.TextBox9.ControlTipText = "Tekenaar"
frmKaderlogo.TextBox10.ControlTipText = "Schaal": frmKaderlogo.TextBox11.ControlTipText = "Tekendatum"
frmKaderlogo.TextBox12.ControlTipText = "Wijzigingsdatum 1": frmKaderlogo.TextBox13.ControlTipText = "Wijzigingsdatum 2"
frmKaderlogo.TextBox14.ControlTipText = "Wijzigingsdatum 3": frmKaderlogo.TextBox15.ControlTipText = "Wijzigingsdatum 4"
frmKaderlogo.TextBox16.ControlTipText = "Wijzigingsdatum 5": frmKaderlogo.TextBox17.ControlTipText = "Wijzigingsdatum 6"
frmKaderlogo.TextBox18.ControlTipText = "Wijzigingsdatum 7"
frmKaderlogo.ComboBox1.ControlTipText = "Bladnummer"
End Sub




























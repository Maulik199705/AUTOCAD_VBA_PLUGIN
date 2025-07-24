VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_About 
   Caption         =   "About WTH-leidingleg programma"
   ClientHeight    =   6705
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8484.001
   OleObjectBlob   =   "F_About.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ListBox1.Clear
    ListBox1.AddItem ""
    ListBox1.AddItem "Leidingtekenprogramma WTH"
    ListBox1.AddItem ""
    ListBox1.AddItem "Door E.Breedveld"
    ListBox1.AddItem ""
    ListBox1.AddItem "CAD E&D Education Developing"
    ListBox1.AddItem "www.caded.nl"
    ListBox1.AddItem "tel: 078-6911342"
    ListBox1.AddItem "mob: 06-28817856"
    ListBox1.AddItem ""
    ListBox1.AddItem ""
    ListBox1.AddItem "Versie 1.0    1 april 2003"
    ListBox1.AddItem "Versie 1.1    21 mei 2003 (finetunen gehele programma)"
    ListBox1.AddItem "Versie 1.2    27 mei 2003 (maken aanpassingen nav bezoek 26 mei)"
    ListBox1.AddItem "Versie 1.3    31 juli 2003 (maken aanpassingen/ uitbreiding slingerafronding nav bezoek 11 juli)"
    ListBox1.AddItem "Versie 1.4    31 jan 2004 (finetunen voor produktie)"
    ListBox1.AddItem "Versie 1.5    10 feb 2004 (aanpassen offset duo-leiding)"
    ListBox1.AddItem "Versie 1.6    14 feb 2004 (aanpassen offset duo-leiding)"
    ListBox1.AddItem "Versie 2.0    27 mei Finetuning en toevoeging lengtemonitor"
    ListBox1.AddItem "Versie 2.1    14 juni Finetuning n.a.v. demo aan gebruikers 27 mei 2004"
    
End Sub

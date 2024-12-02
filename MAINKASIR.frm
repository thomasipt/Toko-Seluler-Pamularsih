VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MAINKASIR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU KASIR"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4365
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "MAINKASIR.frx":0000
      Top             =   2040
   End
End
Attribute VB_Name = "MAINKASIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
    With StatusBar1.Panels
        .Item(1).Style = sbrText
        .Item(1).Text = "USERCODE : " & Operator
        .Item(1).AutoSize = sbrSpring
        .Item(2).Style = sbrText
        .Item(2).AutoSize = sbrSpring
        .Item(2).Text = "TANGGAL SYSTEM  : " & TglS
        .Item(3).Style = sbrText
        .Item(3).AutoSize = sbrSpring
        .Item(3).Text = "Copyright® EDP IT SOLUTION"
    End With
End Sub

Private Sub K_Click(Index As Integer)
cepat = 1000
While Top - Height < Screen.Height
    DoEvents
    Top = Top + cepat
Wend
Hide
Unload Me
LOGIN.Show 1
End Sub


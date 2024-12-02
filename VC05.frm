VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VC05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL BARANG"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "SEMUA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7890
      TabIndex        =   6
      Top             =   105
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   105
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1575
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   105
      Width           =   4155
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3967
      TabIndex        =   0
      Top             =   5775
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7140
      OleObjectBlob   =   "VC05.frx":0000
      Top             =   4830
   End
   Begin VB.PictureBox Picture1 
      Height          =   1230
      Left            =   -735
      ScaleHeight     =   1170
      ScaleWidth      =   11025
      TabIndex        =   1
      Top             =   5565
      Width           =   11085
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "VC05.frx":0234
      TabIndex        =   2
      Top             =   180
      Width           =   1410
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4860
      Left            =   30
      TabIndex        =   3
      Top             =   525
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   8573
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "VC05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi, Pusing As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Command1_Click()
grid.Clear
Call SiapkanGrid2
Call IsiGrid2
End Sub

Private Sub Command2_Click()
Call SiapkanGrid
Call IsiGrid
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""

Call SiapkanGrid
Call IsiGrid
grid.Refresh

SSPL = "Select NamaBR From B003"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
If RSPL.RowCount <> 0 Then
    RSPL.MoveFirst
    Do While Not RSPL.EOF
        Combo1.AddItem RSPL("NamaBR")
    RSPL.MoveNext
    Loop
    RSPL.Close
    Set RSPL = Nothing
    Combo1.ListIndex = 0
End If


End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "BELI": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1000: .Text = "JUAL": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "SALDO": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "H.JUAL": .CellAlignment = 4
End With
End Sub

Private Sub SiapkanGrid2()
grid.Rows = 2
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "BELI": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1000: .Text = "JUAL": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "SALDO": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "H.JUAL": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SKTG = "Select * From B003 order by KodeInd,KodeBR Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("KodeBR"): .CellAlignment = 1
              .Col = 2: .Text = RKTG("JD"): .CellAlignment = 4
              .Col = 3: .Text = RKTG("JC"): .CellAlignment = 4
              .Col = 4: .Text = RKTG("JAkhir"): .CellAlignment = 4
              .Col = 5: .Text = Format(RKTG("HJual"), "##,###.00")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub IsiGrid2()
SKTG = "Select * From B003 where NamaBR = '" + Trim(Combo1) + "'"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("KodeBR"): .CellAlignment = 1
              .Col = 2: .Text = RKTG("JD"): .CellAlignment = 4
              .Col = 3: .Text = RKTG("JC"): .CellAlignment = 4
              .Col = 4: .Text = RKTG("JAkhir"): .CellAlignment = 4
              .Col = 5: .Text = Format(RKTG("HJual"), "##,###.00")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

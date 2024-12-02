VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form JS03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVICE KELUAR"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleMode       =   0  'User
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4268
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "JS03.frx":0000
      Top             =   3330
      Width           =   4740
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2310
      Left            =   90
      TabIndex        =   14
      Top             =   1710
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   4075
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SPAREPART"
      TabPicture(0)   =   "JS03.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SERVICE"
      TabPicture(1)   =   "JS03.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid4"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid grid3 
         Height          =   1710
         Left            =   90
         TabIndex        =   15
         Top             =   495
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   3016
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   16777152
         GridColor       =   0
         Enabled         =   -1  'True
         TextStyle       =   3
         TextStyleFixed  =   3
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grid4 
         Height          =   1710
         Left            =   -74910
         TabIndex        =   16
         Top             =   495
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   3016
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   16777152
         GridColor       =   0
         Enabled         =   -1  'True
         TextStyle       =   3
         TextStyleFixed  =   3
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3690
      TabIndex        =   11
      Top             =   105
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1155
      TabIndex        =   9
      Text            =   "1"
      Top             =   105
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5288
      TabIndex        =   8
      Top             =   4095
      Width           =   3600
   End
   Begin VB.CommandButton cmdEDIT 
      Caption         =   "SIMPAN"
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
      Left            =   203
      TabIndex        =   7
      Top             =   4095
      Width           =   3600
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   4268
      TabIndex        =   0
      Top             =   1755
      Width           =   4740
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         TabIndex        =   3
         Text            =   "12"
         Top             =   1065
         Width           =   2520
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         TabIndex        =   2
         Text            =   "11"
         Top             =   645
         Width           =   2520
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         TabIndex        =   1
         Text            =   "10"
         Top             =   225
         Width           =   2520
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   300
         Left            =   315
         OleObjectBlob   =   "JS03.frx":003C
         TabIndex        =   4
         Top             =   315
         Width           =   1665
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   300
         Left            =   315
         OleObjectBlob   =   "JS03.frx":009D
         TabIndex        =   5
         Top             =   735
         Width           =   1665
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   300
         Left            =   315
         OleObjectBlob   =   "JS03.frx":0106
         TabIndex        =   6
         Top             =   1155
         Width           =   1665
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9330
      OleObjectBlob   =   "JS03.frx":017B
      Top             =   945
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   195
      Left            =   90
      OleObjectBlob   =   "JS03.frx":03AF
      TabIndex        =   10
      Top             =   150
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1170
      Left            =   3690
      TabIndex        =   12
      Top             =   105
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   2064
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      GridColor       =   0
      Enabled         =   -1  'True
      TextStyle       =   3
      TextStyleFixed  =   3
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   1170
      Left            =   90
      TabIndex        =   13
      Top             =   450
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   2064
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      GridColor       =   0
      Enabled         =   -1  'True
      TextStyle       =   3
      TextStyleFixed  =   3
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "JS03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String
Dim KodeK As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl, RDel2 As rdoResultset
Private SDel, SDel2 As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String

Private Sub cmdEDIT_Click()
SCari = "Select * From JS02 where NOTA = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.Edit
    RCari("AMBIL") = "SELESAI"
RCari.Update
RCari.Close
Set RCari = Nothing

Call CTK

Unload Me
JS03.Show 1
End Sub

Private Sub CTK()
Dim tanya
tanya = MsgBox("CETAK FAKTUR", vbOKCancel, "KONFIRMASI")
    If tanya = vbOK Then
        Printer.Font = "Courier New"
        Printer.FontSize = 9
        Printer.Print Tab(5); ""
        Printer.Print Tab(5); ""
        Printer.Print Tab(5); ""
        Printer.Print Tab(6); "TOTAL."; Spc(3); Text10
        Printer.Print Tab(6); "DP."; Spc(3); Text11
        Printer.Print Tab(6); "SISA."; Spc(3); Text12
        Printer.Print Tab(6); "TANGGAL."; Spc(3); Now
        Printer.EndDoc
    Else
        Exit Sub
    End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
grid1.Visible = True
Command3.Visible = False
Text1 = ""
End Sub

Private Sub Form_Load()
Dim tanya
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Call SiapkanGrid1
Call IsiGrid1

Call SiapkanGrid2

grid1.Visible = False

SSTab1.Tab = 0

End Sub

Private Sub SiapkanGrid1()
With grid1
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "NOTA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 3000: .Text = "NAMA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid1()
SCari = "Select * From JS02 where AMBIL = 'BELUM SELESAI' order by NOTA Asc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid1.Rows = B + 1
      grid1.Row = B
         With grid1
            .Col = 0: .Text = B: .CellAlignment = 4
            .Col = 1: .Text = RCari("Nota"): .CellAlignment = 4
            .Col = 2: .Text = RCari("Nama")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub SiapkanGrid2()
With grid2
    .Cols = 11
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "NOTA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2000: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 2500: .Text = "ALAMAT": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1000: .Text = "MASUK": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1000: .Text = "JENIS": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 1000: .Text = "SERI": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 1500: .Text = "KERUSAKAN": .CellAlignment = 4
    .Col = 8: .ColWidth(8) = 1500: .Text = "KETERANGAN": .CellAlignment = 4
    .Col = 9: .ColWidth(9) = 1000: .Text = "SPAREPART": .CellAlignment = 4
    .Col = 10: .ColWidth(10) = 1000: .Text = "SERVICE": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid2()
SCari = "Select * From JS02 where NOTA= '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid2.Rows = B + 1
      grid2.Row = B
         With grid2
            .Col = 0: .Text = B: .CellAlignment = 4
            .Col = 1: .Text = RCari("Nota"): .CellAlignment = 4
            .Col = 2: .Text = RCari("Nama")
            .Col = 3: .Text = RCari("Alamat")
            .Col = 4: .Text = RCari("Masuk")
            .Col = 5: .Text = RCari("Jenis")
            .Col = 6: .Text = RCari("Seri")
            .Col = 7: .Text = RCari("Kerusakan")
            .Col = 8: .Text = RCari("Keterangan")
            .Col = 9: .Text = Format(RCari("Sparepart"), "##,###.00")
            .Col = 10: .Text = Format(RCari("Servis"), "##,###.00")
            Text10 = Format(RCari("TOTAL"), "##,###.00")
            Text11 = Format(RCari("BAYAR"), "##,###.00")
            Text12 = Format(RCari("SISA"), "##,###.00")
            Text2 = RCari("TERBILANG")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid1_dblClick()
Text1 = (grid1.TextMatrix(grid1.Row, 1))
grid1.Visible = False
Command3.Visible = True

Call IsiGrid2

SiapkanGrid3
Call IsiGrid3

SiapkanGrid4
Call IsiGrid4

End Sub

Private Sub SiapkanGrid3()
With grid3
    .Cols = 4
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 500: .Text = "JML": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "HARGA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid3()
SCari3 = "Select * From JS02HIS where NOTA= '" + Trim(Text1) + "' and STS='SP'"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurReadOnly)
If RCari3.RowCount <> 0 Then
   RCari3.MoveFirst
   B = 1
   Do Until RCari3.EOF
      grid3.Rows = B + 1
      grid3.Row = B
         With grid3
            .Col = 0: .Text = RCari3("Kode"): .CellAlignment = 4
            .Col = 1: .Text = RCari3("Nama")
            .Col = 2: .Text = RCari3("Jumlah"): .CellAlignment = 4
            .Col = 3: .Text = Format(RCari3("Harga"), "##,###.00")
         End With
      B = B + 1
      RCari3.MoveNext
   Loop
End If
RCari3.Close
Set RCari3 = Nothing
End Sub

Private Sub SiapkanGrid4()
With grid4
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "BIAYA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid4()
SCari4 = "Select * From JS02HIS where NOTA= '" + Trim(Text1) + "' and STS='SVC'"
Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurReadOnly)
If RCari4.RowCount <> 0 Then
   RCari4.MoveFirst
   B = 1
   Do Until RCari4.EOF
      grid4.Rows = B + 1
      grid4.Row = B
         With grid4
            .Col = 0: .Text = RCari4("Kode"): .CellAlignment = 4
            .Col = 1: .Text = RCari4("Nama")
            .Col = 2: .Text = Format(RCari4("Biaya"), "##,###.00")
         End With
      B = B + 1
      RCari4.MoveNext
   Loop
End If
RCari4.Close
Set RCari4 = Nothing
End Sub

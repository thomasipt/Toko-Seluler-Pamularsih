VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form VC01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE INDUK VOUCHER"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2385
      TabIndex        =   10
      Text            =   "5"
      Top             =   4410
      Width           =   2025
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1417
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   79
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1417
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   499
      Width           =   3330
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
      Left            =   5047
      TabIndex        =   4
      Top             =   2697
      Width           =   1890
   End
   Begin VB.CommandButton cmdDEL 
      Caption         =   "HAPUS"
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
      Left            =   5047
      TabIndex        =   3
      Top             =   1422
      Width           =   1890
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   5640
      Top             =   4635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   330
      OleObjectBlob   =   "VC01.frx":0000
      Top             =   6525
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   157
      OleObjectBlob   =   "VC01.frx":0234
      TabIndex        =   6
      Top             =   154
      Width           =   930
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   157
      OleObjectBlob   =   "VC01.frx":029A
      TabIndex        =   7
      Top             =   574
      Width           =   930
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2220
      Left            =   157
      TabIndex        =   8
      Top             =   1009
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   3916
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
   Begin VB.CommandButton cmdOK 
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
      Left            =   5032
      TabIndex        =   2
      Top             =   147
      Width           =   1890
   End
   Begin VB.CommandButton cmdEDIT 
      Caption         =   "EDIT"
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
      Left            =   5032
      TabIndex        =   5
      Top             =   147
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   4905
      ScaleHeight     =   3915
      ScaleWidth      =   2655
      TabIndex        =   9
      Top             =   -405
      Width           =   2715
   End
End
Attribute VB_Name = "VC01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String

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

Private Sub cmdDEL_Click()
If Text5 = "0" Then
    SDel = "Delete * From VC01 where INDUK = '" + Trim(KB) + "'"
    Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
    RDEl.Close
    Set RDEl = Nothing
    Unload Me
    VC01.Show 1
Else
    MsgBox "DEPOSIT UNIT MASIH TERISI " + Format(Trim(Text5), "##,###"), vbCritical, "KONFIRMASI"
    Unload Me
    VC01.Show 1
End If
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

SSave = "Select * From VC01 where INDUK = '" + Trim(KB) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.Edit
    RSave("INDUK") = Trim(Text1)
    RSave("NAMA_INDUK") = Trim(Text2)
RSave.Update
RSave.Close
Set RSave = Nothing

MsgBox "UPDATE KODE INDUK", vbCritical, "KONFIRMASI"

Unload Me
VC01.Show 1

End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

If Text1 = "" Then Exit Sub
SQLTiket = "Select * from VC01 where INDUK = '" + Trim(Text1) + "'"
Set RSLTiket = RDCO.OpenResultset(SQLTiket, rdOpenDynamic, rdConcurRowVer)
If RSLTiket.RowCount <> 0 Then
    MsgBox "KODE SUDAH ADA", vbCritical, "KONFIRMASI"
    Text1 = ""
    Text1.SetFocus
Else
    Call Simpan
End If
RSLTiket.Close
Set RSLTiket = Nothing

MsgBox "UPDATE KODE INDUK", vbCritical, "KONFIRMASI"

Unload Me
VC01.Show 1
End Sub

Private Sub Simpan()
SSave = "Select * From VC01"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("INDUK") = Trim(Text1)
    RSave("NAMA_INDUK") = Trim(Text2)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Call SiapkanGrid
Call IsiGrid
cmdOK.Visible = True
cmdEDIT.Visible = False
cmdDEL.Visible = False

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 2
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE INDUK": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2000: .Text = "NAMA INDUK": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From VC01 order by INDUK"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari("INDUK"): .CellAlignment = 4
              .Col = 1: .Text = RCari("NAMA_INDUK")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid_dblClick()
grid.Col = 0
KB = ""
Clipboard.SetText (grid.Text)
KB = grid.Text

cmdOK.Visible = False
cmdEDIT.Visible = True
cmdDEL.Visible = True

Call IsiText

End Sub

Private Sub IsiText()
SCari2 = "Select * From VC01 where INDUK = '" + Trim(KB) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Text1 = RCari2("INDUK")
    Text2 = RCari2("NAMA_INDUK")
    Text5 = RCari2("SALDO")
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
Call CekData
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
    Text2 = Format(Text2, ">")
End Sub

Private Sub CekData()
If Text1.Text = "" Then Exit Sub

SCari = "Select * From VC01 where INDUK = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " KODE INDUK SUDAH TERDAFTAR", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub


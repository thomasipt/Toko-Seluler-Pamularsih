VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form VC02 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE ELECTRIC"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Left            =   2745
      TabIndex        =   4
      Text            =   "4"
      Top             =   2025
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   2745
      TabIndex        =   3
      Text            =   "3"
      Top             =   1575
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2745
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   90
      Width           =   1770
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
      Left            =   6420
      TabIndex        =   7
      Top             =   2070
      Width           =   1890
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
      Left            =   6420
      TabIndex        =   8
      Top             =   3713
      Width           =   1890
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
      Left            =   2745
      TabIndex        =   2
      Text            =   "2"
      Top             =   1200
      Width           =   3330
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
      Left            =   2745
      TabIndex        =   1
      Text            =   "1"
      Top             =   825
      Width           =   1275
   End
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
      Left            =   2325
      TabIndex        =   9
      Text            =   "5"
      Top             =   5130
      Width           =   2025
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   5580
      Top             =   5355
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   270
      OleObjectBlob   =   "VC02.frx":0000
      Top             =   6525
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   180
      OleObjectBlob   =   "VC02.frx":0234
      TabIndex        =   10
      Top             =   900
      Width           =   2385
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   180
      OleObjectBlob   =   "VC02.frx":02AC
      TabIndex        =   11
      Top             =   1275
      Width           =   2385
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2130
      Left            =   135
      TabIndex        =   12
      Top             =   2505
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   3757
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   165
      Left            =   180
      OleObjectBlob   =   "VC02.frx":0312
      TabIndex        =   14
      Top             =   165
      Width           =   2385
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   225
      Left            =   180
      OleObjectBlob   =   "VC02.frx":0394
      TabIndex        =   15
      Top             =   1650
      Width           =   2385
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   300
      Left            =   2745
      OleObjectBlob   =   "VC02.frx":03FC
      TabIndex        =   16
      Top             =   450
      Width           =   3330
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   225
      Left            =   180
      OleObjectBlob   =   "VC02.frx":046E
      TabIndex        =   17
      Top             =   2100
      Width           =   2385
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
      Left            =   6420
      TabIndex        =   5
      Top             =   450
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
      Left            =   6420
      TabIndex        =   6
      Top             =   443
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   5550
      Left            =   6195
      ScaleHeight     =   5490
      ScaleWidth      =   2655
      TabIndex        =   13
      Top             =   -240
      Width           =   2715
   End
End
Attribute VB_Name = "VC02"
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
SDel = "Delete * From VC02 where KODE = '" + Trim(KB) + "'"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

MsgBox "UPDATE KODE ELECTRIC", vbCritical, "KONFIRMASI"

Unload Me
VC02.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

SSave = "Select * From VC02 where KODE = '" + Trim(KB) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.Edit
    RSave("INDUK") = Trim(Combo1)
    RSave("KODE") = Trim(Text1)
    RSave("NAMA") = Trim(Text2)
    RSave("HARGA") = CCur(Text3)
    RSave("UNIT") = CCur(Text4)
RSave.Update
RSave.Close
Set RSave = Nothing

MsgBox "UPDATE KODE ELECTRIC", vbCritical, "KONFIRMASI"

Unload Me
VC02.Show 1

End Sub

Private Sub cmdOK_Click()
If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

If Text1 = "" Then Exit Sub
SQLTiket = "Select * from VC02 where KODE = '" + Trim(Text1) + "'"
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

MsgBox "UPDATE KODE ELECTRIC", vbCritical, "KONFIRMASI"

Unload Me
VC02.Show 1
End Sub

Private Sub Simpan()
SSave = "Select * From VC02"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("INDUK") = Trim(Combo1)
    RSave("KODE") = Trim(Text1)
    RSave("NAMA") = Trim(Text2)
    RSave("HARGA") = CCur(Text3)
    RSave("UNIT") = CCur(Text4)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
SCari2 = "Select * From VC01 where INDUK = '" + Combo1 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel5 = RCari2("NAMA_INDUK")
Else
    MsgBox "KODE INDUK BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo1.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
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

Combo1 = ""

SSPL = "Select INDUK From VC01 order by INDUK"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
RSPL.MoveFirst
Do While Not RSPL.EOF
    Combo1.AddItem RSPL("INDUK")
RSPL.MoveNext
Loop
RSPL.Close
Set RSPL = Nothing

SkinLabel5 = ""

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 5
    .Col = 0: .ColWidth(0) = 1000: .Text = "INDUK": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1250: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1250: .Text = "HARGA": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "UNIT": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From VC02 order by INDUK"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari("INDUK"): .CellAlignment = 4
              .Col = 1: .Text = RCari("KODE"): .CellAlignment = 4
              .Col = 2: .Text = RCari("NAMA")
              .Col = 3: .Text = Format(RCari("HARGA"), "##,##.00")
              .Col = 4: .Text = Format(RCari("UNIT"), "##,##")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid_dblClick()
grid.Col = 1
KB = ""
Clipboard.SetText (grid.Text)
KB = grid.Text

cmdOK.Visible = False
cmdEDIT.Visible = True
cmdDEL.Visible = True

Call IsiText

End Sub

Private Sub IsiText()
SCari2 = "Select * From VC02 where KODE = '" + Trim(KB) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Combo1 = RCari2("INDUK")
    Text1 = RCari2("KODE")
    Text2 = RCari2("NAMA")
    Text3 = Format(RCari2("HARGA"), "##,###.00")
    Text4 = Format(RCari2("UNIT"), "##,###")
    Text5 = RCari2("KODE")
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

SCari = "Select * From VC02 where KODE = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " KODE ELECTRIC SUDAH TERDAFTAR", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        Text3 = Format(Text3, "##,###.00")
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        Text4 = Format(Text4, "##,###")
    End If
End Sub

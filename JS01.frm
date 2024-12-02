VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form JS01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE JASA SERVICE"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4965
      TabIndex        =   7
      Top             =   2355
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
      Left            =   2632
      TabIndex        =   6
      Top             =   2355
      Width           =   1890
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1935
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1620
      Width           =   2850
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1935
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1125
      Width           =   2850
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1935
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   630
      Width           =   5190
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1935
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   135
      Width           =   1995
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5310
      OleObjectBlob   =   "JS01.frx":0000
      Top             =   45
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "JS01.frx":0234
      TabIndex        =   8
      Top             =   225
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "JS01.frx":029A
      TabIndex        =   9
      Top             =   1215
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "JS01.frx":0304
      TabIndex        =   10
      Top             =   720
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "JS01.frx":036A
      TabIndex        =   11
      Top             =   1710
      Width           =   1560
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2220
      Left            =   67
      TabIndex        =   12
      Top             =   3195
      Width           =   7020
      _ExtentX        =   12383
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
      Left            =   300
      TabIndex        =   4
      Top             =   2355
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   -315
      ScaleHeight     =   795
      ScaleWidth      =   8145
      TabIndex        =   13
      Top             =   2205
      Width           =   8205
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
      Left            =   300
      TabIndex        =   5
      Top             =   2355
      Width           =   1890
   End
End
Attribute VB_Name = "JS01"
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
SDel = "Delete * From JS01 where Kode = '" + Trim(KB) + "'"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing
Unload Me
JS01.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

SSave = "Select * From JS01 where Kode = '" + Trim(KB) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.Edit
    RSave("Kode") = Trim(Text1)
    RSave("Nama") = Trim(Text2)
    RSave("Komisi") = CCur(Text3)
    RSave("Biaya") = CCur(Text4)
RSave.Update
RSave.Close
Set RSave = Nothing

Unload Me
JS01.Show 1

End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

If Text1 = "" Then Exit Sub
SQLTiket = "Select * from JS01 where Kode = '" + Trim(Text1) + "'"
Set RSLTiket = RDCO.OpenResultset(SQLTiket, rdOpenDynamic, rdConcurRowVer)
If RSLTiket.RowCount <> 0 Then
    MsgBox "KODE JASA SUDAH ADA", vbCritical, "KONFIRMASI"
    Text1 = ""
    Text1.SetFocus
Else
    Call Simpan
End If
RSLTiket.Close
Set RSLTiket = Nothing

Unload Me
JS01.Show 1
End Sub

Private Sub Simpan()
SSave = "Select * From JS01"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("Kode") = Trim(Text1)
        RSave("Nama") = Trim(Text2)
        RSave("Komisi") = CCur(Text3)
        RSave("Biaya") = CCur(Text4)
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

ClearTextBoxes JS01
Call SiapkanGrid
Call IsiGrid
cmdOK.Visible = True
cmdEDIT.Visible = False
cmdDEL.Visible = False
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 4
    .Col = 0: .ColWidth(0) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "KOMISI": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "BIAYA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From JS01 order by KODE"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari("Kode"): .CellAlignment = 4
              .Col = 1: .Text = RCari("Nama")
              .Col = 2: .Text = Format(RCari("Komisi"), "##,###.00")
              .Col = 3: .Text = Format(RCari("Biaya"), "##,###.00")
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
SCari2 = "Select * From JS01 where Kode = '" + Trim(KB) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Text1 = RCari2("Kode")
    Text2 = RCari2("Nama")
    Text3 = Format(RCari2("Komisi"), "##,###.00")
    Text4 = Format(RCari2("Biaya"), "##,###.00")
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, "##,###.00")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
Text4 = Format(Text4, "##,###.00")
End Sub

Private Sub CekData()
If Text1.Text = "" Then Exit Sub

SCari = "Select * From JS01 where Kode = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " KODE JASA SUDAH TERDAFTAR", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    Else
       Text2.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

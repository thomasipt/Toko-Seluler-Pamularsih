VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form B003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE BARANG"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1350
      Width           =   2580
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   2400
      TabIndex        =   7
      Top             =   2655
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   635
      _Version        =   393216
      Format          =   58458113
      CurrentDate     =   39620
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   488
      Width           =   2580
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   885
      Width           =   2580
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
      Left            =   7251
      TabIndex        =   9
      Top             =   3285
      Width           =   1890
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
      Left            =   242
      TabIndex        =   8
      Top             =   3285
      Width           =   1890
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2220
      Width           =   2580
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1800
      Width           =   2580
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "Combo4"
      Top             =   90
      Width           =   2580
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5100
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   2220
      Width           =   1680
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   390
      OleObjectBlob   =   "B003.frx":0000
      Top             =   9000
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "B003.frx":0234
      TabIndex        =   10
      Top             =   960
      Width           =   2115
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "B003.frx":02B6
      TabIndex        =   11
      Top             =   165
      Width           =   2565
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "B003.frx":0336
      TabIndex        =   12
      Top             =   2295
      Width           =   2205
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "B003.frx":03A8
      TabIndex        =   13
      Top             =   1830
      Width           =   2190
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "B003.frx":041A
      TabIndex        =   14
      Top             =   570
      Width           =   1980
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   240
      Left            =   6900
      OleObjectBlob   =   "B003.frx":048E
      TabIndex        =   15
      Top             =   2280
      Width           =   330
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2445
      Left            =   60
      TabIndex        =   16
      Top             =   4125
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      AllowUserResizing=   3
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   165
      Left            =   5130
      OleObjectBlob   =   "B003.frx":04EE
      TabIndex        =   17
      Top             =   165
      Width           =   4170
   End
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   -360
      ScaleHeight     =   765
      ScaleWidth      =   10125
      TabIndex        =   18
      Top             =   3150
      Width           =   10185
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   165
      Left            =   90
      OleObjectBlob   =   "B003.frx":054C
      TabIndex        =   19
      Top             =   2760
      Width           =   2205
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "B003.frx":05B8
      TabIndex        =   20
      Top             =   1455
      Width           =   2205
   End
End
Attribute VB_Name = "B003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RSave, RSave2, REdit, RKTG, RKTG2, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SQLUser, SSave, SSave2, SEdit, SKTG, SKTG2, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub cmdCLOSE_Click()
cepat = 1000
While Top - Height < Screen.Height
    DoEvents
    Top = Top + cepat
Wend
Hide
Unload Me
End Sub

Private Sub cmdHAPUS_Click()
MsgBox "ANDA AKAN MENGHAPUS DATA OBAT", vbCritical, "KONFIRMASI"

End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text5 = "" Or Text4 = "" Or Text6 = "" Or Combo4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Combo4.SetFocus
Exit Sub
End If

If Text1 = "" Then Exit Sub
SQLTiket = "Select * from B003 where KodeBR = '" + Trim(Text1) + "'"
Set RSLTiket = RDCO.OpenResultset(SQLTiket, rdOpenDynamic, rdConcurRowVer)
If RSLTiket.RowCount <> 0 Then
    MsgBox "KODE BARANG SUDAH ADA", vbCritical, "KONFIRMASI"
    Text1 = ""
    Text1.SetFocus
Else
    Call Simpan
    Call KOSONG
    Call SiapkanGrid
    Call IsiGrid
End If
RSLTiket.Close
Set RSLTiket = Nothing
Text1.SetFocus

Unload Me
B003.Show 1
End Sub

Private Sub Simpan()
SSave = "Select * From B003"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("KodeInd") = Combo4
        RSave("KodeBR") = Trim(Text1)
        RSave("NamaBR") = Trim(Text2)
        RSave("JAkhir") = 0
        RSave("HBeli") = CCur(Text5)
        RSave("HJual") = CCur(Text4)
        RSave("Persen") = CCur(Text6)
        RSave("Satuan") = Combo1
        RSave("Tanggal") = DTPicker1
        RSave("Status") = 1
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hwnd
Call KOSONG

Call SiapkanGrid
Call IsiGrid
grid.Refresh

SSPL = "Select Kode From B001 order by KODE"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
RSPL.MoveFirst
Do While Not RSPL.EOF
    Combo4.AddItem RSPL("Kode")
RSPL.MoveNext
Loop
RSPL.Close
Set RSPL = Nothing
Combo4.ListIndex = 0

DTPicker1 = Date

End Sub

Private Sub KOSONG()

ClearTextBoxes Me

Combo1 = ""
'Combo4 = ""

SkinLabel6 = ""

End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 1000: .Text = "INDUK": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2700: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 750: .Text = "JUMLAH": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1500: .Text = "JUAL Rp.": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "GARANSI": .CellAlignment = 4
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
              .Col = 0: .Text = RKTG("KodeInd"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 2: .Text = RKTG("NamaBR")
              .Col = 3: .Text = RKTG("JAkhir"): .CellAlignment = 4
              .Col = 4: .Text = Format(RKTG("HJual"), "##,###.00")
              .Col = 5: .Text = RKTG("Tanggal"): .CellAlignment = 4
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
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
    Text3 = Format(Text3, ">")
End Sub

Private Sub text4_gotfocus()
Text4 = ""
Text6 = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        Text4 = Format(Text4, "##,###.00")
        Text6 = (CCur(Text4) - CCur(Text5)) / (CCur(Text5) / 100)
    End If
End Sub

Private Sub Text5_GotFocus()
Text5 = ""
Text4 = ""
Text6 = ""
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        Text5 = Format(Text5, "##,###.00")
    End If
End Sub

Private Sub text6_GotFocus()
Text6 = ""
Text4 = ""
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        Z = Format(CCur(Text5) + ((CCur(Text5) / 100) * CCur(Text6)), "##,###")
        ZZ = Format(Z, "##,###.00")
        A = Right(Z, 2)
        If A > 0 And A < 100 Then
            MsgBox " HARGA JUAL Rp. " + ZZ + " SISTEM AKAN MEMBULATKAN Rp 100", vbCritical, "KONFIRMASI"
            B = 100 - A
            Text4 = Format(Z + B, "##,###.00")
            Combo1.SetFocus
            Exit Sub
        Else
            Text4 = Format(Z, "##,###.00")
            Combo1.SetFocus
        End If
    End If
End Sub

Private Sub CekData()
If Text1.Text = "" Then Exit Sub

SCari = "Select * From B003 where KodeBR = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " KODE BARANG SUDAH TERDAFTAR", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    Else
       Text2.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
ComboSearch Combo4, KeyAscii
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo4_LostFocus()
If Combo4 = "" Then Exit Sub
SCari2 = "Select * From B001 where Kode = '" + Combo4 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel6 = RCari2("Nama")
Else
    MsgBox "KODE INDUK BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo4.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub grid_dblClick()
grid.Col = 1
BR = ""
Clipboard.SetText (grid.Text)
BR = grid.Text

If BR = "" Then Exit Sub

Unload Me
B003EDIT.Show 1
End Sub





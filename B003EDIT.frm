VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form B003EDIT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT BARANG"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3847
      TabIndex        =   17
      Top             =   2265
      Width           =   1890
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5257
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   1605
      Width           =   1680
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2557
      TabIndex        =   0
      Text            =   "Combo4"
      Top             =   105
      Width           =   1440
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2557
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   1230
      Width           =   2580
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2557
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1605
      Width           =   2580
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2557
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   855
      Width           =   2580
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2557
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2580
   End
   Begin VB.TextBox Text3 
      Height          =   765
      Left            =   1305
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   5025
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3390
      OleObjectBlob   =   "B003EDIT.frx":0000
      Top             =   5145
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   7350
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   210
      Left            =   345
      OleObjectBlob   =   "B003EDIT.frx":0234
      TabIndex        =   9
      Top             =   930
      Width           =   1980
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   315
      Left            =   345
      OleObjectBlob   =   "B003EDIT.frx":02A8
      TabIndex        =   10
      Top             =   180
      Width           =   2565
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   210
      Left            =   345
      OleObjectBlob   =   "B003EDIT.frx":0328
      TabIndex        =   11
      Top             =   1680
      Width           =   2205
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   210
      Left            =   345
      OleObjectBlob   =   "B003EDIT.frx":039A
      TabIndex        =   12
      Top             =   1305
      Width           =   2190
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   210
      Left            =   345
      OleObjectBlob   =   "B003EDIT.frx":040C
      TabIndex        =   13
      Top             =   555
      Width           =   1980
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   240
      Left            =   7050
      OleObjectBlob   =   "B003EDIT.frx":0480
      TabIndex        =   14
      Top             =   1605
      Width           =   330
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   165
      Left            =   4200
      OleObjectBlob   =   "B003EDIT.frx":04E0
      TabIndex        =   15
      Top             =   180
      Width           =   5115
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
      Left            =   7312
      TabIndex        =   7
      Top             =   2280
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
      Left            =   303
      TabIndex        =   6
      Top             =   2265
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   960
      Left            =   -308
      ScaleHeight     =   900
      ScaleWidth      =   10140
      TabIndex        =   16
      Top             =   2130
      Width           =   10200
   End
End
Attribute VB_Name = "B003EDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RSave, RSave2, REdit, RKTG, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SQLUser, SSave, SSave2, SEdit, SKTG, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub IPT()
Text1 = BR
Text3 = BR

SCari3 = "Select * From B003 where KodeBR = '" + Trim(Text1) + "'"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurReadOnly)
If RCari3.RowCount <> 0 Then
    Combo4 = RCari3("KodeInd")
    Text1 = RCari3("KodeBR")
    Text2 = RCari3("namaBR")
    Text5 = RCari3("HBeli")
    Text4 = RCari3("HJual")
    Text6 = RCari3("Persen")
    Combo1 = RCari3("Satuan")
    DTPicker1 = RCari3("Tanggal")
End If
RCari3.Close
Set RCari3 = Nothing
End Sub

Private Sub cmdDEL_Click()
SDel = "Delete * From B003 where KodeBR = '" + Trim(BR) + "'"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing
MsgBox "DATABASE TELAH DI UPDATE", vbCritical, "KONFIRMASI"
Unload Me
B003.Show 1
End Sub

Private Sub cmdEDIT_Click()
Dim tanya
tanya = MsgBox("YAKIN AKAN MERUBAH DATA", vbOKCancel, "KONFIRMASI")
If tanya = vbOK Then
    SCari4 = "Select * From B003 where KodeBR = '" + Trim(Text3) + "'"
    Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenDynamic, rdConcurRowVer)
    RCari4.Edit
        RCari4("KodeInd") = Combo4
        RCari4("KodeBR") = Text1
        RCari4("namaBR") = Text2
        RCari4("HBeli") = CCur(Text5)
        RCari4("HJual") = CCur(Text4)
        RCari4("Persen") = CCur(Text6)
        RCari4("Status") = 1
    RCari4.Update
    RCari4.Close
    MsgBox "DATABASE TELAH DI UPDATE", vbCritical, "KONFIRMASI"
End If
Unload Me
B003.Show 1
End Sub

Private Sub cmdCLOSE_Click()
Unload Me
B003.Show 1
End Sub

Private Sub Combo4_LostFocus()
Call Cari2
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

SSPL = "Select Kode From B001"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
RSPL.MoveFirst
Do While Not RSPL.EOF
    Combo4.AddItem RSPL("Kode")
RSPL.MoveNext
Loop
RSPL.Close
Set RSPL = Nothing

Call IPT
Call Cari2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
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
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
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
        cmdEDIT.SetFocus
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

Private Sub Cari2()
SCari2 = "Select * From B001 where Kode = '" + Combo4 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel6 = RCari2("Nama")
Else
    MsgBox "KODE INDUK BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


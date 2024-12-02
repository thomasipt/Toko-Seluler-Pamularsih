VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form VC03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENJUALAN PULSA ELETRONIK"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2010
      TabIndex        =   5
      Text            =   "2"
      Top             =   2835
      Width           =   5010
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2010
      TabIndex        =   4
      Text            =   "1"
      Top             =   2430
      Width           =   1545
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2025
      TabIndex        =   0
      Text            =   "6"
      Top             =   90
      Width           =   1545
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2010
      TabIndex        =   3
      Text            =   "5"
      Top             =   2040
      Width           =   5010
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2010
      TabIndex        =   2
      Text            =   "4"
      Top             =   1650
      Width           =   5010
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
      Left            =   5033
      TabIndex        =   7
      Top             =   3555
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2025
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   540
      Width           =   2040
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4725
      OleObjectBlob   =   "VC03.frx":0000
      Top             =   3600
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":0234
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   2070
      OleObjectBlob   =   "VC03.frx":029A
      TabIndex        =   9
      Top             =   990
      Width           =   1965
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
      Left            =   248
      TabIndex        =   6
      Top             =   3555
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   -315
      ScaleHeight     =   1515
      ScaleWidth      =   8145
      TabIndex        =   10
      Top             =   3345
      Width           =   8205
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   225
      Left            =   4095
      OleObjectBlob   =   "VC03.frx":0300
      TabIndex        =   11
      Top             =   990
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":0368
      TabIndex        =   12
      Top             =   1695
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":03D6
      TabIndex        =   13
      Top             =   2100
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":0440
      TabIndex        =   14
      Top             =   150
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":04A6
      TabIndex        =   15
      Top             =   2490
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":0510
      TabIndex        =   16
      Top             =   2895
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   225
      Left            =   2085
      OleObjectBlob   =   "VC03.frx":0582
      TabIndex        =   17
      Top             =   1305
      Width           =   1965
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   2010
      TabIndex        =   18
      Top             =   855
      Width           =   5010
   End
End
Attribute VB_Name = "VC03"
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

Private Sub cmdOK_Click()
If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "MASIH ADA DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Dim tanya
tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI PENJUALAN", vbSystemModal, "KONFIRMASI")
If tanya = vbOK Then
    Call Simpan
End If

Unload Me
VC03.Show 1
End Sub

Private Sub Simpan()
SSave = "Select * From VC01 where INDUK = '" + Trim(KB) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
If RSave("SALDO") < CCur(SkinLabel6) Then
    MsgBox "SALDO SUDAH HABIS / TIDAK CUKUP", vbCritical, "TRANSAKSI GAGAL"
    Combo1.SetFocus
    Exit Sub
Else
    RSave.Edit
        RSave("CREDIT") = RSave("CREDIT") + CCur(SkinLabel6)
        RSave("SALDO") = RSave("SALDO") - CCur(SkinLabel6)
        RSave("TANGGAL") = Date
    
        SSave2 = "Select * From VC03"
        Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
        RSave2.AddNew
            RSave2("NOTA") = Trim(Text6)
            RSave2("INDUK") = Trim(KB)
            RSave2("KODE") = Trim(Combo1)
            RSave2("CREDIT") = CCur(SkinLabel6)
            RSave2("SALDO") = RSave("SALDO")
            RSave2("CUSTOMER") = Trim(Text4)
            RSave2("ALAMAT") = Trim(Text5)
            RSave2("HP") = Trim(Text1)
            RSave2("KETERANGAN") = Trim(Text2)
            RSave2("TANGGAL") = Date
        RSave2.Update
        RSave2.Close
        Set RSave2 = Nothing
        
    RSave.Update
    RSave.Close
    Set RSave = Nothing
End If
End Sub

Private Sub EditVC01()
Dim Stock As String
Dim HBeli As String

SSave2 = "Select * From VC01 where Kode = '" + Trim(Combo1) + "'"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    Stock = RSave2("Stokbel")
    HBeli = RSave2("Satuan")
RSave2.Edit
    RSave2("Stokbel") = CCur(Stock) - (CCur(Text1) * CCur(HBeli))
    RSave2("Jumlah") = CCur(Pusing) - CCur(Text1)
RSave2.Update
RSave2.Close
Set RSave2 = Nothing
End Sub

Private Sub EditNoBukti()
SCari9 = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RCari9 = RDCO.OpenResultset(SCari9, rdOpenKeyset, rdConcurRowVer)
    TOGEL = RCari9("NoJual")
    RCari9.Edit
        RCari9("NoJual") = TOGEL + 1
    RCari9.Update
    RCari9.Close
    Set RCari9 = Nothing
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
SCari2 = "Select * From VC02 where KODE = '" + Combo1 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel3 = RCari2("NAMA")
    SkinLabel7 = Format(RCari2("HARGA"), "##,###.00")
    SkinLabel6 = Format(RCari2("UNIT"), "##,###")
    KB = RCari2("INDUK")
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


Combo1 = ""

SSPL = "Select KODE From VC02 order by KODE"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
RSPL.MoveFirst
Do While Not RSPL.EOF
    Combo1.AddItem RSPL("KODE")
RSPL.MoveNext
Loop
RSPL.Close
Set RSPL = Nothing

SkinLabel3 = ""
SkinLabel7 = ""
SkinLabel6 = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
    Text2 = Format(Text2, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
    Text4 = Format(Text4, ">")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
    Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
Call CekData
End Sub

Private Sub CekData()
If Text6.Text = "" Then Exit Sub

SCari = "Select * From VC03 where NOTA = '" + Trim(Text6) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox "NO NOTA TELAH DIGUNAKAN", vbCritical, "KONFIRMASI"
        Text6 = ""
        Text6.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

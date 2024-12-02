VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form VC00 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPOSIT PULSA"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
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
      Left            =   142
      TabIndex        =   2
      Top             =   1650
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
      Left            =   4222
      TabIndex        =   3
      Top             =   1650
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1770
   End
   Begin VB.TextBox Text1 
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
      Left            =   1890
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   945
      Width           =   2130
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   135
      OleObjectBlob   =   "VC00.frx":0000
      TabIndex        =   4
      Top             =   165
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   135
      OleObjectBlob   =   "VC00.frx":0076
      TabIndex        =   5
      Top             =   990
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   300
      Left            =   1890
      OleObjectBlob   =   "VC00.frx":00EA
      TabIndex        =   6
      Top             =   540
      Width           =   3330
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5490
      OleObjectBlob   =   "VC00.frx":015C
      Top             =   810
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   -90
      ScaleHeight     =   1515
      ScaleWidth      =   8145
      TabIndex        =   7
      Top             =   1485
      Width           =   8205
   End
End
Attribute VB_Name = "VC00"
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

Private Sub cmdOK_Click()
If Combo1 = "" Or Text1 = "" Then
    MsgBox "MASIH ADA DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Dim tanya
tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI DEPOSIT PULSA", vbSystemModal, "KONFIRMASI")
If tanya = vbOK Then
    Call Simpan
End If

Unload Me
VC00.Show 1

End Sub

Private Sub Simpan()
SSave = "Select * From VC01 where INDUK = '" + Combo1 + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.Edit
    RSave("DEBET") = RSave("DEBET") + CCur(Text1)
    RSave("SALDO") = RSave("SALDO") + CCur(Text1)
    RSave("TANGGAL") = Date

    SSave2 = "Select * From VC03"
    Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
    RSave2.AddNew
        RSave2("NOTA") = "DEPOSIT"
        RSave2("INDUK") = Trim(Combo1)
        RSave2("DEBET") = CCur(Text1)
        RSave2("SALDO") = RSave("SALDO")
        RSave2("TANGGAL") = Date
    RSave2.Update
    RSave2.Close
    Set RSave2 = Nothing
    
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        Text1 = Format(Text1, "##,###")
    End If
End Sub

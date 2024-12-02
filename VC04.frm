VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form VC04 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUTASI ELECTRIC"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   202
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1575
      Width           =   1770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MUTASI"
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
      Left            =   202
      TabIndex        =   3
      Top             =   2415
      Width           =   4185
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   6510
      Top             =   1995
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1349
      TabIndex        =   0
      Top             =   3495
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   1230
      Left            =   -180
      ScaleHeight     =   1170
      ScaleWidth      =   5985
      TabIndex        =   2
      Top             =   3255
      Width           =   6045
   End
   Begin VB.CommandButton cmdCTK 
      Caption         =   "SALDO VOUCHER"
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
      Left            =   202
      TabIndex        =   1
      Top             =   210
      Width           =   4185
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10260
      OleObjectBlob   =   "VC04.frx":0000
      Top             =   3285
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   300
      Left            =   202
      OleObjectBlob   =   "VC04.frx":0234
      TabIndex        =   4
      Top             =   1995
      Width           =   4185
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   420
      Left            =   2092
      TabIndex        =   5
      Top             =   1515
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   22216705
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   165
      Left            =   2632
      OleObjectBlob   =   "VC04.frx":02A6
      TabIndex        =   6
      Top             =   1230
      Width           =   1230
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   165
      Left            =   472
      OleObjectBlob   =   "VC04.frx":0312
      TabIndex        =   8
      Top             =   1230
      Width           =   1230
   End
   Begin VB.PictureBox Picture2 
      Height          =   1230
      Left            =   -735
      ScaleHeight     =   1170
      ScaleWidth      =   5985
      TabIndex        =   9
      Top             =   -210
      Width           =   6045
   End
End
Attribute VB_Name = "VC04"
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

Private Sub cmdCTK_Click()
crpt.ReportFileName = "C:\WINDOWS\ReportSELULER\LapVC.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
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

Private Sub Command2_Click()
If Combo1 = "" Then Exit Sub

SDel = "Delete * From VC03A"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

Dim tanya
tanya = MsgBox("CETAK MUTASI", vbOKCancel, "KONFIRMASI")
    If tanya = vbOK Then
        Call Simpan
        MsgBox "SELESAI", vbCritical, "KONFIRMASI"
        Call CETAK
    Else
        Exit Sub
    End If
End Sub

Private Sub CETAK()
crpt.ReportFileName = "C:\WINDOWS\ReportSELULER\VC03A.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Simpan()
SSave = "Select * From VC03 where INDUK = '" + Combo1 + "' and TANGGAL like '" + Trim(DTPicker3) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.MoveFirst
Do While Not RSave.EOF

    SSave2 = "Select * From VC03A"
    Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
    RSave2.AddNew
        RSave2("DEBET") = RSave("debet")
        RSave2("CREDIT") = RSave("credit")
        RSave2("SALDO") = RSave("saldo")
        On Error Resume Next
            RSave2("HP") = RSave("hp")
            RSave2("CUSTOMER") = RSave("customer")
            RSave2("ALAMAT") = RSave("alamat")
            RSave2("KETERANGAN") = RSave("keterangan")
    RSave2.Update
    RSave2.Close
    Set RSave2 = Nothing

RSave.MoveNext
Loop
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

DTPicker1 = Date
DTPicker2 = Date + 1
DTPicker3 = Date

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



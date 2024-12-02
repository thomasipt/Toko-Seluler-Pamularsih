VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form TGLFAK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TANGGAL FAKTUR"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3735
      OleObjectBlob   =   "TGLFAK.frx":0000
      Top             =   555
   End
   Begin VB.CommandButton cmdCTK 
      Caption         =   "CETAK"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2220
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
      Left            =   2265
      TabIndex        =   2
      Top             =   2220
      Width           =   1890
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   1035
      TabIndex        =   0
      Top             =   458
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
      Format          =   22085633
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   127
      OleObjectBlob   =   "TGLFAK.frx":0234
      TabIndex        =   1
      Top             =   128
      Width           =   4140
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   120
      OleObjectBlob   =   "TGLFAK.frx":02BA
      TabIndex        =   4
      Top             =   1125
      Width           =   4140
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   -135
      ScaleHeight     =   855
      ScaleWidth      =   4695
      TabIndex        =   5
      Top             =   2100
      Width           =   4755
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   420
      Left            =   1035
      TabIndex        =   6
      Top             =   1455
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
      Format          =   22085633
      CurrentDate     =   39286
      MinDate         =   39083
   End
End
Attribute VB_Name = "TGLFAK"
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

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdCTK_Click()
TglFuck = DTPicker1
TglFuck2 = DTPicker2
Indikator = 1
Call Seleksi

Dim tanya
tanya = MsgBox("CETAK TANGGAL", vbOKCancel, "KONFIRMASI")
    If tanya = vbOK Then
        Unload Me
    Else
        Exit Sub
    End If
End Sub

Private Sub Seleksi()
SDel = "Delete * From B005CTK"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

SCari1 = "Select * From B005 where Tgl_Fak >= datevalue('" + Trim(DTPicker1) + "') and Tgl_Fak <= datevalue('" + Trim(DTPicker2) + "')"
Set RCari1 = RDCO.OpenResultset(SCari1, rdOpenDynamic, rdConcurRowVer)
RCari1.MoveFirst
Do While Not RCari1.EOF

    SCari2 = "Select * From B005CTK"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
        RCari2("Status") = 1
        RCari2("KODE_TRANS") = RCari1("Kode_Trans")
        RCari2("KODE_JNS") = RCari1("KODE_JNS")
        RCari2("NAMA_JNS") = RCari1("NAMA_JNS")
        RCari2("NO_FAKTUR") = RCari1("NO_FAKTUR")
        RCari2("NO_BUKTI") = RCari1("NO_BUKTI")
        RCari2("KETERANGAN") = RCari1("KETERANGAN")
        RCari2("JML_DBT") = RCari1("JML_DBT")
        RCari2("JML_CRD") = RCari1("JML_CRD")
        RCari2("JML_AKHIR") = RCari1("JML_AKHIR")
        RCari2("MUTASI_DBT") = RCari1("MUTASI_DBT")
        RCari2("MUTASI_CRT") = RCari1("MUTASI_CRT")
        RCari2("SALDO_AKHIR") = RCari1("SALDO_AKHIR")
        RCari2("H_POKOK") = RCari1("H_POKOK")
        RCari2("NOMDISC") = RCari1("NOMDISC")
        RCari2("SPCDISC") = RCari1("SPCDISC")
        RCari2("LABA") = RCari1("LABA")
        RCari2("KAS") = RCari1("KAS")
        RCari2("TGL_S") = RCari1("TGL_S")
        RCari2("TGL_FAK") = RCari1("TGL_FAK")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing

RCari1.MoveNext
Loop
RCari1.Close
Set RCari1 = Nothing

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

Indikator = 0
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

DTPicker1 = Date
DTPicker2 = Date + 1
End Sub

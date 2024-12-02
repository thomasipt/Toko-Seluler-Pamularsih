VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form JS001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR SERVICE"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CETAK SELESAI"
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
      Left            =   2767
      TabIndex        =   3
      Top             =   855
      Width           =   2250
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Text            =   "2"
      Top             =   90
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1410
      TabIndex        =   0
      Text            =   "1"
      Top             =   90
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK BELUM SELESAI"
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
      Left            =   187
      TabIndex        =   2
      Top             =   855
      Width           =   2250
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
      Left            =   172
      TabIndex        =   4
      Top             =   1530
      Width           =   4860
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   5310
      Top             =   675
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5445
      OleObjectBlob   =   "JS001.frx":0000
      Top             =   1350
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   540
      OleObjectBlob   =   "JS001.frx":0234
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   2430
      OleObjectBlob   =   "JS001.frx":029C
      TabIndex        =   6
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Picture1 
      Height          =   4065
      Left            =   -360
      ScaleHeight     =   4005
      ScaleWidth      =   5895
      TabIndex        =   7
      Top             =   585
      Width           =   5955
   End
End
Attribute VB_Name = "JS001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String
Dim Hari, Rumus As String

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

Private D, M, Y

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Command1_Click()

If Text1 = "" Or Text2 = "" Then Exit Sub

Call Seleksi

Y = Text2
M = Text1
D = 1
D1 = 31

Dim tanya
tanya = MsgBox("CETAK LAPORAN", vbOKCancel, "KONFIRMASI")
    If tanya = vbOK Then
        crpt.ReportFileName = "c:\windows\ReportSELULER\LapService.rpt"
        crpt.WindowState = crptMaximized
        crpt.WindowMaxButton = True
        crpt.WindowMinButton = True
        crpt.Action = 1
        crpt.Reset
    Else
        Exit Sub
    End If

End Sub

Private Sub Seleksi()
SDel = "Delete * From JS02CTK"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

SCari1 = "Select * From JS02 where BULAN = '" + Trim(Text1) + "' and TAHUN = '" + Trim(Text2) + "'and AMBIL = 'BELUM SELESAI'"
Set RCari1 = RDCO.OpenResultset(SCari1, rdOpenDynamic, rdConcurRowVer)
RCari1.MoveFirst
Do While Not RCari1.EOF

    SCari2 = "Select * From JS02CTK"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
        RCari2("Status") = 1
        RCari2("NOTA") = RCari1("Nota")
        RCari2("NAMA") = RCari1("Nama")
        RCari2("ALAMAT") = RCari1("Alamat")
        RCari2("MASUK") = RCari1("MASUK")
        RCari2("KELUAR") = RCari1("Keluar")
        RCari2("SPAREPART") = RCari1("Sparepart")
        RCari2("SERVIS") = RCari1("Servis")
        RCari2("AMBIL") = RCari1("Ambil")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing

RCari1.MoveNext
Loop
RCari1.Close
Set RCari1 = Nothing

ErrorHandler:
Select Case Err.Number
    Case 40060
    RCari2("KELUAR") = RCari1("Masuk")
End Select

End Sub

Private Sub Seleksi2()
SDel = "Delete * From JS02CTK"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

SCari1 = "Select * From JS02 where BULAN = '" + Trim(Text1) + "' and TAHUN = '" + Trim(Text2) + "'and AMBIL = 'SELESAI'"
Set RCari1 = RDCO.OpenResultset(SCari1, rdOpenDynamic, rdConcurRowVer)
RCari1.MoveFirst
Do While Not RCari1.EOF

    SCari2 = "Select * From JS02CTK"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
        RCari2("Status") = 1
        RCari2("NOTA") = RCari1("Nota")
        RCari2("NAMA") = RCari1("Nama")
        RCari2("ALAMAT") = RCari1("Alamat")
        RCari2("MASUK") = RCari1("MASUK")
        RCari2("KELUAR") = RCari1("Keluar")
        RCari2("SPAREPART") = RCari1("Sparepart")
        RCari2("SERVIS") = RCari1("Servis")
        RCari2("AMBIL") = RCari1("Ambil")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing

RCari1.MoveNext
Loop
RCari1.Close
Set RCari1 = Nothing

ErrorHandler:
Select Case Err.Number
    Case 40060
    RCari2("KELUAR") = RCari1("Masuk")
End Select

End Sub

Private Sub Command2_Click()

If Text1 = "" Or Text2 = "" Then Exit Sub

Call Seleksi2

Y = Text2
M = Text1
D = 1
D1 = 31

Dim tanya
tanya = MsgBox("CETAK LAPORAN", vbOKCancel, "KONFIRMASI")
    If tanya = vbOK Then
        crpt.ReportFileName = "c:\windows\ReportSELULER\LapService2.rpt"
        crpt.WindowState = crptMaximized
        crpt.WindowMaxButton = True
        crpt.WindowMinButton = True
        crpt.Action = 1
        crpt.Reset
    Else
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

ClearTextBoxes Me

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
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

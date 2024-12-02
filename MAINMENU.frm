VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form MAINMENU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU UTAMA"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "TABEL BARANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3668
      TabIndex        =   18
      Top             =   2175
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ELECTRIC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3668
      TabIndex        =   17
      Top             =   1710
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      Height          =   2025
      Left            =   4823
      Picture         =   "MAINMENU.frx":0000
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   298
      TabIndex        =   16
      Top             =   2700
      Width           =   4530
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   600
      Left            =   240
      OleObjectBlob   =   "MAINMENU.frx":35EF
      TabIndex        =   13
      Top             =   3015
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PEMBELIAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   7077
      TabIndex        =   11
      Top             =   1710
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PENJUALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   259
      TabIndex        =   10
      Top             =   1710
      Width           =   2115
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
      Left            =   98
      TabIndex        =   9
      Top             =   6585
      Width           =   9255
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6293
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6060
      Width           =   3060
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3195
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6060
      Width           =   3060
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   98
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6060
      Width           =   3060
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   98
      TabIndex        =   2
      Top             =   4755
      Width           =   9255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":3664
         TabIndex        =   4
         Top             =   480
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   98
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":36D4
         TabIndex        =   5
         Top             =   1110
         Width           =   8775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":3744
         TabIndex        =   1
         Top             =   135
         Width           =   8775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":37B4
         TabIndex        =   3
         Top             =   735
         Width           =   8775
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "MAINMENU.frx":3824
      Top             =   9900
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   525
      Top             =   9030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   300
      Left            =   240
      OleObjectBlob   =   "MAINMENU.frx":3A58
      TabIndex        =   14
      Top             =   3780
      Width           =   3540
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   300
      Left            =   240
      OleObjectBlob   =   "MAINMENU.frx":3AE7
      TabIndex        =   15
      Top             =   4140
      Width           =   3540
   End
   Begin VB.Frame Frame2 
      Height          =   2100
      Left            =   98
      TabIndex        =   12
      Top             =   2625
      Width           =   4530
   End
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   98
      TabIndex        =   19
      Top             =   1575
      Width           =   9255
   End
   Begin VB.Menu P 
      Caption         =   "PENJUALAN"
      Index           =   1
      Begin VB.Menu PJ 
         Caption         =   "TRANSAKSI PENJUALAN"
         Index           =   11
      End
      Begin VB.Menu PJ 
         Caption         =   "TRANSAKSI PULSA"
         Index           =   12
      End
      Begin VB.Menu PJ 
         Caption         =   "-"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "TRANSAKSI PIUTANG"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "DAFTAR PIUTANG"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "-"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "CETAK ULANG NOTA"
         Index           =   17
         Visible         =   0   'False
      End
   End
   Begin VB.Menu B 
      Caption         =   "PEMBELIAN"
      Index           =   2
      Begin VB.Menu PB 
         Caption         =   "TRANSAKSI PEMBELIAN"
         Index           =   21
      End
      Begin VB.Menu PB 
         Caption         =   "-"
         Index           =   22
      End
      Begin VB.Menu PB 
         Caption         =   "DEPOSIT PULSA"
         Index           =   23
      End
      Begin VB.Menu PB 
         Caption         =   "REKAP DEPOSIT"
         Index           =   24
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "DAFTAR HUTANG"
         Index           =   25
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "-"
         Index           =   26
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "RETURN"
         Index           =   27
         Visible         =   0   'False
      End
   End
   Begin VB.Menu D 
      Caption         =   "DATA"
      Index           =   31
      Begin VB.Menu DS 
         Caption         =   "KODE KATEGORI BARANG"
         Index           =   31
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "KODE BARANG"
         Index           =   32
      End
      Begin VB.Menu DS 
         Caption         =   "KODE PELANGGAN"
         Index           =   33
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "KODE DISTRIBUTOR"
         Index           =   34
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "-"
         Index           =   35
      End
      Begin VB.Menu DS 
         Caption         =   "JASA"
         Index           =   36
      End
      Begin VB.Menu DS 
         Caption         =   "-"
         Index           =   37
      End
      Begin VB.Menu DS 
         Caption         =   "ELECTRIC"
         Index           =   38
         Begin VB.Menu DSS 
            Caption         =   "INDUK VOUCHER"
            Index           =   381
         End
         Begin VB.Menu DSS 
            Caption         =   "KODE VOUCHER"
            Index           =   382
         End
      End
   End
   Begin VB.Menu T 
      Caption         =   "TOOLS"
      Index           =   4
      Begin VB.Menu TS 
         Caption         =   "SETING TOKO"
         Index           =   41
      End
      Begin VB.Menu TS 
         Caption         =   "GANTI PASSWORD"
         Index           =   42
      End
      Begin VB.Menu TS 
         Caption         =   "USER BARU"
         Index           =   43
         Visible         =   0   'False
      End
   End
   Begin VB.Menu L 
      Caption         =   "LAPORAN"
      Index           =   5
      Begin VB.Menu LS 
         Caption         =   "MUTASI BARANG"
         Index           =   501
      End
      Begin VB.Menu LS 
         Caption         =   "MUTASI ELECTRIC"
         Index           =   502
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   503
      End
      Begin VB.Menu LS 
         Caption         =   "LAP PEMBELIAN"
         Index           =   504
      End
      Begin VB.Menu LS 
         Caption         =   "LAP JUMLAH PEMBELIAN"
         Index           =   505
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   506
      End
      Begin VB.Menu LS 
         Caption         =   "LAP PENJUALAN"
         Index           =   507
      End
      Begin VB.Menu LS 
         Caption         =   "LAP JUMLAH PENJUALAN"
         Index           =   508
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   509
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "LABA RUGI"
         Index           =   510
         Visible         =   0   'False
      End
   End
   Begin VB.Menu S 
      Caption         =   "SERVICE"
      Index           =   6
      Begin VB.Menu SS 
         Caption         =   "MASUK"
         Index           =   60
      End
      Begin VB.Menu SS 
         Caption         =   "KELUAR"
         Index           =   61
      End
      Begin VB.Menu SS 
         Caption         =   "-"
         Index           =   62
      End
      Begin VB.Menu SS 
         Caption         =   "DAFTAR SERVICE"
         Index           =   63
      End
      Begin VB.Menu SS 
         Caption         =   "-"
         Index           =   64
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "NOTA SERVICE MASUK"
         Index           =   65
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "NOTA SERVICE KELUAR"
         Index           =   66
         Visible         =   0   'False
      End
   End
   Begin VB.Menu A 
      Caption         =   "ABOUT"
      Index           =   7
      Begin VB.Menu AA 
         Caption         =   "LISENSI"
         Index           =   70
      End
   End
End
Attribute VB_Name = "MAINMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private Sub AA_Click(Index As Integer)
Select Case Index
    Case 70
        L001.Show 1
End Select
End Sub

Private Sub cmdCLOSE_Click()
Unload Me
LOGIN.Show
End Sub

Private Sub Command1_Click()
JL001.Show 1
End Sub

Private Sub Command2_Click()
BL001.Show 1
End Sub

Private Sub Command3_Click()
VC03.Show 1
End Sub

Private Sub Command4_Click()
VC05.Show 1
End Sub

Private Sub DS_Click(Index As Integer)
Select Case Index
    Case 31
        B001.Show 1
    Case 32
        B003.Show 1
    Case 33
        P001.Show 1
    Case 34
        D001.Show 1
    Case 36
        JS01.Show 1
    Case 37
        VC01.Show 1
End Select
End Sub

Private Sub DSS_Click(Index As Integer)
Select Case Index
    Case 381
        VC01.Show 1
    Case 382
        VC02.Show 1
End Select
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

Text1 = "USER : " + Operator
Text2 = Date
Text3 = "Copyrighted 2008 - IPT"

SkinLabel1 = NTOKO
SkinLabel4 = NAlamat
SkinLabel5 = NMOtto
SkinLabel6 = NTelepon
Me.Left = 0
Me.Top = 0
End Sub

Private Sub LS_Click(Index As Integer)
Select Case Index
    Case 501
        Call LapBR
    Case 502
        VC04.Show 1
    Case 504
        Indikator = 0
        TglFuck = ""
        TGLFAK.Show 1
        If Indikator = 1 Then
            Call LapTransBeli
        Else
            Exit Sub
        End If
    Case 505
    Case 506
    Case 507
        Indikator = 0
        TglFuck = ""
        TglFuck2 = ""
        TGLFAK.Show 1
        If Indikator = 1 Then
            Call LapTransJual
        Else
            Exit Sub
        End If
    Case 508
    Case 509
    Case 510
        LR.Show 1
End Select
End Sub

Private Sub LapBR()
Crpt.ReportFileName = "C:\WINDOWS\ReportSELULER\LapBR.rpt"
Crpt.WindowState = crptMaximized
Crpt.WindowMaxButton = False
Crpt.WindowMinButton = False
Crpt.Action = 1
End Sub

Private Sub LapTransBeli()
Crpt.ReportFileName = "C:\WINDOWS\ReportSELULER\TransBeli.rpt"
Crpt.WindowState = crptMaximized
Crpt.WindowMaxButton = False
Crpt.WindowMinButton = False
Crpt.Action = 1
End Sub

Private Sub LapTransJual()
Crpt.ReportFileName = "C:\WINDOWS\ReportSELULER\TransJual.rpt"
Crpt.WindowState = crptMaximized
Crpt.WindowMaxButton = False
Crpt.WindowMinButton = False
Crpt.Action = 1
End Sub

Private Sub PB_Click(Index As Integer)
Select Case Index
    Case 21
        BL001.Show 1
    Case 23
        VC00.Show 1
End Select
End Sub

Private Sub PJ_Click(Index As Integer)
Select Case Index
    Case 11
        JL001.Show 1
    Case 12
        VC03.Show 1
    Case 14
    Case 15
End Select
End Sub

Private Sub SS_Click(Index As Integer)
Select Case Index
    Case 60
        JS02.Show 1
    Case 61
        JS03.Show 1
    Case 63
        JS001.Show 1
End Select
End Sub

Private Sub TS_Click(Index As Integer)
Select Case Index
    Case 41
        NAMA.Show 1
    Case 42
        GPASS.Show 1
    Case 43
        User.Show 1
End Select
End Sub

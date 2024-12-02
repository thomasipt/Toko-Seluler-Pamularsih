VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form L001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERATURAN LISENSI PROGRAM"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":0000
      TabIndex        =   14
      Top             =   5625
      Width           =   2970
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
      Left            =   3735
      TabIndex        =   0
      Top             =   4860
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "L001.frx":0083
      Top             =   6120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   285
      Left            =   720
      OleObjectBlob   =   "L001.frx":02B7
      TabIndex        =   1
      Top             =   135
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":03C2
      TabIndex        =   2
      Top             =   135
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":041D
      TabIndex        =   3
      Top             =   495
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":0478
      TabIndex        =   4
      Top             =   855
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":04D3
      TabIndex        =   5
      Top             =   1440
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":052E
      TabIndex        =   6
      Top             =   2295
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":0589
      TabIndex        =   7
      Top             =   2880
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   285
      Left            =   720
      OleObjectBlob   =   "L001.frx":05E4
      TabIndex        =   8
      Top             =   495
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   510
      Left            =   720
      OleObjectBlob   =   "L001.frx":06E1
      TabIndex        =   9
      Top             =   855
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   780
      Left            =   720
      OleObjectBlob   =   "L001.frx":0854
      TabIndex        =   10
      Top             =   1440
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   510
      Left            =   720
      OleObjectBlob   =   "L001.frx":0A5B
      TabIndex        =   11
      Top             =   2295
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   285
      Left            =   720
      OleObjectBlob   =   "L001.frx":0BCA
      TabIndex        =   12
      Top             =   2880
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   285
      Left            =   6300
      OleObjectBlob   =   "L001.frx":0CD7
      TabIndex        =   13
      Top             =   4320
      Width           =   2880
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   285
      Left            =   90
      OleObjectBlob   =   "L001.frx":0D52
      TabIndex        =   15
      Top             =   3240
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   510
      Left            =   720
      OleObjectBlob   =   "L001.frx":0DAD
      TabIndex        =   16
      Top             =   3240
      Width           =   8550
   End
End
Attribute VB_Name = "L001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim tanya
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
End Sub

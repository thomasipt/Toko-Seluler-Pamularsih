VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form JS02 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVICE MASUK"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridB 
      Height          =   1170
      Left            =   150
      TabIndex        =   73
      Top             =   4575
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   2064
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
      Appearance      =   0
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
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   1170
      Left            =   150
      TabIndex        =   24
      Top             =   4575
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   2064
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
      Appearance      =   0
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
   Begin VB.Frame Frame2 
      Caption         =   "EDIT JUMLAH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   1973
      TabIndex        =   44
      Top             =   2820
      Width           =   4755
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   54
         Text            =   "20"
         Top             =   900
         Width           =   1995
      End
      Begin VB.CommandButton Command6 
         Caption         =   "BATAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   53
         Top             =   1215
         Width           =   1245
      End
      Begin VB.CommandButton Command5 
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
         Height          =   285
         Left            =   3240
         TabIndex        =   52
         Top             =   765
         Width           =   1245
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   48
         Text            =   "23"
         Top             =   585
         Width           =   1995
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   47
         Text            =   "21"
         Top             =   1215
         Width           =   1995
      End
      Begin VB.CommandButton Command4 
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
         Height          =   285
         Left            =   3240
         TabIndex        =   46
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   45
         Text            =   "19"
         Top             =   270
         Width           =   1995
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":0000
         TabIndex        =   49
         Top             =   630
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":005F
         TabIndex        =   50
         Top             =   1260
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":00C2
         TabIndex        =   51
         Top             =   315
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":0121
         TabIndex        =   55
         Top             =   945
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SERVICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2595
      Left            =   -6585
      TabIndex        =   58
      Top             =   4500
      Width           =   8580
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   61
         Text            =   "28"
         Top             =   495
         Width           =   1995
      End
      Begin VB.TextBox Text27 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4125
         TabIndex        =   60
         Text            =   "27"
         Top             =   180
         Width           =   1995
      End
      Begin VB.CommandButton Command8 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7650
         TabIndex        =   13
         Top             =   135
         Width           =   525
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   59
         Text            =   "24"
         Top             =   180
         Width           =   1995
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":0182
         TabIndex        =   62
         Top             =   540
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
         Height          =   195
         Left            =   3390
         OleObjectBlob   =   "JS02.frx":01E1
         TabIndex        =   63
         Top             =   225
         Width           =   1035
      End
      Begin MSFlexGridLib.MSFlexGrid grid4 
         Height          =   1710
         Left            =   105
         TabIndex        =   64
         Top             =   810
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   3016
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   16777152
         AllowUserResizing=   3
         Appearance      =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":0242
         TabIndex        =   65
         Top             =   225
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ESTIMASI BIAYA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1680
      Left            =   3908
      TabIndex        =   66
      Top             =   6210
      Width           =   4740
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         TabIndex        =   68
         Text            =   "10"
         Top             =   360
         Width           =   2520
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         TabIndex        =   10
         Text            =   "11"
         Top             =   780
         Width           =   2520
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         TabIndex        =   67
         Text            =   "12"
         Top             =   1200
         Width           =   2520
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   300
         Left            =   315
         OleObjectBlob   =   "JS02.frx":02A1
         TabIndex        =   69
         Top             =   450
         Width           =   1665
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   300
         Left            =   315
         OleObjectBlob   =   "JS02.frx":0302
         TabIndex        =   70
         Top             =   870
         Width           =   1665
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   300
         Left            =   315
         OleObjectBlob   =   "JS02.frx":036B
         TabIndex        =   71
         Top             =   1290
         Width           =   1665
      End
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   57
      Text            =   "DATA INDUK"
      Top             =   60
      Width           =   1365
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6128
      TabIndex        =   27
      Text            =   "9"
      Top             =   5775
      Width           =   2520
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "PART"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   7703
      TabIndex        =   7
      Top             =   2895
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SERVICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   7688
      TabIndex        =   9
      Top             =   4575
      Width           =   945
   End
   Begin VB.CommandButton cmdEDIT 
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
      Left            =   180
      TabIndex        =   11
      Top             =   6345
      Width           =   3600
   End
   Begin VB.CommandButton Command1 
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
      Left            =   150
      TabIndex        =   12
      Top             =   7335
      Width           =   3600
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      TabIndex        =   6
      Text            =   "7"
      Top             =   2460
      Width           =   6930
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      TabIndex        =   5
      Text            =   "6"
      Top             =   2070
      Width           =   3150
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Text            =   "5"
      Top             =   1680
      Width           =   2310
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      TabIndex        =   3
      Text            =   "4"
      Top             =   1680
      Width           =   2310
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1230
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "JS02.frx":03E0
      Top             =   1050
      Width           =   7350
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1230
      TabIndex        =   1
      Text            =   "2"
      Top             =   735
      Width           =   3885
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1230
      TabIndex        =   0
      Text            =   "1"
      Top             =   405
      Width           =   1995
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9360
      OleObjectBlob   =   "JS02.frx":03E2
      Top             =   930
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "JS02.frx":0616
      TabIndex        =   14
      Top             =   450
      Width           =   1035
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "JS02.frx":0682
      TabIndex        =   15
      Top             =   780
      Width           =   1035
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "JS02.frx":06E8
      TabIndex        =   16
      Top             =   1095
      Width           =   1035
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "JS02.frx":0752
      TabIndex        =   19
      Top             =   1770
      Width           =   1350
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   195
      Left            =   4170
      OleObjectBlob   =   "JS02.frx":07B3
      TabIndex        =   20
      Top             =   1770
      Width           =   1350
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "JS02.frx":0812
      TabIndex        =   21
      Top             =   2160
      Width           =   1350
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "JS02.frx":087B
      TabIndex        =   22
      Top             =   2550
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1170
      Left            =   150
      TabIndex        =   23
      Top             =   2895
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   2064
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
      Appearance      =   0
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   4500
      TabIndex        =   25
      Top             =   390
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   529
      _Version        =   393216
      Format          =   20185089
      CurrentDate     =   39620
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   7230
      TabIndex        =   26
      Top             =   390
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   529
      _Version        =   393216
      Format          =   20185089
      CurrentDate     =   39620
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   195
      Left            =   3390
      OleObjectBlob   =   "JS02.frx":08E6
      TabIndex        =   17
      Top             =   450
      Width           =   1245
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   195
      Left            =   5985
      OleObjectBlob   =   "JS02.frx":0956
      TabIndex        =   18
      Top             =   450
      Width           =   1245
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   195
      Left            =   2565
      OleObjectBlob   =   "JS02.frx":09C8
      TabIndex        =   29
      Top             =   4185
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   195
      Left            =   4770
      OleObjectBlob   =   "JS02.frx":0A3D
      TabIndex        =   30
      Top             =   5865
      Width           =   1350
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6128
      TabIndex        =   28
      Text            =   "8"
      Top             =   4095
      Width           =   2520
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4545
      TabIndex        =   43
      Text            =   "18"
      Top             =   4095
      Width           =   1530
   End
   Begin VB.CommandButton Command7 
      Height          =   8895
      Left            =   -337
      TabIndex        =   56
      Top             =   -645
      Width           =   9375
   End
   Begin MSFlexGridLib.MSFlexGrid gridA 
      Height          =   1170
      Left            =   150
      TabIndex        =   72
      Top             =   2895
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   2064
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
      Appearance      =   0
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
   Begin VB.Frame Frame1 
      Caption         =   "SPAREPART"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2955
      Left            =   -4200
      TabIndex        =   31
      Top             =   2820
      Width           =   8580
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   41
         Text            =   "17"
         Top             =   180
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7665
         TabIndex        =   8
         Top             =   315
         Width           =   525
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   39
         Text            =   "16"
         Top             =   735
         Width           =   1995
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   37
         Text            =   "15"
         Top             =   315
         Width           =   1995
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   36
         Text            =   "14"
         Top             =   825
         Width           =   1995
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   34
         Text            =   "13"
         Top             =   495
         Width           =   1995
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":0AAE
         TabIndex        =   32
         Top             =   540
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":0B0D
         TabIndex        =   33
         Top             =   870
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   195
         Left            =   3885
         OleObjectBlob   =   "JS02.frx":0B6E
         TabIndex        =   35
         Top             =   360
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   195
         Left            =   3885
         OleObjectBlob   =   "JS02.frx":0BD1
         TabIndex        =   38
         Top             =   780
         Width           =   1035
      End
      Begin MSFlexGridLib.MSFlexGrid grid3 
         Height          =   1710
         Left            =   105
         TabIndex        =   40
         Top             =   1155
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   3016
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   16777152
         AllowUserResizing=   3
         Appearance      =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   195
         Left            =   420
         OleObjectBlob   =   "JS02.frx":0C32
         TabIndex        =   42
         Top             =   225
         Width           =   1035
      End
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "JS02"
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

Private BBB, TTT

Private Sub cmdCLOSE_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "DATA INDUK TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

Call Tutup

Frame1.Left = 38
Frame1.ZOrder

Text17 = ""
Text13 = ""
Text14 = ""
Text15 = ""
Text16 = ""

With grid3
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "HARGA": .CellAlignment = 4
End With

SKTG = "Select * From B003 where KodeInd = '105' order by KodeBR Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid3.Rows = B + 1
      grid3.Row = B
         With grid3
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("NamaBR")
              .Col = 2: .Text = Format(RKTG("HJual"), "##,###.00")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing

End Sub

Private Sub cmdEDIT_Click()
Dim tanya
tanya = MsgBox("TRANSAKSI SELESAI", vbOKCancel, "KONFIRMASI")
    If tanya = vbOK Then
        Call SimpanJS02
        Call SimpanJS02HIS_A
        Call SimpanJS02HIS_B
        NOTAFAK = ""
        NOTAFAK = Text1
    Else
        Exit Sub
    End If

Unload Me
'JS02FAK.Show 1
End Sub

Private Sub SimpanJS02()
SSave = "Select * From JS02"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("Status") = 1
        RSave("Nota") = Trim(Text1)
        RSave("Nama") = Trim(Text2)
        RSave("Alamat") = Trim(Text3)
        RSave("Jenis") = Trim(Text4)
        RSave("Seri") = Trim(Text5)
        RSave("Kerusakan") = Trim(Text6)
        RSave("Keterangan") = Trim(Text7)
        RSave("Nota") = Trim(Text1)
        
        RSave("Masuk") = DTPicker1
        RSave("Keluar") = DTPicker1
        RSave("Prediksi") = DTPicker2
        
        RSave("Sparepart") = CCur(Text8)
        RSave("Servis") = CCur(Text9)
        
        RSave("Total") = CCur(Text10)
        RSave("Bayar") = CCur(Text11)
        RSave("Sisa") = CCur(Text12)
        
        RSave("Ambil") = "BELUM SELESAI"
        RSave("Terbilang") = Terbilang(Text10)
        
        RSave("Bulan") = BBB
        RSave("Tahun") = TTT
        
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub SimpanJS02HIS_A()
SCari1 = "Select * From JS02A"
Set RCari1 = RDCO.OpenResultset(SCari1, rdOpenDynamic, rdConcurRowVer)
RCari1.MoveFirst
Do While Not RCari1.EOF
    SCari2 = "Select * From JS02HIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
        RCari2("NOTA") = Trim(Text1)
        RCari2("KODE") = RCari1("Kode")
        RCari2("NAMA") = RCari1("Nama")
        RCari2("HARGA") = RCari1("Harga")
        RCari2("JUMLAH") = RCari1("Jumlah")
        RCari2("TOTAL") = RCari1("Total")
        RCari2("STS") = "SP"
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
RCari1.MoveNext
Loop
RCari1.Close
Set RCari1 = Nothing
End Sub

Private Sub SimpanJS02HIS_B()
SCari1 = "Select * From JS02B"
Set RCari1 = RDCO.OpenResultset(SCari1, rdOpenDynamic, rdConcurRowVer)
RCari1.MoveFirst
Do While Not RCari1.EOF
    SCari2 = "Select * From JS02HIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
        RCari2("NOTA") = Trim(Text1)
        RCari2("KODE") = RCari1("Kode")
        RCari2("NAMA") = RCari1("Nama")
        RCari2("BIAYA") = RCari1("Biaya")
        RCari2("STS") = "SVC"
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
RCari1.MoveNext
Loop
RCari1.Close
Set RCari1 = Nothing
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text2 = "" Or Text2 = "" Or Text2 = "" Or Text2 = "" Or Text2 = "" Then
    MsgBox "DATA INDUK TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

Call Tutup

Frame3.Left = 38
Frame3.ZOrder

Text24 = ""
Text27 = ""
Text28 = ""

With grid4
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "BIAYA": .CellAlignment = 4
End With

SKTG = "Select * From JS01 order by Kode Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid4.Rows = B + 1
      grid4.Row = B
         With grid4
              .Col = 0: .Text = RKTG("Kode"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("Nama")
              .Col = 2: .Text = Format(RKTG("Biaya"), "##,###.00")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub Command3_Click()
If Text15 = "" Then Exit Sub

Text16 = Format(CCur(Text15) * CCur(Text14), "##,###.00")

If Text17 = "" Or Text13 = "" Then
    Frame1.Left = 10000
Else
    Frame1.Left = 10000
    
    SSave = "Select * From JS02A"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.AddNew
            RSave("Kode") = Trim(Text17)
            RSave("Nama") = Trim(Text13)
            RSave("Harga") = CCur(Text14)
            RSave("Jumlah") = CCur(Text15)
            RSave("Total") = CCur(Text16)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
    
    Call IsiGrid1
    Call TotalIsiGrid1
    cmdCLOSE.SetFocus
End If

Call Buka
Call Estimasi
End Sub

Private Sub TotalIsiGrid1()
SCari2 = "Select * From TotalJS02A"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    On Error GoTo ErrorHandler
    Text18 = RCari2("SumOfJumlah")
    Text8 = Format(RCari2("SumOfTotal"), "##,###.00")
End If
RCari2.Close
Set RCari2 = Nothing

If Text18 > 0 Then
    gridA.Visible = False
End If

ErrorHandler:
Select Case Err.Number
    Case 94
    Text18 = "0"
    Text8 = "0,00"
    gridA.Visible = True
    gridA.ZOrder
End Select

End Sub

Private Sub TotalIsiGrid2()
SCari2 = "Select * From TotalJS02B"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    'On Error GoTo ErrorHandler
    Text9 = Format(RCari2("SumOfBiaya"), "##,###.00")
End If
RCari2.Close
Set RCari2 = Nothing

If Text9 = "" Then
    gridB.Visible = True
    gridB.ZOrder
Else
    gridB.Visible = False
End If

'ErrorHandler:
'Select Case Err.Number
'    Case 94
'    Text9 = "0,00"
'    gridB.Visible = True
'    gridB.ZOrder
'End Select

End Sub

Private Sub Command4_Click()
SCari = "Select * From JS02A where Kode = '" + Trim(KodeK) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.Edit
    RCari("Jumlah") = Trim(Text21)
    RCari("Total") = CCur(Text20) * CCur(Text21)
RCari.Update
RCari.Close
Set RCari = Nothing

Frame2.Visible = False
cmdCLOSE.SetFocus
Call IsiGrid1
Call TotalIsiGrid1
Call Buka
Call Estimasi
End Sub

Private Sub Command5_Click()
SDel = "Delete * From JS02A where Kode = '" + Trim(KodeK) + "'"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

Frame2.Visible = False
cmdCLOSE.SetFocus
Call SiapkanGrid1
Call IsiGrid1
Call TotalIsiGrid1
Call Buka
Call Estimasi

End Sub

Private Sub Command6_Click()
Frame2.Visible = False
cmdCLOSE.SetFocus
Call Buka
End Sub

Private Sub Command7_Click()
Call IsiGrid1
Call IsiGrid2
End Sub

Private Sub Command8_Click()
If Text24 = "" Or Text28 = "" Then
    Frame3.Left = 10000
Else
    Frame3.Left = 10000
    
    SSave = "Select * From JS02B"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.AddNew
            RSave("Kode") = Trim(Text24)
            RSave("Nama") = Trim(Text28)
            RSave("Biaya") = CCur(Text27)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
    
    Call IsiGrid2
    Call TotalIsiGrid2
    cmdCLOSE.SetFocus
End If

Call Buka
Call Estimasi
End Sub

Private Sub Form_Load()
Dim tanya
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Call SiapkanGrid1
Call SiapkanGrid2

DTPicker1 = Date
DTPicker2 = Date + 7

Frame1.Left = 10000
Frame3.Left = 10000
Frame2.Visible = False

Call DelJS02A

Call Buka

Text22 = "DATA INDUK"
Text8 = "0,00"
Text9 = "0,00"
Text11 = "0,00"
Text12 = "0,00"

gridA.Visible = False
gridB.Visible = False

BBB = Month(Date)
TTT = Year(Date)

End Sub

Private Sub DelJS02A()
SDel = "Delete * From JS02A"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

SDel2 = "Delete * From JS02B"
Set RDel2 = RDCO.OpenResultset(SDel2, rdOpenDynamic, rdConcurRowVer)
RDel2.Close
Set RDel2 = Nothing

End Sub

Private Sub SiapkanGrid1()
With grid1
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "HARGA": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1000: .Text = "JUMLAH": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "TOTAL": .CellAlignment = 4
End With

With gridA
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "HARGA": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1000: .Text = "JUMLAH": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "TOTAL": .CellAlignment = 4
End With

End Sub

Private Sub SiapkanGrid2()
With grid2
    .Cols = 4
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "BIAYA": .CellAlignment = 4
End With

With gridB
    .Cols = 4
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "BIAYA": .CellAlignment = 4
End With

End Sub

Private Sub IsiGrid1()
SCari = "Select * From JS02A order by NO Asc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid1.Rows = B + 1
      grid1.Row = B
         With grid1
            .Col = 0: .Text = B: .CellAlignment = 4
            .Col = 1: .Text = RCari("Kode"): .CellAlignment = 4
            .Col = 2: .Text = RCari("Nama")
            .Col = 3: .Text = Format(RCari("Harga"), "##,###.00")
            .Col = 4: .Text = RCari("Jumlah")
            .Col = 5: .Text = Format(RCari("Total"), "##,###.00")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiGrid2()
SCari2 = "Select * From JS02B order by NO Asc"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurReadOnly)
If RCari2.RowCount <> 0 Then
   RCari2.MoveFirst
   B = 1
   Do Until RCari2.EOF
      grid2.Rows = B + 1
      grid2.Row = B
         With grid2
            .Col = 0: .Text = B: .CellAlignment = 4
            .Col = 1: .Text = RCari2("Kode"): .CellAlignment = 4
            .Col = 2: .Text = RCari2("Nama")
            .Col = 3: .Text = Format(RCari2("Biaya"), "##,###.00")
         End With
      B = B + 1
      RCari2.MoveNext
   Loop
End If
RCari2.Close
Set RCari2 = Nothing

End Sub

Private Sub CekData()
If Text1.Text = "" Then Exit Sub

SCari = "Select * From JS02 where Nota = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " KODE NOTA TELAH DIGUNAKAN", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    Else
       Text2.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid1_dblClick()
Dim tanya
If (grid1.TextMatrix(grid1.Row, 1)) = "" Then Exit Sub

tanya = MsgBox("EDIT TABEL SPAREPART", vbCritical, "KONFIRMASI")
    If tanya = vbOK Then
        Call Tutup
        Frame2.Visible = True
        Frame2.ZOrder
        Text19 = (grid1.TextMatrix(grid1.Row, 1))
        Text23 = (grid1.TextMatrix(grid1.Row, 2))
        Text20 = (grid1.TextMatrix(grid1.Row, 3))
        Text21 = (grid1.TextMatrix(grid1.Row, 4))
        KodeK = Text19
        Text21.SetFocus
    Else
        MsgBox "CANCEL", vbCritical, "KONFIRMASI"
    End If
End Sub

Private Sub grid2_dblClick()
Dim tanya
If (grid2.TextMatrix(grid2.Row, 1)) = "" Then Exit Sub

tanya = MsgBox("HAPUS DATA SERVICE", vbCritical, "KONFIRMASI")
    If tanya = vbOK Then
        Call Tutup
        KodeK = (grid2.TextMatrix(grid2.Row, 1))
        
        SDel = "Delete * From JS02B where Kode = '" + Trim(KodeK) + "'"
        Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
        RDEl.Close
        Set RDEl = Nothing
        
        Call SiapkanGrid2
        Call IsiGrid2
        Call TotalIsiGrid2
        Call Buka
        Call Estimasi

    Else
        MsgBox "CANCEL", vbCritical, "KONFIRMASI"
    End If
End Sub

Private Sub grid3_dblClick()
Text17 = (grid3.TextMatrix(grid3.Row, 0))
Text13 = (grid3.TextMatrix(grid3.Row, 1))
Text14 = Format((grid3.TextMatrix(grid3.Row, 2)), "##,###.00")
Text15.SetFocus
End Sub

Private Sub grid4_dblClick()
Text24 = (grid4.TextMatrix(grid4.Row, 0))
Text28 = (grid4.TextMatrix(grid4.Row, 1))
Text27 = Format((grid4.TextMatrix(grid4.Row, 2)), "##,###.00")
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


Private Sub Text11_GotFocus()
If CCur(Text11) = 0 Then Text11 = ""
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If Text10 = "" Then Exit Sub

If KeyAscii = 13 Then

If Text11 = "" Then Text11.SetFocus

If Not IsNumeric(Text11) Then
    Text11.SetFocus
    Text11 = ""
    MsgBox "NOMINAL HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If

    Text12 = Format(CCur(Text10) - CCur(Text11), "##,###.00")
    
    If CCur(Text10) < CCur(Text11) Then
        MsgBox "NOMINAL MELEBIHI ESTIMASI BIAYA", vbCritical, "WARNING"
        Text11.SetFocus
        Exit Sub
    End If
    
Text11 = Format(Text11, "##,###.00")
SendKeys vbTab
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text15_LostFocus()
Text15 = Format(Text15, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
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
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub

Private Sub Tutup()
Command7.Visible = True
Command7.ZOrder
End Sub

Private Sub Buka()
Command7.Visible = False
End Sub

Private Sub Estimasi()
If Text9 = "" Then
    Text9 = "0,00"
End If

If Text8 = "" Then
    Text8 = "0,00"
End If

Text10 = Format(CCur(Text8) + CCur(Text9), "##,###.00")
End Sub

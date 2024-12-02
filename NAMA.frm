VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form NAMA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETING TOKO"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3225
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1530
      Width           =   2850
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   105
      TabIndex        =   7
      Top             =   15
      Width           =   7410
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1275
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1275
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   495
         Width           =   6015
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1275
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   750
         Width           =   6015
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1275
         TabIndex        =   3
         Text            =   "Text4"
         Top             =   1005
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "NAMA.frx":0000
         TabIndex        =   8
         Top             =   255
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "NAMA.frx":0070
         TabIndex        =   9
         Top             =   510
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "NAMA.frx":00DA
         TabIndex        =   10
         Top             =   765
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "NAMA.frx":0142
         TabIndex        =   11
         Top             =   1020
         Width           =   1560
      End
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
      Left            =   105
      TabIndex        =   5
      Top             =   2160
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
      Left            =   5625
      TabIndex        =   6
      Top             =   2175
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "NAMA.frx":01AE
      Top             =   4200
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   225
      Left            =   1545
      OleObjectBlob   =   "NAMA.frx":03E2
      TabIndex        =   12
      Top             =   1575
      Width           =   1560
   End
End
Attribute VB_Name = "NAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSkin, RTOKO, RTgl, RHapus, RDEl, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private SSkin, STOKO, STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, Kode As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
SCari = "Select * From A001"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.Edit
    RCari("TOKO") = Trim(Text1)
    RCari("Alamat") = Trim(Text2)
    RCari("Motto") = Trim(Text3)
    RCari("Telepon") = Trim(Text4)
    RCari("S") = Trim(Combo1)
RCari.Update
RCari.Close
Set RCari = Nothing

STgl = "Select * from A001"
Set RTgl = RDCO.OpenResultset(STgl, rdOpenDynamic, rdConcurRowVer)
If RTgl.RowCount <> 0 Then
    TglS = RTgl("Tanggal")
    Skin = RTgl("S")
    NTOKO = RTgl("TOKO")
    NAlamat = RTgl("ALamat")
    NMOtto = RTgl("Motto")
    NTelepon = RTgl("Telepon")
Else
End If
RTgl.Close
Set RTgl = Nothing

Unload Me
Unload MAINMENU
MAINMENU.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
ClearTextBoxes NAMA
Call Isi
Call IsiCombo
End Sub

Private Sub Isi()
STOKO = "Select * from A001"
Set RTOKO = RDCO.OpenResultset(STOKO, rdOpenDynamic, rdConcurRowVer)
If RTOKO.RowCount <> 0 Then
    Text1 = Format(RTOKO("TOKO"), ">")
    Text2 = Format(RTOKO("ALamat"), ">")
    Text3 = Format(RTOKO("Motto"), ">")
    Text4 = RTOKO("Telepon")
    Combo1 = Format(RTOKO("S"), ">")
End If
RTOKO.Close
Set RTOKO = Nothing
End Sub

Private Sub IsiCombo()
SSkin = "Select * from S001 order by nama asc"
Set RSkin = RDCO.OpenResultset(SSkin, rdOpenDynamic, rdOpenKeyset)
RSkin.MoveFirst
Do While Not RSkin.EOF
    Combo1.AddItem RSkin("Nama")
RSkin.MoveNext
Loop
RSkin.Close
Set RSkin = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text1_LostFocus()
If Text1.Text = "" Then
    Text1 = Format(NTOKO, ">")
    Text1.BackColor = RGB(255, 255, 255)
    Text1 = Format(Text1, ">")
Else
    Text1.BackColor = RGB(255, 255, 255)
    Text1 = Format(Text1, ">")
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text2_LostFocus()
If Text2.Text = "" Then
    Text2 = Format(NAlamat, ">")
    Text2.BackColor = RGB(255, 255, 255)
    Text2 = Format(Text2, ">")
Else
    Text2.BackColor = RGB(255, 255, 255)
    Text2 = Format(Text2, ">")
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text3_GotFocus()
Text3.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text3_LostFocus()
If Text3.Text = "" Then
    Text3 = Format(NMOtto, ">")
    Text3.BackColor = RGB(255, 255, 255)
    Text3 = Format(Text3, ">")
Else
    Text3.BackColor = RGB(255, 255, 255)
    Text3 = Format(Text3, ">")
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub text4_gotfocus()
Text4.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text4_LostFocus()
If Text4.Text = "" Then
    Text4 = Format(NTelepon, ">")
    Text4.BackColor = RGB(255, 255, 255)
    Text4 = Format(Text4, ">")
Else
    Text4.BackColor = RGB(255, 255, 255)
    Text4 = Format(Text4, ">")
End If
End Sub

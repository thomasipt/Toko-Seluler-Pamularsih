VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form P001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE PELANGGAN"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
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
      Left            =   117
      TabIndex        =   6
      Top             =   1830
      Width           =   1890
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
      Left            =   1508
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   420
      Width           =   5895
   End
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
      Left            =   1508
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   105
      Width           =   2850
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
      Left            =   5641
      TabIndex        =   7
      Top             =   1845
      Width           =   1890
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
      Left            =   117
      TabIndex        =   5
      Top             =   1830
      Width           =   1890
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
      Left            =   1508
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   735
      Width           =   5895
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
      Left            =   1508
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1050
      Width           =   5895
   End
   Begin VB.TextBox Text5 
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
      Left            =   1508
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1365
      Width           =   5895
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   194
      OleObjectBlob   =   "P001.frx":0000
      Top             =   6308
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   258
      OleObjectBlob   =   "P001.frx":0234
      TabIndex        =   8
      Top             =   113
      Width           =   930
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   258
      OleObjectBlob   =   "P001.frx":029A
      TabIndex        =   9
      Top             =   435
      Width           =   930
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2220
      Left            =   123
      TabIndex        =   10
      Top             =   2565
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   3916
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   258
      OleObjectBlob   =   "P001.frx":0300
      TabIndex        =   11
      Top             =   750
      Width           =   930
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   225
      Left            =   258
      OleObjectBlob   =   "P001.frx":036A
      TabIndex        =   12
      Top             =   1065
      Width           =   930
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   225
      Left            =   258
      OleObjectBlob   =   "P001.frx":03D6
      TabIndex        =   13
      Top             =   1380
      Width           =   1140
   End
End
Attribute VB_Name = "P001"
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

Private Sub cmdEDIT_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
    MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Else
    SSave = "Select * From P001 where Kode = '" + Trim(KB) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
    RSave.Edit
        RSave("Kode") = Trim(Text1)
        RSave("Nama") = Format(Trim(Text2), ">")
        RSave("Alamat") = Format(Trim(Text2), ">")
        RSave("Telepon") = Trim(Text2)
        RSave("Rekening") = Trim(Text2)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
    Call IsiGrid
    ClearTextBoxes P001
    Text1.SetFocus
    cmdOK.Visible = True
    cmdEDIT.Visible = False
End If
End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
    MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Else
    SSave2 = "Select * From P001"
    Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    RSave2.AddNew
        RSave2("Kode") = Trim(Text1)
        RSave2("Nama") = Format(Text2, ">")
        RSave2("Alamat") = Format(Text2, ">")
        RSave2("Telepon") = Format(Text2, ">")
        RSave2("Rekening") = Format(Text2, ">")
    RSave2.Update
    RSave2.Close
    Set RSave2 = Nothing
    Call IsiGrid
    ClearTextBoxes P001
    Text1.SetFocus
    cmdOK.Visible = True
    cmdEDIT.Visible = False
End If
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
ClearTextBoxes P001
Call SiapkanGrid
Call IsiGrid
cmdOK.Visible = True
cmdEDIT.Visible = False
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 5
    .Col = 0: .ColWidth(0) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 3000: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 3000: .Text = "ALAMAT": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 2000: .Text = "TELEPON": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 2000: .Text = "REKENING": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From P001"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari("Kode"): .CellAlignment = 4
              .Col = 1: .Text = RCari("Nama")
              .Col = 2: .Text = RCari("Alamat")
              .Col = 3: .Text = RCari("Telepon"): .CellAlignment = 4
              .Col = 4: .Text = RCari("Rekening"): .CellAlignment = 4
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid_doubleClick()
grid.Col = 0
KB = ""
Clipboard.SetText (grid.Text)
KB = grid.Text

cmdOK.Visible = False
cmdEDIT.Visible = True

Call IsiText

End Sub

Private Sub IsiText()
SCari2 = "Select * From P001 where Kode = '" + Trim(KB) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Text1 = RCari2("Kode")
    Text2 = RCari2("Nama")
    Text3 = RCari2("Alamat")
    Text4 = RCari2("Telepon")
    Text5 = RCari2("Rekening")
End If
RCari2.Close
Set RCari2 = Nothing
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
Text1.BackColor = RGB(255, 255, 255)
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
Text2.BackColor = RGB(255, 255, 255)
Text2 = Format(Text2, ">")
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
Text3.BackColor = RGB(255, 255, 255)
Text3 = Format(Text3, ">")
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
Text4.BackColor = RGB(255, 255, 255)
Text4 = Format(Text4, ">")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text5_GotFocus()
Text5.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text5_LostFocus()
Text5.BackColor = RGB(255, 255, 255)
Text5 = Format(Text5, ">")
End Sub






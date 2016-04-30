VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Laporan_Bulanan 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbBulan 
      Height          =   315
      ItemData        =   "Laporan_Bulanan.frx":0000
      Left            =   1800
      List            =   "Laporan_Bulanan.frx":0028
      TabIndex        =   4
      Text            =   "April"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox tahun 
      Height          =   315
      ItemData        =   "Laporan_Bulanan.frx":008F
      Left            =   1800
      List            =   "Laporan_Bulanan.frx":00B4
      TabIndex        =   3
      Text            =   "2016"
      Top             =   1920
      Width           =   1215
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "D:\Digital Printing 3\lapbulanan.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lihat Laporan"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Bulan"
      BeginProperty Font 
         Name            =   "GeoSlab703 Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "LAPORAN BULANAN"
      BeginProperty Font 
         Name            =   "GeoSlab703 Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Menu mnPemesanan 
      Caption         =   "Pemesanan"
   End
   Begin VB.Menu mnLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnLaporanBulanan 
         Caption         =   "Laporan Bulanan"
      End
      Begin VB.Menu mnLaporanHarian 
         Caption         =   "Laporan Harian"
      End
   End
   Begin VB.Menu mnExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Laporan_Bulanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

On Error GoTo pesan:

If tahun.Text = "2011" Then
thn = 2011
ElseIf tahun.Text = "2012" Then
thn = 2012
ElseIf tahun.Text = "2013" Then
thn = 2013
ElseIf tahun.Text = "2014" Then
thn = 2014
ElseIf tahun.Text = "2015" Then
thn = 2015
ElseIf tahun.Text = "2016" Then
thn = 2016
ElseIf tahun.Text = "2017" Then
thn = 2017
ElseIf tahun.Text = "2018" Then
thn = 2018
ElseIf tahun.Text = "2019" Then
thn = 2019
ElseIf tahun.Text = "2020" Then
thn = 2020
ElseIf tahun.Text = "2021" Then
thn = 2021
ElseIf tahun.Text = "2022" Then
thn = 2022
End If

If cmbBulan.Text = "Januari" Then
bulan = 1
ElseIf cmbBulan.Text = "Februari" Then
bulan = 2
ElseIf cmbBulan.Text = "Maret" Then
bulan = 3
ElseIf cmbBulan.Text = "April" Then
bulan = 4
ElseIf cmbBulan.Text = "Mei" Then
bulan = 5
ElseIf cmbBulan.Text = "Juni" Then
bulan = 6
ElseIf cmbBulan.Text = "Juli" Then
bulan = 7
ElseIf cmbBulan.Text = "Agustus" Then
bulan = 8
ElseIf cmbBulan.Text = "Semptember" Then
bulan = 9
ElseIf cmbBulan.Text = "Oktober" Then
bulan = 10
ElseIf cmbBulan.Text = "November" Then
bulan = 11
ElseIf cmbBulan.Text = "Desember" Then
bulan = 12

End If


CrystalReport1.ReportFileName = App.Path & "\lapbulanan.rpt"
CrystalReport1.RetrieveDataFiles

CrystalReport1.SelectionFormula = _
    " YEAR({Pemesanan.Tanggal})= " & thn & _
    " and month({Pemesanan.Tanggal})= " & bulan & ""

CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
pesan:
If (Err.Number = 20533) Then
MsgBox "Report Gak Connect Ke Database...!", vbCritical, App.Title
Exit Sub
End If
End Sub

Private Sub mnExit_Click()
Unload Me
End Sub

Private Sub mnLaporanBulanan_Click()
Laporan_Bulanan.Show
End Sub

Private Sub mnLaporanHarian_Click()
Laporan_Harian.Show
Unload Me
End Sub

Private Sub mnPemesanan_Click()
Form_Pemesanan.Show
Unload Me
End Sub

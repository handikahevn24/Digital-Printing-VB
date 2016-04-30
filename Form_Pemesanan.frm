VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Pemesanan 
   Caption         =   "Pemesanan"
   ClientHeight    =   10320
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8280
      Top             =   120
   End
   Begin MSAdodcLib.Adodc Ado_Pemesanan 
      Height          =   615
      Left            =   3000
      Top             =   8400
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Digital Printing 3\Digital_Printing.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Digital Printing 3\Digital_Printing.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Pemesanan"
      Caption         =   "Ado_Pemesanan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form_Pemesanan.frx":0000
      Height          =   1455
      Left            =   1200
      TabIndex        =   35
      Top             =   9120
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "URAIAN"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8895
      Begin VB.TextBox txtTotalbayar 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   38
         Top             =   6240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker tanggal 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   120717315
         CurrentDate     =   42481
      End
      Begin VB.TextBox txtNominaldiskon 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         TabIndex        =   33
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtQuantity 
         Height          =   405
         Left            =   2400
         TabIndex        =   12
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox txtUkuranl 
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         TabIndex        =   7
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtUkuranp 
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtJumlah 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   13
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox txtKeterangan 
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   11
         Top             =   4680
         Width           =   6135
      End
      Begin VB.TextBox txtDp 
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   10
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox txtBiayaedit 
         Height          =   405
         Left            =   2400
         TabIndex        =   8
         Top             =   3360
         Width           =   1935
      End
      Begin VB.ComboBox cmbDiskon 
         Height          =   315
         ItemData        =   "Form_Pemesanan.frx":001C
         Left            =   2400
         List            =   "Form_Pemesanan.frx":0029
         TabIndex        =   9
         Text            =   "0%"
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox cmbBahan 
         Height          =   315
         ItemData        =   "Form_Pemesanan.frx":003C
         Left            =   2400
         List            =   "Form_Pemesanan.frx":0049
         TabIndex        =   4
         Text            =   "Pilih Bahan"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "SIMPAN"
         Height          =   375
         Left            =   7200
         TabIndex        =   28
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "PRINT"
         Height          =   375
         Left            =   7200
         TabIndex        =   27
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton btnHapus 
         Caption         =   "HAPUS LAYAR"
         Height          =   375
         Left            =   5400
         TabIndex        =   26
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmdHitung 
         Caption         =   "HITUNG"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txtHargapermeter 
         Height          =   405
         Left            =   2400
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtNama 
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtNonota 
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "TOTAL BAYAR"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   31
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   30
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "JUMLAH HARGA"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "KETERANGAN"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "DP"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "DISKON"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "BIAYA EDIT"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "UKURAN"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "HARGA PARAMETER"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "BAHAN"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "NAMA"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "TANGGAL MASUK"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "NO. NOTA"
         BeginProperty Font 
            Name            =   "GeoSlab703 Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label TitleApp 
      Caption         =   "DIGITAL PRINTING"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   37
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lbltanggal 
      Caption         =   "Jam"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label jam 
      Caption         =   "Jam"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   720
      Width           =   1095
   End
   Begin VB.Menu mnPemesanan 
      Caption         =   "Pemesanan"
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnLaporanbulanan 
         Caption         =   "Laporan Bulanan"
      End
      Begin VB.Menu mnlaporanharian 
         Caption         =   "Laporan Harian"
      End
   End
   Begin VB.Menu mnExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form_Pemesanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hapuslayar()
txtNonota.Text = ""
tanggal.Value = Now
txtNama.Text = ""
cmbBahan.Text = "Pilih Bahan"
txtHargapermeter.Text = ""
txtQuantity.Text = ""
txtUkuranp.Text = ""
txtUkuranl.Text = ""
txtBiayaedit.Text = ""
cmbDiskon.Text = "0%"
txtDp.Text = ""
txtNominaldiskon.Text = ""
txtJumlah.Text = ""
txtKeterangan.Text = ""
End Sub
Sub cekdb()

End Sub

Private Sub simpan()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
rs.CursorType = adOpenDynamic

conn.Provider = "microsoft.jet.oledb.4.0"
conn.CursorLocation = adUseClient
conn.Open App.Path & "\Digital_Printing.mdb"
rs.Open "Select * From Pemesanan", conn, , , adCmdText
rs.Filter = " No_Nota= '" & txtNonota.Text & "'"

If Not rs.EOF Then
MsgBox "Data sudah Ada", vbInformation, Info
txtNonota.SetFocus
Else

On Error Resume Next
With Ado_Pemesanan.Recordset
    'Simpan data
    .AddNew
    !No_Nota = txtNonota.Text
    !tanggal = tanggal.Value
    !Pelanggan = txtNama.Text
    !Jenis_Bahan = cmbBahan.Text
    !Quantity = txtQuantity.Text
    !Ukuran = txtUkuranp.Text + " M x " + txtUkuranl.Text + " M"
    !Nama_Edit = txtNama.Text
    !Biaya_Edit = txtBiayaedit.Text
    !Jumlah_Harga = txtJumlah.Text
    !Keterangan = txtKeterangan.Text
    .Update
    MsgBox "Data Berhasil Disimpan", vbInformation, "Pemesanan"
End With
On Error GoTo 0
End If
End Sub
Private Sub btnHapus_Click()
Call hapuslayar
End Sub

Private Sub btnPrint_Click()
Call cetak
End Sub


Sub cek()
If txtNonota.Text = "" Or txtNama.Text = "" Or cmbBahan.Text = "Pilih Bahan" Or txtHargapermeter.Text = "" Or txtUkuranp.Text = "" Or txtUkuranl.Text = "" Or txtQuantity.Text = "" Then
MsgBox "Data tidak lengkap", vbInformation

    If txtNonota.Text = "" Then
    txtNonota.BackColor = vbRed
    txtNonota.SetFocus
            
    ElseIf txtNama.Text = "" Then
    txtNama.BackColor = vbRed
    txtNama.SetFocus
    
    ElseIf cmbBahan.Text = "Pilih Bahan" Then
    cmbBahan.BackColor = vbRed
    cmbBahan.SetFocus
    
    ElseIf txtHargapermeter.Text = "" Then
    txtHargapermeter.BackColor = vbRed
    txtHargapermeter.SetFocus
    
    ElseIf txtUkuranp.Text = "" Then
    txtUkuranp.BackColor = vbRed
    txtUkuranp.SetFocus
    
    ElseIf txtUkuranl.Text = "" Then
    txtUkuranl.BackColor = vbRed
    txtUkuranl.SetFocus
    
    ElseIf txtQuantity.Text = "" Then
    txtQuantity.BackColor = vbRed
    txtQuantity.SetFocus
    End If
End If

If Len(txtNonota.Text) > 0 Then
txtNonota.BackColor = &H8000000F
End If
If Len(txtNama.Text) > 0 Then
txtNama.BackColor = &H8000000F
End If
If Len(cmbBahan.Text) > 0 Then
cmbBahan.BackColor = &H8000000F
End If
If Len(txtHargapermeter.Text) > 0 Then
txtHargapermeter.BackColor = &H8000000F
End If
If Len(txtUkuranp.Text) > 0 Then
txtUkuranp.BackColor = &H8000000F
End If
If Len(txtUkuranl.Text) > 0 Then
txtUkuranl.BackColor = &H8000000F
End If
If Len(txtQuantity.Text) > 0 Then
txtQuantity.BackColor = &H8000000F
End If
End Sub

Private Sub cmdHitung_Click()
Dim harga_bahan, hargapermeter  As Integer
Dim jumlah_diskon As Double
jbahan = cmbBahan.Text
hmeter = Val(txtHargapermeter.Text)
Diskon = cmbDiskon.Text
DP = Val(txtDp.Text)
Biaya_Edit = Val(txtBiayaedit.Text)
Jumlah_Harga = txtJumlah.Text
nominaldiskon = txtNominaldiskon.Text

Call cek

If jbahan = "Spanslik" Then
harga_bahan = 10000
ElseIf jbahan = "Korea" Then
harga_bahan = 15000
ElseIf jbahan = "Jerman" Then
harga_bahan = 20000
End If

If Diskon = "10%" Then
jumlah_diskon = 0.1
ElseIf Diskon = "20%" Then
jumlah_diskon = 0.2
ElseIf Diskon = "50%" Then
jumlah_diskon = 0.5
End If



'Rumus Perhitungan'
meter = Val(txtUkuranp.Text) * Val(txtUkuranl.Text)
harga = hmeter * meter
Jumlah_Harga = harga_bahan + harga + Biaya_Edit - DP
nominaldiskon = Jumlah_Harga * jumlah_diskon
txtNominaldiskon.Text = nominaldiskon
total_harga = Jumlah_Harga - nominaldiskon
txtJumlah.Text = total_harga
txtTotalbayar = Val(txtJumlah.Text) * Val(txtQuantity.Text)
txtJumlah.BackColor = vbYellow



End Sub


Public Sub lapbulanan()
Dim bln As Date
On Error GoTo pesan:
bln = DTPicker1.Value
CrystalReport1.ReportFileName = App.Path & "\lapbulanan.rpt"
CrystalReport1.RetrieveDataFiles

CrystalReport1.SelectionFormula = _
    " YEAR({Pemesanan.Tanggal})= " & Year(bln) & _
    " and month({Pemesanan.Tanggal})= " & Month(bln) & ""

CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 0
pesan:
If (Err.Number = 20533) Then
MsgBox "Report Gak Connect Ke Database...!", vbCritical, App.Title
Exit Sub
End If
End Sub



Private Sub cmdSimpan_Click()
Call simpan
End Sub


Private Sub Form_Load()
tanggal.Value = Now
lbltanggal.Caption = Format(Date, "Long Date")
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
End Sub

Sub cetak()
No_Nota = txtNonota.Text
Nama = txtNama.Text
tanggal = tanggal.Value
Bahan = cmbBahan.Text
Ukuran = txtUkuranp.Text & " M x " & txtUkuranl.Text & " M"
BiayaEdit = Val(txtBiayaedit.Text)
Diskon = Val(txtNominaldiskon.Text)
DP = Val(txtDp.Text)
Quantity = Val(txtQuantity.Text)
Jumlah = Val(txtJumlah.Text)
Keterangan = txtKeterangan.Text
Total = Val(txtTotalbayar.Text)

Form_Print.Font = "courier new"
Form_Print.Show
     Form_Print.CurrentX = 0
     Form_Print.CurrentY = 0
     Form_Print.FontSize = 9.5
     Form_Print.Print Tab(6); Tab(15); "Digital Printing";
     Form_Print.Print Tab(3); "No Nota: "; No_Nota; "              ";
     Form_Print.Print Tab(2); "==========================================";
     Form_Print.Print Tab(3); "NAMA"; Tab(20); ": "; Tab(23); Nama;
     Form_Print.Print Tab(3); "TANGGAL"; Tab(20); ": "; Tab(23); tanggal;
     Form_Print.Print Tab(3); "BAHAN"; Tab(20); ": "; Tab(23); Bahan;
     Form_Print.Print Tab(3); "UKURAN"; Tab(20); ": "; Tab(23); Ukuran;
     Form_Print.Print Tab(3); "BIAYA EDIT"; Tab(20); ":"; Tab(23); "Rp."; BiayaEdit;
     Form_Print.Print Tab(3); "DISKON"; Tab(20); ":"; Tab(23); "Rp."; Diskon;
     Form_Print.Print Tab(3); "DP"; Tab(20); ":"; Tab(23); "Rp."; DP;
     Form_Print.Print Tab(3); "JUMLAH"; Tab(20); ":"; Tab(23); "Rp."; Jumlah;
     Form_Print.Print Tab(3); "QUANTITY"; Tab(20); ":"; Tab(22); Quantity;
     Form_Print.Print Tab(3); "TOTAL BAYAR"; Tab(20); ":"; Tab(23); "Rp."; Total;
     Form_Print.Print Tab(3); "KETERANGAN"; Tab(20); ": "; Tab(23); Keterangan;
     
     Form_Print.Print Tab(2); "==========================================";
Form_Print.Font = "Courier New"
    Form_Print.Print Tab(2); "==========================================";
    Form_Print.FontSize = 12
Form_Print.Print Tab(13); "Terimakasih";
End Sub

Private Sub Timer1_Timer()
jam.Caption = Time
End Sub


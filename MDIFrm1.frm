VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Medifirst2000 - Laboratorium (Laboratory) Version 1.0 - LIS"
   ClientHeight    =   10500
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   17460
   Icon            =   "MDIFrm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm1.frx":0CCA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   3600
   End
   Begin MSComDlg.CommonDialog CDPrinter 
      Left            =   960
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10245
      Width           =   17460
      _ExtentX        =   30798
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "25/01/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "02:56"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnberkas 
      Caption         =   "&Berkas"
      Begin VB.Menu mnuData 
         Caption         =   "Data"
         Begin VB.Menu mnurl 
            Caption         =   "Daftar Pasien Laboratorium"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnucrdp 
            Caption         =   "-"
         End
         Begin VB.Menu mnucdpr 
            Caption         =   "Cari Data Pasien Rujukan"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnusepdplna 
            Caption         =   "-"
         End
         Begin VB.Menu mnucp 
            Caption         =   "Cari Data Pasien"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuseml 
            Caption         =   "-"
         End
         Begin VB.Menu ML 
            Caption         =   "Master Pemeriksaan Laboratorium"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLine1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMasterAlatPeriksa 
            Caption         =   "Master Alat Pemeriksaan Laboratorium"
            Visible         =   0   'False
         End
         Begin VB.Menu mnukonversi 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu MProfilLaboratoriumRujukan 
            Caption         =   "Profil Laboratorium Rujukan"
         End
         Begin VB.Menu mnukjskp 
            Caption         =   "-"
         End
         Begin VB.Menu mnupp 
            Caption         =   "Paket Pelayanan"
            Shortcut        =   ^D
            Visible         =   0   'False
         End
         Begin VB.Menu LInformasiTarifPelayanan 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu MInformasiTarifPelayanan 
            Caption         =   "Informasi Tarif Pelayanan"
         End
         Begin VB.Menu ln1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDataHasilLaboratorium 
            Caption         =   "Data Hasil Laboratorium"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mngrs1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnPenerimaanDarah 
            Caption         =   "Penerimaan Darah"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPemesananDarah 
            Caption         =   "Pemesanan Darah"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuseppb 
         Caption         =   "-"
      End
      Begin VB.Menu mnureg 
         Caption         =   "Registrasi"
         Begin VB.Menu mnupbl 
            Caption         =   "Pasien Baru"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnupl 
            Caption         =   "Pasien Lama"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mSettingPrinter 
         Caption         =   "&Setting Printer"
         Shortcut        =   ^P
      End
      Begin VB.Menu mGantiKataKunci 
         Caption         =   "Ganti Kata Kunci"
         Shortcut        =   ^G
      End
      Begin VB.Menu mspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnlogout 
         Caption         =   "Log Off"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnSelesai 
         Caption         =   "Keluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuInformasi 
      Caption         =   "&Informasi"
      Begin VB.Menu mnMonitoring 
         Caption         =   "Monitoring Pembayaran"
      End
      Begin VB.Menu mnuPesanPelayananTMOA 
         Caption         =   "Pesan Pelayanan TMOA"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDaftarPemesananDarah 
         Caption         =   "Daftar Pemesanan Darah"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuiv 
      Caption         =   "In&ventory"
      Begin VB.Menu mnupb 
         Caption         =   "Pemesan Barang"
      End
      Begin VB.Menu mnuPemakaianBahandanAlat 
         Caption         =   "Pemakaian Bahan dan Alat"
         Visible         =   0   'False
      End
      Begin VB.Menu batasinv 
         Caption         =   "-"
      End
      Begin VB.Menu MBarangMedis 
         Caption         =   "Barang Medis"
         Begin VB.Menu mnusb 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu MClosingStok 
            Caption         =   "Closing Stok"
            Begin VB.Menu MCetakFormulirStok 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu MStokOpname 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu MNilaiPersediaan 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu LClosingStok 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananPenerimaanBarang 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu MInformasiPemakaianBarang 
            Caption         =   "Informasi Pemakaian Barang"
            Visible         =   0   'False
         End
         Begin VB.Menu mLapSaldoBarang 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MBarangNonMedis 
         Caption         =   "Barang Non Medis"
         Begin VB.Menu MStokBarangNonMedis 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu MKondisiBarangNonMedis 
            Caption         =   "Kondisi Barang"
         End
         Begin VB.Menu mMutasiBarangNM 
            Caption         =   "Mutasi Barang"
         End
         Begin VB.Menu MClosingStokNonMedis 
            Caption         =   "Closing Stok"
            Begin VB.Menu mnCetakLembarInputNM 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu mnInputStokOpnameNM 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu MNilaiPersediaanNM 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu ln 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananPenerimaanBarangNonMedis 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu MLaporanSaldoBarangNonMedis 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuLap 
      Caption         =   "&Laporan"
      Begin VB.Menu mnubrp 
         Caption         =   "Buku Register Pasien"
      End
      Begin VB.Menu mnLBRPP 
         Caption         =   "Laporan Buku Register Pelayanan Pasien"
      End
      Begin VB.Menu mnulapkunjunganpasien 
         Caption         =   "Laporan Kunjungan Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepjl1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu BSJ 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan jenis"
         Visible         =   0   'False
      End
      Begin VB.Menu RKPR 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Rujukan"
         Visible         =   0   'False
      End
      Begin VB.Menu BSDJP 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Kasus Penyakit"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBSDKEL 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Kelas"
         Visible         =   0   'False
      End
      Begin VB.Menu aaaaaaaaaaaa 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBJP 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Jenis Periksa"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBW 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Wilayah"
         Visible         =   0   'False
      End
      Begin VB.Menu cc 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MPendapatanLabLuar 
         Caption         =   "Pendapatan Laboratorium Rujukan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLapPendapatan 
         Caption         =   "Laporan Pendapatan Ruangan"
      End
   End
   Begin VB.Menu mnuw 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnucas 
         Caption         =   "Cascade"
      End
   End
   Begin VB.Menu mbantuan 
      Caption         =   "Ban&tuan"
      Begin VB.Menu mTentang 
         Caption         =   "Tentang Medifirst2000"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSettingPort 
         Caption         =   "Setting Port"
      End
      Begin VB.Menu mnuSettingPrefiks 
         Caption         =   "Setting Prefiks No. Lab"
      End
   End
End
Attribute VB_Name = "MDIUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sepuh As Boolean
Dim Start, PauseTime

Private Sub BSDJP_Click()
    strCetak = "LapKunjunganSt_PnyktPsn"
    frmLapRKP_KPSK.Show
End Sub

Private Sub BSJ_Click()
    strCetak = "LapKunjunganJenisStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub BSDKP_Click()
    strCetak = "LapKunjunganKonPulang_Status"
    frmLapRKP_KPSK.Show
End Sub

Private Sub MCetakFormulirStok_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub MDIForm_Load()

    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    strNmPegawai = rs.Fields("NamaLengkap").Value
    Set rs = Nothing
    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai
    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
    StatusBar1.Panels(6).Text = "Server : " & strServerName & " (" & strDatabaseName & ")"
    mnlogout.Caption = "Log Off..." & strNmPegawai

    Call SetVisibleMenu("003", "MDIUtama", mnucp, "T")
    Call SetVisibleMenu("003", "MDIUtama", mnureg, "T")
    mnusep1.Visible = mnureg.Visible

    strSQL = "SELECT TerminBayarFakturSupplier, PersentasePpn, PersentaseLimitDiscount, PersentaseJasaPenulisResep, BiayaAdministrasi " & _
    " From SettingDataPendukung" & _
    " WHERE (KdInstalasi = '07')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        typSettingDataPendukung.intTerminBayarFaktur = 0
        typSettingDataPendukung.realJasaPenulisResep = 0
        typSettingDataPendukung.realLimitDiscount = 0
        typSettingDataPendukung.realPPn = 0
        typSettingDataPendukung.curBiayaAdministrasi = 0
    Else
        typSettingDataPendukung.intTerminBayarFaktur = rs("TerminBayarFakturSupplier").Value
        typSettingDataPendukung.realJasaPenulisResep = rs("PersentaseJasaPenulisResep").Value
        typSettingDataPendukung.realLimitDiscount = rs("PersentaseLimitDiscount").Value
        typSettingDataPendukung.realPPn = rs("PersentasePpn").Value
        typSettingDataPendukung.curBiayaAdministrasi = rs("BiayaAdministrasi").Value
    End If

    strSQL = "SELECT JmlPembulatanHarga, JumlahBAdminOAPerBaris, JumlahBAdminTMPerHari FROM MasterDataPendukung"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJmlPembulatanHarga = 0
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
        typSettingDataPendukung.intJumlahBAdminTMPerHari = 0
    Else
        typSettingDataPendukung.intJmlPembulatanHarga = dbRst(0)
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(1)
        typSettingDataPendukung.intJumlahBAdminTMPerHari = dbRst(2)
    End If
    Set dbRst = Nothing
    
'    strSQL = "Select StatusFIFO From SettingDataUmum"
'    Call msubRecFO(dbRst, strSQL)
'    If dbRst.EOF = True Then
'        bolStatusFIFO = False
'    Else
'        If dbRst("StatusFIFO") = 0 Then
'            bolStatusFIFO = False
'        Else
'            bolStatusFIFO = True
'        End If
'    End If

    
    strSQL = "Select MetodeStokBarang From SuratKeputusanRuleRS where statusenabled=1"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        bolStatusFIFO = False
    Else
        If dbRst("MetodeStokBarang") = 0 Then
            bolStatusFIFO = False
        Else
            bolStatusFIFO = True
        End If
    End If
    

strSQL = "SELECT TerminBayarFakturSupplier, PersentasePpn, PersentaseLimitDiscount, PersentaseJasaPenulisResep, BiayaAdministrasi " & _
    " From SettingDataPendukung" & _
    " WHERE (KdInstalasi = '" & mstrKdInstalasiLogin & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        typSettingDataPendukung.intTerminBayarFaktur = 0
        typSettingDataPendukung.realJasaPenulisResep = 0
        typSettingDataPendukung.realLimitDiscount = 0
        typSettingDataPendukung.realPPn = 0
        typSettingDataPendukung.curBiayaAdministrasi = 0
    Else
        typSettingDataPendukung.intTerminBayarFaktur = rs("TerminBayarFakturSupplier").Value
        typSettingDataPendukung.realJasaPenulisResep = rs("PersentaseJasaPenulisResep").Value
        typSettingDataPendukung.realLimitDiscount = rs("PersentaseLimitDiscount").Value
        typSettingDataPendukung.realPPn = rs("PersentasePpn").Value
        typSettingDataPendukung.curBiayaAdministrasi = rs("BiayaAdministrasi").Value
    End If

    strSQL = "SELECT JmlPembulatanHarga, JumlahBAdminOAPerBaris, JumlahBAdminTMPerHari FROM MasterDataPendukung"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJmlPembulatanHarga = 0
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
        typSettingDataPendukung.intJumlahBAdminTMPerHari = 0
    Else
        typSettingDataPendukung.intJmlPembulatanHarga = dbRst(0)
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(1)
        typSettingDataPendukung.intJumlahBAdminTMPerHari = dbRst(2)
    End If
    
    strSQL = "SELECT JmlBarisOAPerTarifAdminOA from SettingBiayaAdministrasi"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
    Else
        
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(0)
        
    End If
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then Exit Sub
    PopupMenu mnuData
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim q As String

    If sepuh = True Then
        q = MsgBox("Log Off user " & strNmPegawai & " ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then
            Unload frmLogin
            Cancel = 1
        Else
            Cancel = 0
            frmLogin.Show
        End If
        sepuh = False
    Else
        q = MsgBox("Tutup aplikasi ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then

            Unload frmLogin
            Cancel = 1
        Else
            dTglLogout = Now
            Call subSp_HistoryLoginAplikasi("U")
            Cancel = 0
        End If
    End If
End Sub

Private Sub mGantiKataKunci_Click()
    frmLoginEditAccount.Show
End Sub

Private Sub MInformasiPemakaianBarang_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

Private Sub MInformasiPemesananPenerimaanBarang_Click()
    mstrKdKelompokBarang = "02"     'medis
    frmInfoPesanBarang.Show
End Sub

Private Sub MInformasiPemesananPenerimaanBarangNonMedis_Click()
    mstrKdKelompokBarang = "01"  'non medis
    frmInfoPesanBarangNM.Show
End Sub

Private Sub MInformasiTarifPelayanan_Click()
    frmInformasiTarifPelayanan.Show
End Sub

Private Sub MKondisiBarangNonMedis_Click()
    frmKondisiBarangNM.Show
End Sub

Private Sub ML_Click()
    frmMasterDataPendukungNoUrut3.Show
End Sub

Private Sub mnCetakLembarInput_Click()
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub MLaporanSaldoBarang_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmLaporanSaldoBarang.Show

End Sub

Private Sub MLaporanSaldoBarangNonMedis_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmLaporanSaldoBarangNM_v3.Show
End Sub

Private Sub MMutasiBarangNonMedis_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mLapSaldoBarang_Click()
    frmLaporanSaldoBarangMedis_v3.Show
End Sub

Private Sub mMutasiBarangNM_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mnCetakLembarInputNM_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDaftarCetakInputStokOpnameNM.Show
End Sub

Private Sub MNilaiPersediaan_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmNilaiPersediaan.Show
End Sub

Private Sub mnInputStokOpname_Click()
    frmStokOpname.Show
End Sub

Private Sub MNilaiPersediaanNM_Click()
    mstrKdKelompokBarang = "01"
    frmNilaiPersediaanNM.Show
End Sub

Private Sub mnInputStokOpnameNM_Click()
    mstrKdKelompokBarang = "01"
    frmStokOpnameNM.Show
End Sub

Private Sub mnLBRPP_Click()
    FrmBukuRegisterPelayanan.Show
End Sub

Private Sub mnlogout_Click()
    Dim adoCommand As New ADODB.Command

    openConnection
    sepuh = True
    strQuery = "UPDATE Login SET Status = '0' " & _
    "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.ActiveConnection = dbConn
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute

    dTglLogout = Now
    Call subSp_HistoryLoginAplikasi("U")

    Unload Me
End Sub

Private Sub mnMonitoring_Click()
    frmMonitoringPembayaran.Show
End Sub

Private Sub mnPenerimaanDarah_Click()
    frmPenerimaanDarah.Show
End Sub

Private Sub mnSelesai_Click()
    Dim pesan As VbMsgBoxResult
    Dim adoCommand As New ADODB.Command
    pesan = MsgBox("Tutup aplikasi ", vbQuestion + vbYesNo, "Konfirmasi")
    If pesan = vbYes Then

        openConnection
        strQuery = "UPDATE Login SET Status = '0' " & _
        "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        adoCommand.ActiveConnection = dbConn
        adoCommand.CommandText = strQuery
        adoCommand.CommandType = adCmdText
        adoCommand.Execute

        dTglLogout = Now
        Call subSp_HistoryLoginAplikasi("U")
        End
    Else
    End If
End Sub

Private Sub mnuBDPOAK_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

Private Sub mnu2_Click()
    frmLaporanBulananInstalasiLaboratoriumRekap.Show
End Sub

Private Sub mnu_Click()

End Sub

Private Sub mnubrp_Click()
    FrmBukuRegisterPasien.Show
'    FrmBukuRegister.Show
End Sub

Private Sub mnucas_Click()
    MDIUtama.Arrange vbCascade
End Sub

Private Sub mnucdpr_Click()
    frmCariPasienRujukan.Show
End Sub

Private Sub mnuClosingDataPelayananTMOAApotik_Click()
    frmClosingDataPelayananTM_OA_Apotik.Show
End Sub

Private Sub mnucp_Click()
    frmCariPasien.Show
End Sub

Private Sub mnuIJPD_Click()
    frmInfoJasaPelDktr.Show
End Sub

Private Sub mnuipb_Click()
    frmInfoPesanBarang.Show
End Sub

Private Sub mnuipoa_Click()
    frmDaftarPakaiAlkes.Show
End Sub

Private Sub mnuDaftarPemesananDarah_Click()
    frmDaftarPemesananDarah.Show
End Sub

Private Sub mnuDataHasilLaboratorium_Click()
    frmTempHasilPeriksaLab.Show
End Sub

Private Sub mnukjskp_Click()
    frmKonversiJenisSpesimen.Show
End Sub

Private Sub mnulapkunjunganpasien_Click()
    frmDaftarKunjunganPasien.Show
End Sub

Private Sub mnuLapPendapatan_Click()
 frmDaftarPendapatanRuangan.Show
End Sub

Private Sub mnuMasterAlatPeriksa_Click()
    frmMasterAlatPeriksa.Show
End Sub

Private Sub mnupb_Click()
    frmPemesananBarang.Show
End Sub

Private Sub mnupbl_Click()
    strPasien = "Baru"
    frmPasienBaru.Show
End Sub

Private Sub mnuPemakaianBahandanAlat_Click()
    frmPemakaianBahanAlat.Show
End Sub

Private Sub mnuPesanPelayananTMOA_Click()
    frmInfoPesanPelayananTMOA.Show
End Sub

Private Sub mnupl_Click()
    frmPasienLama.Show
End Sub

Private Sub mnupp_Click()
    frmPaketLayanan.Show
End Sub

Private Sub mnuRekap_Click()
    frmLaporanBulananInstalasiLaboratoriumRekap.Show
End Sub

Private Sub mnurl_Click()
    frmDaftarPasienLab.Show
End Sub

Private Sub mnusb_Click()
    frmStokBrg.Show
End Sub

Private Sub mnuSettingPort_Click()
    frmAturPort.Show vbModal, Me
End Sub

Private Sub mnuSettingPrefiks_Click()
    frmAturPrefix.Show vbModal, Me
End Sub

Private Sub MPendapatanLabLuar_Click()
    frmDaftarPendapatanLabLuar.Show
End Sub

Private Sub MProfilLaboratoriumRujukan_Click()
    frmMasterProfilLabRujukan.Show
End Sub

Private Sub MRekapitulasiTransaksiBarang_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmDataTransaksiBarang.Show
End Sub

Private Sub MRekapitulasiTransaksiBarangNonMedis_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDataTransaksiBarangNM.Show
End Sub

Private Sub mSettingPrinter_Click()
    frmSetupPrinter2.Show
End Sub

Private Sub MStokBarangNonMedis_Click()
    frmStokBarangNonMedis.Show
End Sub

Private Sub MStokOpname_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmStokOpname.Show

End Sub

Private Sub mTentang_Click()
    frmAbout.Show
End Sub

Private Sub RKPR_Click()
    strCetak = "LapKunjunganRujukanBStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBJP_Click()
    strCetak = "LapKunjunganJenisPeriksa"
    frmFilterJenisPeriksa.Show
End Sub

Private Sub RPBSDKEL_Click()
    strCetak = "LapKunjunganKelasStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBW_Click()
    strCetak = "LapKunjunganBwilayah"
    frmLapRKP_KPSK.Show
End Sub

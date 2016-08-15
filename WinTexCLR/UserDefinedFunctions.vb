Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Collections
Imports System.Diagnostics
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.VisualBasic

Partial Public Class UserDefinedFunctions

    Private Structure oSiparis
        Dim cSiparisNo As String
        Dim dSiparisTarihi As Date
        Dim dSonSevkTarihi As Date
        Dim cMusteriNo As String
        Dim cSorumlu As String
        Dim cModelNo As String
        Dim cRenk As String
        Dim cSiparisTakipNotlari As String
    End Structure

    Private Class SiparisDurumu1

        Public Siparis_Planlanan_Svk_Hft As SqlDecimal
        Public Siparis_Son_Svk_Hft As SqlDecimal
        Public Siparis_Gelis_Tarihi As SqlDateTime
        Public Siparis_Planlanan_Svk_Tarihi As SqlDateTime
        Public Son_Sevk_Tarihi As SqlDateTime
        Public Sevkiyat_Siparis_Gun_Frk As SqlDecimal
        Public Musteri As SqlString
        Public Musteri_Temsilcisi As SqlString
        Public Siparis_No As SqlString
        Public Model_No As SqlString
        Public Renk As SqlString
        Public Siparis_Durumu As SqlString
        Public Kumas_Tedarikcisi As SqlString
        Public Kumas_Termini As SqlDateTime
        Public Kumas_Gelis_Tarihi As SqlDateTime
        Public Kumas_Tamamlanma_Yuzdesi As SqlDecimal
        Public Aksesuar_Durumu As SqlString
        Public Kesim_OK_Tarihi As SqlDateTime
        Public Kesim_Tarihi As SqlDateTime
        Public Siparis_Adedi As SqlDecimal
        Public Kesim_Adedi As SqlDecimal
        Public Dikim_Giren As SqlDecimal
        Public Dikim_Adedi As SqlDecimal
        Public Sevk_Adedi As SqlDecimal
        Public Kesim_Sevkiyat_Farki As SqlDecimal
        Public Dikim_Sevkiyat_Farki As SqlDecimal
        Public Sevkiyat_Siparis_Farki As SqlDecimal
        Public Kesim_Yuzdesi As SqlDecimal
        Public Dikim_Yuzdesi As SqlDecimal
        Public Sevkiyat_Yuzdesi As SqlDecimal
        Public Sevkiyat_Yuzdesi2 As SqlDecimal
        Public Kesim_Atolyesi As SqlString
        Public Dikim_Atolyesi As SqlString
        Public Gerceklesen_Sevkiyat As SqlDateTime
        Public Ikinci_Kalite As SqlDecimal
        Public Atolye_Iade As SqlDecimal
        Public Uretim_Kaybi As SqlDecimal
        Public SiparisTakipNotlari As SqlString

        Public Sub New(Optional Siparis_Planlanan_Svk_Hft_1 As SqlDecimal = Nothing, _
                       Optional Siparis_Son_Svk_Hft_1 As SqlDecimal = Nothing, _
                       Optional Siparis_Gelis_Tarihi_1 As SqlDateTime = Nothing, _
                       Optional Siparis_Planlanan_Svk_Tarihi_1 As SqlDateTime = Nothing, _
                       Optional Son_Sevk_Tarihi_1 As SqlDateTime = Nothing, _
                       Optional Sevkiyat_Siparis_Gun_Frk_1 As SqlDecimal = Nothing, _
                       Optional Musteri_1 As SqlString = Nothing, _
                       Optional Musteri_Temsilcisi_1 As SqlString = Nothing, _
                       Optional Siparis_No_1 As SqlString = Nothing, _
                       Optional Model_No_1 As SqlString = Nothing, _
                       Optional Renk_1 As SqlString = Nothing, _
                       Optional Siparis_Durumu_1 As SqlString = Nothing, _
                       Optional Kumas_Tedarikcisi_1 As SqlString = Nothing, _
                       Optional Kumas_Termini_1 As SqlDateTime = Nothing, _
                       Optional Kumas_Gelis_Tarihi_1 As SqlDateTime = Nothing, _
                       Optional Kumas_Tamamlanma_Yuzdesi_1 As SqlDecimal = Nothing, _
                       Optional Aksesuar_Durumu_1 As SqlString = Nothing, _
                       Optional Kesim_OK_Tarihi_1 As SqlDateTime = Nothing, _
                       Optional Kesim_Tarihi_1 As SqlDateTime = Nothing, _
                       Optional Siparis_Adedi_1 As SqlDecimal = Nothing, _
                       Optional Kesim_Adedi_1 As SqlDecimal = Nothing, _
                       Optional Dikim_Giren_1 As SqlDecimal = Nothing, _
                       Optional Dikim_Adedi_1 As SqlDecimal = Nothing, _
                       Optional Sevk_Adedi_1 As SqlDecimal = Nothing, _
                       Optional Kesim_Sevkiyat_Farki_1 As SqlDecimal = Nothing, _
                       Optional Dikim_Sevkiyat_Farki_1 As SqlDecimal = Nothing, _
                       Optional Sevkiyat_Siparis_Farki_1 As SqlDecimal = Nothing, _
                       Optional Kesim_Yuzdesi_1 As SqlDecimal = Nothing, _
                       Optional Dikim_Yuzdesi_1 As SqlDecimal = Nothing, _
                       Optional Sevkiyat_Yuzdesi_1 As SqlDecimal = Nothing, _
                       Optional Sevkiyat_Yuzdesi2_1 As SqlDecimal = Nothing, _
                       Optional Kesim_Atolyesi_1 As SqlString = Nothing, _
                       Optional Dikim_Atolyesi_1 As SqlString = Nothing, _
                       Optional Gerceklesen_Sevkiyat_1 As SqlDateTime = Nothing, _
                       Optional Ikinci_Kalite_1 As SqlDecimal = Nothing, _
                       Optional Atolye_Iade_1 As SqlDecimal = Nothing, _
                       Optional Uretim_Kaybi_1 As SqlDecimal = Nothing, _
                       Optional SiparisTakipNotlari_1 As SqlString = Nothing)

            Siparis_Planlanan_Svk_Hft = Siparis_Planlanan_Svk_Hft_1
            Siparis_Son_Svk_Hft = Siparis_Son_Svk_Hft_1
            Siparis_Gelis_Tarihi = Siparis_Gelis_Tarihi_1
            Siparis_Planlanan_Svk_Tarihi = Siparis_Planlanan_Svk_Tarihi_1
            Son_Sevk_Tarihi = Son_Sevk_Tarihi_1
            Sevkiyat_Siparis_Gun_Frk = Sevkiyat_Siparis_Gun_Frk_1
            Musteri = Musteri_1
            Musteri_Temsilcisi = Musteri_Temsilcisi_1
            Siparis_No = Siparis_No_1
            Model_No = Model_No_1
            Renk = Renk_1
            Siparis_Durumu = Siparis_Durumu_1
            Kumas_Tedarikcisi = Kumas_Tedarikcisi_1
            Kumas_Termini = Kumas_Termini_1
            Kumas_Gelis_Tarihi = Kumas_Gelis_Tarihi_1
            Kumas_Tamamlanma_Yuzdesi = Kumas_Tamamlanma_Yuzdesi_1
            Aksesuar_Durumu = Aksesuar_Durumu_1
            Kesim_OK_Tarihi = Kesim_OK_Tarihi_1
            Kesim_Tarihi = Kesim_Tarihi_1
            Siparis_Adedi = Siparis_Adedi_1
            Kesim_Adedi = Kesim_Adedi_1
            Dikim_Giren = Dikim_Giren_1
            Dikim_Adedi = Dikim_Adedi_1
            Sevk_Adedi = Sevk_Adedi_1
            Kesim_Sevkiyat_Farki = Kesim_Sevkiyat_Farki_1
            Dikim_Sevkiyat_Farki = Dikim_Sevkiyat_Farki_1
            Sevkiyat_Siparis_Farki = Sevkiyat_Siparis_Farki_1
            Kesim_Yuzdesi = Kesim_Yuzdesi_1
            Dikim_Yuzdesi = Dikim_Yuzdesi_1
            Sevkiyat_Yuzdesi = Sevkiyat_Yuzdesi_1
            Sevkiyat_Yuzdesi2 = Sevkiyat_Yuzdesi2_1
            Kesim_Atolyesi = Kesim_Atolyesi_1
            Dikim_Atolyesi = Dikim_Atolyesi_1
            Gerceklesen_Sevkiyat = Gerceklesen_Sevkiyat_1
            Ikinci_Kalite = Ikinci_Kalite_1
            Atolye_Iade = Atolye_Iade_1
            Uretim_Kaybi = Uretim_Kaybi_1
            SiparisTakipNotlari = SiparisTakipNotlari_1

        End Sub
    End Class

    <SqlFunction(DataAccess:=DataAccessKind.Read, _
        FillRowMethodName:="SiparisDurumu1_FillRow", _
        TableDefinition:="Siparis_Planlanan_Svk_Hft  decimal(10,0) , " + _
                        "Siparis_Son_Svk_Hft  decimal(10,0) , " + _
                        "Siparis_Gelis_Tarihi  datetime , " + _
                        "Siparis_Planlanan_Svk_Tarihi  datetime , " + _
                        "Son_Sevk_Tarihi  datetime , " + _
                        "Sevkiyat_Siparis_Gun_Frk  decimal(10,0) , " + _
                        "Musteri  nvarchar(30) , " + _
                        "Musteri_Temsilcisi  nvarchar(30) , " + _
                        "Siparis_No  nvarchar(30) , " + _
                        "Model_No  nvarchar(30) , " + _
                        "Renk  nvarchar(30) , " + _
                        "Siparis_Durumu  nvarchar(30) , " + _
                        "Kumas_Tedarikcisi  nvarchar(30) , " + _
                        "Kumas_Termini  datetime , " + _
                        "Kumas_Gelis_Tarihi  datetime , " + _
                        "Kumas_Tamamlanma_Yuzdesi  decimal(10,0) , " + _
                        "Aksesuar_Durumu  nvarchar(30) , " + _
                        "Kesim_OK_Tarihi  datetime , " + _
                        "Kesim_Tarihi  datetime , " + _
                        "Siparis_Adedi  decimal(10,0) , " + _
                        "Kesim_Adedi  decimal(10,0) , " + _
                        "Dikim_Giren  decimal(10,0) , " + _
                        "Dikim_Adedi  decimal(10,0) , " + _
                        "Sevk_Adedi  decimal(10,0) , " + _
                        "Kesim_Sevkiyat_Farki  decimal(10,0) , " + _
                        "Dikim_Sevkiyat_Farki  decimal(10,0) , " + _
                        "Sevkiyat_Siparis_Farki  decimal(10,0) , " + _
                        "Kesim_Yuzdesi  decimal(10,0) , " + _
                        "Dikim_Yuzdesi  decimal(10,0) , " + _
                        "Sevkiyat_Yuzdesi  decimal(10,0) , " + _
                        "Sevkiyat_Yuzdesi2  decimal(10,0) , " + _
                        "Kesim_Atolyesi  nvarchar(30) , " + _
                        "Dikim_Atolyesi  nvarchar(30) , " + _
                        "Gerceklesen_Sevkiyat  datetime, " + _
                        "Ikinci_Kalite  decimal(10,0) , " + _
                        "Atolye_Iade  decimal(10,0) , " + _
                        "Uretim_Kaybi  decimal(10,0) , " + _
                        "SiparisTakipNotlari nvarchar(250)")> _
    Public Shared Function RaporSiparisDurumu1(Optional nOpenClose As Integer = 1, Optional nMonth As Integer = 0) As IEnumerable
        ' nOpenClose = 1    , Açýk
        ' nOpenClose = 2    , Kapali
        ' nOpenClose = 3    , iptal (kapatýlmýþ fakat çeki listesi yok)
        ' nMonth = 0        , Bütün aylar
        ' nMonth = 1..12    , Ocak..Aralýk
        Dim resultCollection As New ArrayList()

        Dim Siparis_Planlanan_Svk_Hft As SqlDecimal
        Dim Siparis_Son_Svk_Hft As SqlDecimal
        Dim Siparis_Gelis_Tarihi As SqlDateTime
        Dim Siparis_Planlanan_Svk_Tarihi As SqlDateTime
        Dim Son_Sevk_Tarihi As SqlDateTime
        Dim Sevkiyat_Siparis_Gun_Frk As SqlDecimal
        Dim Musteri As SqlString
        Dim Musteri_Temsilcisi As SqlString
        Dim Siparis_No As SqlString
        Dim Model_No As SqlString
        Dim Renk As SqlString
        Dim Siparis_Durumu As SqlString
        Dim Kumas_Tedarikcisi As SqlString
        Dim Kumas_Termini As SqlDateTime
        Dim Kumas_Gelis_Tarihi As SqlDateTime
        Dim Kumas_Tamamlanma_Yuzdesi As SqlDecimal
        Dim Aksesuar_Durumu As SqlString
        Dim Kesim_OK_Tarihi As SqlDateTime
        Dim Kesim_Tarihi As SqlDateTime
        Dim Siparis_Adedi As SqlDecimal
        Dim Kesim_Adedi As SqlDecimal
        Dim Dikim_Giren As SqlDecimal
        Dim Dikim_Adedi As SqlDecimal
        Dim Sevk_Adedi As SqlDecimal
        Dim Kesim_Sevkiyat_Farki As SqlDecimal
        Dim Dikim_Sevkiyat_Farki As SqlDecimal
        Dim Sevkiyat_Siparis_Farki As SqlDecimal
        Dim Kesim_Yuzdesi As SqlDecimal
        Dim Dikim_Yuzdesi As SqlDecimal
        Dim Sevkiyat_Yuzdesi As SqlDecimal
        Dim Sevkiyat_Yuzdesi2 As SqlDecimal
        Dim Kesim_Atolyesi As SqlString
        Dim Dikim_Atolyesi As SqlString
        Dim Gerceklesen_Sevkiyat As SqlDateTime
        Dim Ikinci_Kalite As SqlDecimal
        Dim Atolye_Iade As SqlDecimal
        Dim Uretim_Kaybi As SqlDecimal
        Dim SiparisTakipNotlari As SqlString

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim aSiparis() As oSiparis = Nothing
        Dim oSiparisReader As SqlDataReader
        Dim nCnt As Integer = -1

        Dim cUretim As String = ""
        Dim nSipAdet As Double = 0
        Dim nKesim As Double = 0
        Dim nDikim As Double = 0
        Dim nSevkiyat As Double = 0
        Dim cFirma As String = ""
        Dim cAnaKumFirma As String = ""
        Dim nFiyat As Double = 0
        Dim dSevkiyat As Date
        Dim dTermin As Date
        Dim nIhtiyac As Double = 0
        Dim nKarsilanan As Double = 0
        Dim dGiris As Date
        Dim cAksesuar As String = ""
        Dim cSevkDept As String = ""
        Dim dPlSevk As Date
        Dim cAnaKumas As String = ""
        Dim cAnaKumasRenk As String = ""
        Dim cMTF As String = ""
        Dim dKesim_OK_Tarihi As Date
        Dim dKesim_Tarihi As Date
        Dim nDikimGiren As Double = 0
        Dim nIkinciKalite As Double = 0
        Dim nAtolyeIade As Double = 0
        Dim nUretimKaybi As Double = 0
        Dim cKesim As String = ""
        Dim nKesimFiyat As Double = 0
        Dim dGercekSevkTarihi As Date
        Dim lOK As Boolean = False

        Try
            ConnYage = OpenConn()

            cSevkDept = GetSysParConnected("sevkstokdepartmani", ConnYage)

            cSQL = "select distinct a.kullanicisipno, a.siparistarihi, a.sonsevktarihi, a.musterino, a.sorumlu, " + _
                    " b.modelno, b.renk, " + _
                    " gerceksevktarihi = (select top 1 z.sevktar " + _
                                       " from sevkformlines x, sevkformlinesrba y, sevkform z " + _
                                       " where x.sevkformno = y.sevkformno " + _
                                       " and x.sevkformno = z.sevkformno " + _
                                       " and x.ulineno = y.ulineno " + _
                                       " and z.ok = 'E' " + _
                                       " and x.siparisno = a.kullanicisipno " + _
                                       " and x.modelno = b.modelno " + _
                                       " and y.renk = b.renk " + _
                                       " order by z.sevktar) " + _
                    " from siparis a, sipmodel b "

            cSQL = cSQL + _
                    " where a.kullanicisipno = b.siparisno " + _
                    " and substring(a.kullanicisipno,1,2) <> 'KL' "

            Select Case nOpenClose
                Case 1
                    cSQL = cSQL + " and (a.dosyakapandi is null or a.dosyakapandi = 'H' or a.dosyakapandi = '') "
                Case 2, 3
                    cSQL = cSQL + " and a.dosyakapandi = 'E' "
            End Select

            cSQL = cSQL + _
                    " order by a.sonsevktarihi, a.kullanicisipno, b.modelno, b.renk "

            oSiparisReader = GetSQLReader(cSQL, ConnYage)

            Do While oSiparisReader.Read
                dGercekSevkTarihi = SQLReadDate(oSiparisReader, "gerceksevktarihi")
                If dGercekSevkTarihi = #1/1/1950# Then
                    dGercekSevkTarihi = SQLReadDate(oSiparisReader, "sonsevktarihi")
                End If

                lOK = False
                If nMonth = 0 Then
                    lOK = True
                Else
                    If (Month(dGercekSevkTarihi) = nMonth) And (Year(dGercekSevkTarihi) = Year(Now)) Then
                        lOK = True
                    End If
                End If

                Select Case nOpenClose
                    Case 2
                        ' kapalý sipariþlerde iptalleri ayýkla
                        If SQLReadDate(oSiparisReader, "gerceksevktarihi") = #1/1/1950# Then
                            lOK = False
                        End If
                    Case 3
                        If SQLReadDate(oSiparisReader, "gerceksevktarihi") = #1/1/1950# And (Year(SQLReadDate(oSiparisReader, "sonsevktarihi")) = Year(Now)) Then
                            lOK = True
                        Else
                            lOK = False
                        End If
                End Select

                If lOK Then
                    nCnt = nCnt + 1
                    ReDim Preserve aSiparis(nCnt)
                    aSiparis(nCnt).cSiparisNo = SQLReadString(oSiparisReader, "kullanicisipno")
                    aSiparis(nCnt).cModelNo = SQLReadString(oSiparisReader, "modelno")
                    aSiparis(nCnt).cRenk = SQLReadString(oSiparisReader, "renk")
                    aSiparis(nCnt).cMusteriNo = SQLReadString(oSiparisReader, "musterino")
                    aSiparis(nCnt).cSorumlu = SQLReadString(oSiparisReader, "sorumlu")
                    aSiparis(nCnt).dSiparisTarihi = SQLReadDate(oSiparisReader, "siparistarihi")
                    aSiparis(nCnt).dSonSevkTarihi = SQLReadDate(oSiparisReader, "sonsevktarihi")
                End If
            Loop
            oSiparisReader.Close()
            oSiparisReader = Nothing

            For nCnt = 0 To UBound(aSiparis)

                cSQL = "select SiparisTakipNotlari " + _
                        " from siparis " + _
                        " where kullanicisipno = '" + aSiparis(nCnt).cSiparisNo + "' "

                oSiparisReader = GetSQLReader(cSQL, ConnYage)

                If oSiparisReader.Read Then
                    aSiparis(nCnt).cSiparisTakipNotlari = SQLReadString(oSiparisReader, "SiparisTakipNotlari")
                End If
                oSiparisReader.Close()
                oSiparisReader = Nothing

                Siparis_No = aSiparis(nCnt).cSiparisNo
                Model_No = aSiparis(nCnt).cModelNo
                Renk = aSiparis(nCnt).cRenk
                Musteri = aSiparis(nCnt).cMusteriNo
                Musteri_Temsilcisi = aSiparis(nCnt).cSorumlu
                Siparis_Gelis_Tarihi = aSiparis(nCnt).dSiparisTarihi
                Son_Sevk_Tarihi = aSiparis(nCnt).dSonSevkTarihi
                SiparisTakipNotlari = StrStrip2(aSiparis(nCnt).cSiparisTakipNotlari)

                cSQL = "select max(bitistarihi) " + _
                        " from uretpllines " + _
                        " where ModelNo = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and departman = '" + cSevkDept + "' " + _
                        " and uretimtakipno in (select uretimtakipno " + _
                                                " from sipmodel " + _
                                                " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "' " + _
                                                " and modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                                                " and renk = '" + aSiparis(nCnt).cRenk + "') "

                dPlSevk = SQLGetDateConnected(cSQL, ConnYage) ' planlanan sevkiyat tarihi

                cSQL = "select sum(coalesce(adet,0)) " + _
                        " from sipmodel " + _
                        " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and renk = '" + aSiparis(nCnt).cRenk + "' "

                nSipAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = " select sum(coalesce(c.adet,0)) " + _
                        " from uretharfis a, uretharfislines b, uretharrba c " + _
                        " where a.uretfisno = b.uretfisno " + _
                        " and a.cikisdept like '%KESIM%' " + _
                        " and b.ulineno = c.ulineno " + _
                        " and c.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and c.renk = '" + aSiparis(nCnt).cRenk + "' " + _
                        " and c.uretimtakipno in (select uretimtakipno " + _
                                                " from sipmodel " + _
                                                " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "') "

                nKesim = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = " select sum(coalesce(c.adet,0)) " + _
                        " from uretharfis a, uretharfislines b, uretharrba c " + _
                        " where a.uretfisno = b.uretfisno " + _
                        " and a.girisdept like '%DIKIM%' " + _
                        " and b.ulineno = c.ulineno " + _
                        " and c.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and c.renk = '" + aSiparis(nCnt).cRenk + "' " + _
                        " and c.uretimtakipno in (select uretimtakipno " + _
                                                " from sipmodel " + _
                                                " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "') "

                nDikimGiren = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = " select sum(coalesce(c.adet,0)) " + _
                        " from uretharfis a, uretharfislines b, uretharrba c " + _
                        " where a.uretfisno = b.uretfisno " + _
                        " and a.cikisdept like '%DIKIM%' " + _
                        " and b.ulineno = c.ulineno " + _
                        " and b.harekettipi <> '01 IADE' " + _
                        " and c.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and c.renk = '" + aSiparis(nCnt).cRenk + "' " + _
                        " and c.uretimtakipno in (select uretimtakipno " + _
                                                " from sipmodel " + _
                                                " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "') "

                nDikim = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = " select sum(coalesce(c.adet,0)) " + _
                        " from uretharfis a, uretharfislines b, uretharrba c " + _
                        " where a.uretfisno = b.uretfisno " + _
                        " and a.cikisdept like '%DIKIM%' " + _
                        " and b.ulineno = c.ulineno " + _
                        " and b.harekettipi = '01 IADE' " + _
                        " and c.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and c.renk = '" + aSiparis(nCnt).cRenk + "' " + _
                        " and c.uretimtakipno in (select uretimtakipno " + _
                                                " from sipmodel " + _
                                                " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "') "

                nAtolyeIade = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = " select sum(coalesce(c.adet,0)) " + _
                        " from uretharfis a, uretharfislines b, uretharrba c " + _
                        " where a.uretfisno = b.uretfisno " + _
                        " and a.cikisdept = 'UTU&PAKET' " + _
                        " and a.girisdept = 'SEVKIYAT' " + _
                        " and b.ulineno = c.ulineno " + _
                        " and b.harekettipi = '11 IKINCI KALITE' " + _
                        " and c.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and c.renk = '" + aSiparis(nCnt).cRenk + "' " + _
                        " and c.uretimtakipno in (select uretimtakipno " + _
                                                " from sipmodel " + _
                                                " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "') "

                nIkinciKalite = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = "select sum((koliend - kolibeg + 1) * b.adet) " + _
                       " from sevkformlines a, sevkformlinesrba b, sevkform c " + _
                       " where a.sevkformno = b.sevkformno " + _
                       " and a.sevkformno = c.sevkformno " + _
                       " and a.ulineno = b.ulineno " + _
                       " and c.ok = 'E' " + _
                       " and a.siparisno = '" + aSiparis(nCnt).cSiparisNo + "' " + _
                       " and a.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                       " and b.renk = '" + aSiparis(nCnt).cRenk + "' "

                nSevkiyat = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = "select top 1 c.sevktar " + _
                       " from sevkformlines a, sevkformlinesrba b, sevkform c " + _
                       " where a.sevkformno = b.sevkformno " + _
                       " and a.sevkformno = c.sevkformno " + _
                       " and a.ulineno = b.ulineno " + _
                       " and c.ok = 'E' " + _
                       " and a.siparisno = '" + aSiparis(nCnt).cSiparisNo + "' " + _
                       " and a.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                       " and b.renk = '" + aSiparis(nCnt).cRenk + "' " + _
                       " order by c.sevktar "

                dSevkiyat = SQLGetDateConnected(cSQL, ConnYage) ' gerçekleþen sevkiyat tarihi

                cSQL = " select top 1 b.fistarihi " + _
                        " from uretharrba a, uretharfis b " + _
                        " where a.uretfisno = b.uretfisno " + _
                        " and a.modelno = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and a.renk = '" + aSiparis(nCnt).cRenk + "' " + _
                        " and b.cikisdept like '%KESIM%' " + _
                        " and a.uretimtakipno in (select uretimtakipno " + _
                                                " from sipmodel " + _
                                                " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "') " + _
                        " order by b.fistarihi "

                dKesim_Tarihi = SQLGetDateConnected(cSQL, ConnYage)

                cSQL = "select oktar2 " + _
                        " from sipok " + _
                        " where siparisno = '" + aSiparis(nCnt).cSiparisNo + "' " + _
                        " and modelkodu = '" + aSiparis(nCnt).cModelNo + "' " + _
                        " and (renk = '" + aSiparis(nCnt).cRenk + "' or renk = 'HEPSI') " + _
                        " and oktipi = 'ÖN IMALAT NUMUNESÝ' "

                dKesim_OK_Tarihi = SQLGetDateConnected(cSQL, ConnYage)

                nUretimKaybi = nDikimGiren - (nSevkiyat + nAtolyeIade + nIkinciKalite)

                cFirma = ""
                cKesim = ""
                dTermin = CDate("01.01.1950")
                nIhtiyac = 0
                nKarsilanan = 0
                dGiris = CDate("01.01.1950")

                GetSipAnaKumTed(aSiparis(nCnt).cSiparisNo, aSiparis(nCnt).cModelNo, aSiparis(nCnt).cRenk, cAnaKumFirma, dTermin, nIhtiyac, nKarsilanan, dGiris)
                cUretim = GetUretimDurumu(aSiparis(nCnt).cSiparisNo, aSiparis(nCnt).cModelNo, aSiparis(nCnt).cRenk)
                GetDeptFason(aSiparis(nCnt).cSiparisNo, cKesim, nKesimFiyat, "KESIM")
                GetDeptFason(aSiparis(nCnt).cSiparisNo, cFirma, nFiyat, "DIKIM")
                cAksesuar = GetSipAksDurum(aSiparis(nCnt).cSiparisNo, aSiparis(nCnt).cModelNo, aSiparis(nCnt).cRenk)

                Siparis_Planlanan_Svk_Tarihi = dPlSevk
                Siparis_Planlanan_Svk_Hft = DatePart("ww", dPlSevk)
                Siparis_Son_Svk_Hft = DatePart("ww", aSiparis(nCnt).dSonSevkTarihi)
                Sevkiyat_Siparis_Gun_Frk = DateDiff(Microsoft.VisualBasic.DateInterval.Day, aSiparis(nCnt).dSiparisTarihi, aSiparis(nCnt).dSonSevkTarihi)
                Siparis_Durumu = cUretim
                Kumas_Tedarikcisi = cAnaKumFirma
                If dTermin = CDate("01.01.1950") Then
                    Kumas_Termini = SqlDateTime.Null
                Else
                    Kumas_Termini = dTermin
                End If
                If dGiris = CDate("01.01.1950") Then
                    Kumas_Gelis_Tarihi = SqlDateTime.Null
                Else
                    Kumas_Gelis_Tarihi = dGiris
                End If
                If nIhtiyac = 0 Then
                    Kumas_Tamamlanma_Yuzdesi = 0
                Else
                    Kumas_Tamamlanma_Yuzdesi = CType(nKarsilanan / nIhtiyac * 100, SqlDecimal)
                End If
                Aksesuar_Durumu = cAksesuar
                If dKesim_OK_Tarihi = CDate("01.01.1950") Then
                    Kesim_OK_Tarihi = SqlDateTime.Null
                Else
                    Kesim_OK_Tarihi = dKesim_OK_Tarihi
                End If
                If dKesim_Tarihi = CDate("01.01.1950") Then
                    Kesim_Tarihi = SqlDateTime.Null
                Else
                    Kesim_Tarihi = dKesim_Tarihi
                End If
                Siparis_Adedi = CType(nSipAdet, SqlDecimal)
                Kesim_Adedi = CType(nKesim, SqlDecimal)
                Dikim_Giren = CType(nDikimGiren, SqlDecimal)
                Dikim_Adedi = CType(nDikim, SqlDecimal)
                Atolye_Iade = CType(nAtolyeIade, SqlDecimal)
                Ikinci_Kalite = CType(nIkinciKalite, SqlDecimal)
                Sevk_Adedi = CType(nSevkiyat, SqlDecimal)
                Uretim_Kaybi = CType(nUretimKaybi, SqlDecimal)
                Kesim_Sevkiyat_Farki = CType(nKesim - nSevkiyat, SqlDecimal)
                Dikim_Sevkiyat_Farki = CType(nDikim - nSevkiyat, SqlDecimal)
                Sevkiyat_Siparis_Farki = CType(nSevkiyat - nSipAdet, SqlDecimal)
                If nSipAdet = 0 Then
                    Kesim_Yuzdesi = 0
                    Dikim_Yuzdesi = 0
                    Sevkiyat_Yuzdesi = 0
                Else
                    Kesim_Yuzdesi = CType(nKesim / nSipAdet * 100, SqlDecimal)
                    Dikim_Yuzdesi = CType(nDikim / nSipAdet * 100, SqlDecimal)
                    Sevkiyat_Yuzdesi = CType(nSevkiyat / nSipAdet * 100, SqlDecimal)
                End If
                If nKesim = 0 Then
                    Sevkiyat_Yuzdesi2 = 0
                Else
                    Sevkiyat_Yuzdesi2 = CType(nSevkiyat / nKesim * 100, SqlDecimal)
                End If
                Kesim_Atolyesi = cKesim
                Dikim_Atolyesi = cFirma
                If dSevkiyat = CDate("01.01.1950") Then
                    Gerceklesen_Sevkiyat = SqlDateTime.Null
                Else
                    Gerceklesen_Sevkiyat = dSevkiyat
                End If

                resultCollection.Add(New SiparisDurumu1(Siparis_Planlanan_Svk_Hft, _
                                                        Siparis_Son_Svk_Hft, _
                                                        Siparis_Gelis_Tarihi, _
                                                        Siparis_Planlanan_Svk_Tarihi, _
                                                        Son_Sevk_Tarihi, _
                                                        Sevkiyat_Siparis_Gun_Frk, _
                                                        Musteri, _
                                                        Musteri_Temsilcisi, _
                                                        Siparis_No, _
                                                        Model_No, _
                                                        Renk, _
                                                        Siparis_Durumu, _
                                                        Kumas_Tedarikcisi, _
                                                        Kumas_Termini, _
                                                        Kumas_Gelis_Tarihi, _
                                                        Kumas_Tamamlanma_Yuzdesi, _
                                                        Aksesuar_Durumu, _
                                                        Kesim_OK_Tarihi, _
                                                        Kesim_Tarihi, _
                                                        Siparis_Adedi, _
                                                        Kesim_Adedi, _
                                                        Dikim_Giren, _
                                                        Dikim_Adedi, _
                                                        Sevk_Adedi, _
                                                        Kesim_Sevkiyat_Farki, _
                                                        Dikim_Sevkiyat_Farki, _
                                                        Sevkiyat_Siparis_Farki, _
                                                        Kesim_Yuzdesi, _
                                                        Dikim_Yuzdesi, _
                                                        Sevkiyat_Yuzdesi, _
                                                        Sevkiyat_Yuzdesi2, _
                                                        Kesim_Atolyesi, _
                                                        Dikim_Atolyesi, _
                                                        Gerceklesen_Sevkiyat, _
                                                        Ikinci_Kalite, _
                                                        Atolye_Iade, _
                                                        Uretim_Kaybi, _
                                                        SiparisTakipNotlari))
            Next

            CloseConn(ConnYage)

            Return resultCollection

        Catch ex As Exception
            ErrDisp("RaporSiparisDurumu1 : " + ex.Message.Trim + vbCrLf + cSQL)
            Return Nothing
        End Try
    End Function

    Public Shared Sub SiparisDurumu1_FillRow(ObjSiparisDurumu1 As Object, _
                                             <Out()> ByRef Siparis_Planlanan_Svk_Hft As SqlDecimal, _
                                             <Out()> ByRef Siparis_Son_Svk_Hft As SqlDecimal, _
                                             <Out()> ByRef Siparis_Gelis_Tarihi As SqlDateTime, _
                                             <Out()> ByRef Siparis_Planlanan_Svk_Tarihi As SqlDateTime, _
                                             <Out()> ByRef Son_Sevk_Tarihi As SqlDateTime, _
                                             <Out()> ByRef Sevkiyat_Siparis_Gun_Frk As SqlDecimal, _
                                             <Out()> ByRef Musteri As SqlString, _
                                             <Out()> ByRef Musteri_Temsilcisi As SqlString, _
                                             <Out()> ByRef Siparis_No As SqlString, _
                                             <Out()> ByRef Model_No As SqlString, _
                                             <Out()> ByRef Renk As SqlString, _
                                             <Out()> ByRef Siparis_Durumu As SqlString, _
                                             <Out()> ByRef Kumas_Tedarikcisi As SqlString, _
                                             <Out()> ByRef Kumas_Termini As SqlDateTime, _
                                             <Out()> ByRef Kumas_Gelis_Tarihi As SqlDateTime, _
                                             <Out()> ByRef Kumas_Tamamlanma_Yuzdesi As SqlDecimal, _
                                             <Out()> ByRef Aksesuar_Durumu As SqlString, _
                                             <Out()> ByRef Kesim_OK_Tarihi As SqlDateTime, _
                                             <Out()> ByRef Kesim_Tarihi As SqlDateTime, _
                                             <Out()> ByRef Siparis_Adedi As SqlDecimal, _
                                             <Out()> ByRef Kesim_Adedi As SqlDecimal, _
                                             <Out()> ByRef Dikim_Giren As SqlDecimal, _
                                             <Out()> ByRef Dikim_Adedi As SqlDecimal, _
                                             <Out()> ByRef Sevk_Adedi As SqlDecimal, _
                                             <Out()> ByRef Kesim_Sevkiyat_Farki As SqlDecimal, _
                                             <Out()> ByRef Dikim_Sevkiyat_Farki As SqlDecimal, _
                                             <Out()> ByRef Sevkiyat_Siparis_Farki As SqlDecimal, _
                                             <Out()> ByRef Kesim_Yuzdesi As SqlDecimal, _
                                             <Out()> ByRef Dikim_Yuzdesi As SqlDecimal, _
                                             <Out()> ByRef Sevkiyat_Yuzdesi As SqlDecimal, _
                                             <Out()> ByRef Sevkiyat_Yuzdesi2 As SqlDecimal, _
                                             <Out()> ByRef Kesim_Atolyesi As SqlString, _
                                             <Out()> ByRef Dikim_Atolyesi As SqlString, _
                                             <Out()> ByRef Gerceklesen_Sevkiyat As SqlDateTime, _
                                             <Out()> ByRef Ikinci_Kalite As SqlDecimal, _
                                             <Out()> ByRef Atolye_Iade As SqlDecimal, _
                                             <Out()> ByRef Uretim_Kaybi As SqlDecimal, _
                                             <Out()> ByRef SiparisTakipNotlari As SqlString)

        Dim siparisDurumu1 As SiparisDurumu1 = DirectCast(ObjSiparisDurumu1, SiparisDurumu1)

        Siparis_Planlanan_Svk_Hft = siparisDurumu1.Siparis_Planlanan_Svk_Hft
        Siparis_Son_Svk_Hft = siparisDurumu1.Siparis_Son_Svk_Hft
        Siparis_Gelis_Tarihi = siparisDurumu1.Siparis_Gelis_Tarihi
        Siparis_Planlanan_Svk_Tarihi = siparisDurumu1.Siparis_Planlanan_Svk_Tarihi
        Son_Sevk_Tarihi = siparisDurumu1.Son_Sevk_Tarihi
        Sevkiyat_Siparis_Gun_Frk = siparisDurumu1.Sevkiyat_Siparis_Gun_Frk
        Musteri = siparisDurumu1.Musteri
        Musteri_Temsilcisi = siparisDurumu1.Musteri_Temsilcisi
        Siparis_No = siparisDurumu1.Siparis_No
        Model_No = siparisDurumu1.Model_No
        Renk = siparisDurumu1.Renk
        Siparis_Durumu = siparisDurumu1.Siparis_Durumu
        Kumas_Tedarikcisi = siparisDurumu1.Kumas_Tedarikcisi
        Kumas_Termini = siparisDurumu1.Kumas_Termini
        Kumas_Gelis_Tarihi = siparisDurumu1.Kumas_Gelis_Tarihi
        Kumas_Tamamlanma_Yuzdesi = siparisDurumu1.Kumas_Tamamlanma_Yuzdesi
        Aksesuar_Durumu = siparisDurumu1.Aksesuar_Durumu
        Kesim_OK_Tarihi = siparisDurumu1.Kesim_OK_Tarihi
        Kesim_Tarihi = siparisDurumu1.Kesim_Tarihi
        Siparis_Adedi = siparisDurumu1.Siparis_Adedi
        Kesim_Adedi = siparisDurumu1.Kesim_Adedi
        Dikim_Giren = siparisDurumu1.Dikim_Giren
        Dikim_Adedi = siparisDurumu1.Dikim_Adedi
        Sevk_Adedi = siparisDurumu1.Sevk_Adedi
        Kesim_Sevkiyat_Farki = siparisDurumu1.Kesim_Sevkiyat_Farki
        Dikim_Sevkiyat_Farki = siparisDurumu1.Dikim_Sevkiyat_Farki
        Sevkiyat_Siparis_Farki = siparisDurumu1.Sevkiyat_Siparis_Farki
        Kesim_Yuzdesi = siparisDurumu1.Kesim_Yuzdesi
        Dikim_Yuzdesi = siparisDurumu1.Dikim_Yuzdesi
        Sevkiyat_Yuzdesi = siparisDurumu1.Sevkiyat_Yuzdesi
        Sevkiyat_Yuzdesi2 = siparisDurumu1.Sevkiyat_Yuzdesi2
        Kesim_Atolyesi = siparisDurumu1.Kesim_Atolyesi
        Dikim_Atolyesi = siparisDurumu1.Dikim_Atolyesi
        Gerceklesen_Sevkiyat = siparisDurumu1.Gerceklesen_Sevkiyat
        Ikinci_Kalite = siparisDurumu1.Ikinci_Kalite
        Atolye_Iade = siparisDurumu1.Atolye_Iade
        Uretim_Kaybi = siparisDurumu1.Uretim_Kaybi
        SiparisTakipNotlari = siparisDurumu1.SiparisTakipNotlari
    End Sub

    <SqlFunction()> _
    Public Shared Function ThumbFileName(cFileName As String) As SqlString

        Dim aFileName() As String
        Dim nMax As Integer = 0
        Dim cFileWithExtension As String = ""
        Dim cFile As String = ""
        Dim cExtension As String = ""
        Dim cPath As String = ""

        ThumbFileName = ""

        If cFileName = "" Then Exit Function

        Try
            cFile = ""
            If InStr(cFileName, "\") > 0 Then
                aFileName = Split(cFileName, "\")
                nMax = UBound(aFileName)
                cFile = LCase(Trim(aFileName(nMax)))
                cFileWithExtension = cFile

                If InStr(cFile, ".") > 0 Then
                    aFileName = Split(cFile, ".")
                    nMax = UBound(aFileName)
                    cFile = LCase(Trim(aFileName(nMax - 1)))
                End If
            End If

            cExtension = ""
            If InStr(cFileName, ".") > 0 Then
                aFileName = Split(cFileName, ".")
                nMax = UBound(aFileName)
                cExtension = LCase(Trim(aFileName(nMax)))
            End If

            cPath = Replace(cFileName, cFileWithExtension, "")

            ThumbFileName = cPath + "thumbs" + "\" + cFileWithExtension

        Catch ex As Exception
            ErrDisp("ThumbFileName : " + ex.Message.Trim + vbCrLf + cFileName)
            Return ""
        End Try
    End Function

    <SqlFunction()> _
    Public Shared Function gettasarimkarisim(cTasarimNo As String) As System.Data.SqlTypes.SqlString

        gettasarimkarisim = ""

        Try
            gettasarimkarisim = GetTasarimKarisim1(cTasarimNo)
        Catch ex As Exception
            ErrDisp("gettasarimkarisim : " + ex.Message.Trim)
        End Try
    End Function

End Class
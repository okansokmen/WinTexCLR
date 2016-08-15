Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module MTF
    Public Const G_isemriicinGelenGiris As String = " (stokhareketkodu = '04 Mlz Uretimden Giris' " +
       " or stokhareketkodu = '06 Tamirden Giris' " +
       " or stokhareketkodu = '02 Tedarikten Giris'  " +
       " or stokhareketkodu = '05 Diger Giris' ) "

    Public Const G_isemriicinGelenCikis As String = " (stokhareketkodu = '02 Tedarikten iade' " +
           " or stokhareketkodu = '06 Tamire Cikis' " +
           " or stokhareketkodu = '04 Mlz Uretime iade' " +
           " or stokhareketkodu = '05 Diger Cikis' ) "

    Public Const G_isemriHariciGelenGiris As String = " (stokhareketkodu = '05 Diger Giris' " +
           " or stokhareketkodu = '02 Tedarikten Giris' " +
           " or stokhareketkodu = '04 Mlz Uretimden Giris' " +
           " or stokhareketkodu = '06 Tamirden Giris' " +
           " or stokhareketkodu = '55 Kontrol Oncesi Giris' " +
           " or stokhareketkodu = '77 Top Bolme Giris' " +
           " or stokhareketkodu = '77 Aksesuar Bolme Giris' " +
           " or stokhareketkodu = '08 SAYIM GIRIS' " +
           " or stokhareketkodu = '90 Trans/Rezv Giris') "

    Public Const G_isemriHariciGelenCikis As String = " (stokhareketkodu = '05 Diger Cikis' " +
           " or stokhareketkodu = '02 Tedarikten iade' " +
           " or stokhareketkodu = '04 Mlz Uretime iade' " +
           " or stokhareketkodu = '06 Tamire Cikis' " +
           " or stokhareketkodu = '55 Kontrol Oncesi Cikis' " +
           " or stokhareketkodu = '77 Top Bolme Cikis' " +
           " or stokhareketkodu = '77 Aksesuar Bolme Cikis' " +
           " or stokhareketkodu = '08 SAYIM CIKIS' " +
           " or stokhareketkodu = '90 Trans/Rezv Cikis') "

    Public Const G_uretimicincikis As String = " stokhareketkodu = '01 Uretime Cikis' "

    Public Const G_uretimdeniade As String = " stokhareketkodu = '01 Uretimden iade' "

    Private Structure ModelHammadde
        Dim cReceteNo As String
        Dim cModelNo As String
        Dim cModelRenk As String
        Dim cModelBeden As String
        Dim cHammaddeKodu As String
        Dim cHammaddeRenk As String
        Dim cHammaddeBeden As String
        Dim cMalTakipEsasi As String
        Dim cTeminDept As String
        Dim cUretimDepartmani As String
        Dim nKullanimMiktari As Double
        Dim nFire As Double
        Dim cHesaplama As String
        Dim cMiktarBirimi As String
        Dim cMalzemeTakipNo As String
    End Structure

    Private Structure oMRBA
        Dim cUTF As String
        Dim cModelNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cReceteNo As String
        Dim nKesilen As Double
        Dim nAdet As Double
        Dim nAdet2 As Double
    End Structure

    Private Structure MRB
        Dim cModelNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cReceteNo As String
        Dim nAdet As Double
        Dim nKesimIsEmriAdet As Double
        Dim nKesimAdet As Double
        Dim nSiraNo As Double
    End Structure

    Private Structure MTF
        Dim cStokNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cUDept As String
        Dim cMDept As String
        Dim nMiktar As Double
        Dim nMiktarKesimIsemri As Double
        Dim nMiktarKesim As Double
        Dim cBirim As String
        Dim nFire As Double
        Dim cHesaplama As String
        Dim nSiraNo As Double
        Dim nKarsilanan As Double
        Dim nYuvarla As Double
        Dim cTamSayiYap As String
    End Structure

    Public Function MTFFastGenerateMulti(ByVal cFilter As String) As Integer

        Dim cSQL As String = ""
        Dim aMTF() As String = Nothing
        Dim nCnt As Integer = 0
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""

        MTFFastGenerateMulti = 0

        Try
            cFilter = Replace(cFilter, "||", "'").Trim

            JustForLog("MTFFastGenerateMulti Begin " + cFilter)

            lAltModelDetay = (GetSysPar("altmodeltakibi") = "1")

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            cSQL = "select distinct b.malzemetakipno " +
                    " from siparis a, " + cSipModelTableName + " b, ymodel c " +
                    " where a.kullanicisipno = b.siparisno " +
                    " and b.modelno = c.modelno " +
                    " and b.malzemetakipno is not null " +
                    " and b.malzemetakipno <> '' " +
                    cFilter +
                    " order by b.malzemetakipno "

            If CheckExists(cSQL) Then
                aMTF = SQLtoStringArray(cSQL)
                For nCnt = 0 To UBound(aMTF)
                    MTKFastGenerate(aMTF(nCnt))
                Next
            End If

            MTFFastGenerateMulti = 1

            JustForLog("MTFFastGenerateMulti END")

        Catch ex As Exception
            ErrDisp(ex.Message, "MTFFastGenerateMulti", cSQL)
        End Try
    End Function

    Public Sub MTFFastGenerateAll()

        Dim cSQL As String = ""
        Dim aMTF() As String = Nothing
        Dim nCnt As Integer = 0
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""

        Try

            lAltModelDetay = (GetSysPar("altmodeltakibi") = "1")

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            cSQL = "select distinct a.malzemetakipno " +
                    " from " + cSipModelTableName + " a, siparis b  " +
                    " where a.siparisno = b.kullanicisipno " +
                    " and a.malzemetakipno is not null " +
                    " and a.malzemetakipno <> '' " +
                    " and (b.dosyakapandi = 'H' or b.dosyakapandi = '' or b.dosyakapandi is null) " +
                    " order by a.malzemetakipno "

            If CheckExists(cSQL) Then
                aMTF = SQLtoStringArray(cSQL)
                For nCnt = 0 To UBound(aMTF)
                    MTKFastGenerate(aMTF(nCnt))
                Next
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "MTFFastGenerateAll", cSQL)
        End Try
    End Sub

    Public Function MTKFastGenerate(ByVal cMTF As String) As Integer

        Dim nCnt1 As Integer = 0
        Dim nCnt2 As Integer = 0
        Dim nCnt3 As Integer = 0
        Dim nCnt4 As Integer = 0
        Dim aMRBA() As MRB = Nothing
        Dim aMTF() As MTF = Nothing
        Dim aMH() As ModelHammadde = Nothing
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cSQL As String = ""
        Dim nCnt As Integer = 0
        Dim cMusteri As String = ""
        Dim nKesilenAdet As Double = 0
        Dim nSiparisAdet As Double = 0
        Dim nReceteAdet As Double = 0
        Dim nUretimToleransi As Double = 0
        Dim nFire As Double = 0
        Dim nFireKesim As Double = 0
        Dim nMalzemeFireFactor As Double = 1
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""
        Dim lFirst As Boolean = True
        Dim nFound As Integer = -1
        Dim cStokNo As String = ""
        Dim cRenk As String = ""
        Dim cBeden As String = ""
        Dim cMDept As String = ""
        Dim cUDept As String = ""
        Dim cBirim As String = ""
        Dim nMiktar As Double = 0
        Dim nMiktarKesimIsemri As Double
        Dim nMiktarKesim As Double
        Dim cParameters As String = ""
        Dim cModelRenk As String
        Dim cModelBeden As String = ""
        Dim lKesimIsEmrineGore As Boolean = False
        Dim lKesimIsEmriOK As Boolean = False
        Dim lKesileneGore As Boolean = False
        Dim lKesimOK As Boolean = False
        Dim nKesimIsEmriAdet As Double = 0

        MTKFastGenerate = 0

        Try
            If cMTF.Trim = "" Then Exit Function

            JustForLog("MTKFastGenerate Begin " + cMTF)

            ErrDispTable("Baslangic", cMTF.Trim)

            ConnYage = OpenConn()

            lKesileneGore = (GetSysParConnected("mtfkesilenegore", ConnYage) = "1")
            lKesimIsEmrineGore = (GetSysParConnected("mtfkesisemrinegore", ConnYage) = "1")
            lAltModelDetay = (GetSysParConnected("altmodeltakibi", ConnYage) = "1")

            If lKesileneGore Then
                cParameters = cParameters + " Kesilen Adet,"
            End If
            If lKesimIsEmrineGore Then
                cParameters = cParameters + " Kesim Emri Adet,"
            End If
            If lAltModelDetay Then
                cParameters = cParameters + " Alt Model,"
            End If
            If cParameters.Trim <> "" Then
                ErrDispConnected(ConnYage, "Hesaplama parametreleri : " + cParameters.Trim, cMTF.Trim)
            End If

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            cSQL = "select top 1 a.musterino " +
                    " from siparis a, sipmodel b " +
                    " where a.kullanicisipno = b.siparisno " +
                    " and b.malzemetakipno = '" + cMTF.Trim + "' " +
                    " and a.musterino is not null " +
                    " and a.musterino <> '' "

            cMusteri = SQLGetStringConnected(cSQL, ConnYage)

            cSQL = "select malzemetakipno " +
                    " from mtkfis " +
                     " where malzemetakipno = '" + cMTF.Trim + "' "

            If CheckExistsConnected(cSQL, ConnYage) Then
                If cMusteri <> "" Then

                    cSQL = "update mtkfis " +
                            " set musteri = '" + SQLWriteString(cMusteri, 30) + "' " +
                            " where malzemetakipno = '" + cMTF.Trim + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Else
                cSQL = "insert into mtkfis " +
                        " (malzemetakipno, dosyakapandi, musteri) " +
                        " values ('" + cMTF.Trim + "', " +
                        " 'H', " +
                        " '" + SQLWriteString(cMusteri, 30) + "') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            ' kilitlenmemiş satırların ihtiyaçlarını sıfırla
            cSQL = "update mtkfislines " +
                    " set ihtiyac = 0 " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (kilitle is null or kilitle = 'H' or kilitle = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set hesaplananihtiyac = 0 " +
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "select max(a.uretimtoleransi) " +
                    " from modeluretim a, sipmodel b " +
                    " where a.modelno = b.modelno " +
                    " and b.malzemetakipno = '" + cMTF.Trim + "' "

            nUretimToleransi = SQLGetDoubleConnected(cSQL, ConnYage)

            ErrDispConnected(ConnYage, "Uretim toleransi : " + nUretimToleransi.ToString, cMTF.Trim)

            ' model adetlerini REÇETE BAZINDA siparişten hesapla

            cSQL = "select modelno, renk, beden, receteno, " +
                    " adet = sum(coalesce(adet,0)) " +
                    " from " + cSipModelTableName +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and adet is not null " +
                    " and adet <> 0 " +
                    " group by modelno, renk, beden, receteno " +
                    " order by modelno, renk, beden, receteno "

            nCnt = -1

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ' Model Reçeteye Göre RBA
                nCnt = nCnt + 1
                ReDim Preserve aMRBA(nCnt)
                aMRBA(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aMRBA(nCnt).cRenk = SQLReadString(oReader, "renk")
                aMRBA(nCnt).cBeden = SQLReadString(oReader, "beden")
                aMRBA(nCnt).cReceteNo = SQLReadString(oReader, "receteno")
                aMRBA(nCnt).nAdet = SQLReadDouble(oReader, "adet")  ' recete siparis adedi
                aMRBA(nCnt).nKesimIsEmriAdet = 0                    ' recete kesim is emri adedi
                aMRBA(nCnt).nKesimAdet = 0                          ' recete kesim adedi
            Loop
            oReader.Close()
            oReader = Nothing

            If nCnt = -1 Then
                ' sipariş adetleri girilmemiş
                ErrDispConnected(ConnYage, "Bitis : Siparis adetleri bulunamadi", cMTF.Trim)
                ConnYage.Close()
                Exit Function
            End If

            ErrDispConnected(ConnYage, "Siparis adetleri okundu", cMTF.Trim)

            ' Kesim ve Kesim İşemri Adetleri
            cSQL = "select w.modelno, w.renk, w.beden, " +
                    " sipadet = sum(coalesce(w.sipadet,0)), " +
                    " kesisemriadet = sum(coalesce(w.kesisemriadet, 0)), " +
                    " kesadet = sum(coalesce(w.kesadet, 0)) "

            cSQL = cSQL +
                    " from (select uretimtakipno, modelno, renk, beden, " +
                            " sipadet = sum(coalesce(adet,0)), " +
                            " kesisemriadet = (select sum(coalesce(x.adet,0)) " +
                                        " from uretimisrba x, uretimisemri y " +
                                        " where x.isemrino = y.isemrino " +
                                        " and x.uretimtakipno = y.uretimtakipno " +
                                        " and x.uretimtakipno = sipmodel.uretimtakipno  " +
                                        " and y.departman like '%KES%'  " +
                                        " and x.modelno = sipmodel.modelno " +
                                        " and x.renk = sipmodel.renk  " +
                                        " and x.beden = sipmodel.beden ), " +
                            " kesadet = (select sum(coalesce(x.adet,0))  " +
                                        " from uretharrba x, uretharfis y  " +
                                        " where x.uretfisno = y.uretfisno " +
                                        " and x.uretimtakipno = sipmodel.uretimtakipno " +
                                        " and y.cikisdept like '%KES%' " +
                                        " and x.modelno = sipmodel.modelno " +
                                        " and x.renk = sipmodel.renk " +
                                        " and x.beden = sipmodel.beden ) " +
                            " from sipmodel " +
                            " where malzemetakipno = '" + cMTF.Trim + "' " +
                            " and adet is not null " +
                            " and adet <> 0  " +
                            " group by uretimtakipno, modelno, renk, beden) w  "
            cSQL = cSQL +
                    " group by w.modelno, w.renk, w.beden " +
                    " order by w.modelno, w.renk, w.beden "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                nKesimIsEmriAdet = SQLReadDouble(oReader, "kesisemriadet")  ' toplam kesim işemri adedi
                nKesilenAdet = SQLReadDouble(oReader, "kesadet")            ' toplam kesilen adet
                nSiparisAdet = SQLReadDouble(oReader, "sipadet")            ' toplam sipariş adedi

                For nCnt = 0 To UBound(aMRBA)
                    If aMRBA(nCnt).cModelNo = SQLReadString(oReader, "modelno") And
                        aMRBA(nCnt).cRenk = SQLReadString(oReader, "renk") And
                        aMRBA(nCnt).cBeden = SQLReadString(oReader, "beden") Then
                        ' reçeteye bağlı sipariş adedi
                        nReceteAdet = aMRBA(nCnt).nAdet
                        ' recete için verilen kesim işemri adedi
                        aMRBA(nCnt).nKesimIsEmriAdet = nKesimIsEmriAdet * nReceteAdet / nSiparisAdet
                        ' eksik malzeme hesaplanmasın diye en azından sipariş adedine göre hesaplama yapılır
                        If aMRBA(nCnt).nKesimIsEmriAdet = 0 Then
                            aMRBA(nCnt).nKesimIsEmriAdet = nSiparisAdet
                        End If
                        ' recete icin kesilen adet
                        aMRBA(nCnt).nKesimAdet = nKesilenAdet * nReceteAdet / nSiparisAdet
                        ' eksik malzeme hesaplanmasın diye en azından sipariş adedine göre hesaplama yapılır
                        If aMRBA(nCnt).nKesimAdet = 0 Then
                            aMRBA(nCnt).nKesimAdet = nSiparisAdet
                        End If
                    End If
                Next
            Loop
            oReader.Close()
            oReader = Nothing

            ' kesim işemrilerinin HEPSI onaylandıysa
            ' önce, en az 1 adet onaylı kesim işemri varsa

            cSQL = "select top 1 a.isemrino " +
                    " from uretimisemri a, sipmodel b " +
                    " where a.uretimtakipno = b.uretimtakipno " +
                    " and b.malzemetakipno = '" + cMTF.Trim + "' " +
                    " and a.departman like '%KES%' " +
                    " and a.onay = 'E' "

            If CheckExistsConnected(cSQL, ConnYage) Then
                ErrDispConnected(ConnYage, "Kesim emirleri onayli gibi", cMTF.Trim)
                ' sonra, onaysız bir kesim işemri yoksa - bütün kesim işemrileri onaylıysa
                cSQL = "select top 1 a.isemrino " +
                        " from uretimisemri a, sipmodel b " +
                        " where a.uretimtakipno = b.uretimtakipno " +
                        " and b.malzemetakipno = '" + cMTF.Trim + "' " +
                        " and a.departman like '%KES%' " +
                        " and (a.onay is null or a.onay = '' or a.onay = 'H') "

                If Not CheckExistsConnected(cSQL, ConnYage) Then
                    lKesimIsEmriOK = True
                    ErrDispConnected(ConnYage, "Kesim emirleri onayli", cMTF.Trim)
                End If
            End If

            ErrDispConnected(ConnYage, "Kesim emirlerine gore hesaplama, durum : " + lKesimIsEmriOK.ToString)

            ' kesim tamalandıysa kesilene göre adetleri al
            ' önce , en az 1 tane kapanmış kesim planlama satırı var mı

            cSQL = "select top 1 a.departman, a.okbilgisi " +
                    " from uretpllines a, sipmodel b " +
                    " where a.uretimtakipno = b.uretimtakipno " +
                    " and b.malzemetakipno = '" + cMTF.Trim + "' " +
                    " and a.departman like '%KES%' " +
                    " and a.okbilgisi = 'E' "

            If CheckExistsConnected(cSQL, ConnYage) Then
                ErrDispConnected(ConnYage, "Kesim tamamlanmis gibi", cMTF.Trim)
                ' sonra , kapanmamış kesim planlama satırı yoksa
                cSQL = "select top 1 a.departman, a.okbilgisi " +
                        " from uretpllines a, sipmodel b " +
                        " where a.uretimtakipno = b.uretimtakipno " +
                        " and b.malzemetakipno = '" + cMTF.Trim + "' " +
                        " and a.departman like '%KES%' " +
                        " and (a.okbilgisi is null or a.okbilgisi = '' or a.okbilgisi = 'H') "

                If Not CheckExistsConnected(cSQL, ConnYage) Then
                    ' bütün kesimler kapanmıştır
                    ' model, renk ve bedene göre kesilen adet tablosu
                    lKesimOK = True
                    ErrDispConnected(ConnYage, "Kesim tamamlanmis, hesaplaniyor", cMTF.Trim)
                End If
            End If

            ErrDispConnected(ConnYage, "Kesim adedine gore hesaplama, durum : " + lKesimOK.ToString)

            ' Buffer BOM
            cSQL = "select receteno = '', a.modelno, a.modelrenk, a.modelbeden, a.hammaddekodu, a.hammadderenk, a.hammaddebeden, b.maltakipesasi, a.temindept, " +
                    " a.uretimdepartmani, a.kullanimmiktari, a.fire, a.hesaplama, a.miktarbirimi, a.malzemetakipno " +
                    " from modelhammadde a, stok b " +
                    " where a.hammaddekodu = b.stokno " +
                    " and a.modelno in (select modelno from sipmodel where malzemetakipno = '" + cMTF.Trim + "') " +
                    " union all " +
                    " select a.receteno, a.modelno, a.modelrenk, a.modelbeden, a.hammaddekodu, a.hammadderenk, a.hammaddebeden, b.maltakipesasi, a.temindept, " +
                    " a.uretimdepartmani, a.kullanimmiktari, a.fire, a.hesaplama, a.miktarbirimi, a.malzemetakipno " +
                    " from modelhammadde2 a, stok b " +
                    " where a.hammaddekodu = b.stokno " +
                    " and a.modelno in (select modelno from sipmodel where malzemetakipno = '" + cMTF.Trim + "') "

            nCnt = -1

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aMH(nCnt)
                aMH(nCnt).cReceteNo = SQLReadString(oReader, "receteno")
                aMH(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aMH(nCnt).cModelRenk = SQLReadString(oReader, "modelrenk")
                aMH(nCnt).cModelBeden = SQLReadString(oReader, "modelbeden")
                aMH(nCnt).cHammaddeKodu = SQLReadString(oReader, "hammaddekodu")
                aMH(nCnt).cHammaddeRenk = SQLReadString(oReader, "hammadderenk")
                aMH(nCnt).cHammaddeBeden = SQLReadString(oReader, "hammaddebeden")
                aMH(nCnt).cMalTakipEsasi = SQLReadString(oReader, "maltakipesasi")
                aMH(nCnt).cTeminDept = SQLReadString(oReader, "temindept")
                aMH(nCnt).cUretimDepartmani = SQLReadString(oReader, "uretimdepartmani")
                aMH(nCnt).nKullanimMiktari = SQLReadDouble(oReader, "kullanimmiktari")
                aMH(nCnt).nFire = SQLReadDouble(oReader, "fire")
                aMH(nCnt).cHesaplama = SQLReadString(oReader, "hesaplama")
                aMH(nCnt).cMiktarBirimi = SQLReadString(oReader, "miktarbirimi")
                aMH(nCnt).cMalzemeTakipNo = SQLReadString(oReader, "malzemetakipno")
            Loop
            oReader.Close()
            oReader = Nothing

            If nCnt = -1 Then
                ' reçeteler girilmemiş
                ErrDispConnected(ConnYage, "Bitis : Receteler Bulunamadi", cMTF.Trim)
                CloseConn(ConnYage)
                Exit Function
            End If

            ErrDispConnected(ConnYage, "BOM okundu", cMTF.Trim)

            ' RBA yı BOM ile çarp
            lFirst = True
            For nCnt = 0 To UBound(aMRBA)
                If aMRBA(nCnt).nAdet > 0 Then
                    For nCnt3 = 0 To UBound(aMH)
                        cModelRenk = IIf(aMH(nCnt3).cModelRenk = "HEPSI", aMRBA(nCnt).cRenk, aMH(nCnt3).cModelRenk).ToString
                        cModelBeden = IIf(aMH(nCnt3).cModelBeden = "HEPSI", aMRBA(nCnt).cBeden, aMH(nCnt3).cModelBeden).ToString

                        If aMH(nCnt3).cReceteNo = aMRBA(nCnt).cReceteNo And
                        aMH(nCnt3).cModelNo = aMRBA(nCnt).cModelNo And
                        cModelRenk = aMRBA(nCnt).cRenk And
                        cModelBeden = aMRBA(nCnt).cBeden Then

                            cStokNo = aMH(nCnt3).cHammaddeKodu
                            cMDept = aMH(nCnt3).cTeminDept
                            cUDept = aMH(nCnt3).cUretimDepartmani
                            cBirim = aMH(nCnt3).cMiktarBirimi
                            cRenk = IIf(aMH(nCnt3).cHammaddeRenk = "HEPSI", cModelRenk, aMH(nCnt3).cHammaddeRenk).ToString
                            cBeden = IIf(aMH(nCnt3).cHammaddeBeden = "HEPSI", cModelBeden, aMH(nCnt3).cHammaddeBeden).ToString

                            Select Case aMH(nCnt3).cMalTakipEsasi
                                Case "1"
                                    cRenk = "HEPSI"
                                    cBeden = "HEPSI"
                                Case "2"
                                    cBeden = "HEPSI"
                                Case "3"
                                    cRenk = "HEPSI"
                            End Select

                            If cRenk.Trim = "" Then
                                cRenk = "HEPSI"
                            End If

                            If cBeden.Trim = "" Then
                                cBeden = "HEPSI"
                            End If

                            nFire = aMH(nCnt3).nFire
                            nFireKesim = 0

                            If aMH(nCnt3).nFire >= nUretimToleransi Then
                                nFireKesim = aMH(nCnt3).nFire - nUretimToleransi
                            End If

                            ' fire carpana cevriliyor
                            Select Case aMH(nCnt3).cHesaplama
                                Case "1"
                                    ' Yukardan asagıya
                                    nFire = (1.0# + (nFire / 100.0#))
                                    nFireKesim = (1.0# + (nFireKesim / 100.0#))
                                Case "2"
                                    ' asagidan yukari hesaplansin
                                    If nFire <> 100 Then
                                        nFire = 1 / (1.0# - (nFire / 100.0#))
                                    End If
                                    If nFireKesim <> 100 Then
                                        nFireKesim = 1 / (1.0# - (nFireKesim / 100.0#))
                                    End If
                                Case Else
                                    ' Yukardan asagıya
                                    nFire = (1.0# + (nFire / 100.0#))
                                    nFireKesim = (1.0# + (nFireKesim / 100.0#))
                            End Select

                            ' fire, fire çarpanına dönüştüğünde 0 olamaz, en az 1 olabilir
                            If nFire = 0 Then
                                nFire = 1
                            End If

                            If nFireKesim = 0 Then
                                nFireKesim = 1
                            End If

                            If aMH(nCnt3).cMalzemeTakipNo = "" Then
                                If lKesimIsEmrineGore And lKesimIsEmriOK Then
                                    ' kesim işemrine göre
                                    nMiktar = aMH(nCnt3).nKullanimMiktari * nFireKesim * aMRBA(nCnt).nKesimIsEmriAdet
                                ElseIf lKesileneGore And lKesimOK Then
                                    ' kesilen adede göre
                                    nMiktar = aMH(nCnt3).nKullanimMiktari * nFireKesim * aMRBA(nCnt).nKesimAdet
                                Else
                                    ' sipariş adedine göre
                                    nMiktar = aMH(nCnt3).nKullanimMiktari * nFire * aMRBA(nCnt).nAdet
                                End If

                                nMiktarKesimIsemri = aMH(nCnt3).nKullanimMiktari * nFireKesim * aMRBA(nCnt).nKesimIsEmriAdet
                                nMiktarKesim = aMH(nCnt3).nKullanimMiktari * nFireKesim * aMRBA(nCnt).nKesimAdet
                            Else
                                ' eğer satıra MTF yazılmışsa toplam miktardır adet ile çarpılmaz
                                nMiktar = aMH(nCnt3).nKullanimMiktari * nFire
                                nMiktarKesimIsemri = aMH(nCnt3).nKullanimMiktari * nFireKesim
                                nMiktarKesim = aMH(nCnt3).nKullanimMiktari * nFireKesim
                            End If

                            If lFirst Then
                                ReDim aMTF(0)
                                aMTF(0).cStokNo = cStokNo
                                aMTF(0).cRenk = cRenk
                                aMTF(0).cBeden = cBeden
                                aMTF(0).cMDept = cMDept
                                aMTF(0).cUDept = cUDept
                                aMTF(0).cBirim = cBirim
                                aMTF(0).nMiktar = nMiktar
                                aMTF(0).nMiktarKesim = nMiktarKesim
                                aMTF(0).nMiktarKesimIsemri = nMiktarKesimIsemri
                                lFirst = False
                            Else
                                nFound = -1
                                For nCnt4 = 0 To UBound(aMTF)
                                    If aMTF(nCnt4).cStokNo = cStokNo And
                                    aMTF(nCnt4).cRenk = cRenk And
                                    aMTF(nCnt4).cBeden = cBeden And
                                    aMTF(nCnt4).cMDept = cMDept And
                                    aMTF(nCnt4).cUDept = cUDept And
                                    aMTF(nCnt4).cBirim = cBirim Then
                                        nFound = nCnt4
                                        Exit For
                                    End If
                                Next
                                If nFound = -1 Then
                                    nFound = UBound(aMTF) + 1
                                    ReDim Preserve aMTF(nFound)
                                    aMTF(nFound).cStokNo = cStokNo
                                    aMTF(nFound).cRenk = cRenk
                                    aMTF(nFound).cBeden = cBeden
                                    aMTF(nFound).cMDept = cMDept
                                    aMTF(nFound).cUDept = cUDept
                                    aMTF(nFound).cBirim = cBirim
                                    aMTF(nFound).nMiktar = nMiktar
                                    aMTF(nFound).nMiktarKesim = nMiktarKesim
                                    aMTF(nFound).nMiktarKesimIsemri = nMiktarKesimIsemri
                                Else
                                    aMTF(nFound).nMiktar = aMTF(nFound).nMiktar + nMiktar
                                    aMTF(nFound).nMiktarKesim = aMTF(nFound).nMiktarKesim + nMiktarKesim
                                    aMTF(nFound).nMiktarKesimIsemri = aMTF(nFound).nMiktarKesimIsemri + nMiktarKesimIsemri
                                End If
                            End If
                        End If
                    Next
                End If
            Next

            If lFirst Then
                ' RBA x BOM yapılamıyor
                ErrDispConnected(ConnYage, "Bitis : RBA x BOM yapilamiyor", cMTF.Trim)
                CloseConn(ConnYage)
                Exit Function
            End If

            ErrDispConnected(ConnYage, "RBA x BOM hesaplandi", cMTF.Trim)

            For nCnt1 = 0 To UBound(aMTF)
                UpdateMTKFisLines(ConnYage, cMTF, aMTF(nCnt1).cStokNo, aMTF(nCnt1).cRenk, aMTF(nCnt1).cBeden, aMTF(nCnt1).nMiktar, aMTF(nCnt1).cBirim, aMTF(nCnt1).cUDept, aMTF(nCnt1).cMDept,,, aMTF(nCnt1).nMiktarKesim, aMTF(nCnt1).nMiktarKesimIsemri)
            Next

            ErrDispConnected(ConnYage, "Ihtiyaclar yazildi", cMTF.Trim)

            CloseConn(ConnYage)
            ErrDispTable("MTF ihtiyac hesaplama bitti", cMTF.Trim)

            ' Kayip malzemeleri bul
            MTKFindLost(cMTF.Trim)
            ErrDispTable("MTF kayip malzemeler bitti", cMTF.Trim)

            ' Post Process
            MTKPostProcess(cMTF.Trim, cSipModelTableName)
            ErrDispTable("MTF post process bitti", cMTF.Trim)

            ' işemri kontrol
            G_IsemriDeptKontrol(cMTF.Trim)
            ErrDispTable("MTF isemri kontrol bitti", cMTF.Trim)

            ' Calc
            MTKLinesTopl(cMTF.Trim)
            ErrDispTable("MTF isemri verilen, gelen, vs hesaplamalari bitti", cMTF.Trim)

            ' dokuma ön maliyet çalışması fiyatları planlama fiyatları olarak al
            MTKOnMaliyet(cMTF.Trim)
            ErrDispTable("MTF dokuma on maliyet fiyatlari bitti", cMTF.Trim)

            ' temizlik
            MTKCleanUP(cMTF.Trim)
            ErrDispTable("Bitis", cMTF.Trim)

            ' The END
            MTKFastGenerate = 1

            JustForLog("MTKFastGenerate END " + cMTF)

        Catch Err As Exception
            ErrDisp(Err.Message, "MTKFastGenerate", cSQL)
        End Try
    End Function

    Private Sub MTKPostProcess(cMTF As String, cSipModelTableName As String)

        Dim nYuvarla As Double = 0
        Dim cSQL As String = ""
        Dim cBirim As String = ""
        Dim cTamSayiYap As String = ""
        Dim cSipList As String = ""
        Dim nCnt As Integer = 0
        Dim oReader As SqlDataReader
        Dim aMTF() As MTF = Nothing
        Dim ConnYage As SqlConnection

        Try
            If cMTF.Trim = "" Then Exit Sub

            ConnYage = OpenConn()

            ' hesaplanan temel ihtiyaç hesaplananihtiyac hanesinde saklanır
            cSQL = "update mtkfislines " +
                    " set hesaplananihtiyac = coalesce(ihtiyac,0), " +
                    " musteriihtiyac = 0, " +
                    " ihtiyatiihtiyac = 0 " +
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set musteriihtiyac = (select sum(coalesce(b.miktar,0)) " +
                                        " from mtkeklefis a, mtkeklefislines b " +
                                        " where a.mtkeklefisno = b.mtkeklefisno " +
                                        " and a.malzemetakipno = mtkfislines.malzemetakipno " +
                                        " and b.stokno = mtkfislines.stokno " +
                                        " and b.renk = mtkfislines.renk " +
                                        " and b.beden = mtkfislines.beden " +
                                        " and b.departman = mtkfislines.departman) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set ihtiyac = coalesce(musteriihtiyac,0) + coalesce(ihtiyatiihtiyac,0) + coalesce(hesaplananihtiyac,0) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = 0

            cSQL = "select distinct a.stokno, b.yuvarla, b.birim1, y2 = c.yuvarla  " +
                    " from mtkfislines a, stok b, birim c " +
                    " where a.stokno = b.stokno " +
                    " and b.birim1 = c.birim " +
                    " and a.malzemetakipno = '" + cMTF.Trim + "' " +
                    " order by a.stokno "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aMTF(nCnt)
                aMTF(nCnt).cStokNo = SQLReadString(oReader, "stokno")
                aMTF(nCnt).nYuvarla = SQLReadDouble(oReader, "yuvarla")
                aMTF(nCnt).cBirim = SQLReadString(oReader, "birim1")
                aMTF(nCnt).cTamSayiYap = SQLReadString(oReader, "y2")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aMTF)
                If aMTF(nCnt).cTamSayiYap = "E" Then
                    cSQL = "update mtkfislines " +
                            " set ihtiyac = ceiling(ihtiyac), " +
                            " kesilenihtiyac = ceiling(kesilenihtiyac), " +
                            " kesimisemriihtiyac = ceiling(kesimisemriihtiyac) " +
                            " where malzemetakipno = '" + cMTF.Trim + "' " +
                            " and stokno = '" + aMTF(nCnt).cStokNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    cSQL = "update mtkfislines " +
                            " set ihtiyac = round(ihtiyac," + SQLWriteDecimal(aMTF(nCnt).nYuvarla) + "), " +
                            " kesilenihtiyac = round(kesilenihtiyac," + SQLWriteDecimal(aMTF(nCnt).nYuvarla) + "), " +
                            " kesimisemriihtiyac = round(kesimisemriihtiyac," + SQLWriteDecimal(aMTF(nCnt).nYuvarla) + ") " +
                            " where malzemetakipno = '" + cMTF.Trim + "' " +
                            " and stokno = '" + aMTF(nCnt).cStokNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            cSQL = "select notlar " +
                    " from mtkfis " +
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            cSipList = SQLGetStringConnected(cSQL, ConnYage)

            cSQL = "select distinct siparisno " +
                    " from " + cSipModelTableName +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and adet is not null " +
                    " and adet <> 0 "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                If cSipList = "" Then
                    cSipList = SQLReadString(oReader, "siparisno")
                Else
                    If InStr(cSipList, SQLReadString(oReader, "siparisno")) = 0 Then
                        cSipList = cSipList + "," + SQLReadString(oReader, "siparisno")
                    End If
                End If
            Loop
            oReader.Close()
            oReader = Nothing

            cSQL = "update mtkfis " +
                    " set notlar = '" + cSipList.Trim + "' " +
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)
            ' Malzeme zaman ve bütçe ön planlaması için
            ' Stok kartlarından temin süresi öndeğerleri alınır

            cSQL = "update mtkfislines " +
                    " set teminsuresi = (select top 1 gelisgun " +
                                        " from stok " +
                                        " where stokno = mtkfislines.stokno) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (teminsuresi is null or teminsuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set teminsuresi = (select top 1 a.gelisgun " +
                                        " from stoktipi a, stok b " +
                                        " where a.kod = b.stoktipi " +
                                        " and b.stokno = mtkfislines.stokno) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (teminsuresi is null or teminsuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set avanssuresi = (select top 1 avanssuresi " +
                                        " from stok " +
                                        " where stokno = mtkfislines.stokno) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (avanssuresi is null or avanssuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set avanssuresi = (select top 1 a.avanssuresi " +
                                        " from stoktipi a, stok b " +
                                        " where a.kod = b.stoktipi " +
                                        " and b.stokno = mtkfislines.stokno) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (avanssuresi is null or avanssuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set oktipi = (select top 1 a.oktipi " +
                                    " from stoktipi a, stok b " +
                                    " where a.kod = b.stoktipi " +
                                    " and b.stokno = mtkfislines.stokno) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (oktipi is null or oktipi = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            'Stokno, renk, beden aynı çıkış üretim departmanı farklıysa
            'gelen malzemeyi bölüştürmek için
            'ihtiyaç miktarıyla doğru orantılı bir katsayı çarpanı kullanılır
            cSQL = "update MtkFisLines " +
                     " Set KatSayi = 1 " +
                     " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update MtkFisLines " +
                    " set katsayi = coalesce(ihtiyac,0) / (select sum(coalesce(a.ihtiyac,0)) " +
                                                            " from mtkfislines a " +
                                                            " Where a.malzemetakipno = mtkfislines.malzemetakipno " +
                                                            " and a.stokno = mtkfislines.stokno " +
                                                            " and a.renk = mtkfislines.renk " +
                                                            " and a.beden = mtkfislines.beden " +
                                                            " and a.temindept = mtkfislines.temindept) " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and ihtiyac is not null " +
                    " and ihtiyac <> 0 "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update MtkFisLines " +
                    " Set katsayi = 1 " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (katsayi = 0 or katsayi is null) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set hedefmlzbirimfiyati = (select top 1 fiyat " +
                                    " From maliyetkumas " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)), "
            cSQL = cSQL +
                    " hedefmlzdovizi = (select top 1 doviz " +
                                    " From maliyetkumas " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)) "
            cSQL = cSQL +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (hedefmlzbirimfiyati is null or hedefmlzbirimfiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set plfirma = (select top 1 firma " +
                                    " From maliyetkumas " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)) "
            cSQL = cSQL +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (plfirma is null or plfirma = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' Dikim Malzemeleri

            cSQL = "update mtkfislines " +
                    " set hedefmlzbirimfiyati = (select top 1 fiyat " +
                                    " From maliyetdikim " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and beden = mtkfislines.beden " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)), "
            cSQL = cSQL +
                    " hedefmlzdovizi = (select top 1 doviz " +
                                    " From maliyetdikim " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and beden = mtkfislines.beden " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)) "
            cSQL = cSQL +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (hedefmlzbirimfiyati is null or hedefmlzbirimfiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set plfirma = (select top 1 firma " +
                                    " From maliyetdikim " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and beden = mtkfislines.beden " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)) "
            cSQL = cSQL +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (plfirma is null or plfirma = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' Paketleme Malzemeleri

            cSQL = "update mtkfislines " +
                    " set hedefmlzbirimfiyati = (select top 1 fiyat " +
                                    " From maliyetpaketleme " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and beden = mtkfislines.beden " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)), "
            cSQL = cSQL +
                    " hedefmlzdovizi = (select top 1 doviz " +
                                    " From maliyetpaketleme " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and beden = mtkfislines.beden " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)) "
            cSQL = cSQL +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (hedefmlzbirimfiyati is null or hedefmlzbirimfiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set plfirma = (select top 1 firma " +
                                    " From maliyetpaketleme " +
                                    " Where stokno = mtkfislines.stokno " +
                                    " and renk = mtkfislines.renk " +
                                    " and beden = mtkfislines.beden " +
                                    " and fiyat is not null " +
                                    " and fiyat <> 0 " +
                                    " and calismano in (select b.maliyetcalismano " +
                                                        " from sipmodel a, ymodel b " +
                                                        " Where a.ModelNo = b.ModelNo " +
                                                        " and a.malzemetakipno = mtkfislines.malzemetakipno)) "
            cSQL = cSQL +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (plfirma is null or plfirma = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "MTKPostProcess", cSQL)
        End Try
    End Sub

    Private Sub UpdateMTKFisLines(ConnYage As SqlConnection, cMTF As String, cStokno As String, cHRenk As String, cHBeden As String, nMiktar As Double,
                                  Optional cBirim As String = "", Optional cUDept As String = "", Optional cMDept As String = "",
                                  Optional cTable As String = "", Optional lMalTakipEsasi As Boolean = True, Optional nMiktarKesim As Double = 0, Optional nMiktarKesimIsemri As Double = 0)
        Dim cSQL As String = ""
        Dim oReader As SqlDataReader
        Dim nMiktar2 As Double = 0
        Dim cTakipEsasi As String = ""
        Dim aMTF() As MTF = Nothing
        Dim nCnt1 As Integer = 0
        Dim nFire As Double = 0
        Dim nMalzemeFireFactor As Double = 1

        Try
            If cStokno.Trim = "" Then Exit Sub

            If cTable = "" Then
                cTable = "mtkfislines"
            End If

            cSQL = "select top 1 maltakipesasi, paratakipesasi, temindepartmani, uretimdepartmani, birim1 " +
                    " from stok " +
                    " where stokno = '" + cStokno.Trim + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                If lMalTakipEsasi Then
                    cTakipEsasi = SQLReadString(oReader, "maltakipesasi")
                Else
                    cTakipEsasi = SQLReadString(oReader, "paratakipesasi")
                End If

                Select Case cTakipEsasi
                    Case "1"
                        cHRenk = "HEPSI"
                        cHBeden = "HEPSI"
                    Case "2"
                        cHBeden = "HEPSI"
                    Case "3"
                        cHRenk = "HEPSI"
                End Select
                If cMDept.Trim = "" Then
                    cMDept = SQLReadString(oReader, "temindepartmani")
                End If
                If cUDept.Trim = "" Then
                    cUDept = SQLReadString(oReader, "uretimdepartmani")
                End If
                If cBirim.Trim = "" Then
                    cBirim = SQLReadString(oReader, "birim1")
                End If
            End If
            oReader.Close()
            oReader = Nothing

            cSQL = "select top 1 malzemetakipno " +
                    " from " + cTable.Trim +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and stokno = '" + cStokno.Trim + "' " +
                    " and renk = '" + cHRenk.Trim + "' " +
                    " and beden = '" + cHBeden.Trim + "' " +
                    " and temindept = '" + cMDept.Trim + "' " +
                    " and departman = '" + cUDept.Trim + "' "

            If CheckExistsConnected(cSQL, ConnYage) Then
                cSQL = "update " + cTable.Trim +
                        " set ihtiyac = coalesce(ihtiyac,0) + " + SQLWriteDecimal(nMiktar) + ", " +
                        " kesilenihtiyac = coalesce(kesilenihtiyac,0) + " + SQLWriteDecimal(nMiktarKesim) + ", " +
                        " kesimisemriihtiyac = coalesce(kesimisemriihtiyac,0) + " + SQLWriteDecimal(nMiktarKesimIsemri) +
                        " where malzemetakipno = '" + cMTF.Trim + "' " +
                        " and stokno = '" + cStokno.Trim + "' " +
                        " and renk = '" + cHRenk.Trim + "' " +
                        " and beden = '" + cHBeden.Trim + "' " +
                        " and temindept = '" + cMDept.Trim + "' " +
                        " and departman = '" + cUDept.Trim + "' " +
                        " and (kilitle is null or kilitle = 'H' or kilitle = '') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "insert into " + cTable.Trim + " (malzemetakipno, stokno, renk, beden, ihtiyac, birim, departman, temindept, " +
                        " isemriverilen, isemriicingelen, isemriharicigelen, isemriicingiden, isemriharicigiden, uretimicincikis, uretimdeniade, " +
                        " hedefmlzbirimfiyati, hedefiscilikbirimfiyati, uretimecikisfiyati, musteriihtiyac, ihtiyatiihtiyac, hesaplananihtiyac, kesilenihtiyac, kesimisemriihtiyac) " +
                        " values ('" + cMTF.Trim + "', " +
                        " '" + cStokno.Trim + "', " +
                        " '" + cHRenk.Trim + "', " +
                        " '" + cHBeden.Trim + "', " +
                        SQLWriteDecimal(nMiktar) + ", " +
                        " '" + cBirim.Trim + "', " +
                        " '" + cUDept.Trim + "', " +
                        " '" + cMDept.Trim + "', " +
                        " 0,0,0,0,0,0,0,0,0,0,0,0,0, " +
                        SQLWriteDecimal(nMiktarKesim) + ", " +
                        SQLWriteDecimal(nMiktarKesimIsemri) + " ) "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            ' Recursive olarak hammadde ağacını hesapla
            cSQL = "select a.hammaddekodu, a.hamrenk, a.hambeden, a.fire, a.miktar, a.hesaplama, b.maltakipesasi, b.birim1 " +
                    " from strecete a, stok b " +
                    " where a.hammaddekodu = b.stokno " +
                    " And a.anahammadde = '" + cStokno.Trim + "' " +
                    " and a.hammaddekodu <> '" + cStokno.Trim + "' " +
                    " and (a.anarenk = '" + cHRenk.Trim + "' or a.anarenk = 'HEPSI') " +
                    " and (a.anabeden = '" + cHBeden.Trim + "' or a.anabeden = 'HEPSI') "

            If CheckExistsConnected(cSQL, ConnYage) Then

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read

                    ReDim Preserve aMTF(nCnt1)

                    aMTF(nCnt1).cStokNo = SQLReadString(oReader, "hammaddekodu")
                    aMTF(nCnt1).cRenk = IIf(SQLReadString(oReader, "hammadderenk") = "HEPSI", cHRenk.Trim, SQLReadString(oReader, "hammadderenk")).ToString
                    aMTF(nCnt1).cBeden = IIf(SQLReadString(oReader, "hammaddebeden") = "HEPSI", cHBeden.Trim, SQLReadString(oReader, "hammaddebeden")).ToString

                    Select Case SQLReadString(oReader, "maltakipesasi")
                        Case "1"
                            aMTF(nCnt1).cRenk = "HEPSI"
                            aMTF(nCnt1).cBeden = "HEPSI"
                        Case "2"
                            aMTF(nCnt1).cBeden = "HEPSI"
                        Case "3"
                            aMTF(nCnt1).cRenk = "HEPSI"
                    End Select

                    aMTF(nCnt1).nMiktar = SQLReadDouble(oReader, "miktar")
                    aMTF(nCnt1).nMiktarKesim = SQLReadDouble(oReader, "miktar")
                    aMTF(nCnt1).nMiktarKesimIsemri = SQLReadDouble(oReader, "miktar")
                    aMTF(nCnt1).nFire = SQLReadDouble(oReader, "fire")
                    aMTF(nCnt1).cHesaplama = SQLReadString(oReader, "hesaplama")
                    aMTF(nCnt1).cBirim = SQLReadString(oReader, "birim1")

                    nFire = aMTF(nCnt1).nFire

                    nMalzemeFireFactor = 1
                    Select Case SQLReadString(oReader, "hesaplama")
                        Case "1"
                            nMalzemeFireFactor = (1.0# + (nFire / 100.0#))       ' Yukardan asagıya
                        Case "2"
                            If nFire <> 100 Then
                                nMalzemeFireFactor = 1 / (1.0# - (nFire / 100.0#)) ' ' asagidan yukari hesaplansin
                            End If
                        Case Else
                            nMalzemeFireFactor = (1.0# + (nFire / 100.0#))    ' Yukardan asagıya
                    End Select

                    nFire = nMalzemeFireFactor
                    ' fire, fire çarpanına dönüştüğünde 0 olamaz, en az 1 olabilir
                    If nFire = 0 Then
                        nFire = 1
                    End If

                    aMTF(nCnt1).nMiktar = aMTF(nCnt1).nMiktar * nFire * nMiktar
                    aMTF(nCnt1).nMiktarKesim = aMTF(nCnt1).nMiktarKesim * nFire * nMiktar
                    aMTF(nCnt1).nMiktarKesimIsemri = aMTF(nCnt1).nMiktarKesimIsemri * nFire * nMiktar

                    nCnt1 = nCnt1 + 1
                Loop
                oReader.Close()
                oReader = Nothing

                For nCnt1 = 0 To UBound(aMTF)
                    ' recurse and recalc
                    UpdateMTKFisLines(ConnYage, cMTF,
                                     aMTF(nCnt1).cStokNo,
                                     aMTF(nCnt1).cRenk,
                                     aMTF(nCnt1).cBeden,
                                     aMTF(nCnt1).nMiktar, , , , cTable, lMalTakipEsasi, aMTF(nCnt1).nMiktarKesim, aMTF(nCnt1).nMiktarKesimIsemri)
                Next
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "UpdateMTKFisLines", cSQL)
        End Try
    End Sub

    Private Sub MTKLinesTopl(Optional cMTFNo As String = "")
        ' cMTFNo boş ise bütün MTF lerin durumu yeniden hesaplanır
        Dim cSQL As String = ""
        Dim cSql1 As String = ""
        Dim cSql2 As String = ""
        Dim cSql3 As String = ""
        Dim cSql4 As String = ""
        Dim cTempView As String = ""
        Dim G_MTKUretCikNoDept As Boolean = False
        Dim lSatisDusmesin As Boolean = False
        Dim ConnYage As SqlConnection

        Try
            ConnYage = OpenConn()

            G_MTKUretCikNoDept = (GetSysParConnected("mtkuretciknodept", ConnYage) = "1")
            lSatisDusmesin = (GetSysParConnected("mtfsatisdusmesin", ConnYage) = "1")

            cSQL = "update mtkfislines " +
                    " set " +
                    " isemriverilen = 0, " +
                    " isemriicingelen = 0, " +
                    " isemriharicigelen = 0, " +
                    " isemriicingiden = 0, " +
                    " isemriharicigiden = 0, " +
                    " uretimicincikis = 0, " +
                    " uretimdeniade = 0 " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSql1 = " select stokno = coalesce(b.stokno,'') , " +
                            " renk = coalesce(b.renk ,'') , " +
                            " beden = coalesce(b.beden,'') , " +
                            " malzemetakipkodu = coalesce(b.malzemetakipkodu ,'') , " +
                            " isemrino = coalesce(b.isemrino,'') , " +
                            " giris = sum(coalesce(b.netmiktar1,0)) , " +
                            " cikis = 0 , " +
                            " stokhareketkodu = coalesce(b.stokhareketkodu,'') , " +
                            " departman = coalesce(a.departman,'') " +
                    " from stokfis a , stokfislines b " +
                    " Where a.stokfisno = b.stokfisno " +
                            IIf(cMTFNo.Trim = "", " and b.malzemetakipkodu is not null and b.malzemetakipkodu <> '' ", " and b.malzemetakipkodu = '" + cMTFNo.Trim + "' ").ToString +
                            " and (a.iptal is null or a.iptal = '' or a.iptal = 'H') " +
                            " and a.stokfistipi in ('Giris','02 Satis Iade','03 Defolu iade') " +
                            IIf(lSatisDusmesin, " and not b.stokhareketkodu in ('07 Satis Iade','07 Satis') ", "").ToString +
                    " group by b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.isemrino, b.stokhareketkodu, a.departman "

            cSql2 = " select stokno = coalesce(b.stokno,'') , " +
                            " renk = coalesce(b.renk ,'') , " +
                            " beden = coalesce(b.beden,'') , " +
                            " malzemetakipkodu = coalesce(b.malzemetakipkodu ,'') , " +
                            " isemrino = coalesce(b.isemrino,'') , " +
                            " giris = 0 , " +
                            " cikis = sum(coalesce(b.netmiktar1,0)) , " +
                            " stokhareketkodu = coalesce(b.stokhareketkodu,'') , " +
                            " departman = coalesce(a.departman,'') " +
                    " from stokfis a , stokfislines b " +
                    " Where a.stokfisno = b.stokfisno " +
                            IIf(cMTFNo.Trim = "", " and b.malzemetakipkodu is not null and b.malzemetakipkodu <> '' ", " and b.malzemetakipkodu = '" + cMTFNo.Trim + "' ").ToString +
                            " and (a.iptal is null or a.iptal = '' or a.iptal = 'H') " +
                            " and a.stokfistipi in ('Cikis','01 Satis') " +
                            IIf(lSatisDusmesin, " and not b.stokhareketkodu in ('07 Satis Iade','07 Satis') ", "").ToString +
                    " group by b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.isemrino, b.stokhareketkodu, a.departman "

            cSql3 = " select stokno = coalesce(stokno,'') , " +
                            " renk = coalesce(renk ,'') , " +
                            " beden = coalesce(beden ,'') , " +
                            " malzemetakipkodu = coalesce(hedefmalzemetakipno,'') , " +
                            " isemrino = '' , " +
                            " giris = sum(coalesce(netmiktar1,0)), " +
                            " cikis = 0 , " +
                            " stokhareketkodu = '90 Trans/Rezv Giris' , " +
                            " departman = '' " +
                    " From StokTransfer " +
                    IIf(cMTFNo.Trim = "", " where hedefmalzemetakipno is not null and hedefmalzemetakipno <> '' ", " where hedefmalzemetakipno = '" + cMTFNo.Trim + "' ").ToString +
                    " group by stokno, renk, beden, hedefmalzemetakipno "

            cSql4 = " select stokno = coalesce(stokno,'') , " +
                            " renk = coalesce(renk ,'') , " +
                            " beden = coalesce(beden,'') , " +
                            " malzemetakipkodu = coalesce(kaynakmalzemetakipno,'') , " +
                            " isemrino = '' , " +
                            " giris = 0 , " +
                            " cikis = sum(coalesce(netmiktar1,0)) , " +
                            " stokhareketkodu = '90 Trans/Rezv Cikis' , " +
                            " departman = '' " +
                    " From StokTransfer " +
                    IIf(cMTFNo.Trim = "", " where kaynakmalzemetakipno is not null and kaynakmalzemetakipno <> '' ", " where kaynakmalzemetakipno = '" + cMTFNo.Trim + "' ").ToString +
                    " group by stokno, renk, beden, kaynakmalzemetakipno "

            cSQL = cSql1 + " Union All " +
                   cSql2 + " Union All " +
                   cSql3 + " Union All " +
                   cSql4

            cTempView = CreateTempView(ConnYage, cSQL)

            ' hareket kodlarina gore update

            ' 01 Uretime Cikis
            ' üretim departmanı belli olduğu için MTF de ilgili üretim departmanlı satır için çıkış olur
            cSQL = "update mtkfislines " +
                    " set uretimicincikis  = (select coalesce(sum(coalesce(cikis,0)),0) " +
                                                " from " + cTempView + " b " +
                                                " Where mtkfislines.stokno = b.stokno  " +
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu" +
                                                " and mtkfislines.renk = b.renk" +
                                                " and mtkfislines.beden = b.beden " +
                                                IIf(G_MTKUretCikNoDept, "", " and mtkfislines.departman = b.departman ").ToString +
                                                " and " + G_uretimicincikis + " ) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' 01 uretimden iade
            ' üretim departmanı belli olduğu için MTF de ilgili üretim departmanlı satıra iade olur
            cSQL = "update mtkfislines " +
                    " set uretimdeniade  = (select coalesce(sum(coalesce(giris,0)),0) " +
                                                " from " + cTempView + " b " +
                                                " Where mtkfislines.stokno = b.stokno  " +
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu" +
                                                " and mtkfislines.renk = b.renk" +
                                                " and mtkfislines.beden = b.beden " +
                                                IIf(G_MTKUretCikNoDept, "", " and mtkfislines.departman = b.departman ").ToString +
                                                " and " + G_uretimdeniade + " ) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' İşEmriİçinGelen alanı tedarikten yapılmış NET girişi gösterir
            ' 02 Tedarikten Giris
            ' 04 Mlz Uretimden Giris
            ' 05 Diger Giris            -> aslinda isemri no girdikten sonra diger giris olmamali, 02 veya 04 yapilmali
            ' 06 Tamirden Giris
            cSQL = "update mtkfislines " +
                    " set isemriicingelen  = (select coalesce(sum(coalesce(giris,0)),0) " +
                                        " from " + cTempView + " b " +
                                        " Where mtkfislines.stokno = b.stokno  " +
                                        " and mtkfislines.malzemetakipno = b.malzemetakipkodu" +
                                        " and mtkfislines.renk = b.renk" +
                                        " and mtkfislines.beden= b.beden" +
                                        " and isemrino is not null " +
                                        " and isemrino <> '' " +
                                        " and " + G_isemriicinGelenGiris + " ) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' 02 Tedarikten iade
            ' 04 Mlz Uretime iade
            ' 05 Diger Cikis
            ' 06 Tamirden Giris
            cSQL = "update mtkfislines " +
                    " set isemriicingelen  = coalesce(isemriicingelen,0) - (select coalesce(sum(coalesce(cikis,0)),0) " +
                                                                            " from " + cTempView + " b " +
                                                                            " Where mtkfislines.stokno = b.stokno  " +
                                                                            " and mtkfislines.malzemetakipno = b.malzemetakipkodu" +
                                                                            " and mtkfislines.renk = b.renk" +
                                                                            " and mtkfislines.beden= b.beden" +
                                                                            " and isemrino is not null " +
                                                                            " and isemrino <> '' " +
                                                                            " and " + G_isemriicinGelenCikis + " ) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set isemriharicigelen  = (select coalesce(sum(coalesce(giris,0)),0) " +
                                            " from " + cTempView + " b " +
                                            " Where mtkfislines.stokno = b.stokno  " +
                                            " and mtkfislines.renk = b.renk " +
                                            " and mtkfislines.beden= b.beden " +
                                            " and mtkfislines.malzemetakipno = b.malzemetakipkodu) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set isemriharicigelen  = coalesce(isemriharicigelen,0) - (select coalesce(sum(coalesce(cikis,0)),0) " +
                                                                                " from " + cTempView + " b " +
                                                                                " Where mtkfislines.stokno = b.stokno  " +
                                                                                " and mtkfislines.renk = b.renk " +
                                                                                " and mtkfislines.beden= b.beden " +
                                                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' Sonuç denklemde, Karşılanan = İşEmriİçinGelen + İşEmriHariciGelen
            ' Yani, Karşılanan = Üretime Net Çıkan + Elimizdeki Net Rezerve Miktar
            ' İşEmriHariciGelen = (ÜretimeÇıkan NET miktar)+ (MTF ye yapılmış NET rezervasyonlar (elimizdeki rezerve malzeme)) - (İşEmriİçinGelen NET miktar)

            cSQL = "update mtkfislines " +
                    " set isemriharicigelen  = (coalesce(uretimicincikis,0) - coalesce(uretimdeniade,0)) " +
                                                " + coalesce(isemriharicigelen,0) " +
                                                " - coalesce(isemriicingelen,0) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set isemriicingiden  = (select coalesce(sum(coalesce(cikis,0)),0) " +
                                                " from " + cTempView + " b " +
                                                " Where mtkfislines.stokno = b.stokno " +
                                                " and mtkfislines.renk = b.renk " +
                                                " and mtkfislines.beden= b.beden " +
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu " +
                                                " and isemrino is not null " +
                                                " and isemrino <> '' ) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set isemriharicigiden  = (select coalesce(sum(coalesce(cikis,0)),0) " +
                                                " from " + cTempView + " b " +
                                                " Where mtkfislines.stokno = b.stokno " +
                                                " and mtkfislines.renk = b.renk " +
                                                " and mtkfislines.beden= b.beden " +
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu " +
                                                " and (isemrino is null or isemrino = '') ) " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)
            ' üretim departmanına göre işemri verilen adet
            ' her işemri satırında üretim departmanı olması gerekiyor
            cSQL = "update mtkfislines " +
                    " set isemriverilen = (select coalesce(sum(coalesce(miktar1,0)),0) " +
                                            " from isemrilines b " +
                                            " Where mtkfislines.stokno = b.stokno  " +
                                            " and mtkfislines.malzemetakipno = b.malzemetakipno " +
                                            " and mtkfislines.renk = b.renk " +
                                            " and mtkfislines.beden = b.beden " +
                                            " and coalesce(mtkfislines.departman,'') = coalesce(b.departman,'') " +
                                            " and b.isemrino is not null " +
                                            " and b.isemrino <> '') " +
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' katsayılar doğru olarak hesaplanmış olmalıdır
            If G_MTKUretCikNoDept Then
                cSQL = "update mtkfislines " +
                        " set isemriicingelen = coalesce(isemriicingelen,0) * coalesce(katsayi,0), " +
                        " isemriharicigelen = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " +
                        " isemriicingiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " +
                        " isemriharicigiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " +
                        " uretimicincikis = coalesce(uretimicincikis,0) * coalesce(katsayi,0), " +
                        " uretimdeniade = coalesce(uretimdeniade,0) * coalesce(katsayi,0) " +
                        IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString
            Else
                cSQL = "update mtkfislines " +
                        " set isemriicingelen = coalesce(isemriicingelen,0) * coalesce(katsayi,0), " +
                        " isemriharicigelen = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " +
                        " isemriicingiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " +
                        " isemriharicigiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0) " +
                        IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' eğer yeni duruma göre ihtiyaç > ÜRETIME ÇIKAN ise ilgili satırları aç
            cSQL = "update mtkfislines " +
                    " set kapandi = 'H' " +
                    " where coalesce(ihtiyac,0) > coalesce(uretimicincikis,0) - coalesce(uretimdeniade,0) " +
                    " and kapandi in ('E','e') " +
                    " and musteriihtiyac is not null " +
                    " and musteriihtiyac <> 0 " +
                    " and malzemetakipno is not null " +
                    " and malzemetakipno <> '' " +
                    IIf(cMTFNo.Trim = "", "", " and malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropView(cTempView, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "MTKLinesTopl", cSQL)
        End Try
    End Sub

    Public Sub G_IsemriDeptKontrol(Optional cMTF As String = "")
        ' işemrinde üretime çıkış departmanı yoksa
        ' ilgili MTF deki en büyük ihtiyaç miktarına sahip üretim departmanını işemrindeki boş üretim departmanına atıyoruz
        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection

        Try
            ConnYage = OpenConn()
            ' departmanı BOŞ atılmış satırları tamamlar
            cSQL = "update isemrilines " +
                    " set departman = (select top 1 departman " +
                                        " from mtkfislines " +
                                        " where malzemetakipno = isemrilines.malzemetakipno " +
                                        " and stokno = isemrilines.stokno " +
                                        " and renk = isemrilines.renk " +
                                        " and beden = isemrilines.beden " +
                                        " order by ihtiyac desc) " +
                    " where (departman is null or departman = '') "

            If cMTF.Trim = "" Then
                cSQL = cSQL +
                    " and malzemetakipno is not null " +
                    " and malzemetakipno <> '' "
            Else
                cSQL = cSQL +
                    " and malzemetakipno = '" + cMTF.Trim + "' "
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' departmanı yanlış girilmiş satırları toparlar
            cSQL = "update isemrilines " +
                    " set departman = (select top 1 departman " +
                                        " from mtkfislines " +
                                        " where malzemetakipno = isemrilines.malzemetakipno " +
                                        " and stokno = isemrilines.stokno " +
                                        " and renk = isemrilines.renk " +
                                        " and beden = isemrilines.beden " +
                                        " order by ihtiyac desc) " +
                    " where departman is not null " +
                    " and departman <> '' " +
                    " and not exists (select malzemetakipno " +
                                    " from mtkfislines " +
                                    " where StokNo = isemrilines.StokNo " +
                                    " and renk = isemrilines.renk " +
                                    " and beden = isemrilines.beden " +
                                    " and coalesce(departman,'') = coalesce(isemrilines.departman,'')) "
            If cMTF.Trim = "" Then
                cSQL = cSQL +
                    " and malzemetakipno is not null " +
                    " and malzemetakipno <> '' "
            Else
                cSQL = cSQL +
                    " and malzemetakipno = '" + cMTF.Trim + "' "
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "G_IsemriDeptKontrol", cSQL)
        End Try
    End Sub

    Public Sub MTKFindLost(Optional cMTF As String = "")

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection

        Try
            ConnYage = OpenConn()

            cSQL = "insert mtkfislines (malzemetakipno, stokno, renk, beden, ihtiyac, birim, departman, temindept) "

            cSQL = cSQL +
                    " select distinct b.malzemetakipkodu, " +
                    " b.stokno, b.renk, b.beden,  " +
                    " ihtiyac = 0, " +
                    " birim = c.birim1, " +
                    " departman = c.uretimdepartmani, " +
                    " temindept = c.temindepartmani " +
                    " from stokfis a, stokfislines b, stok c  " +
                    " where a.stokfisno = b.stokfisno " +
                    " and b.stokno = c.stokno " +
                    " and b.malzemetakipkodu = '" + cMTF.Trim + "' " +
                    " and not exists (select malzemetakipno " +
                                    " from mtkfislines " +
                                    " where malzemetakipno = b.malzemetakipkodu " +
                                    " and stokno = b.stokno " +
                                    " and renk = b.renk " +
                                    " and beden = b.beden) " +
                    " union "

            cSQL = cSQL +
                    " select distinct b.malzemetakipno, " +
                    " b.stokno, b.renk, b.beden,  " +
                    " ihtiyac = 0, " +
                    " birim = c.birim1, " +
                    " departman = c.uretimdepartmani, " +
                    " temindept = c.temindepartmani " +
                    " from isemri a, isemrilines b, stok c  " +
                    " where a.isemrino = b.isemrino " +
                    " and b.stokno = c.stokno " +
                    " and b.malzemetakipno = '" + cMTF.Trim + "' " +
                    " and not exists (select malzemetakipno " +
                                    " from mtkfislines " +
                                    " where malzemetakipno = b.malzemetakipno " +
                                    " and stokno = b.stokno " +
                                    " and renk = b.renk " +
                                    " and beden = b.beden) " +
                    " union "

            cSQL = cSQL +
                    " select distinct malzemetakipno = a.kaynakmalzemetakipno , " +
                    " a.stokno, a.renk, a.beden,  " +
                    " ihtiyac = 0, " +
                    " birim = c.birim1, " +
                    " departman = c.uretimdepartmani, " +
                    " temindept = c.temindepartmani " +
                    " from stoktransfer a, stok c  " +
                    " where a.stokno = c.stokno " +
                    " and a.kaynakmalzemetakipno = '" + cMTF.Trim + "' " +
                    " and not exists (select malzemetakipno " +
                                    " from mtkfislines " +
                                    " where malzemetakipno = a.kaynakmalzemetakipno " +
                                    " and stokno = a.stokno " +
                                    " and renk = a.renk " +
                                    " and beden = a.beden) " +
                    " union "

            cSQL = cSQL +
                    " select distinct malzemetakipno = a.hedefmalzemetakipno , " +
                    " a.stokno, a.renk, a.beden,  " +
                    " ihtiyac = 0, " +
                    " birim = c.birim1, " +
                    " departman = c.uretimdepartmani, " +
                    " temindept = c.temindepartmani " +
                    " from stoktransfer a, stok c  " +
                    " where a.stokno = c.stokno " +
                    " and a.hedefmalzemetakipno = '" + cMTF.Trim + "' " +
                    " and not exists (select malzemetakipno " +
                                    " from mtkfislines " +
                                    " where malzemetakipno = a.hedefmalzemetakipno " +
                                    " and stokno = a.stokno " +
                                    " and renk = a.renk " +
                                    " and beden = a.beden) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "MTKFindLost", cSQL)
        End Try
    End Sub

    Private Sub MTKCleanUP(Optional cMTF As String = "")

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection

        Try
            If cMTF.Trim = "" Then Exit Sub

            ConnYage = OpenConn()

            cSQL = "delete mtkfislines " +
                    " where malzemetakipno = '" + cMTF.Trim + "' " +
                    " and (ihtiyac = 0 or ihtiyac is null) " +
                    " and (musteriihtiyac = 0 or musteriihtiyac is null) " +
                    " and (ihtiyatiihtiyac = 0 or ihtiyatiihtiyac is null) " +
                    " and (hesaplananihtiyac = 0 or hesaplananihtiyac is null) " +
                    " and (isemriicingiden = 0 or isemriicingiden is null) " +
                    " and (isemriharicigiden = 0 or isemriharicigiden is null) " +
                    " and (uretimicincikis = 0 or uretimicincikis is null) " +
                    " and (uretimdeniade = 0 or uretimdeniade is null) " +
                    " and (isemriverilen = 0 or isemriverilen is null) " +
                    " and (isemriicingelen = 0 or isemriicingelen is null) " +
                    " and (isemriharicigelen = 0 or isemriharicigelen is null) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()
        Catch ex As Exception
            ErrDisp(ex.Message, "MTKCleanUP", cSQL)
        End Try
    End Sub

    Public Sub GetSonStokGiristenFiyat(cStokNo As String, cRenk As String, ByRef nFiyat As Double, ByRef cDoviz As String, Optional ByRef dTarih As Date = #1/1/1950#, Optional cBeden As String = "", Optional cMTF As String = "")

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        Try
            ConnYage = OpenConn()

            nFiyat = 0
            cDoviz = ""

            cSQL = "select a.birimfiyat, a.dovizcinsi, b.fistarihi " +
               " from stokfislines a, stokfis b " +
               " where a.stokfisno = b.stokfisno " +
               " and a.stokno = '" + Trim(cStokNo) + "' " +
               " and a.renk = '" + Trim(cRenk) + "' " +
               " and a.birimfiyat is not null " +
               " and a.birimfiyat > 0 " +
               " and a.stokhareketkodu in ('02 Tedarikten Giris','04 Mlz Uretimden Giris','05 Diger Giris') " +
               IIf(Trim(cBeden) = "", "", " and a.beden = '" + Trim(cBeden) + "' ").ToString +
               IIf(Trim(cMTF) = "", "", " and a.malzemetakipkodu = '" + Trim(cMTF) + "' ").ToString +
               IIf(dTarih = #1/1/1950#, "", " and b.fistarihi <= '" + CStr(dTarih) + "' ").ToString +
               " order by b.fistarihi desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "birimfiyat")
                cDoviz = SQLReadString(oReader, "dovizcinsi")
                dTarih = SQLReadDate(oReader, "fistarihi")
            End If
            oReader.Close()

            If nFiyat = 0 Then

                cSQL = "select a.birimfiyat, a.dovizcinsi, b.fistarihi " +
                   " from stokfislines a, stokfis b " +
                   " where a.stokfisno = b.stokfisno " +
                   " and a.stokno = '" + Trim(cStokNo) + "' " +
                   " and a.renk = '" + Trim(cRenk) + "' " +
                   " and a.birimfiyat is not null " +
                   " and a.birimfiyat > 0 " +
                   " and a.stokhareketkodu in ('02 Tedarikten Giris','04 Mlz Uretimden Giris','05 Diger Giris') " +
                   IIf(Trim(cBeden) = "", "", " and a.beden = '" + Trim(cBeden) + "' ").ToString +
                   IIf(dTarih = #1/1/1950#, "", " and b.fistarihi <= '" + CStr(dTarih) + "' ").ToString +
                   " order by b.fistarihi desc "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "birimfiyat")
                    cDoviz = SQLReadString(oReader, "dovizcinsi")
                    dTarih = SQLReadDate(oReader, "fistarihi")
                End If
                oReader.Close()
            End If

            If nFiyat <> 0 Then
                If Trim(cDoviz) = "" Then
                    cDoviz = "TL"
                End If
            End If

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSonStokGiristenFiyat", cSQL)
        End Try
    End Sub

    Private Sub MTKOnMaliyet(Optional cMTFNo As String = "")
        ' Dokuma Ön Maliyet Çalışması fiyatları Hedef Fiyat Olarak Alınır
        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection

        Try
            If cMTFNo.Trim = "" Then Exit Sub

            ConnYage = OpenConn()

            cSQL = "update mtkfislines " +
                    " set hedefmlzbirimfiyati = (select top 1 fiyat  " +
                            " from maliyetkumas " +
                            " where stokno = mtkfislines.stokno  " + 
							" and renk = mtkfislines.renk " +
							" and fiyat is not null " +
							" and fiyat <> 0 " +
							" and calismano in (select b.maliyetcalismano " +
												" from sipmodel a, ymodel b " +
												" where a.modelno = b.modelno " +
												" and a.malzemetakipno = mtkfislines.malzemetakipno) " +
							" order by fiyat), "
            cSQL = cSQL +
                    " hedefmlzdovizi = (select top 1 doviz " +
                            " from maliyetkumas " +
                            " where stokno = mtkfislines.stokno " +
                            " and renk = mtkfislines.renk " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano  " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno)  " +
                            " order by fiyat), "
            cSQL = cSQL +
                    " plfirma = (select top 1 firma " +
                            " from maliyetkumas " +
                            " where stokno = mtkfislines.stokno " +
                            " and renk = mtkfislines.renk " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano  " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno)  " +
                            " order by fiyat) "
            cSQL = cSQL +
                " where malzemetakipno = '" + cMTFNo.Trim + "' " +
                " and (hedefmlzbirimfiyati is null or hedefmlzbirimfiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set hedefmlzbirimfiyati = (select top 1 fiyat  " +
                            " from maliyetdikim " +
                            " where stokno = mtkfislines.stokno  " +
                            " and renk = mtkfislines.renk " +
                            " and beden = mtkfislines.beden " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno) " +
                            " order by fiyat), "
            cSQL = cSQL +
                    " hedefmlzdovizi = (select top 1 doviz " +
                            " from maliyetdikim " +
                            " where stokno = mtkfislines.stokno " +
                            " and renk = mtkfislines.renk " +
                            " and beden = mtkfislines.beden " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano  " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno)  " +
                            " order by fiyat), "
            cSQL = cSQL +
                    " plfirma = (select top 1 firma " +
                            " from maliyetdikim " +
                            " where stokno = mtkfislines.stokno " +
                            " and renk = mtkfislines.renk " +
                            " and beden = mtkfislines.beden " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano  " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno)  " +
                            " order by fiyat) "
            cSQL = cSQL +
                " where malzemetakipno = '" + cMTFNo.Trim + "' " +
                " and (hedefmlzbirimfiyati is null or hedefmlzbirimfiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " +
                    " set hedefmlzbirimfiyati = (select top 1 fiyat  " +
                            " from maliyetpaketleme " +
                            " where stokno = mtkfislines.stokno  " +
                            " and renk = mtkfislines.renk " +
                            " and beden = mtkfislines.beden " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno) " +
                            " order by fiyat), "
            cSQL = cSQL +
                    " hedefmlzdovizi = (select top 1 doviz " +
                            " from maliyetpaketleme " +
                            " where stokno = mtkfislines.stokno " +
                            " and renk = mtkfislines.renk " +
                            " and beden = mtkfislines.beden " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano  " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno)  " +
                            " order by fiyat), "
            cSQL = cSQL +
                    " plfirma = (select top 1 firma " +
                            " from maliyetpaketleme " +
                            " where stokno = mtkfislines.stokno " +
                            " and renk = mtkfislines.renk " +
                            " and beden = mtkfislines.beden " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano  " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.malzemetakipno = mtkfislines.malzemetakipno)  " +
                            " order by fiyat) "
            cSQL = cSQL +
                " where malzemetakipno = '" + cMTFNo.Trim + "' " +
                " and (hedefmlzbirimfiyati is null or hedefmlzbirimfiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines  " +
                    " set uretimecikisfiyati = hedefmlzbirimfiyati, " +
                    " uretimecfdovizi = hedefmlzdovizi " +
                    " where malzemetakipno = '" + cMTFNo.Trim + "' " +
                    " and (uretimecikisfiyati is null Or uretimecikisfiyati = 0) " +
                    " and hedefmlzbirimfiyati is Not null " +
                    " and hedefmlzbirimfiyati <> 0 "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "MTKOnMaliyet", cSQL)
        End Try
    End Sub

End Module

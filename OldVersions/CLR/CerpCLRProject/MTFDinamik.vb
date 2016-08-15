Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module MTFDinamik

    Public Structure oStokRB
        Dim cStokno As String
        Dim cAciklama As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cAnaStokGrubu As String
        Dim cStokTipi As String
        Dim cBirim As String
        Dim cMTF As String
        Dim nIhtiyac As Double
        Dim nUretimeCikan As Double
        Dim nStokMiktari As Double
        Dim nGelecek As Double
        Dim nSerbestStokMiktari As Double
        Dim nSerbestGelecek As Double
        Dim dTermin As Date
        Dim lKapali As Boolean
        Dim lSecildi As Boolean
        Dim cDepartman As String
        Dim nIsemriVerilen As Double
    End Structure

    Public Structure oSRB
        Dim cStokno As String
        Dim cRenk As String
        Dim cBeden As String
        Dim nSerbestStokMiktari As Double   ' toplam stok miktarı
        Dim nSerbestGelecek As Double       ' toplam gelecek
        Dim nSGelecek As Double             ' serbest stoklara gelecek
    End Structure

    Public Function GetToplamSiparisView_0(Optional cFilter As String = "", Optional ByRef cTable As String = "") As Integer

        Dim ConnYage As SqlConnection

        Try
            ConnYage = OpenConn()
            GetToplamSiparisView_0 = GetToplamSiparisView_1(cFilter, cTable, ConnYage)
            CloseConn(ConnYage)
        Catch ex As Exception
            ErrDisp(ex.Message, "GetToplamSiparisView_0")
        End Try
    End Function

    Public Function GetToplamSiparisView_1(Optional cFilter As String = "", Optional ByRef cTable As String = "", Optional ConnYage As SqlConnection = Nothing) As Integer
        ' cFilter -> b.stokno, b.renk, b.beden olabilir
        Dim cSQL As String = ""
        Dim cStokView As String = ""
        Dim cMTFView As String = ""
        Dim nSYETolerans As Double = 0

        GetToplamSiparisView_1 = 0

        Try
            JustForLog("GetToplamSiparisView_1 beg : " + cTable)

            nSYETolerans = CDbl(GetSysParConnected("syetolerans", ConnYage))

            ' transfer fişleri hesaba katılmaz
            ' son stok durumu view haline getiriliyor
            cSQL = "select stokno = coalesce(b.stokno,''), " + _
                    " renk = coalesce(b.renk,''), " + _
                    " beden = coalesce(b.beden,''), " + _
                    " giris = sum(coalesce(b.netmiktar1,0)), " + _
                    " cikis = 0 " + _
                    " from stokfis a , stokfislines b " + _
                    " where a.stokfisno = b.stokfisno " + _
                    " and a.stokfistipi in ('Giris','02 Satis Iade','03 Defolu iade') " + _
                    " and (a.iptal <> 'E' or a.iptal is null) " + _
                    cFilter + _
                    " group by b.stokno, b.renk, b.beden " + _
                    " Union All "

            cSQL = cSQL + _
                    " select stokno = coalesce(b.stokno,''), " + _
                    " renk = coalesce(b.renk,''), " + _
                    " beden = coalesce(b.beden,''), " + _
                    " giris = 0, " + _
                    " cikis = sum(coalesce(b.netmiktar1,0)) " + _
                    " from stokfis a , stokfislines b " + _
                    " where a.stokfisno = b.stokfisno  " + _
                    " and a.stokfistipi in ('Cikis','01 Satis') " + _
                    " and (a.iptal <> 'E' or a.iptal is null) " + _
                    cFilter + _
                    " group by b.stokno, b.renk, b.beden "

            cStokView = CreateTempView(ConnYage, cSQL)

            If nSYETolerans > 0 Then
                ' üretime çıkan miktar ihtiyacın örneğin 80% ından fazlaysa o satırın satınalmasını kapatabiliriz
                cSQL = "update mtkfislines " + _
                        " set kapandi = 'E' " + _
                        " where (kapandi is null or kapandi = 'H' or kapandi = '') " + _
                        " and (coalesce(ihtiyac,0) * " + SQLWriteDecimal(nSYETolerans / 100) + ") <= coalesce(uretimicincikis,0) " + _
                        " and (coalesce(musteriihtiyac,0) + coalesce(ihtiyatiihtiyac,0) = 0) "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            cSQL = "select b.stokno, b.renk, b.beden, " + _
                    " ihtiyac = sum(coalesce(b.ihtiyac,0)), "
            ' üretime çıkan malzemeler
            ' üretime çıkan malzemeler MTF ye bağlanmak zorundadır
            cSQL = cSQL + _
                    " uretimecikan = ( " + _
                    " select coalesce(sum(coalesce(y.netmiktar1,0)),0) " + _
                            " from stokfis z, stokfislines y " + _
                            " where z.stokfisno = y.stokfisno " + _
                            " and y.stokhareketkodu = '01 Uretime Cikis' " + _
                            " and y.stokno = b.stokno " + _
                            " and y.renk = b.renk " + _
                            " and y.beden = b.beden " + _
                            " and not (y.malzemetakipkodu is null or y.malzemetakipkodu = '') "
            ' kapanmamış MTF satırlarından üretime çıkanlar alınır
            cSQL = cSQL + _
                            " and exists (select malzemetakipno " + _
                                        " from mtkfislines " + _
                                        " where y.malzemetakipkodu = malzemetakipno " + _
                                        " and y.stokno = stokno " + _
                                        " and y.renk = renk " + _
                                        " and y.beden = beden " + _
                                        " and z.departman = departman " + _
                                        " and (kapandi is null or kapandi = 'H' or kapandi = '')) "
            ' kapanmamış MTF lerden üretime çıkanlar alınır
            cSQL = cSQL + _
                            " and exists (select malzemetakipno " + _
                                        " from mtkfis " + _
                                        " where y.malzemetakipkodu = malzemetakipno " + _
                                        " and (dosyakapandi is null or dosyakapandi = 'H' or dosyakapandi = '')) " + _
                    " ), "
            ' üretimden iade edilen malzeme satırında bir açıklama (sakat kodu) yoksa
            ' bu miktarı üretime çıkan miktardan düşmüyoruz (jeanci)
            cSQL = cSQL + _
                    " uretimdeniade = ( " + _
                    " select coalesce(sum(coalesce(y.netmiktar1,0)),0) " + _
                            " from stokfis z, stokfislines y " + _
                            " where z.stokfisno = y.stokfisno " + _
                            " and y.stokhareketkodu = '01 Uretimden iade' " + _
                            " and y.stokno = b.stokno " + _
                            " and y.renk = b.renk " + _
                            " and y.beden = b.beden " + _
                            " and not (y.malzemetakipkodu is null or y.malzemetakipkodu = '') " + _
                            " and y.sakatkodu is not null " + _
                            " and y.sakatkodu <> '' "
            ' kapanmamış MTF satırlarından üretime çıkanlar alınır
            cSQL = cSQL + _
                            " and exists (select malzemetakipno " + _
                                        " from mtkfislines " + _
                                        " where y.malzemetakipkodu = malzemetakipno " + _
                                        " and y.stokno = stokno " + _
                                        " and y.renk = renk " + _
                                        " and y.beden = beden " + _
                                        " and z.departman = departman " + _
                                        " and (kapandi is null or kapandi = 'H' or kapandi = '')) "
            ' kapanmamış MTF lerden üretime çıkanlar alınır
            cSQL = cSQL + _
                            " and exists (select malzemetakipno " + _
                                        " from mtkfis " + _
                                        " where y.malzemetakipkodu = malzemetakipno " + _
                                        " and (dosyakapandi is null or dosyakapandi = 'H' or dosyakapandi = '')) " + _
                    " ), "
            ' Açık işemirlerinden serbest stoklara gelecek satınalmalar
            cSQL = cSQL + _
                    " serbestgelecek = ( " + _
                    " select sum(coalesce(y.miktar1,0) - coalesce(y.uretimgelen,0) - coalesce(y.tedarikgelen,0)) " + _
                            " from isemri z, isemrilines y " + _
                            " where z.isemrino = y.isemrino " + _
                            " and y.stokno = b.stokno " + _
                            " and y.renk = b.renk " + _
                            " and y.beden = b.beden " + _
                            " and (z.isemriok is null or z.isemriok = 'H' or z.isemriok = '') " + _
                            " and (coalesce(y.miktar1,0) - coalesce(y.uretimgelen,0) - coalesce(y.tedarikgelen,0) > 0) "
            ' malzeme serbeste gelecekse veya
            ' kapanmış MTF ye gelecekse veya
            ' kapanmış MTF satırına gelecekse serbest gibi dağıtıma girecek demektir
            cSQL = cSQL + _
                            " and ( " + _
                                    " y.malzemetakipno is null " + _
                                    " or y.malzemetakipno = '' " + _
                                    " or exists (select malzemetakipno " + _
                                                " from mtkfis " + _
                                                " where y.malzemetakipno = malzemetakipno " + _
                                                " and dosyakapandi = 'E') " + _
                                    " or not exists (select malzemetakipno " + _
                                                " from mtkfislines " + _
                                                " where y.malzemetakipno = malzemetakipno " + _
                                                " and y.stokno = stokno " + _
                                                " and y.renk = renk " + _
                                                " and y.beden = beden " + _
                                                " and (kapandi = 'H' or kapandi = '' or kapandi is null)) " + _
                                    ") " + _
                    " ), "
            ' Açık işemirlerinden toplam gelecek satınalmalar
            cSQL = cSQL + _
                    " gelecek = ( " + _
                    " select sum(coalesce(y.miktar1,0) - coalesce(y.uretimgelen,0) - coalesce(y.tedarikgelen,0)) " + _
                            " from isemri z, isemrilines y " + _
                            " where z.isemrino = y.isemrino " + _
                            " and y.stokno = b.stokno " + _
                            " and y.renk = b.renk " + _
                            " and y.beden = b.beden " + _
                            " and (z.isemriok is null or z.isemriok = 'H' or z.isemriok = '') " + _
                            " and (coalesce(y.miktar1,0) - coalesce(y.uretimgelen,0) - coalesce(y.tedarikgelen,0) > 0) " + _
                    " ), "
            ' toplam stok miktarı (MTF + serbest)
            ' rezerve veya değil bütün stoklar serbest muamelesi görür
            cSQL = cSQL + _
                    " stokmiktari = (select sum(coalesce(z.giris,0) - coalesce(z.cikis,0)) " + _
                                " from " + cStokView + " z " + _
                                " where z.stokno = b.stokno " + _
                                " and z.renk = b.renk " + _
                                " and z.beden = b.beden) "
            ' açık MTF ve MTF satırları hesaba dahil edilir
            cSQL = cSQL + _
                    " from mtkfis a, mtkfislines b " + _
                    " where a.malzemetakipno = b.malzemetakipno " + _
                    " and (a.dosyakapandi is null or a.dosyakapandi = 'H' or a.dosyakapandi = '') " + _
                    " and (b.kapandi is null or b.kapandi = 'H' or b.kapandi = '') " + _
                    cFilter + _
                    " group by b.stokno, b.renk, b.beden "

            cMTFView = CreateTempView(ConnYage, cSQL)

            cSQL = " (stokno char(30) null, " + _
                    " cinsaciklamasi char(255) null, " + _
                    " anastokgrubu char(30) null, " + _
                    " stoktipi char(30) null, " + _
                    " birim char(30) null, " + _
                    " renk char(30) null, " + _
                    " beden char(30) null, " + _
                    " ihtiyac decimal(18,2) null, " + _
                    " uretimecikan decimal(18,2) null, " + _
                    " gelecek decimal(18,2) null, " + _
                    " stokmiktari decimal(18,2) null, " + _
                    " serbestgelecek decimal(18,2) null) "

            cTable = CreateTempTable(ConnYage, cSQL, cTable)

            cSQL = "insert into " + cTable + _
                    " (stokno, cinsaciklamasi, stoktipi, anastokgrubu, birim, " + _
                    " renk, beden, ihtiyac, uretimecikan, gelecek, " + _
                    " stokmiktari, serbestgelecek) "

            cSQL = cSQL + _
                    " select a.stokno, b.cinsaciklamasi, b.stoktipi, b.anastokgrubu, b.birim1, " + _
                    " a.renk, a.beden, a.ihtiyac, uretimecikan = coalesce(a.uretimecikan,0) - coalesce(a.uretimdeniade,0), a.gelecek, " + _
                    " a.stokmiktari, a.serbestgelecek " + _
                    " from " + cMTFView + " a , stok b " + _
                    " where a.stokno = b.stokno " + _
                    " order by a.stokno, a.renk, a.beden "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropView(cMTFView, ConnYage)
            DropView(cStokView, ConnYage)

            GetToplamSiparisView_1 = 1

            JustForLog("GetToplamSiparisView_1 end : " + cTable)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetToplamSiparisView_1", cSQL)
        End Try
    End Function

    Public Function MTFHesaplax_0(Optional cFilter1 As String = "", Optional cFilter2 As String = "", Optional cTSip As String = "", Optional ByRef cMTFHesaplaX As String = "") As Integer

        Dim ConnYage As SqlConnection

        Try
            ConnYage = OpenConn()
            MTFHesaplax_0 = MTFHesaplax_1(cFilter1, cFilter2, cTSip, cMTFHesaplaX, ConnYage)
            CloseConn(ConnYage)
        Catch ex As Exception
            ErrDisp(ex.Message, "MTFHesaplax_0")
        End Try

    End Function

    Public Function MTFHesaplax_1(Optional cFilter1 As String = "", Optional cFilter2 As String = "", Optional cTSip As String = "", Optional ByRef cMTFHesaplaX As String = "", Optional ConnYage As SqlConnection = Nothing) As Integer
        ' cFilter1 ->
        ' cFilter2 -> stok, renk, beden filitresi olabilir
        Dim cSQL As String = ""
        Dim cMTFView As String = ""
        Dim cTable As String = ""
        Dim aSRB() As oSRB = Nothing
        Dim aStokRB() As oStokRB = Nothing
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim nFound As Integer = 0
        Dim nIhtiyac As Double = 0
        Dim nHarcanan As Double = 0
        Dim lDropTSip As Boolean = False
        Dim oReader As SqlDataReader
        Dim nSonuc As Integer = 0

        MTFHesaplax_1 = 0

        Try
            JustForLog("MTFHesaplax_1 beg : " + cMTFHesaplaX)

            If cTSip = "" Then
                nSonuc = GetToplamSiparisView_1(cFilter2, cTSip, ConnYage)
                If nSonuc = 0 Then
                    Exit Function
                End If
                lDropTSip = True
            End If

            cSQL = "select b.malzemetakipno, b.stokno, b.renk, b.beden, b.departman, " + _
                    " ihtiyac = sum(coalesce(b.ihtiyac,0)), "

            cSQL = cSQL + _
                    " uretimecikan =((select coalesce(sum(coalesce(y.netmiktar1,0)),0) " + _
                                    " from stokfis z, stokfislines y " + _
                                    " where z.stokfisno = y.stokfisno " + _
                                    " and y.stokhareketkodu = '01 Uretime Cikis' " + _
                                    " and y.stokno = b.stokno " + _
                                    " and y.renk = b.renk " + _
                                    " and y.beden = b.beden " + _
                                    " and z.departman = b.departman " + _
                                    " and y.malzemetakipkodu = b.malzemetakipno) " + _
                                    " - "
            cSQL = cSQL + _
                                    " (select coalesce(sum(coalesce(y.netmiktar1,0)),0) " + _
                                    " from stokfis z, stokfislines y " + _
                                    " where z.stokfisno = y.stokfisno " + _
                                    " and y.stokhareketkodu = '01 Uretimden iade' " + _
                                    " and y.sakatkodu is not null " + _
                                    " and y.sakatkodu <> '' " + _
                                    " and y.stokno = b.stokno " + _
                                    " and y.renk = b.renk " + _
                                    " and y.beden = b.beden " + _
                                    " and z.departman = b.departman " + _
                                    " and y.malzemetakipkodu = b.malzemetakipno)) "
            cSQL = cSQL + _
                    " from mtkfis a, mtkfislines b  " + _
                    " where a.malzemetakipno = b.malzemetakipno " + _
                    " and (a.dosyakapandi is null or a.dosyakapandi = 'H' or a.dosyakapandi = '') " + _
                    " and (b.kapandi is null or b.kapandi = 'H' or b.kapandi = '') " + _
                    cFilter2 + _
                    " and exists (select stokno " + _
                                " from " + cTSip + _
                                " where stokno = b.stokno " + _
                                " and renk = b.renk " + _
                                " and beden = b.beden) " + _
                    " group by b.malzemetakipno, b.stokno, b.renk, b.beden, b.departman "

            cMTFView = CreateTempView(ConnYage, cSQL)

            cSQL = " (malzemetakipkodu char(30) null, " + _
                    " stokno char(30) null, " + _
                    " cinsaciklamasi char(255) null, " + _
                    " anastokgrubu char(30) null, " + _
                    " stoktipi char(30) null, " + _
                    " birim char(30) null, " + _
                    " renk char(30) null, " + _
                    " beden char(30) null, " + _
                    " ihtiyac decimal(18,2) null, " + _
                    " uretimecikan decimal(18,2) null, " + _
                    " gelecek decimal(18,2) null, " + _
                    " stokmiktari decimal(18,2) null, " + _
                    " termin datetime null) "

            cTable = CreateTempTable(ConnYage, cSQL)

            cSQL = "insert into " + cTable + _
                    " (malzemetakipkodu, stokno, cinsaciklamasi, anastokgrubu, stoktipi, " + _
                    " birim, renk, beden, ihtiyac, uretimecikan, " + _
                    " gelecek, stokmiktari, termin) "

            cSQL = cSQL + _
                    " select malzemetakipkodu = a.malzemetakipno, a.stokno, b.cinsaciklamasi, b.anastokgrubu, b.stoktipi, " + _
                    " b.birim1, a.renk, a.beden, " + _
                    " ihtiyac = sum(coalesce(a.ihtiyac,0)), " + _
                    " uretimecikan = sum(coalesce(a.uretimecikan,0)), " + _
                    " gelecek = 0, " + _
                    " stokmiktari = 0, " + _
                    " termin = '01.01.1950' " + _
                    " from " + cMTFView + " a , stok b " + _
                    " where a.stokno = b.stokno " + _
                    " group by a.malzemetakipno, a.stokno, b.cinsaciklamasi, b.anastokgrubu, b.stoktipi, " + _
                    " b.birim1, a.renk, a.beden "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update " + cTable + _
                    " set termin = (select top 1 bitistarihi " + _
                                    " from mtkfislines " + _
                                    " where malzemetakipno = " + cTable + ".malzemetakipkodu " + _
                                    " and stokno = " + cTable + ".stokno " + _
                                    " and renk = " + cTable + ".renk " + _
                                    " and beden =  " + cTable + ".beden " + _
                                    " and bitistarihi is not null " + _
                                    " and bitistarihi  > '01.01.2000' " + _
                                    " order by bitistarihi) " + _
                    " where (termin is null or termin <= '01.01.2000') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update " + cTable + _
                    " set termin = (select top 1 a.ilksevktarihi " + _
                                    " from siparis a, sipmodel b  " + _
                                    " where b.malzemetakipno = " + cTable + ".malzemetakipkodu " + _
                                    " and a.kullanicisipno = b.siparisno  " + _
                                    " and a.ilksevktarihi is not null " + _
                                    " and a.ilksevktarihi  > '01.01.2000' " + _
                                    " order by a.ilksevktarihi) " + _
                    " where (termin is null or termin <= '01.01.2000') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update " + cTable + _
                    " set termin = '01.01.2099' " + _
                    " where (termin is null or termin <= '01.01.2000') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = 0

            cSQL = "select stokno, renk, beden, gelecek, stokmiktari, serbestgelecek " + _
                    " from " + cTSip + _
                    " order by stokno, renk, beden "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aSRB(nCnt)
                aSRB(nCnt).cStokno = SQLReadString(oReader, "stokno")
                aSRB(nCnt).cRenk = SQLReadString(oReader, "renk")
                aSRB(nCnt).cBeden = SQLReadString(oReader, "beden")
                ' toplam gelecek (MTF ye bağlı veya serbest depoya)
                aSRB(nCnt).nSerbestGelecek = SQLReadDouble(oReader, "gelecek")
                ' toplam stok (MTF ye bağlı veya serbest depo)
                aSRB(nCnt).nSerbestStokMiktari = SQLReadDouble(oReader, "stokmiktari")
                ' sadece serbeste gelecek (MTFye bağlı olmayan, kapanmış MTF ye bağlı olanlar, kapanmış MTF satırına bağlı olanlar)
                aSRB(nCnt).nSGelecek = SQLReadDouble(oReader, "serbestgelecek")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            nCnt = 0

            cSQL = "select stokno, renk, beden, cinsaciklamasi, anastokgrubu, " + _
                    " stoktipi, birim, malzemetakipkodu, ihtiyac, uretimecikan, " + _
                    " gelecek, stokmiktari, termin,  " + _
                    " isemriverilen = (select sum(coalesce(b.miktar1,0)) " + _
                                    " from isemri a, isemrilines b " + _
                                    " where a.isemrino = b.isemrino " + _
                                    " and b.stokno = " + cTable + ".stokno " + _
                                    " and b.renk = " + cTable + ".renk " + _
                                    " and b.beden = " + cTable + ".beden " + _
                                    " and b.malzemetakipno = " + cTable + ".malzemetakipkodu ) " + _
                    " from " + cTable + _
                    " order by termin, malzemetakipkodu, stokno, renk, beden "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aStokRB(nCnt)
                aStokRB(nCnt).cStokno = SQLReadString(oReader, "stokno")
                aStokRB(nCnt).cRenk = SQLReadString(oReader, "renk")
                aStokRB(nCnt).cBeden = SQLReadString(oReader, "beden")
                aStokRB(nCnt).cAciklama = SQLReadString(oReader, "cinsaciklamasi")
                aStokRB(nCnt).cAnaStokGrubu = SQLReadString(oReader, "anastokgrubu")
                aStokRB(nCnt).cStokTipi = SQLReadString(oReader, "stoktipi")
                aStokRB(nCnt).cBirim = SQLReadString(oReader, "birim")
                aStokRB(nCnt).cMTF = SQLReadString(oReader, "malzemetakipkodu")
                aStokRB(nCnt).nIhtiyac = SQLReadDouble(oReader, "ihtiyac")
                aStokRB(nCnt).nUretimeCikan = SQLReadDouble(oReader, "uretimecikan")
                aStokRB(nCnt).nGelecek = SQLReadDouble(oReader, "gelecek")
                aStokRB(nCnt).nStokMiktari = SQLReadDouble(oReader, "stokmiktari")
                aStokRB(nCnt).dTermin = SQLReadDate(oReader, "termin")
                aStokRB(nCnt).nIsemriVerilen = SQLReadDouble(oReader, "isemriverilen")
                aStokRB(nCnt).cDepartman = ""
                aStokRB(nCnt).lKapali = False
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

             ' önce serbest stoktan rezervasyon simülasyonu yap

            For nCnt = 0 To UBound(aStokRB)

                nIhtiyac = aStokRB(nCnt).nIhtiyac - aStokRB(nCnt).nUretimeCikan - aStokRB(nCnt).nGelecek - aStokRB(nCnt).nStokMiktari

                If nIhtiyac > 0 And Not aStokRB(nCnt).lKapali Then

                    nFound = -1
                    For nCnt1 = 0 To UBound(aSRB)
                        If aStokRB(nCnt).cStokno = aSRB(nCnt1).cStokno And _
                            aStokRB(nCnt).cRenk = aSRB(nCnt1).cRenk And _
                            aStokRB(nCnt).cBeden = aSRB(nCnt1).cBeden Then
                            nFound = nCnt1
                            Exit For
                        End If
                    Next
                    If nFound > -1 Then
                        If aSRB(nFound).nSerbestStokMiktari > 0 Then
                            If nIhtiyac > aSRB(nFound).nSerbestStokMiktari Then
                                nHarcanan = aSRB(nFound).nSerbestStokMiktari
                            Else
                                nHarcanan = nIhtiyac
                            End If
                            aSRB(nFound).nSerbestStokMiktari = aSRB(nFound).nSerbestStokMiktari - nHarcanan
                            aStokRB(nCnt).nStokMiktari = aStokRB(nCnt).nStokMiktari + nHarcanan
                        End If
                    End If
                End If
            Next

            ' serbest satınalma dağıtım simülasyonu yapılıyor

            For nCnt = 0 To UBound(aStokRB)

                nIhtiyac = aStokRB(nCnt).nIhtiyac - aStokRB(nCnt).nUretimeCikan - aStokRB(nCnt).nGelecek - aStokRB(nCnt).nStokMiktari

                If nIhtiyac > 0 And Not aStokRB(nCnt).lKapali Then

                    nFound = -1
                    For nCnt1 = 0 To UBound(aSRB)
                        If aStokRB(nCnt).cStokno = aSRB(nCnt1).cStokno And _
                            aStokRB(nCnt).cRenk = aSRB(nCnt1).cRenk And _
                            aStokRB(nCnt).cBeden = aSRB(nCnt1).cBeden Then
                            nFound = nCnt1
                            Exit For
                        End If
                    Next
                    If nFound > -1 Then
                        If aSRB(nFound).nSerbestGelecek > 0 Then
                            If nIhtiyac > aSRB(nFound).nSerbestGelecek Then
                                nHarcanan = aSRB(nFound).nSerbestGelecek
                            Else
                                nHarcanan = nIhtiyac
                            End If
                            aSRB(nFound).nSerbestGelecek = aSRB(nFound).nSerbestGelecek - nHarcanan
                            aStokRB(nCnt).nGelecek = aStokRB(nCnt).nGelecek + nHarcanan
                        End If
                    End If
                End If
            Next

            ' son kontrol ve kapatmaları yap

            For nCnt = 0 To UBound(aStokRB)

                nIhtiyac = aStokRB(nCnt).nIhtiyac - aStokRB(nCnt).nUretimeCikan - aStokRB(nCnt).nGelecek - aStokRB(nCnt).nStokMiktari

                If nIhtiyac <= 0.01 Then
                    aStokRB(nCnt).lKapali = True
                Else
                    aStokRB(nCnt).lSecildi = False ' True
                End If
            Next

            cSQL = " (stokno char(30) null, " + _
                    " renk char(30) null, " + _
                    " beden char(30) null, " + _
                    " cinsaciklamasi char(250) null, " + _
                    " anastokgrubu char(30) null, " + _
                    " stoktipi char(30) null, " + _
                    " birim char(30) null, " + _
                    " malzemetakipkodu char(30) null, " + _
                    " ihtiyac decimal(18,2) null, " + _
                    " uretimecikan decimal(18,2) null, " + _
                    " gelecek decimal(18,2) null, " + _
                    " stokmiktari decimal(18,2) null, " + _
                    " termin datetime null, " + _
                    " kapandi char(1) null, " + _
                    " secildi char(1) null, " + _
                    " isemriverilen decimal(18,2) null, " + _
                    " departman char(30) null, " + _
                    " imalatci char(30) null, " + _
                    " isemrimtf char(30) null) "

            cMTFHesaplaX = CreateTempTable(ConnYage, cSQL, cMTFHesaplaX)

            For nCnt = 0 To UBound(aStokRB)

                cSQL = "set dateformat dmy " + _
                        " insert into " + cMTFHesaplaX + _
                        " (stokno, renk, beden, cinsaciklamasi, anastokgrubu, " + _
                        " stoktipi, birim, malzemetakipkodu, ihtiyac, uretimecikan, " + _
                        " gelecek, stokmiktari, termin, kapandi, secildi, " + _
                        " isemriverilen, departman) "

                cSQL = cSQL + _
                        " values ('" + SQLWriteString(aStokRB(nCnt).cStokno) + "', " + _
                        " '" + SQLWriteString(aStokRB(nCnt).cRenk) + "', " + _
                        " '" + SQLWriteString(aStokRB(nCnt).cBeden) + "', " + _
                        " '" + SQLWriteString(aStokRB(nCnt).cAciklama) + "', " + _
                        " '" + SQLWriteString(aStokRB(nCnt).cAnaStokGrubu) + "', "

                cSQL = cSQL + _
                        " '" + SQLWriteString(aStokRB(nCnt).cStokTipi) + "', " + _
                        " '" + SQLWriteString(aStokRB(nCnt).cBirim) + "', " + _
                        " '" + SQLWriteString(aStokRB(nCnt).cMTF) + "', " + _
                        SQLWriteDecimal(aStokRB(nCnt).nIhtiyac) + ", " + _
                        SQLWriteDecimal(aStokRB(nCnt).nUretimeCikan) + ", "

                cSQL = cSQL + _
                        SQLWriteDecimal(aStokRB(nCnt).nGelecek) + ", " + _
                        SQLWriteDecimal(aStokRB(nCnt).nStokMiktari) + ", " + _
                        " '" + SQLWriteDate(aStokRB(nCnt).dTermin) + "', " + _
                        " '" + IIf(aStokRB(nCnt).lKapali, "E", "H").ToString + "', " + _
                        " '" + IIf(aStokRB(nCnt).lSecildi, "E", "H").ToString + "', "

                cSQL = cSQL + _
                        SQLWriteDecimal(aStokRB(nCnt).nIsemriVerilen) + ", " + _
                        " '" + SQLWriteString(aStokRB(nCnt).cDepartman) + "') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next

            ' üretim departmanına çıkış için departman konuldu
            ' birden fazla departmana çıkış için bir formül bulunmalı
            cSQL = "update " + cMTFHesaplaX + _
                    " set departman = (select top 1 departman " + _
                                        " from mtkfislines " + _
                                        " where malzemetakipno = " + cMTFHesaplaX + ".malzemetakipkodu " + _
                                        " and stokno = " + cMTFHesaplaX + ". stokno " + _
                                        " and renk = " + cMTFHesaplaX + ". renk " + _
                                        " and beden = " + cMTFHesaplaX + ". beden) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropView(cMTFView, ConnYage)
            DropTable(cTable, ConnYage)

            If lDropTSip Then DropTable(cTSip, ConnYage)


            MTFHesaplax_1 = 1

            JustForLog("MTFHesaplax_1 end : " + cMTFHesaplaX)

        Catch ex As Exception
            ErrDisp(ex.Message, "MTFHesaplax_1", cSQL)
        End Try
    End Function

    Public Function DMTFIsemriKilavuzuHesapla(cDetayIhtiyacTable As String) As Integer

        Dim cSQL As String = ""

        DMTFIsemriKilavuzuHesapla = 0

        Try
            If cDetayIhtiyacTable.Trim = "" Then Exit Function
            ' işemri verirken daha önceden aynı mtf ye işemri verildiyse statik mtf de işemri verilen rakkamları yüksek çıkmasın diye
            ' başka mtf numaralarına işemri geçilir
            ' önce işemri verilmemiş MTF ler seçilir
            cSQL = "update " + cDetayIhtiyacTable + _
                    " set isemrimtf = malzemetakipkodu " + _
                    " where secildi = 'E' " + _
                    " and coalesce(isemriverilen,0) = 0 " + _
                    " and (isemrimtf is null or isemrimtf = '') "

            ExecuteSQLCommand(cSQL)

            ' sonra kendi ihtiyacı, verilecek işemrinden fazla olan MTF ler seçilir
            ' yani ihtiyaç >= VerilmişİşEmri + VerilecekİşEmri
            cSQL = "update " + cDetayIhtiyacTable + _
                    " set isemrimtf = malzemetakipkodu " + _
                    " where secildi = 'E' " + _
                    " and coalesce(ihtiyac,0) >= coalesce(isemriverilen,0) + (coalesce(ihtiyac,0) - coalesce(uretimecikan,0) - coalesce(gelecek,0) - coalesce(stokmiktari,0)) " + _
                    " and (isemrimtf is null or isemrimtf = '') "

            ExecuteSQLCommand(cSQL)

            ' bu turda iş emri verilmemişler arasından işemri verilecek miktardan fazla MUTLAK ihtiyacı olan MTF ler seçilir
            cSQL = "update " + cDetayIhtiyacTable + _
                    " set isemrimtf = (select top 1 x.malzemetakipkodu " + _
                                        " From " + cDetayIhtiyacTable + " x " + _
                                        " Where x.malzemetakipkodu <> " + cDetayIhtiyacTable + ".malzemetakipkodu " + _
                                        " and x.stokno = " + cDetayIhtiyacTable + ".stokno " + _
                                        " and x.renk = " + cDetayIhtiyacTable + ".renk " + _
                                        " and x.beden = " + cDetayIhtiyacTable + ".beden " + _
                                        " and x.isemriverilen = 0 " + _
                                        " and (coalesce(x.ihtiyac,0) - coalesce(x.uretimecikan,0)) >= (coalesce(" + cDetayIhtiyacTable + ".ihtiyac,0) - coalesce(" + cDetayIhtiyacTable + ".uretimecikan,0) - coalesce(" + cDetayIhtiyacTable + ".gelecek,0) - coalesce(" + cDetayIhtiyacTable + ".stokmiktari,0)) " + _
                                        " order by x.termin) " + _
                    " where secildi = 'E' " + _
                    " and coalesce(isemriverilen,0) > 0 " + _
                    " and (isemrimtf is null or isemrimtf = '') "

            ExecuteSQLCommand(cSQL)

            ' bu turda iş emri verilmemiş MTF ler seçilir
            cSQL = "update " + cDetayIhtiyacTable + _
                    " set isemrimtf = (select top 1 x.malzemetakipkodu " + _
                                        " From " + cDetayIhtiyacTable + " x " + _
                                        " Where x.malzemetakipkodu <> " + cDetayIhtiyacTable + ".malzemetakipkodu " + _
                                        " and x.stokno = " + cDetayIhtiyacTable + ".stokno " + _
                                        " and x.renk = " + cDetayIhtiyacTable + ".renk " + _
                                        " and x.beden = " + cDetayIhtiyacTable + ".beden " + _
                                        " and x.isemriverilen = 0 " + _
                                        " order by x.termin ) " + _
                    " where secildi = 'E' " + _
                    " and coalesce(isemriverilen,0) > 0 " + _
                    " and (isemrimtf is null or isemrimtf = '') "

            ExecuteSQLCommand(cSQL)

            ' Kalan MTF ler kendine eşitlenir
            cSQL = "update " + cDetayIhtiyacTable + _
                    " set isemrimtf = malzemetakipkodu " + _
                    " where secildi = 'E' " + _
                    " and (isemrimtf is null or isemrimtf = '') "

            ExecuteSQLCommand(cSQL)

            ' imalatçı firmaları (imalat ülkelerini) yaz

            cSQL = "update " + cDetayIhtiyacTable + _
                    " set imalatci = (select top 1 a.imalatci " + _
                                    " from siparis a, sipmodel b " + _
                                    " where a.kullanicisipno = b.siparisno " + _
                                    " and b.malzemetakipno = " + cDetayIhtiyacTable + ".isemrimtf) " + _
                    " where secildi = 'E' "

            ExecuteSQLCommand(cSQL)

            DMTFIsemriKilavuzuHesapla = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "DMTFIsemriKilavuzuHesapla", cSQL)
        End Try

    End Function
End Module

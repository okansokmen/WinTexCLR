Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module utilMasterPlan

    Private Structure oCP
        Dim dPlGonderiTarihi As Date
        Dim dPlTarihi As Date
        Dim cSiparisNo As String
        Dim cModelKodu As String
        Dim cOkTipi As String
        Dim cRenk As String
        Dim cBeden As String
        Dim dOkTar As Date
        Dim dOkTar2 As Date
    End Structure

    Private Structure oSTF
        Dim dIlkSevkTar As Date
        Dim cSevkiyatTakipNo As String
        Dim cSevkEmriNo As String
        Dim nToplam As Double
        Dim nPlanlanan As Double
        Dim nGiden As Double
        Dim dGerSevkTar As Date
    End Structure

    Private Structure oUTF
        Dim cUretimTakipNo As String
        Dim dBaslamaTarihi As Date
        Dim dBitisTarihi As Date
        Dim cDepartman As String
        Dim cModelNo As String
        Dim cPlFirma As String
        Dim nToplamAdet As Double
        Dim nIsEmriVerilen As Double
        Dim nGelen As Double
        Dim nGiden As Double
    End Structure

    Private Structure oMTF
        Dim cMalzemeTakipNo As String
        Dim dBaslamaTarihi As Date
        Dim dBitisTarihi As Date
        Dim cStokTipi As String
        Dim cPlFirma As String
        Dim cStokNo As String
        Dim cDepartman As String
        Dim cBirim1 As String
        Dim nIhtiyac As Double
        Dim nIsemriVerilen As Double
        Dim nKarsilanan As Double
    End Structure

    Public Function GetMasterPlanData(cFilter As String) As Integer

        Dim cTable As String = "masterplan"
        Dim cSQL As String = ""
        Dim nRow As Long = 0
        Dim cAciklama As String = ""
        Dim cDurum As String = ""
        Dim cPersonel As String = ""
        Dim cFirma As String = ""
        Dim cSiparisNo As String = ""
        Dim cMusteriNo As String = ""
        Dim cModelNo As String = ""
        Dim cDepartman As String = ""
        Dim cInsertHeader As String = ""
        Dim nDecimal As Integer = 0
        Dim nMiktar As Double = 0
        Dim cColor As String = ""
        Dim cKalipNo As String = ""
        Dim nIhtiyac As Double = 0
        Dim nIsemri As Double = 0
        Dim nBaslanan As Double = 0
        Dim nBiten As Double = 0
        Dim cTipi As String = ""
        Dim cFoyNo As String = ""
        Dim dSevkTarihi As Date = #1/1/1950#
        Dim nPairKey As Double = 0
        Dim cRenk As String = ""
        Dim cBeden As String = ""
        Dim cIhtiyacTable As String = ""
        Dim cDetayIhtiyacTable As String = ""
        Dim lDinamikMTF As Boolean = False
        Dim dGerceklesen As Date = #1/1/1950#
        Dim dGBasla As Date = #1/1/1950#
        Dim dGBitir As Date = #1/1/1950#
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim aSTF() As oSTF = Nothing
        Dim aUTF() As oUTF = Nothing
        Dim aMTF() As oMTF = Nothing
        Dim aCP() As oCP = Nothing
        Dim nSonuc As Integer = 0

        GetMasterPlanData = 0

        Try

            ConnYage = OpenConn()

            JustForLog("Masterplan build start")

            lDinamikMTF = (CDbl(GetSysParConnected("masterplandinamikmtf", ConnYage)) = 1)

            cSQL = "delete " + cTable
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cInsertHeader = "set dateformat dmy " + _
                            " insert into " + cTable + _
                            " (tarih, kategori, aciklama, baslabitir, durum, " + _
                            " personel, firma, siparisno, musterino, modelno, " + _
                            " departman, color, kalipno, ihtiyac, isemri, " + _
                            " baslanan, biten, tipi, foyno, sevktarihi, " + _
                            " pairkey, renk, beden, gercektarih ) "

            ' STF
            nCnt = 0

            cSQL = "select distinct ilksevktar, sevkiyattakipno, sevkemrino, toplam, planlanan, giden, " + _
                    " gersevktar = (select top 1 a.sevktar " + _
                                    " from sevkform a, sevkformlines b " + _
                                    " where a.sevkformno = b.sevkformno " + _
                                    " and b.sevkiyattakipno = sevkplfislines.sevkiyattakipno " + _
                                    " order by a.sevktar desc ) " + _
                    " from sevkplfislines "

            cSQL = cSQL + _
                    " where (ok is null or ok = 'H' or ok = '') " + _
                    " and ilksevktar is not null " + _
                    " and ilksevktar > '01.01.1950' "

            If cFilter.Trim = "" Then
                cSQL = cSQL + _
                        " and sevkiyattakipno in (select y.sevkiyattakipno " + _
                                                " from siparis x, sipmodel y " + _
                                                " where x.kullanicisipno = y.siparisno " + _
                                                " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '')  " + _
                                                " and x.planlamaok = 'E' " + _
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            Else
                cSQL = cSQL + _
                        " and sevkiyattakipno in (select y.sevkiyattakipno " + _
                                                " from siparis x, sipmodel y " + _
                                                " where x.kullanicisipno = y.siparisno " + _
                                                " and x.kullanicisipno in (" + cFilter + ") " + _
                                                " and x.planlamaok = 'E' " + _
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            End If

            cSQL = cSQL + _
                    " order by sevkiyattakipno "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aSTF(nCnt)
                aSTF(nCnt).dIlkSevkTar = SQLReadDate(oReader, "ilksevktar")
                aSTF(nCnt).cSevkiyatTakipNo = SQLReadString(oReader, "sevkiyattakipno")
                aSTF(nCnt).cSevkEmriNo = SQLReadString(oReader, "sevkemrino")
                aSTF(nCnt).nToplam = SQLReadDouble(oReader, "toplam")
                aSTF(nCnt).nPlanlanan = SQLReadDouble(oReader, "planlanan")
                aSTF(nCnt).nGiden = SQLReadDouble(oReader, "giden")
                aSTF(nCnt).dGerSevkTar = SQLReadDate(oReader, "gersevktar")
                nCnt = nCnt + 1
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aSTF)
                cAciklama = aSTF(nCnt).cSevkEmriNo
                cDepartman = "SEVKIYAT"
                cTipi = "SEVKIYAT"
                cFoyNo = aSTF(nCnt).cSevkiyatTakipNo
                nIhtiyac = aSTF(nCnt).nToplam
                nIsemri = aSTF(nCnt).nPlanlanan
                nBaslanan = 0
                nBiten = aSTF(nCnt).nGiden
                dGerceklesen = aSTF(nCnt).dGerSevkTar
                dSevkTarihi = aSTF(nCnt).dIlkSevkTar

                cDurum = ""
                If aSTF(nCnt).nPlanlanan > 0 Then
                    cDurum = cDurum + "Planlanan:" + Microsoft.VisualBasic.Format(aSTF(nCnt).nPlanlanan, G_NumberFormat) + " "
                End If
                If aSTF(nCnt).nGiden > 0 Then
                    cDurum = cDurum + "Sevkedilen:" + Microsoft.VisualBasic.Format(aSTF(nCnt).nGiden, G_NumberFormat) + " "
                End If

                If aSTF(nCnt).nGiden >= aSTF(nCnt).nPlanlanan Then
                    cColor = "YESIL"
                ElseIf aSTF(nCnt).nGiden > 0 Then
                    cColor = "SARI"
                End If

                cSQL = "select distinct siparisno " + _
                        " from sipmodel " + _
                        " where sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " + _
                        " and siparisno is not null " + _
                        " and siparisno <> '' "

                cSiparisNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct modelno " + _
                        " from sipmodel " + _
                        " where sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " + _
                        " and modelno is not null " + _
                        " and modelno <> '' "

                cModelNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct a.kalipno " + _
                        " from ymodel a, sipmodel b " + _
                        " where a.modelno = b.modelno " + _
                        " and b.sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " + _
                        " and a.modelno is not null " + _
                        " and a.modelno <> '' "

                cKalipNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct a.musterino " + _
                        " from siparis a, sipmodel b " + _
                        " where a.kullanicisipno = b.siparisno " + _
                        " and b.sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " + _
                        " and a.musterino is not null " + _
                        " and a.musterino <> '' "

                cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct a.sorumlu " + _
                        " from siparis a, sipmodel b " + _
                        " where a.kullanicisipno = b.siparisno " + _
                        " and b.sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " + _
                        " and a.sorumlu is not null " + _
                        " and a.sorumlu <> '' "

                cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                cFirma = cMusteriNo

                nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                cSQL = cInsertHeader + _
                        " values ('" + SQLWriteDate(dSevkTarihi) + "', " + _
                        " 'STF', " + _
                        " '" + SQLWriteString(cAciklama, 250) + "'," + _
                        " 'SEVKIYAT', " + _
                        " '" + cDurum + "', " + _
                        " '" + SQLWriteString(cPersonel, 30) + "', " + _
                        " '" + SQLWriteString(cFirma, 250) + "', " + _
                        " '" + SQLWriteString(cSiparisNo, 250) + "', " + _
                        " '" + SQLWriteString(cMusteriNo, 250) + "', " + _
                        " '" + SQLWriteString(cModelNo, 250) + "', " + _
                        " '" + SQLWriteString(cDepartman, 250) + "', " + _
                        " '" + SQLWriteString(cColor, 30) + "', " + _
                        " '" + SQLWriteString(cKalipNo, 250) + "', " + _
                        SQLWriteDecimal(nIhtiyac) + ", " + _
                        SQLWriteDecimal(nIsemri) + ", " + _
                        SQLWriteDecimal(nBaslanan) + ", " + _
                        SQLWriteDecimal(nBiten) + ", " + _
                        " '" + SQLWriteString(cTipi, 30) + "', " + _
                        " '" + SQLWriteString(cFoyNo, 30) + "', " + _
                        " '" + SQLWriteDate(dSevkTarihi) + "', " + _
                        SQLWriteDecimal(nPairKey) + ",'','', " + _
                        " '" + SQLWriteDate(dGerceklesen) + "') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next

            ' UTF
            nCnt = 0

            cSQL = "select uretimtakipno, baslamatarihi, bitistarihi, departman, modelno, plfirma, " + _
                    " toplamadet = sum(coalesce(toplamadet,0)), " + _
                    " isemriverilen = sum(coalesce(isemriverilen,0)), " + _
                    " gelen = sum(coalesce(gelen,0)), " + _
                    " giden = sum(coalesce(giden,0)) " + _
                    " from uretpllines "

            cSQL = cSQL + _
                    " where (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and ((baslamatarihi is not null and baslamatarihi > '01.01.1950') or (bitistarihi is not null and bitistarihi > '01.01.1950')) "

            If cFilter.Trim = "" Then
                cSQL = cSQL + _
                        " and uretimtakipno in (select y.uretimtakipno " + _
                                                " from siparis x, sipmodel y " + _
                                                " where x.kullanicisipno = y.siparisno " + _
                                                " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '')  " + _
                                                " and x.planlamaok = 'E' " + _
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            Else
                cSQL = cSQL + _
                        " and uretimtakipno in (select y.uretimtakipno " + _
                                                " from siparis x, sipmodel y " + _
                                                " where x.kullanicisipno = y.siparisno " + _
                                                " and x.kullanicisipno in (" + cFilter + ") " + _
                                                " and x.planlamaok = 'E' " + _
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            End If


            cSQL = cSQL + _
                    " group by uretimtakipno, baslamatarihi, bitistarihi, departman, modelno, plfirma " + _
                    " order by uretimtakipno "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aUTF(nCnt)
                aUTF(nCnt).cUretimTakipNo = SQLReadString(oReader, "uretimtakipno")
                aUTF(nCnt).dBaslamaTarihi = SQLReadDate(oReader, "baslamatarihi")
                aUTF(nCnt).dBitisTarihi = SQLReadDate(oReader, "bitistarihi")
                aUTF(nCnt).cDepartman = SQLReadString(oReader, "departman")
                aUTF(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aUTF(nCnt).cPlFirma = SQLReadString(oReader, "plfirma")
                aUTF(nCnt).nToplamAdet = SQLReadDouble(oReader, "toplamadet")
                aUTF(nCnt).nIsEmriVerilen = SQLReadDouble(oReader, "isemriverilen")
                aUTF(nCnt).nGelen = SQLReadDouble(oReader, "gelen")
                aUTF(nCnt).nGiden = SQLReadDouble(oReader, "giden")

                nCnt = nCnt + 1
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aUTF)
                cAciklama = aUTF(nCnt).cModelNo
                cTipi = aUTF(nCnt).cDepartman
                cFoyNo = aUTF(nCnt).cUretimTakipNo
                cModelNo = aUTF(nCnt).cModelNo
                cDepartman = aUTF(nCnt).cDepartman

                nIhtiyac = aUTF(nCnt).nToplamAdet
                nIsemri = aUTF(nCnt).nIsEmriVerilen
                nBaslanan = aUTF(nCnt).nGelen
                nBiten = aUTF(nCnt).nGiden

                GetUTFGercekTarih(ConnYage, aUTF(nCnt).cUretimTakipNo, aUTF(nCnt).cDepartman, aUTF(nCnt).cModelNo, dGBasla, dGBitir)

                cSQL = "select kalipno " + _
                        " from ymodel " + _
                        " where modelno = '" + aUTF(nCnt).cModelNo + "' "

                cKalipNo = SQLGetStringConnected(cSQL, ConnYage)

                cDurum = ""
                If aUTF(nCnt).nIsEmriVerilen > 0 Then
                    cDurum = cDurum + "İşemri:" + Microsoft.VisualBasic.Format(aUTF(nCnt).nIsEmriVerilen, G_NumberFormat) + " "
                End If
                If aUTF(nCnt).nGiden > 0 Then
                    cDurum = cDurum + "Üretilen:" + Microsoft.VisualBasic.Format(aUTF(nCnt).nGiden, G_NumberFormat) + " "
                End If

                If aUTF(nCnt).nGiden >= aUTF(nCnt).nIsEmriVerilen Then
                    cColor = "YESIL"
                ElseIf aUTF(nCnt).nGiden > 0 Then
                    cColor = "SARI"
                ElseIf aUTF(nCnt).nIsEmriVerilen = 0 Then
                    cColor = "KIRMIZI"
                End If

                cSQL = "select distinct siparisno " + _
                        " from sipmodel " + _
                        " where uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " + _
                        " and modelno = '" + aUTF(nCnt).cModelNo + "' " + _
                        " and siparisno is not null " + _
                        " and siparisno <> '' "

                cSiparisNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct a.musterino " + _
                        " from siparis a, sipmodel b " + _
                        " where a.kullanicisipno = b.siparisno " + _
                        " and b.uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " + _
                        " and b.modelno = '" + aUTF(nCnt).cModelNo + "' " + _
                        " and a.musterino is not null " + _
                        " and a.musterino <> '' "

                cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                ' üretim işemirlerindeki personel
                cSQL = "select distinct a.eleman " + _
                        " from uretimisemri a, uretimisdetayi b " + _
                        " where a.isemrino = b.isemrino " + _
                        " and a.uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " + _
                        " and a.departman = '" + aUTF(nCnt).cDepartman + "' " + _
                        " and b.modelno = '" + aUTF(nCnt).cModelNo + "' " + _
                        " and a.eleman is not null " + _
                        " and a.eleman <> '' "

                cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                ' üretim işemirlerindeki firmalar
                cSQL = "select distinct a.firma " + _
                        " from uretimisemri a, uretimisdetayi b " + _
                        " where a.isemrino = b.isemrino " + _
                        " and a.uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " + _
                        " and a.departman = '" + aUTF(nCnt).cDepartman + "' " + _
                        " and b.modelno = '" + aUTF(nCnt).cModelNo + "' " + _
                        " and a.firma is not null " + _
                        " and a.firma <> '' "

                cFirma = SQLBuildFilterString2(ConnYage, cSQL, False)

                If Trim(cFirma) = "" Then
                    cFirma = aUTF(nCnt).cPlFirma
                End If

                cSQL = "select min(ilksevktar) " + _
                        " from sevkplfislines " + _
                        " where ilksevktar is not null " + _
                        " and ilksevktar <> '01.01.1950' " + _
                        " and exists (select sevkiyattakipno " + _
                                    " from sipmodel " + _
                                    " where sevkiyattakipno = sevkplfislines.sevkiyattakipno " + _
                                    " and uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " + _
                                    " and modelno = '" + aUTF(nCnt).cModelNo + "') "

                dSevkTarihi = SQLGetDateConnected(cSQL, ConnYage)

                nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                If aUTF(nCnt).dBitisTarihi > #1/1/1950# Then

                    cSQL = cInsertHeader + _
                            " values ('" + SQLWriteDate(aUTF(nCnt).dBitisTarihi) + "', " + _
                            " 'UTF', " + _
                            " '" + SQLWriteString(cAciklama, 250) + "'," + _
                            " 'BITIR', " + _
                            " '" + cDurum + "', " + _
                            " '" + SQLWriteString(cPersonel, 30) + "', " + _
                            " '" + SQLWriteString(cFirma, 250) + "', " + _
                            " '" + SQLWriteString(cSiparisNo, 250) + "', " + _
                            " '" + SQLWriteString(cMusteriNo, 250) + "', " + _
                            " '" + SQLWriteString(cModelNo, 250) + "', " + _
                            " '" + SQLWriteString(cDepartman, 250) + "', " + _
                            " '" + SQLWriteString(cColor, 30) + "', " + _
                            " '" + SQLWriteString(cKalipNo, 250) + "', " + _
                            SQLWriteDecimal(nIhtiyac) + ", " + _
                            SQLWriteDecimal(nIsemri) + ", " + _
                            SQLWriteDecimal(nBaslanan) + ", " + _
                            SQLWriteDecimal(nBiten) + ", " + _
                            " '" + SQLWriteString(cTipi, 30) + "', " + _
                            " '" + SQLWriteString(cFoyNo, 30) + "', " + _
                            " '" + SQLWriteDate(dSevkTarihi) + "', " + _
                            SQLWriteDecimal(nPairKey) + ",'','', " + _
                            " '" + SQLWriteDate(dGBitir) + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                If aUTF(nCnt).dBaslamaTarihi > #1/1/1950# Then

                    cSQL = cInsertHeader + _
                            " values ('" + SQLWriteDate(aUTF(nCnt).dBaslamaTarihi) + "', " + _
                            " 'UTF', " + _
                            " '" + SQLWriteString(cAciklama, 250) + "'," + _
                            " 'BASLA', " + _
                            " '" + cDurum + "', " + _
                            " '" + SQLWriteString(cPersonel, 30) + "', " + _
                            " '" + SQLWriteString(cFirma, 250) + "', " + _
                            " '" + SQLWriteString(cSiparisNo, 250) + "', " + _
                            " '" + SQLWriteString(cMusteriNo, 250) + "', " + _
                            " '" + SQLWriteString(cModelNo, 250) + "', " + _
                            " '" + SQLWriteString(cDepartman, 250) + "', " + _
                            " '" + SQLWriteString(cColor, 30) + "', " + _
                            " '" + SQLWriteString(cKalipNo, 250) + "', " + _
                            SQLWriteDecimal(nIhtiyac) + ", " + _
                            SQLWriteDecimal(nIsemri) + ", " + _
                            SQLWriteDecimal(nBaslanan) + ", " + _
                            SQLWriteDecimal(nBiten) + ", " + _
                            " '" + SQLWriteString(cTipi, 30) + "', " + _
                            " '" + SQLWriteString(cFoyNo, 30) + "', " + _
                            " '" + SQLWriteDate(dSevkTarihi) + "', " + _
                            SQLWriteDecimal(nPairKey) + ",'','', " + _
                            " '" + SQLWriteDate(dGBasla) + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            ' MTF

            If lDinamikMTF Then
                nSonuc = GetToplamSiparisView_1("", cIhtiyacTable, ConnYage)
                nSonuc = MTFHesaplax_1("", "", cIhtiyacTable, cDetayIhtiyacTable, ConnYage)
            End If

            nCnt = 0

            cSQL = "select a.malzemetakipno, a.baslamatarihi, a.bitistarihi, b.stoktipi, a.plfirma, a.stokno, a.departman, b.birim1, " + _
                    " ihtiyac = sum(coalesce(ihtiyac,0)), " + _
                    " isemriverilen = sum(coalesce(isemriverilen,0)), " + _
                    " karsilanan = sum(coalesce(isemriicingelen,0)) + sum(coalesce(isemriharicigelen,0)) " + _
                    " from mtkfislines a, stok b "

            cSQL = cSQL + _
                    " where a.stokno = b.stokno " + _
                    " and (a.kapandi is null or a.kapandi = 'H' or a.kapandi = '') " + _
                    " and ((a.baslamatarihi is not null and a.baslamatarihi > '01.01.1950') or (a.bitistarihi is not null and a.bitistarihi > '01.01.1950')) "

            If cFilter.Trim = "" Then
                cSQL = cSQL + _
                        " and malzemetakipno in (select y.malzemetakipno " + _
                                                " from siparis x, sipmodel y " + _
                                                " where x.kullanicisipno = y.siparisno " + _
                                                " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '') " + _
                                                " and x.planlamaok = 'E' " + _
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            Else
                cSQL = cSQL + _
                        " and malzemetakipno in (select y.malzemetakipno " + _
                                                " from siparis x, sipmodel y " + _
                                                " where x.kullanicisipno = y.siparisno " + _
                                                " and x.kullanicisipno in (" + cFilter + ") " + _
                                                " and x.planlamaok = 'E' " + _
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            End If

            cSQL = cSQL + _
                    " group by a.malzemetakipno, a.baslamatarihi, a.bitistarihi, b.stoktipi, a.plfirma, a.stokno, a.departman, b.birim1 " + _
                    " order by a.malzemetakipno "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aMTF(nCnt)
                aMTF(nCnt).cMalzemeTakipNo = SQLReadString(oReader, "malzemetakipno")
                aMTF(nCnt).dBaslamaTarihi = SQLReadDate(oReader, "baslamatarihi")
                aMTF(nCnt).dBitisTarihi = SQLReadDate(oReader, "bitistarihi")
                aMTF(nCnt).cStokTipi = SQLReadString(oReader, "stoktipi")
                aMTF(nCnt).cPlFirma = SQLReadString(oReader, "plfirma")
                aMTF(nCnt).cStokNo = SQLReadString(oReader, "stokno")
                aMTF(nCnt).cDepartman = SQLReadString(oReader, "departman")
                aMTF(nCnt).cBirim1 = SQLReadString(oReader, "birim1")
                aMTF(nCnt).nIhtiyac = SQLReadDouble(oReader, "ihtiyac")
                aMTF(nCnt).nIsemriVerilen = SQLReadDouble(oReader, "isemriverilen")
                aMTF(nCnt).nKarsilanan = SQLReadDouble(oReader, "karsilanan")

                nCnt = nCnt + 1
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aMTF)

                cAciklama = aMTF(nCnt).cStokNo
                cTipi = aMTF(nCnt).cStokTipi
                cFoyNo = aMTF(nCnt).cMalzemeTakipNo

                GetMTFGercekTarih(ConnYage, aMTF(nCnt).cMalzemeTakipNo, aMTF(nCnt).cStokNo, , , dGBasla, dGBitir)

                If lDinamikMTF Then

                    cSQL = "SELECT ihtiyac = sum(coalesce(ihtiyac,0)), " + _
                            " uretimecikan = sum(coalesce(uretimecikan,0)), " + _
                            " gelecek = sum(coalesce(gelecek,0)), " + _
                            " stokmiktari = sum(coalesce(stokmiktari,0)) " + _
                            " from " + cDetayIhtiyacTable + _
                            " where malzemetakipkodu = '" + aMTF(nCnt).cMalzemeTakipNo + "' " + _
                            " and stokno = '" + aMTF(nCnt).cStokNo + "' " + _
                            " and departman = '" + aMTF(nCnt).cDepartman + "' "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        nIhtiyac = SQLReadDouble(oReader, "ihtiyac")
                        nIsemri = SQLReadDouble(oReader, "gelecek")
                        nBaslanan = SQLReadDouble(oReader, "uretimecikan")
                        nBiten = SQLReadDouble(oReader, "uretimecikan") + SQLReadDouble(oReader, "gelecek") + SQLReadDouble(oReader, "stokmiktari")
                    End If
                    oReader.Close()
                Else
                    nIhtiyac = aMTF(nCnt).nIhtiyac
                    nIsemri = aMTF(nCnt).nIsemriVerilen
                    nBaslanan = 0
                    nBiten = aMTF(nCnt).nKarsilanan
                End If

                cDurum = ""
                If nIhtiyac > 0 Then
                    nMiktar = Yuvarlat(ConnYage, aMTF(nCnt).cBirim1, nIhtiyac, nDecimal)
                    cDurum = cDurum + "İhtiyaç:" + IIf(nDecimal = 0, Microsoft.VisualBasic.Format(nMiktar, G_NumberFormat), Microsoft.VisualBasic.Format(nMiktar, G_Number2Format)).ToString + " "
                End If
                If nIsemri > 0 Then
                    nMiktar = Yuvarlat(ConnYage, aMTF(nCnt).cBirim1, nIsemri, nDecimal)
                    cDurum = cDurum + "İşemri:" + IIf(nDecimal = 0, Microsoft.VisualBasic.Format(nMiktar, G_NumberFormat), Microsoft.VisualBasic.Format(nMiktar, G_Number2Format)).ToString + " "
                End If
                If nBiten > 0 Then
                    nMiktar = Yuvarlat(ConnYage, aMTF(nCnt).cBirim1, nBiten, nDecimal)
                    cDurum = cDurum + "Karşılanan:" + IIf(nDecimal = 0, Microsoft.VisualBasic.Format(nMiktar, G_NumberFormat), Microsoft.VisualBasic.Format(nMiktar, G_Number2Format)).ToString + " "
                End If
                If cDurum <> "" Then
                    cDurum = cDurum + aMTF(nCnt).cBirim1
                End If

                If aMTF(nCnt).nKarsilanan >= aMTF(nCnt).nIhtiyac Then
                    cColor = "YESIL"
                ElseIf aMTF(nCnt).nKarsilanan > 0 Then
                    cColor = "SARI"
                ElseIf aMTF(nCnt).nIsemriVerilen = 0 Then
                    cColor = "KIRMIZI"
                End If

                cDepartman = aMTF(nCnt).cDepartman

                cSQL = "select distinct siparisno " + _
                         " from sipmodel " + _
                         " where malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " + _
                         " and siparisno is not null " + _
                         " and siparisno <> '' "

                cSiparisNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct modelno " + _
                        " from sipmodel " + _
                        " where malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " + _
                        " and modelno is not null " + _
                        " and modelno <> '' "

                cModelNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct a.kalipno " + _
                        " from ymodel a, sipmodel b " + _
                        " where a.modelno = b.modelno " + _
                        " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " + _
                        " and a.modelno is not null " + _
                        " and a.modelno <> '' "

                cKalipNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct a.musterino " + _
                        " from siparis a, sipmodel b " + _
                        " where a.kullanicisipno = b.siparisno " + _
                        " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " + _
                        " and a.musterino is not null " + _
                        " and a.musterino <> '' "

                cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                ' firmaları işemirlerinden al
                cSQL = "select distinct a.firma " + _
                        " from isemri a, isemrilines b " + _
                        " where a.isemrino = b.isemrino " + _
                        " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " + _
                        " and b.stokno = '" + aMTF(nCnt).cStokNo + "' " + _
                        " and a.firma is not null " + _
                        " and a.firma <> '' "

                cFirma = SQLBuildFilterString2(ConnYage, cSQL, False)

                If Trim(cFirma) = "" Then
                    cFirma = aMTF(nCnt).cPlFirma
                End If

                ' personeli işemirlerinden al
                cSQL = "select distinct a.takipelemani " + _
                        " from isemri a, isemrilines b " + _
                        " where a.isemrino = b.isemrino " + _
                        " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " + _
                        " and b.stokno = '" + aMTF(nCnt).cStokNo + "' " + _
                        " and a.takipelemani is not null " + _
                        " and a.takipelemani <> '' "

                cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select min(ilksevktar) " + _
                        " from sevkplfislines " + _
                        " where ilksevktar is not null " + _
                        " and ilksevktar <> '01.01.1950' " + _
                        " and exists (select sevkiyattakipno " + _
                                    " from sipmodel " + _
                                    " where sevkiyattakipno = sevkplfislines.sevkiyattakipno " + _
                                    " and malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "') "

                dSevkTarihi = SQLGetDateConnected(cSQL, ConnYage)

                nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                If aMTF(nCnt).dBitisTarihi > #1/1/1950# Then

                    cSQL = cInsertHeader + _
                            " values ('" + SQLWriteDate(aMTF(nCnt).dBitisTarihi) + "', " + _
                            " 'MTF', " + _
                            " '" + SQLWriteString(cAciklama, 250) + "'," + _
                            " 'DEPOYA GIRIS', " + _
                            " '" + cDurum + "', " + _
                            " '" + SQLWriteString(cPersonel, 30) + "', " + _
                            " '" + SQLWriteString(cFirma, 250) + "', " + _
                            " '" + SQLWriteString(cSiparisNo, 250) + "', " + _
                            " '" + SQLWriteString(cMusteriNo, 250) + "', " + _
                            " '" + SQLWriteString(cModelNo, 250) + "', " + _
                            " '" + SQLWriteString(cDepartman, 250) + "', " + _
                            " '" + SQLWriteString(cColor, 30) + "', " + _
                            " '" + SQLWriteString(cKalipNo, 250) + "', " + _
                            SQLWriteDecimal(nIhtiyac) + ", " + _
                            SQLWriteDecimal(nIsemri) + ", " + _
                            SQLWriteDecimal(nBaslanan) + ", " + _
                            SQLWriteDecimal(nBiten) + ", " + _
                            " '" + SQLWriteString(cTipi, 30) + "', " + _
                            " '" + SQLWriteString(cFoyNo, 30) + "', " + _
                            " '" + SQLWriteDate(dSevkTarihi) + "', " + _
                            SQLWriteDecimal(nPairKey) + ",'','', " + _
                            " '" + SQLWriteDate(dGBitir) + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If


                If aMTF(nCnt).dBaslamaTarihi > #1/1/1950# Then

                    cSQL = cInsertHeader + _
                            " values ('" + SQLWriteDate(aMTF(nCnt).dBaslamaTarihi) + "', " + _
                            " 'MTF', " + _
                            " '" + SQLWriteString(cAciklama, 250) + "'," + _
                            " 'SATINALMA ISEMRI', " + _
                            " '" + cDurum + "', " + _
                            " '" + SQLWriteString(cPersonel, 30) + "', " + _
                            " '" + SQLWriteString(cFirma, 250) + "', " + _
                            " '" + SQLWriteString(cSiparisNo, 250) + "', " + _
                            " '" + SQLWriteString(cMusteriNo, 250) + "', " + _
                            " '" + SQLWriteString(cModelNo, 250) + "', " + _
                            " '" + SQLWriteString(cDepartman, 250) + "', " + _
                            " '" + SQLWriteString(cColor, 30) + "', " + _
                            " '" + SQLWriteString(cKalipNo, 250) + "', " + _
                            SQLWriteDecimal(nIhtiyac) + ", " + _
                            SQLWriteDecimal(nIsemri) + ", " + _
                            SQLWriteDecimal(nBaslanan) + ", " + _
                            SQLWriteDecimal(nBiten) + ", " + _
                            " '" + SQLWriteString(cTipi, 30) + "', " + _
                            " '" + SQLWriteString(cFoyNo, 30) + "', " + _
                            " '" + SQLWriteDate(dSevkTarihi) + "', " + _
                            SQLWriteDecimal(nPairKey) + ",'','', " + _
                            " '" + SQLWriteDate(dGBasla) + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            If lDinamikMTF Then
                DropTable(cIhtiyacTable, ConnYage)
                DropTable(cDetayIhtiyacTable, ConnYage)
            End If

            ' CP
            nCnt = 0

            cSQL = "select distinct plgonderitarihi, pltarihi, siparisno, modelkodu, oktipi, renk, beden, oktar, oktar2  " + _
                    " from sipok "

            cSQL = cSQL + _
                    " where (ok is null or ok = 'H' or ok = '') " + _
                    " and ((PlTarihi is not null and PlTarihi > '01.01.1950') or (plgonderitarihi is not null and plgonderitarihi > '01.01.1950')) "

            If cFilter.Trim = "" Then
                cSQL = cSQL + _
                        " and siparisno in (select kullanicisipno " + _
                                            " from siparis " + _
                                            " where planlamaok = 'E' " + _
                                            " and (dosyakapandi is null or dosyakapandi = 'H' or dosyakapandi = '') " + _
                                            " and planlamaok = 'E' " + _
                                            " and (plkapanis is null or plkapanis = 'H' or plkapanis = '')) "
            Else
                cSQL = cSQL + _
                        " and siparisno in (select kullanicisipno " + _
                                            " from siparis " + _
                                            " where planlamaok = 'E' " + _
                                            " and kullanicisipno in (" + cFilter + ") " + _
                                            " and planlamaok = 'E' " + _
                                            " and (plkapanis is null or plkapanis = 'H' or plkapanis = '')) "
            End If

            cSQL = cSQL + _
                    " order by siparisno, modelkodu "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aCP(nCnt)
                aCP(nCnt).dPlGonderiTarihi = SQLReadDate(oReader, "plgonderitarihi")
                aCP(nCnt).dPlTarihi = SQLReadDate(oReader, "pltarihi")
                aCP(nCnt).cSiparisNo = SQLReadString(oReader, "siparisno")
                aCP(nCnt).cModelKodu = SQLReadString(oReader, "modelkodu")
                aCP(nCnt).cOkTipi = SQLReadString(oReader, "oktipi")
                aCP(nCnt).cRenk = SQLReadString(oReader, "renk")
                aCP(nCnt).cBeden = SQLReadString(oReader, "beden")
                aCP(nCnt).dOkTar = SQLReadDate(oReader, "oktar")
                aCP(nCnt).dOkTar2 = SQLReadDate(oReader, "oktar2")

                nCnt = nCnt + 1
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aCP)

                cAciklama = IIf(aCP(nCnt).cRenk = "HEPSI", "", aCP(nCnt).cRenk + " ").ToString + _
                IIf(aCP(nCnt).cBeden = "HEPSI", "", aCP(nCnt).cBeden + " ").ToString

                cDepartman = "MT"
                cFirma = "DAHILI"
                cTipi = aCP(nCnt).cOkTipi
                cFoyNo = aCP(nCnt).cSiparisNo
                cRenk = aCP(nCnt).cRenk
                cBeden = aCP(nCnt).cBeden
                cDurum = ""
                dGBasla = aCP(nCnt).dOkTar
                dGBitir = aCP(nCnt).dOkTar2

                nIhtiyac = 0
                nIsemri = 0
                nBaslanan = 0
                nBiten = 0

                cSiparisNo = aCP(nCnt).cSiparisNo
                cModelNo = aCP(nCnt).cModelKodu

                cSQL = "select kalipno " + _
                        " from ymodel " + _
                        " where modelno = '" + cModelNo + "' "

                cKalipNo = SQLGetStringConnected(cSQL, ConnYage)

                cSQL = "select distinct musterino " + _
                        " from siparis " + _
                        " where kullanicisipno = '" + cSiparisNo + "' "

                cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select distinct sorumlu " + _
                        " from siparis " + _
                        " where kullanicisipno = '" + cSiparisNo + "' " + _
                        " and sorumlu is not null " + _
                        " and sorumlu <> '' "

                cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                cSQL = "select min(ilksevktar) " + _
                        " from sevkplfislines " + _
                        " where ilksevktar is not null " + _
                        " and ilksevktar <> '01.01.1950' " + _
                        " and exists (select sevkiyattakipno " + _
                                    " from sipmodel " + _
                                    " where sevkiyattakipno = sevkplfislines.sevkiyattakipno " + _
                                    " and siparisno = '" + cSiparisNo + "' " + _
                                    " and modelno = '" + cModelNo + "') "

                dSevkTarihi = SQLGetDateConnected(cSQL, ConnYage)

                nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                If aCP(nCnt).dPlTarihi > #1/1/1950# Then

                    cSQL = cInsertHeader + _
                            " values ('" + SQLWriteDate(aCP(nCnt).dPlTarihi) + "', " + _
                            " 'CP', " + _
                            " '" + SQLWriteString(cAciklama, 250) + "', " + _
                            " 'ONAY', " + _
                            " '" + cDurum + "', " + _
                            " '" + SQLWriteString(cPersonel, 30) + "', " + _
                            " '" + SQLWriteString(cFirma, 250) + "', " + _
                            " '" + SQLWriteString(cSiparisNo, 250) + "', " + _
                            " '" + SQLWriteString(cMusteriNo, 250) + "', " + _
                            " '" + SQLWriteString(cModelNo, 250) + "', " + _
                            " '" + SQLWriteString(cDepartman, 250) + "', " + _
                            " '" + SQLWriteString(cColor, 30) + "', " + _
                            " '" + SQLWriteString(cKalipNo, 250) + "', " + _
                            SQLWriteDecimal(nIhtiyac) + ", " + _
                            SQLWriteDecimal(nIsemri) + ", " + _
                            SQLWriteDecimal(nBaslanan) + ", " + _
                            SQLWriteDecimal(nBiten) + ", " + _
                            " '" + SQLWriteString(cTipi, 30) + "', " + _
                            " '" + SQLWriteString(cFoyNo, 30) + "', " + _
                            " '" + SQLWriteDate(dSevkTarihi) + "', " + _
                            SQLWriteDecimal(nPairKey) + ", " + _
                            " '" + SQLWriteString(cRenk, 30) + "', " + _
                            " '" + SQLWriteString(cBeden, 30) + "', " + _
                            " '" + SQLWriteDate(dGBitir) + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                If aCP(nCnt).dPlGonderiTarihi > #1/1/1950# Then

                    cSQL = cInsertHeader + _
                            " values ('" + SQLWriteDate(aCP(nCnt).dPlGonderiTarihi) + "', " + _
                            " 'CP', " + _
                            " '" + SQLWriteString(cAciklama, 250) + "', " + _
                            " 'YOLLA', " + _
                            " '" + cDurum + "', " + _
                            " '" + SQLWriteString(cPersonel, 30) + "', " + _
                            " '" + SQLWriteString(cFirma, 250) + "', " + _
                            " '" + SQLWriteString(cSiparisNo, 250) + "', " + _
                            " '" + SQLWriteString(cMusteriNo, 250) + "', " + _
                            " '" + SQLWriteString(cModelNo, 250) + "', " + _
                            " '" + SQLWriteString(cDepartman, 250) + "', " + _
                            " '" + SQLWriteString(cColor, 30) + "', " + _
                            " '" + SQLWriteString(cKalipNo, 250) + "', " + _
                            SQLWriteDecimal(nIhtiyac) + ", " + _
                            SQLWriteDecimal(nIsemri) + ", " + _
                            SQLWriteDecimal(nBaslanan) + ", " + _
                            SQLWriteDecimal(nBiten) + ", " + _
                            " '" + SQLWriteString(cTipi, 30) + "', " + _
                            " '" + SQLWriteString(cFoyNo, 30) + "', " + _
                            " '" + SQLWriteDate(dSevkTarihi) + "', " + _
                            SQLWriteDecimal(nPairKey) + ", " + _
                            " '" + SQLWriteString(cRenk, 30) + "', " + _
                            " '" + SQLWriteString(cBeden, 30) + "', " + _
                            " '" + SQLWriteDate(dGBasla) + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            JustForLog("Masterplan build end")

            ConnYage.Close()

            GetMasterPlanData = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "GetMasterPlanData", cSQL)
        End Try
    End Function

    Public Sub GetMTFGercekTarih(ByVal ConnYage As SqlConnection, ByVal cMTF As String, ByVal cStokno As String, Optional ByVal cRenk As String = "", Optional ByVal cBeden As String = "", _
                                 Optional ByRef dBasla As Date = #1/1/1950#, Optional ByRef dBitir As Date = #1/1/1950#)
        Dim cSQL As String = ""

        Try
            ' Malzemede başlangıç tarihi ilk işemri veriliş tarihidir
            cSQL = "select a.tarih " + _
                    " from isemri a, isemrilines b " + _
                    " where a.isemrino = b.isemrino " + _
                    " and b.malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and b.stokno = '" + cStokno.Trim + "' " + _
                    IIf(cRenk.Trim = "", "", " and b.renk = '" + cRenk.Trim + "' ").ToString + _
                    IIf(cBeden.Trim = "", "", " and b.beden = '" + cBeden.Trim + "' ").ToString + _
                    " and a.tarih is not null " + _
                    " order by a.tarih "

            dBasla = SQLGetDateConnected(cSQL, ConnYage)

            ' Malzeme bitiş tarihi en son giriş hareket tarihidir
            cSQL = "select a.fistarihi " + _
                    " from stokfis a, stokfislines b " + _
                    " where a.stokfisno = b.stokfisno " + _
                    " and b.malzemetakipkodu = '" + cMTF.Trim + "' " + _
                    " and b.stokno = '" + cStokno.Trim + "' " + _
                    IIf(cRenk.Trim = "", "", " and b.renk = '" + cRenk.Trim + "' ").ToString + _
                    IIf(cBeden.Trim = "", "", " and b.beden = '" + cBeden.Trim + "' ").ToString + _
                    " and b.stokhareketkodu in ('04 Mlz Uretimden Giris','06 Tamirden Giris','02 Tedarikten Giris','05 Diger Giris','90 Trans/Rezv Giris') " + _
                    " and a.fistarihi is not null " + _
                    " order by a.fistarihi desc "

            dBitir = SQLGetDateConnected(cSQL, ConnYage)

            If dBitir = #1/1/1950# Then
                ' Malzeme bitiş tarihi en son transfer hareket tarihidir
                cSQL = "select tarih " + _
                        " from stoktransfer " + _
                        " where hedefmalzemetakipno = '" + cMTF.Trim + "' " + _
                        " and stokno = '" + cStokno.Trim + "' " + _
                        IIf(cRenk.Trim = "", "", " and renk = '" + cRenk.Trim + "' ").ToString + _
                        IIf(cBeden.Trim = "", "", " and beden = '" + cBeden.Trim + "' ").ToString + _
                        " and tarih is not null " + _
                        " order by tarih desc "

                dBitir = SQLGetDateConnected(cSQL, ConnYage)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "GetMTFGercekTarih", cSQL)
        End Try
    End Sub

    Public Sub GetUTFGercekTarih(ByVal ConnYage As SqlConnection, ByVal cUTF As String, ByVal cDepartman As String, ByVal cModelNo As String, _
                                Optional ByRef dBasla As Date = #1/1/1950#, Optional ByRef dBitir As Date = #1/1/1950#)
        Dim cSQL As String = ""

        Try
            ' UTFde ilk giriş tarihi, KESIM için ilk kumaş çıkış tarihidir
            If cDepartman.Trim = "KESIM" Then
                cSQL = "select a.fistarihi " + _
                        " from stokfis a, stokfislines b " + _
                        " where a.stokfisno = b.stokfisno " + _
                        " and a.departman = '" + cDepartman.Trim + "' " + _
                        " and b.uretimtakipno = '" + cUTF.Trim + "' " + _
                        " and b.modelno = '" + cModelNo.Trim + "' " + _
                        " and b.stokhareketkodu = '01 Uretime Cikis' " + _
                        " and a.fistarihi is not null " + _
                        " order by a.fistarihi "

                dBasla = SQLGetDateConnected(cSQL, ConnYage)
            Else
                cSQL = "select a.fistarihi " + _
                        " from uretharfis a, uretharfislines b " + _
                        " where a.uretfisno = b.uretfisno " + _
                        " and a.girisdept = '" + cDepartman.Trim + "' " + _
                        " and b.uretimtakipno = '" + cUTF.Trim + "' " + _
                        " and b.modelno = '" + cModelNo.Trim + "' " + _
                        " order by a.fistarihi "

                dBasla = SQLGetDateConnected(cSQL, ConnYage)
            End If
            ' UTF bitiş tarihi son çıkış fiş tarihidir
            cSQL = "select a.fistarihi " + _
                    " from uretharfis a, uretharfislines b " + _
                    " where a.uretfisno = b.uretfisno " + _
                    " and a.cikisdept = '" + cDepartman.Trim + "' " + _
                    " and b.uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and b.modelno = '" + cModelNo.Trim + "' " + _
                    " order by a.fistarihi desc "

            dBitir = SQLGetDateConnected(cSQL, ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetUTFGercekTarih", cSQL)
        End Try
    End Sub

    Public Sub GetSTFGercekTarih(ByVal ConnYage As SqlConnection, ByVal cSTF As String, ByRef dGerceklesen As Date, Optional ByVal cSiparisNo As String = "", Optional ByVal cModelNo As String = "")

        Dim cSQL As String = ""

        Try
            ' gerçekleşen tarih son sevkiyat tarihidir
            cSQL = "select a.sevktar " + _
                    " from sevkform a, sevkformlines b " + _
                    " where a.sevkformno = b.sevkformno " + _
                    " and b.sevkiyattakipno = '" + cSTF.Trim + "' " + _
                    IIf(cSiparisNo.Trim = "", "", " and b.siparisno = '" + cSiparisNo.Trim + "' ").ToString + _
                    IIf(cModelNo.Trim = "", "", " and b.modelno = '" + cModelNo.Trim + "' ").ToString + _
                    " order by a.sevktar desc "

            dGerceklesen = SQLGetDateConnected(cSQL, ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSTFGercekTarih", cSQL)
        End Try
    End Sub

    Public Function Yuvarlat(ByVal ConnYage As SqlConnection, ByVal cBirim As String, ByVal nMiktar As Double, Optional ByRef nDecimal As Integer = 0) As Double

        Dim cSQL As String = ""
        Dim cMiktar As String = ""
        Dim aMiktar() As String = Nothing

        Yuvarlat = nMiktar

        Try
            If cBirim = "" Then Exit Function

            cMiktar = nMiktar.ToString

            cSQL = "select kusurat " + _
                        " from birim " + _
                        " where birim = '" + cBirim.Trim + "' "

            nDecimal = CInt(SQLGetDoubleConnected(cSQL, ConnYage))

            If nDecimal = 0 Then
                If InStr(cMiktar, ".") > 0 Then
                    aMiktar = Split(cMiktar, ".")
                    If Val(Mid(aMiktar(1), 1, 1)) >= 5 Then
                        Yuvarlat = Val(aMiktar(0)) + 1
                    Else
                        Yuvarlat = Val(aMiktar(0))
                    End If
                End If
            Else
                Yuvarlat = Math.Round(nMiktar, nDecimal)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "Yuvarlat", cSQL)
        End Try
    End Function
End Module

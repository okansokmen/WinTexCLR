Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module utilMasterPlan

    Private Structure ostisonmaliyet8
        Dim csiparisno As String
        Dim cmodelno As String
        Dim caciklama As String
        Dim nfiyat As Double
        Dim cdoviz As String
        Dim nsiparisadet As Double
        Dim nsiparistutar As Double
        Dim nsevkadet As Double
        Dim nsevktutar As Double
        Dim g_sevktarihi As Date
        Dim g_kumastutar As Double
        Dim g_aksesuartutar As Double
        Dim g_isciliktutar As Double
        Dim g_genelgider As Double
        Dim g_digertutar As Double
        Dim p_sevktarihi As Date
        Dim p_kumastutar As Double
        Dim p_aksesuartutar As Double
        Dim p_isciliktutar As Double
        Dim p_genelgider As Double
        Dim p_digertutar As Double
        Dim cmusterino As sting
        Dim cdosyakapandi As String
    End Structure

    Private Structure oMasraf
        Dim dTarih As Date
        Dim cDoviz As String
        Dim nTutar As Double
        Dim nKur As Double
        Dim nEURKur As Double
    End Structure

    Private Structure oSHKur
        Dim dTarih As Date
        Dim cDoviz As String
        Dim nStokHareketNo As Double
        Dim cKurCinsi As String
    End Structure

    Private Structure oReklamasyon
        Dim dTarih As Date
        Dim cDoviz As String
        Dim nTutar As Double
    End Structure

    Private Structure oTanimlamaData
        Dim cMasraf As String
        Dim cDoviz As String
        Dim nTutar As Double
    End Structure

    Private Structure oStiSonMaliyet7
        Dim cSiparisNo As String
        Dim cTipi As String
        Dim cBirim As String
        Dim nMiktar As Double
        Dim nTutar As Double
        Dim nPlMiktar As Double
        Dim nPlTutar As Double
    End Structure

    Private Structure oOnMaliyet
        Dim cMlzAdi As String
        Dim cMlzCode As String
        Dim cAnaStokGrubu As String
        Dim cStokTipi As String
        Dim cBirim As String
        Dim cDoviz As String
        Dim nBirimMiktar As Double
        Dim nMiktar As Double
        Dim nFiyat As Double
        Dim nEURFiyat As Double
        Dim nEURTutar As Double
        Dim nKur As Double
        Dim nEURKur As Double
    End Structure

    Private Structure oTransfer
        Dim cTransferFisNo As String
        Dim dTarih As Date
        Dim nMiktar As Double
        Dim nFiyat As Double
        Dim cDoviz As String
        Dim cStokTipi As String
        Dim cKaynalMTF As String
        Dim cHedefMTF As String
        Dim nKur As Double
        Dim nEURKur As Double
        Dim nOrjTutar As Double
        Dim nTLTutar As Double
        Dim nEURTutar As Double
    End Structure

    Private Structure oStokFis
        Dim cStokFisNo As String
        Dim dFisTarihi As Date
        Dim cBelgeNo As String
        Dim dFaturaTarihi As Date
        Dim cFaturaNo As String
        Dim cDepartman As String
        Dim cFirma As String
        Dim nMiktar As Double
        Dim nFiyat As Double
        Dim cDoviz As String
        Dim nKur As Double
        Dim nIscilikFiyat As Double
        Dim cIscilikDoviz As String
        Dim nIscilikKur As Double
        Dim nEURKur As Double
        Dim nOrjTutar As Double
        Dim nTLTutar As Double
        Dim nEURTutar As Double
        Dim cStokHareketKodu As String
    End Structure

    Private Structure oMalzeme
        Dim cStokNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cStokTipi As String
        Dim cAnaStokGrubu As String
        Dim cBirim As String
        Dim nIhtiyac As Double
        Dim nUretimeCikan As Double
        Dim nUretimIade As Double
    End Structure

    Private Structure oGenelgider
        Dim nYil As Double
        Dim nAy As Double
        Dim nGenelGiderEUR As Double
        Dim nGumrukGiderEUR As Double
        Dim nToplamSevk As Double
        Dim cSiparisNo As String
        Dim nSevk As Double
    End Structure

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

    Private Structure oUretFis
        Dim cUretFisNo As String
        Dim cModelNo As String
        Dim cDepartman As String
        Dim cFirma As String
        Dim cUTF As String
        Dim cIsemriNo As String
        Dim dTarih As Date
        Dim nMiktar As Double
        Dim cBelgeNo As String
        Dim cFaturaNo As String
        Dim dFaturaTarihi As Date

        Dim nKur As Double
        Dim nTLFiyat As Double
        Dim nTLTutar As Double

        Dim nEURKUR As Double
        Dim nEURFiyat As Double
        Dim nEURTutar As Double

        Dim nFiyat As Double
        Dim nTutar As Double
        Dim cDoviz As String
    End Structure

    Private Structure oUIE
        Dim cSiparisNo As String
        Dim cStokTipi As String
        Dim cDepartman As String
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
        Dim lSTFOK As Boolean = False
        Dim lUTFOK As Boolean = False
        Dim lMTFOK As Boolean = False
        Dim lCPOK As Boolean = False

        GetMasterPlanData = 0

        Try

            ConnYage = OpenConn()

            JustForLog("Masterplan build start")

            lDinamikMTF = (CDbl(GetSysParConnected("masterplandinamikmtf", ConnYage)) = 1)

            cSQL = "delete " + cTable
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cInsertHeader = "Set dateformat dmy " +
                            " insert into " + cTable +
                            " (tarih, kategori, aciklama, baslabitir, durum, " +
                            " personel, firma, siparisno, musterino, modelno, " +
                            " departman, color, kalipno, ihtiyac, isemri, " +
                            " baslanan, biten, tipi, foyno, sevktarihi, " +
                            " pairkey, renk, beden, gercektarih ) "

            ' STF
            JustForLog("STF-Masterplan build start")

            nCnt = 0

            cSQL = "Select distinct ilksevktar, sevkiyattakipno, sevkemrino, toplam, planlanan, giden, " +
                    " gersevktar = (Select top 1 a.sevktar " +
                                    " from sevkform a, sevkformlines b " +
                                    " where a.sevkformno = b.sevkformno " +
                                    " And b.sevkiyattakipno = sevkplfislines.sevkiyattakipno " +
                                    " order by a.sevktar desc ) " +
                    " from sevkplfislines "

            cSQL = cSQL +
                    " where (ok Is null Or ok = 'H' or ok = '') " +
                    " and ilksevktar is not null " +
                    " and ilksevktar > '01.01.1950' "

            If cFilter.Trim = "" Then
                cSQL = cSQL +
                        " and sevkiyattakipno in (select y.sevkiyattakipno " +
                                                " from siparis x, sipmodel y " +
                                                " where x.kullanicisipno = y.siparisno " +
                                                " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '')  " +
                                                " and x.planlamaok = 'E' " +
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            Else
                cSQL = cSQL +
                        " and sevkiyattakipno in (select y.sevkiyattakipno " +
                                                " from siparis x, sipmodel y " +
                                                " where x.kullanicisipno = y.siparisno " +
                                                " and x.kullanicisipno in (" + cFilter + ") " +
                                                " and x.planlamaok = 'E' " +
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            End If

            cSQL = cSQL +
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

                lSTFOK = True
                nCnt = nCnt + 1
            Loop
            oReader.Close()

            If lSTFOK Then
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

                    cSQL = "select distinct siparisno " +
                            " from sipmodel " +
                            " where sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " +
                            " and siparisno is not null " +
                            " and siparisno <> '' "

                    cSiparisNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct modelno " +
                            " from sipmodel " +
                            " where sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " +
                            " and modelno is not null " +
                            " and modelno <> '' "

                    cModelNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct a.kalipno " +
                            " from ymodel a, sipmodel b " +
                            " where a.modelno = b.modelno " +
                            " and b.sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " +
                            " and a.modelno is not null " +
                            " and a.modelno <> '' "

                    cKalipNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct a.musterino " +
                            " from siparis a, sipmodel b " +
                            " where a.kullanicisipno = b.siparisno " +
                            " and b.sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " +
                            " and a.musterino is not null " +
                            " and a.musterino <> '' "

                    cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct a.sorumlu " +
                            " from siparis a, sipmodel b " +
                            " where a.kullanicisipno = b.siparisno " +
                            " and b.sevkiyattakipno = '" + aSTF(nCnt).cSevkiyatTakipNo + "' " +
                            " and a.sorumlu is not null " +
                            " and a.sorumlu <> '' "

                    cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cFirma = cMusteriNo

                    nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                    cSQL = cInsertHeader +
                            " values ('" + SQLWriteDate(dSevkTarihi) + "', " +
                            " 'STF', " +
                            " '" + SQLWriteString(cAciklama, 250) + "'," +
                            " 'SEVKIYAT', " +
                            " '" + cDurum + "', " +
                            " '" + SQLWriteString(cPersonel, 30) + "', " +
                            " '" + SQLWriteString(cFirma, 250) + "', " +
                            " '" + SQLWriteString(cSiparisNo, 250) + "', " +
                            " '" + SQLWriteString(cMusteriNo, 250) + "', " +
                            " '" + SQLWriteString(cModelNo, 250) + "', " +
                            " '" + SQLWriteString(cDepartman, 250) + "', " +
                            " '" + SQLWriteString(cColor, 30) + "', " +
                            " '" + SQLWriteString(cKalipNo, 250) + "', " +
                            SQLWriteDecimal(nIhtiyac) + ", " +
                            SQLWriteDecimal(nIsemri) + ", " +
                            SQLWriteDecimal(nBaslanan) + ", " +
                            SQLWriteDecimal(nBiten) + ", " +
                            " '" + SQLWriteString(cTipi, 30) + "', " +
                            " '" + SQLWriteString(cFoyNo, 30) + "', " +
                            " '" + SQLWriteDate(dSevkTarihi) + "', " +
                            SQLWriteDecimal(nPairKey) + ",'','', " +
                            " '" + SQLWriteDate(dGerceklesen) + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Next
            End If


            ' UTF
            JustForLog("UTF-Masterplan build start")

            nCnt = 0

            cSQL = "select uretimtakipno, baslamatarihi, bitistarihi, departman, modelno, plfirma, " +
                    " toplamadet = sum(coalesce(toplamadet,0)), " +
                    " isemriverilen = sum(coalesce(isemriverilen,0)), " +
                    " gelen = sum(coalesce(gelen,0)), " +
                    " giden = sum(coalesce(giden,0)) " +
                    " from uretpllines "

            cSQL = cSQL +
                    " where (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " +
                    " and ((baslamatarihi is not null and baslamatarihi > '01.01.1950') or (bitistarihi is not null and bitistarihi > '01.01.1950')) "

            If cFilter.Trim = "" Then
                cSQL = cSQL +
                        " and uretimtakipno in (select y.uretimtakipno " +
                                                " from siparis x, sipmodel y " +
                                                " where x.kullanicisipno = y.siparisno " +
                                                " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '')  " +
                                                " and x.planlamaok = 'E' " +
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            Else
                cSQL = cSQL +
                        " and uretimtakipno in (select y.uretimtakipno " +
                                                " from siparis x, sipmodel y " +
                                                " where x.kullanicisipno = y.siparisno " +
                                                " and x.kullanicisipno in (" + cFilter + ") " +
                                                " and x.planlamaok = 'E' " +
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            End If


            cSQL = cSQL +
                    " group by uretimtakipno, baslamatarihi, bitistarihi, departman, modelno, plfirma " +
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

                lUTFOK = True
                nCnt = nCnt + 1
            Loop
            oReader.Close()

            If lUTFOK Then
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

                    cSQL = "select kalipno " +
                            " from ymodel " +
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

                    cSQL = "select distinct siparisno " +
                            " from sipmodel " +
                            " where uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " +
                            " and modelno = '" + aUTF(nCnt).cModelNo + "' " +
                            " and siparisno is not null " +
                            " and siparisno <> '' "

                    cSiparisNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct a.musterino " +
                            " from siparis a, sipmodel b " +
                            " where a.kullanicisipno = b.siparisno " +
                            " and b.uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " +
                            " and b.modelno = '" + aUTF(nCnt).cModelNo + "' " +
                            " and a.musterino is not null " +
                            " and a.musterino <> '' "

                    cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    ' üretim işemirlerindeki personel
                    cSQL = "select distinct a.eleman " +
                            " from uretimisemri a, uretimisdetayi b " +
                            " where a.isemrino = b.isemrino " +
                            " and a.uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " +
                            " and a.departman = '" + aUTF(nCnt).cDepartman + "' " +
                            " and b.modelno = '" + aUTF(nCnt).cModelNo + "' " +
                            " and a.eleman is not null " +
                            " and a.eleman <> '' "

                    cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                    ' üretim işemirlerindeki firmalar
                    cSQL = "select distinct a.firma " +
                            " from uretimisemri a, uretimisdetayi b " +
                            " where a.isemrino = b.isemrino " +
                            " and a.uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " +
                            " and a.departman = '" + aUTF(nCnt).cDepartman + "' " +
                            " and b.modelno = '" + aUTF(nCnt).cModelNo + "' " +
                            " and a.firma is not null " +
                            " and a.firma <> '' "

                    cFirma = SQLBuildFilterString2(ConnYage, cSQL, False)

                    If Trim(cFirma) = "" Then
                        cFirma = aUTF(nCnt).cPlFirma
                    End If

                    cSQL = "select min(ilksevktar) " +
                            " from sevkplfislines " +
                            " where ilksevktar is not null " +
                            " and ilksevktar <> '01.01.1950' " +
                            " and exists (select sevkiyattakipno " +
                                        " from sipmodel " +
                                        " where sevkiyattakipno = sevkplfislines.sevkiyattakipno " +
                                        " and uretimtakipno = '" + aUTF(nCnt).cUretimTakipNo + "' " +
                                        " and modelno = '" + aUTF(nCnt).cModelNo + "') "

                    dSevkTarihi = SQLGetDateConnected(cSQL, ConnYage)

                    nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                    If aUTF(nCnt).dBitisTarihi > #1/1/1950# Then

                        cSQL = cInsertHeader +
                                " values ('" + SQLWriteDate(aUTF(nCnt).dBitisTarihi) + "', " +
                                " 'UTF', " +
                                " '" + SQLWriteString(cAciklama, 250) + "'," +
                                " 'BITIR', " +
                                " '" + cDurum + "', " +
                                " '" + SQLWriteString(cPersonel, 30) + "', " +
                                " '" + SQLWriteString(cFirma, 250) + "', " +
                                " '" + SQLWriteString(cSiparisNo, 250) + "', " +
                                " '" + SQLWriteString(cMusteriNo, 250) + "', " +
                                " '" + SQLWriteString(cModelNo, 250) + "', " +
                                " '" + SQLWriteString(cDepartman, 250) + "', " +
                                " '" + SQLWriteString(cColor, 30) + "', " +
                                " '" + SQLWriteString(cKalipNo, 250) + "', " +
                                SQLWriteDecimal(nIhtiyac) + ", " +
                                SQLWriteDecimal(nIsemri) + ", " +
                                SQLWriteDecimal(nBaslanan) + ", " +
                                SQLWriteDecimal(nBiten) + ", " +
                                " '" + SQLWriteString(cTipi, 30) + "', " +
                                " '" + SQLWriteString(cFoyNo, 30) + "', " +
                                " '" + SQLWriteDate(dSevkTarihi) + "', " +
                                SQLWriteDecimal(nPairKey) + ",'','', " +
                                " '" + SQLWriteDate(dGBitir) + "') "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                    If aUTF(nCnt).dBaslamaTarihi > #1/1/1950# Then

                        cSQL = cInsertHeader +
                                " values ('" + SQLWriteDate(aUTF(nCnt).dBaslamaTarihi) + "', " +
                                " 'UTF', " +
                                " '" + SQLWriteString(cAciklama, 250) + "'," +
                                " 'BASLA', " +
                                " '" + cDurum + "', " +
                                " '" + SQLWriteString(cPersonel, 30) + "', " +
                                " '" + SQLWriteString(cFirma, 250) + "', " +
                                " '" + SQLWriteString(cSiparisNo, 250) + "', " +
                                " '" + SQLWriteString(cMusteriNo, 250) + "', " +
                                " '" + SQLWriteString(cModelNo, 250) + "', " +
                                " '" + SQLWriteString(cDepartman, 250) + "', " +
                                " '" + SQLWriteString(cColor, 30) + "', " +
                                " '" + SQLWriteString(cKalipNo, 250) + "', " +
                                SQLWriteDecimal(nIhtiyac) + ", " +
                                SQLWriteDecimal(nIsemri) + ", " +
                                SQLWriteDecimal(nBaslanan) + ", " +
                                SQLWriteDecimal(nBiten) + ", " +
                                " '" + SQLWriteString(cTipi, 30) + "', " +
                                " '" + SQLWriteString(cFoyNo, 30) + "', " +
                                " '" + SQLWriteDate(dSevkTarihi) + "', " +
                                SQLWriteDecimal(nPairKey) + ",'','', " +
                                " '" + SQLWriteDate(dGBasla) + "') "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                Next
            End If

            ' MTF
            JustForLog("MTF-Masterplan build start")

            If lDinamikMTF Then
                nSonuc = GetToplamSiparisView_1("", cIhtiyacTable, ConnYage)
                nSonuc = MTFHesaplax_1("", "", cIhtiyacTable, cDetayIhtiyacTable, ConnYage)
            End If

            nCnt = 0

            cSQL = "select a.malzemetakipno, a.baslamatarihi, a.bitistarihi, b.stoktipi, a.plfirma, a.stokno, a.departman, b.birim1, " +
                    " ihtiyac = sum(coalesce(ihtiyac,0)), " +
                    " isemriverilen = sum(coalesce(isemriverilen,0)), " +
                    " karsilanan = sum(coalesce(isemriicingelen,0)) + sum(coalesce(isemriharicigelen,0)) " +
                    " from mtkfislines a, stok b "

            cSQL = cSQL +
                    " where a.stokno = b.stokno " +
                    " and (a.kapandi is null or a.kapandi = 'H' or a.kapandi = '') " +
                    " and ((a.baslamatarihi is not null and a.baslamatarihi > '01.01.1950') or (a.bitistarihi is not null and a.bitistarihi > '01.01.1950')) "

            If cFilter.Trim = "" Then
                cSQL = cSQL +
                        " and malzemetakipno in (select y.malzemetakipno " +
                                                " from siparis x, sipmodel y " +
                                                " where x.kullanicisipno = y.siparisno " +
                                                " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '') " +
                                                " and x.planlamaok = 'E' " +
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            Else
                cSQL = cSQL +
                        " and malzemetakipno in (select y.malzemetakipno " +
                                                " from siparis x, sipmodel y " +
                                                " where x.kullanicisipno = y.siparisno " +
                                                " and x.kullanicisipno in (" + cFilter + ") " +
                                                " and x.planlamaok = 'E' " +
                                                " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '')) "
            End If

            cSQL = cSQL +
                    " group by a.malzemetakipno, a.baslamatarihi, a.bitistarihi, b.stoktipi, a.plfirma, a.stokno, a.departman, b.birim1 " +
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

                lMTFOK = True
                nCnt = nCnt + 1
            Loop
            oReader.Close()

            If lMTFOK Then
                For nCnt = 0 To UBound(aMTF)

                    cAciklama = aMTF(nCnt).cStokNo
                    cTipi = aMTF(nCnt).cStokTipi
                    cFoyNo = aMTF(nCnt).cMalzemeTakipNo

                    GetMTFGercekTarih(ConnYage, aMTF(nCnt).cMalzemeTakipNo, aMTF(nCnt).cStokNo, , , dGBasla, dGBitir)

                    If lDinamikMTF Then

                        cSQL = "SELECT ihtiyac = sum(coalesce(ihtiyac,0)), " +
                                " uretimecikan = sum(coalesce(uretimecikan,0)), " +
                                " gelecek = sum(coalesce(gelecek,0)), " +
                                " stokmiktari = sum(coalesce(stokmiktari,0)) " +
                                " from " + cDetayIhtiyacTable +
                                " where malzemetakipkodu = '" + aMTF(nCnt).cMalzemeTakipNo + "' " +
                                " and stokno = '" + aMTF(nCnt).cStokNo + "' " +
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

                    cSQL = "select distinct siparisno " +
                             " from sipmodel " +
                             " where malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " +
                             " and siparisno is not null " +
                             " and siparisno <> '' "

                    cSiparisNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct modelno " +
                            " from sipmodel " +
                            " where malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " +
                            " and modelno is not null " +
                            " and modelno <> '' "

                    cModelNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct a.kalipno " +
                            " from ymodel a, sipmodel b " +
                            " where a.modelno = b.modelno " +
                            " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " +
                            " and a.modelno is not null " +
                            " and a.modelno <> '' "

                    cKalipNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct a.musterino " +
                            " from siparis a, sipmodel b " +
                            " where a.kullanicisipno = b.siparisno " +
                            " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " +
                            " and a.musterino is not null " +
                            " and a.musterino <> '' "

                    cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    ' firmaları işemirlerinden al
                    cSQL = "select distinct a.firma " +
                            " from isemri a, isemrilines b " +
                            " where a.isemrino = b.isemrino " +
                            " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " +
                            " and b.stokno = '" + aMTF(nCnt).cStokNo + "' " +
                            " and a.firma is not null " +
                            " and a.firma <> '' "

                    cFirma = SQLBuildFilterString2(ConnYage, cSQL, False)

                    If Trim(cFirma) = "" Then
                        cFirma = aMTF(nCnt).cPlFirma
                    End If

                    ' personeli işemirlerinden al
                    cSQL = "select distinct a.takipelemani " +
                            " from isemri a, isemrilines b " +
                            " where a.isemrino = b.isemrino " +
                            " and b.malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "' " +
                            " and b.stokno = '" + aMTF(nCnt).cStokNo + "' " +
                            " and a.takipelemani is not null " +
                            " and a.takipelemani <> '' "

                    cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select min(ilksevktar) " +
                            " from sevkplfislines " +
                            " where ilksevktar is not null " +
                            " and ilksevktar <> '01.01.1950' " +
                            " and exists (select sevkiyattakipno " +
                                        " from sipmodel " +
                                        " where sevkiyattakipno = sevkplfislines.sevkiyattakipno " +
                                        " and malzemetakipno = '" + aMTF(nCnt).cMalzemeTakipNo + "') "

                    dSevkTarihi = SQLGetDateConnected(cSQL, ConnYage)

                    nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                    If aMTF(nCnt).dBitisTarihi > #1/1/1950# Then

                        cSQL = cInsertHeader +
                                " values ('" + SQLWriteDate(aMTF(nCnt).dBitisTarihi) + "', " +
                                " 'MTF', " +
                                " '" + SQLWriteString(cAciklama, 250) + "'," +
                                " 'DEPOYA GIRIS', " +
                                " '" + cDurum + "', " +
                                " '" + SQLWriteString(cPersonel, 30) + "', " +
                                " '" + SQLWriteString(cFirma, 250) + "', " +
                                " '" + SQLWriteString(cSiparisNo, 250) + "', " +
                                " '" + SQLWriteString(cMusteriNo, 250) + "', " +
                                " '" + SQLWriteString(cModelNo, 250) + "', " +
                                " '" + SQLWriteString(cDepartman, 250) + "', " +
                                " '" + SQLWriteString(cColor, 30) + "', " +
                                " '" + SQLWriteString(cKalipNo, 250) + "', " +
                                SQLWriteDecimal(nIhtiyac) + ", " +
                                SQLWriteDecimal(nIsemri) + ", " +
                                SQLWriteDecimal(nBaslanan) + ", " +
                                SQLWriteDecimal(nBiten) + ", " +
                                " '" + SQLWriteString(cTipi, 30) + "', " +
                                " '" + SQLWriteString(cFoyNo, 30) + "', " +
                                " '" + SQLWriteDate(dSevkTarihi) + "', " +
                                SQLWriteDecimal(nPairKey) + ",'','', " +
                                " '" + SQLWriteDate(dGBitir) + "') "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If


                    If aMTF(nCnt).dBaslamaTarihi > #1/1/1950# Then

                        cSQL = cInsertHeader +
                                " values ('" + SQLWriteDate(aMTF(nCnt).dBaslamaTarihi) + "', " +
                                " 'MTF', " +
                                " '" + SQLWriteString(cAciklama, 250) + "'," +
                                " 'SATINALMA ISEMRI', " +
                                " '" + cDurum + "', " +
                                " '" + SQLWriteString(cPersonel, 30) + "', " +
                                " '" + SQLWriteString(cFirma, 250) + "', " +
                                " '" + SQLWriteString(cSiparisNo, 250) + "', " +
                                " '" + SQLWriteString(cMusteriNo, 250) + "', " +
                                " '" + SQLWriteString(cModelNo, 250) + "', " +
                                " '" + SQLWriteString(cDepartman, 250) + "', " +
                                " '" + SQLWriteString(cColor, 30) + "', " +
                                " '" + SQLWriteString(cKalipNo, 250) + "', " +
                                SQLWriteDecimal(nIhtiyac) + ", " +
                                SQLWriteDecimal(nIsemri) + ", " +
                                SQLWriteDecimal(nBaslanan) + ", " +
                                SQLWriteDecimal(nBiten) + ", " +
                                " '" + SQLWriteString(cTipi, 30) + "', " +
                                " '" + SQLWriteString(cFoyNo, 30) + "', " +
                                " '" + SQLWriteDate(dSevkTarihi) + "', " +
                                SQLWriteDecimal(nPairKey) + ",'','', " +
                                " '" + SQLWriteDate(dGBasla) + "') "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                Next
            End If

            If lDinamikMTF Then
                DropTable(cIhtiyacTable, ConnYage)
                DropTable(cDetayIhtiyacTable, ConnYage)
            End If

            ' CP
            JustForLog("CP-Masterplan build start")

            nCnt = 0

            cSQL = "select distinct plgonderitarihi, pltarihi, siparisno, modelkodu, oktipi, renk, beden, oktar, oktar2  " +
                    " from sipok "

            cSQL = cSQL +
                    " where (ok is null or ok = 'H' or ok = '') " +
                    " and ((PlTarihi is not null and PlTarihi > '01.01.1950') or (plgonderitarihi is not null and plgonderitarihi > '01.01.1950')) "

            If cFilter.Trim = "" Then
                cSQL = cSQL +
                        " and siparisno in (select kullanicisipno " +
                                            " from siparis " +
                                            " where planlamaok = 'E' " +
                                            " and (dosyakapandi is null or dosyakapandi = 'H' or dosyakapandi = '') " +
                                            " and planlamaok = 'E' " +
                                            " and (plkapanis is null or plkapanis = 'H' or plkapanis = '')) "
            Else
                cSQL = cSQL +
                        " and siparisno in (select kullanicisipno " +
                                            " from siparis " +
                                            " where planlamaok = 'E' " +
                                            " and kullanicisipno in (" + cFilter + ") " +
                                            " and planlamaok = 'E' " +
                                            " and (plkapanis is null or plkapanis = 'H' or plkapanis = '')) "
            End If

            cSQL = cSQL +
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

                lCPOK = True
                nCnt = nCnt + 1
            Loop
            oReader.Close()

            If lCPOK Then
                For nCnt = 0 To UBound(aCP)

                    cAciklama = IIf(aCP(nCnt).cRenk = "HEPSI", "", aCP(nCnt).cRenk + " ").ToString +
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

                    cSQL = "select kalipno " +
                            " from ymodel " +
                            " where modelno = '" + cModelNo + "' "

                    cKalipNo = SQLGetStringConnected(cSQL, ConnYage)

                    cSQL = "select distinct musterino " +
                            " from siparis " +
                            " where kullanicisipno = '" + cSiparisNo + "' "

                    cMusteriNo = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select distinct sorumlu " +
                            " from siparis " +
                            " where kullanicisipno = '" + cSiparisNo + "' " +
                            " and sorumlu is not null " +
                            " and sorumlu <> '' "

                    cPersonel = SQLBuildFilterString2(ConnYage, cSQL, False)

                    cSQL = "select min(ilksevktar) " +
                            " from sevkplfislines " +
                            " where ilksevktar is not null " +
                            " and ilksevktar <> '01.01.1950' " +
                            " and exists (select sevkiyattakipno " +
                                        " from sipmodel " +
                                        " where sevkiyattakipno = sevkplfislines.sevkiyattakipno " +
                                        " and siparisno = '" + cSiparisNo + "' " +
                                        " and modelno = '" + cModelNo + "') "

                    dSevkTarihi = SQLGetDateConnected(cSQL, ConnYage)

                    nPairKey = CDbl(GetFisNoConnected(ConnYage, "pairkey"))

                    If aCP(nCnt).dPlTarihi > #1/1/1950# Then

                        cSQL = cInsertHeader +
                                " values ('" + SQLWriteDate(aCP(nCnt).dPlTarihi) + "', " +
                                " 'CP', " +
                                " '" + SQLWriteString(cAciklama, 250) + "', " +
                                " 'ONAY', " +
                                " '" + cDurum + "', " +
                                " '" + SQLWriteString(cPersonel, 30) + "', " +
                                " '" + SQLWriteString(cFirma, 250) + "', " +
                                " '" + SQLWriteString(cSiparisNo, 250) + "', " +
                                " '" + SQLWriteString(cMusteriNo, 250) + "', " +
                                " '" + SQLWriteString(cModelNo, 250) + "', " +
                                " '" + SQLWriteString(cDepartman, 250) + "', " +
                                " '" + SQLWriteString(cColor, 30) + "', " +
                                " '" + SQLWriteString(cKalipNo, 250) + "', " +
                                SQLWriteDecimal(nIhtiyac) + ", " +
                                SQLWriteDecimal(nIsemri) + ", " +
                                SQLWriteDecimal(nBaslanan) + ", " +
                                SQLWriteDecimal(nBiten) + ", " +
                                " '" + SQLWriteString(cTipi, 30) + "', " +
                                " '" + SQLWriteString(cFoyNo, 30) + "', " +
                                " '" + SQLWriteDate(dSevkTarihi) + "', " +
                                SQLWriteDecimal(nPairKey) + ", " +
                                " '" + SQLWriteString(cRenk, 30) + "', " +
                                " '" + SQLWriteString(cBeden, 30) + "', " +
                                " '" + SQLWriteDate(dGBitir) + "') "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                    If aCP(nCnt).dPlGonderiTarihi > #1/1/1950# Then

                        cSQL = cInsertHeader +
                                " values ('" + SQLWriteDate(aCP(nCnt).dPlGonderiTarihi) + "', " +
                                " 'CP', " +
                                " '" + SQLWriteString(cAciklama, 250) + "', " +
                                " 'YOLLA', " +
                                " '" + cDurum + "', " +
                                " '" + SQLWriteString(cPersonel, 30) + "', " +
                                " '" + SQLWriteString(cFirma, 250) + "', " +
                                " '" + SQLWriteString(cSiparisNo, 250) + "', " +
                                " '" + SQLWriteString(cMusteriNo, 250) + "', " +
                                " '" + SQLWriteString(cModelNo, 250) + "', " +
                                " '" + SQLWriteString(cDepartman, 250) + "', " +
                                " '" + SQLWriteString(cColor, 30) + "', " +
                                " '" + SQLWriteString(cKalipNo, 250) + "', " +
                                SQLWriteDecimal(nIhtiyac) + ", " +
                                SQLWriteDecimal(nIsemri) + ", " +
                                SQLWriteDecimal(nBaslanan) + ", " +
                                SQLWriteDecimal(nBiten) + ", " +
                                " '" + SQLWriteString(cTipi, 30) + "', " +
                                " '" + SQLWriteString(cFoyNo, 30) + "', " +
                                " '" + SQLWriteDate(dSevkTarihi) + "', " +
                                SQLWriteDecimal(nPairKey) + ", " +
                                " '" + SQLWriteString(cRenk, 30) + "', " +
                                " '" + SQLWriteString(cBeden, 30) + "', " +
                                " '" + SQLWriteDate(dGBasla) + "') "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                Next
            End If

            JustForLog("Masterplan build end")

            cSQL = "update " + cTable +
                    " set modelaciklama = (select top 1 aciklama from ymodel where modelno = " + cTable + ".modelno) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

            GetMasterPlanData = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "GetMasterPlanData", cSQL)
        End Try
    End Function

    Public Sub GetMTFGercekTarih(ByVal ConnYage As SqlConnection, ByVal cMTF As String, ByVal cStokno As String, Optional ByVal cRenk As String = "", Optional ByVal cBeden As String = "",
                                 Optional ByRef dBasla As Date = #1/1/1950#, Optional ByRef dBitir As Date = #1/1/1950#)
        Dim cSQL As String = ""

        Try
            ' Malzemede başlangıç tarihi ilk işemri veriliş tarihidir
            cSQL = "select a.tarih " +
                    " from isemri a, isemrilines b " +
                    " where a.isemrino = b.isemrino " +
                    " and b.malzemetakipno = '" + cMTF.Trim + "' " +
                    " and b.stokno = '" + cStokno.Trim + "' " +
                    IIf(cRenk.Trim = "", "", " and b.renk = '" + cRenk.Trim + "' ").ToString +
                    IIf(cBeden.Trim = "", "", " and b.beden = '" + cBeden.Trim + "' ").ToString +
                    " and a.tarih is not null " +
                    " order by a.tarih "

            dBasla = SQLGetDateConnected(cSQL, ConnYage)

            ' Malzeme bitiş tarihi en son giriş hareket tarihidir
            cSQL = "select a.fistarihi " +
                    " from stokfis a, stokfislines b " +
                    " where a.stokfisno = b.stokfisno " +
                    " and b.malzemetakipkodu = '" + cMTF.Trim + "' " +
                    " and b.stokno = '" + cStokno.Trim + "' " +
                    IIf(cRenk.Trim = "", "", " and b.renk = '" + cRenk.Trim + "' ").ToString +
                    IIf(cBeden.Trim = "", "", " and b.beden = '" + cBeden.Trim + "' ").ToString +
                    " and b.stokhareketkodu in ('04 Mlz Uretimden Giris','06 Tamirden Giris','02 Tedarikten Giris','05 Diger Giris','90 Trans/Rezv Giris') " +
                    " and a.fistarihi is not null " +
                    " order by a.fistarihi desc "

            dBitir = SQLGetDateConnected(cSQL, ConnYage)

            If dBitir = #1/1/1950# Then
                ' Malzeme bitiş tarihi en son transfer hareket tarihidir
                cSQL = "select tarih " +
                        " from stoktransfer " +
                        " where hedefmalzemetakipno = '" + cMTF.Trim + "' " +
                        " and stokno = '" + cStokno.Trim + "' " +
                        IIf(cRenk.Trim = "", "", " and renk = '" + cRenk.Trim + "' ").ToString +
                        IIf(cBeden.Trim = "", "", " and beden = '" + cBeden.Trim + "' ").ToString +
                        " and tarih is not null " +
                        " order by tarih desc "

                dBitir = SQLGetDateConnected(cSQL, ConnYage)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "GetMTFGercekTarih", cSQL)
        End Try
    End Sub

    Public Sub GetUTFGercekTarih(ByVal ConnYage As SqlConnection, ByVal cUTF As String, ByVal cDepartman As String, ByVal cModelNo As String,
                                Optional ByRef dBasla As Date = #1/1/1950#, Optional ByRef dBitir As Date = #1/1/1950#)
        Dim cSQL As String = ""

        Try
            ' UTFde ilk giriş tarihi, KESIM için ilk kumaş çıkış tarihidir
            If cDepartman.Trim = "KESIM" Then
                cSQL = "select a.fistarihi " +
                        " from stokfis a, stokfislines b " +
                        " where a.stokfisno = b.stokfisno " +
                        " and a.departman = '" + cDepartman.Trim + "' " +
                        " and b.uretimtakipno = '" + cUTF.Trim + "' " +
                        " and b.modelno = '" + cModelNo.Trim + "' " +
                        " and b.stokhareketkodu = '01 Uretime Cikis' " +
                        " and a.fistarihi is not null " +
                        " order by a.fistarihi "

                dBasla = SQLGetDateConnected(cSQL, ConnYage)
            Else
                cSQL = "select a.fistarihi " +
                        " from uretharfis a, uretharfislines b " +
                        " where a.uretfisno = b.uretfisno " +
                        " and a.girisdept = '" + cDepartman.Trim + "' " +
                        " and b.uretimtakipno = '" + cUTF.Trim + "' " +
                        " and b.modelno = '" + cModelNo.Trim + "' " +
                        " order by a.fistarihi "

                dBasla = SQLGetDateConnected(cSQL, ConnYage)
            End If
            ' UTF bitiş tarihi son çıkış fiş tarihidir
            cSQL = "select a.fistarihi " +
                    " from uretharfis a, uretharfislines b " +
                    " where a.uretfisno = b.uretfisno " +
                    " and a.cikisdept = '" + cDepartman.Trim + "' " +
                    " and b.uretimtakipno = '" + cUTF.Trim + "' " +
                    " and b.modelno = '" + cModelNo.Trim + "' " +
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
            cSQL = "select a.sevktar " +
                    " from sevkform a, sevkformlines b " +
                    " where a.sevkformno = b.sevkformno " +
                    " and b.sevkiyattakipno = '" + cSTF.Trim + "' " +
                    IIf(cSiparisNo.Trim = "", "", " and b.siparisno = '" + cSiparisNo.Trim + "' ").ToString +
                    IIf(cModelNo.Trim = "", "", " and b.modelno = '" + cModelNo.Trim + "' ").ToString +
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

            cSQL = "select kusurat " +
                        " from birim " +
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

    ' Maliyet Hesabı

    Public Function GenelGiderDagitimi() As SqlInt32

        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cFilter As String = ""
        Dim cSQL As String = ""
        Dim aGenelGider() As oGenelgider
        Dim nCnt As Integer
        Dim nGenelGider As Double
        Dim nGumruk As Double

        GenelGiderDagitimi = 0

        Try
            ConnYage = OpenConn()

            If CLng(GetSysParConnected("cekionay", ConnYage)) = 1 Then
                cFilter = " and a.ok = 'E' "
            End If

            cSQL = "update tanimlamadata " +
                    " set s_numeric = 0 " +
                    " where karttipi = 'siparismaliyet' " +
                    " and alan in ('gider1','gider2') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = -1
            ReDim aGenelGider(0)

            cSQL = "select w.yil, w.ay,  w.genelgidereur, w.gumrukgidereur, w.toplamsevk, v.siparisno, v.sevk "

            cSQL = cSQL +
                    " from (select yil, ay, " +
                            " genelgidereur = sum(coalesce(genelgidereur, 0)), " +
                            " gumrukgidereur = sum(coalesce(gumrukgidereur, 0)), " +
                            " toplamsevk = (select sum((b.koliend - b.kolibeg + 1) * c.adet) " +
                                            " From sevkform a, sevkformlines b, sevkformlinesrba c  " +
                                            " Where a.sevkformno = b.sevkformno " +
                                            " And b.sevkformno = c.sevkformno   " +
                                            " And b.ulineno = c.ulineno  " +
                                            " And datepart(yy, a.sevktar) = genelgider.yil  " +
                                            " And datepart(mm, a.sevktar) = genelgider.ay " +
                                            cFilter + "  )  " +
                            " From genelgider " +
                            " Group By yil, ay) w, "
            cSQL = cSQL +
                       " (select b.siparisno, yil = datepart(yy, a.sevktar), ay = datepart(mm, a.sevktar),  " +
                            " sevk = sum((b.koliend - b.kolibeg + 1) * c.adet) " +
                            " From sevkform a, sevkformlines b, sevkformlinesrba c  " +
                            " Where a.sevkformno = b.sevkformno " +
                            " And b.sevkformno = c.sevkformno  " +
                            " And b.ulineno = c.ulineno  " +
                            cFilter +
                            " group by b.siparisno, DatePart(yy, a.sevktar), DatePart(mm, a.sevktar))  v "
            cSQL = cSQL +
                    " where w.yil = v.yil " +
                    " And w.ay = v.ay " +
                    " And w.toplamsevk is not null " +
                    " And w.toplamsevk > 0 " +
                    " And v.sevk is Not null " +
                    " And v.sevk > 0 "

            If Not CheckExistsConnected(cSQL, ConnYage) Then
                ConnYage.Close()
                Exit Function
            End If

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                nCnt = nCnt + 1
                ReDim Preserve aGenelGider(nCnt)

                aGenelGider(nCnt).nYil = SQLReadDouble(oReader, "yil")
                aGenelGider(nCnt).nAy = SQLReadDouble(oReader, "ay")
                aGenelGider(nCnt).nGenelGiderEUR = SQLReadDouble(oReader, "genelgidereur")
                aGenelGider(nCnt).nGumrukGiderEUR = SQLReadDouble(oReader, "gumrukgidereur")
                aGenelGider(nCnt).nToplamSevk = SQLReadDouble(oReader, "toplamsevk")
                aGenelGider(nCnt).cSiparisNo = SQLReadString(oReader, "siparisno")
                aGenelGider(nCnt).nSevk = SQLReadDouble(oReader, "sevk")
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aGenelGider)

                nGenelGider = aGenelGider(nCnt).nGenelGiderEUR * aGenelGider(nCnt).nSevk / aGenelGider(nCnt).nToplamSevk
                nGumruk = aGenelGider(nCnt).nGumrukGiderEUR * aGenelGider(nCnt).nSevk / aGenelGider(nCnt).nToplamSevk

                cSQL = "select kayitno " +
                        " from tanimlamadata " +
                        " where karttipi = 'siparismaliyet' " +
                        " and kayitno = '" + aGenelGider(nCnt).cSiparisNo + "' " +
                        " and alan = 'gider1' " +
                        " and username = 'Admin' "

                If CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "update tanimlamadata " +
                            " set s_numeric = coalesce(s_numeric,0) + " + SQLWriteDecimal(nGenelGider) +
                            " where karttipi = 'siparismaliyet' " +
                            " and kayitno = '" + aGenelGider(nCnt).cSiparisNo + "' " +
                            " and alan = 'gider1' " +
                            " and username = 'Admin' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    cSQL = "insert tanimlamadata (kayitno, karttipi, alan, username, s_numeric) " +
                            " values ('" + aGenelGider(nCnt).cSiparisNo + "', " +
                            " 'siparismaliyet', " +
                            " 'gider1', " +
                            " 'Admin', " +
                            SQLWriteDecimal(nGenelGider) + " ) "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                cSQL = "select kayitno " +
                        " from tanimlamadata " +
                        " where karttipi = 'siparismaliyet' " +
                        " and kayitno = '" + aGenelGider(nCnt).cSiparisNo + "' " +
                        " and alan = 'gider2' " +
                        " and username = 'Admin' "

                If CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "update tanimlamadata " +
                            " set s_numeric = coalesce(s_numeric,0) + " + SQLWriteDecimal(nGumruk) +
                            " where karttipi = 'siparismaliyet' " +
                            " and kayitno = '" + aGenelGider(nCnt).cSiparisNo + "' " +
                            " and alan = 'gider2' " +
                            " and username = 'Admin' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    cSQL = "insert tanimlamadata (kayitno, karttipi, alan, username, s_numeric) " +
                            " values ('" + aGenelGider(nCnt).cSiparisNo + "', " +
                            " 'siparismaliyet', " +
                            " 'gider2', " +
                            " 'Admin', " +
                            SQLWriteDecimal(nGumruk) + " ) "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            ConnYage.Close()

            GenelGiderDagitimi = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "GenelGiderDagit", cSQL)
        End Try
    End Function

    Public Function STISonMaliyetInitialCleanup(cSipFilter As String) As Integer

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection

        STISonMaliyetInitialCleanup = 0

        Try
            ConnYage = OpenConn()

            cSipFilter = Replace(cSipFilter, "||", "'").Trim

            cSQL = "delete stisonmaliyet " +
               " where malzemetakipno in (" + cSipFilter + ") "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' kafa tablo
            cSQL = "delete stisonmaliyet1 " +
                " where siparisno  in (" + cSipFilter + ")  "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' ön maliyet (stok tipi bazinda)
            cSQL = "delete stisonmaliyet2 " +
                " where siparisno  in (" + cSipFilter + ")  "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' faturalar
            cSQL = "delete stisonmaliyet3 " +
                " where malzemetakipno in (" + cSipFilter + ") "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' firma / TL / EUR hakediş
            cSQL = "delete stisonmaliyet4 " +
                " where malzemetakipno in (" + cSipFilter + ") "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' firma / Fatura Dövizi hakediş
            cSQL = "delete stisonmaliyet5 " +
                " where malzemetakipno in (" + cSipFilter + ") "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' birim malzeme harcamaları
            cSQL = "delete stisonmaliyet6 " +
                " where malzemetakipno in (" + cSipFilter + ") "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' özet planlanan ve gerçeklesen, miktar ve tutarlar
            cSQL = "delete stisonmaliyet7 " +
                " where siparisno  in (" + cSipFilter + ")  "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

            STISonMaliyetInitialCleanup = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "STISonMaliyetInitialCleanup", cSQL)
        End Try
    End Function

    Public Function STISMPOpenRecords(cFilter As String) As Integer

        Dim nCnt As Integer = 0
        Dim cSQL As String = ""
        Dim nSipAdet As Double = 0
        Dim nSipTutar As Double = 0
        Dim nKesim As Double = 0
        Dim nDikim As Double = 0
        Dim nSevkAdet As Double = 0
        Dim nSevkTutar As Double = 0
        Dim dSevkTarih As Date = #1/1/1950#
        Dim aSiparis() As String

        STISMPOpenRecords = 0

        Try
            cFilter = Replace(cFilter, "||", "'").Trim

            cSQL = "select distinct a.kullanicisipno  " +
                   " from siparis a, sipmodel b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and a.kullanicisipno  is not null " +
                   " and a.kullanicisipno  <> '' " +
                   cFilter +
                   " order by a.kullanicisipno  "

            If Not CheckExists(cSQL) Then Exit Function

            aSiparis = SQLtoStringArray(cSQL)

            For nCnt = 0 To UBound(aSiparis)
                ' siparis kartı
                nSipAdet = GetSipAdet(CStr(aSiparis(nCnt)))
                nSipTutar = GetSipTutarDvz(CStr(aSiparis(nCnt)), , , , , "EUR")
                nKesim = GetSipUretimCikis(CStr(aSiparis(nCnt)), "KESIM", , , , True)
                nDikim = GetSipUretimCikis(CStr(aSiparis(nCnt)), "DIKIM", , , , True)
                nSevkAdet = GetSevkAdet(, , , , , CStr(aSiparis(nCnt)))
                nSevkTutar = GetSipSevkDVZTutar(CStr(aSiparis(nCnt)), , "EUR")
                dSevkTarih = GetSipSevkTarih2(CStr(aSiparis(nCnt)))

                ' mal sevk olmus fakat dikilmemis ise mali sevk adedi kadar dik
                If nSevkAdet > 0 Then
                    If nDikim = 0 Then
                        nDikim = nSevkAdet
                    End If
                End If

                ' mal dikilmis olmus fakat kesilmemis ise mali dikim adedi kadar kes
                If nDikim > 0 Then
                    If nKesim = 0 Then
                        nKesim = nDikim
                    End If
                End If

                cSQL = "select siparisno " +
                        " from stisonmaliyet1 " +
                        " where siparisno = '" + aSiparis(nCnt).ToString + "' "

                If Not CheckExists(cSQL) Then
                    cSQL = "set dateformat dmy " +
                            " insert stisonmaliyet1 (siparisno, sevktutar, sevkdoviz, siptutar, sipdoviz, " +
                                                    " sevktarih, siparisadet, kesimadet, dikimadet, sevkiyatadet) "
                    cSQL = cSQL +
                            " values ('" + aSiparis(nCnt).ToString + "', " +
                            SQLWriteDecimal(nSevkTutar) + ", " +
                            " 'EUR', " +
                            SQLWriteDecimal(nSipTutar) + ", " +
                            " 'EUR', "

                    cSQL = cSQL +
                            " '" + SQLWriteDate(dSevkTarih) + "', " +
                            SQLWriteDecimal(nSipAdet) + ", " +
                            SQLWriteDecimal(nKesim) + ", " +
                            SQLWriteDecimal(nDikim) + ", " +
                            SQLWriteDecimal(nSevkAdet) + ") "

                    ExecuteSQLCommand(cSQL)
                Else
                    cSQL = "set dateformat dmy " +
                            " update stisonmaliyet1 " +
                            " set sevktutar = " + SQLWriteDecimal(nSevkTutar) + ", " +
                            " sevkdoviz = 'EUR', " +
                            " siptutar = " + SQLWriteDecimal(nSipTutar) + ", " +
                            " sipdoviz = 'EUR', " +
                            " sevktarih = '" + SQLWriteDate(dSevkTarih) + "', "

                    cSQL = cSQL +
                            " siparisadet = " + SQLWriteDecimal(nSipAdet) + ", " +
                            " kesimadet = " + SQLWriteDecimal(nKesim) + ", " +
                            " dikimadet = " + SQLWriteDecimal(nDikim) + ", " +
                            " sevkiyatadet = " + SQLWriteDecimal(nSevkAdet) +
                            " where siparisno = '" + aSiparis(nCnt).ToString + "' "

                    ExecuteSQLCommand(cSQL)
                End If
            Next

            STISMPOpenRecords = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "STISMPOpenRecords", cSQL)
        End Try
    End Function

    Private Sub GetUretimIscilikFiyat(ConnYage As SqlConnection, ByVal cModelNo As String, ByVal cDepartman As String, ByRef nFiyat As Double, ByRef cDoviz As String,
                                 Optional ByVal cFirma As String = "", Optional ByVal cUTF As String = "", Optional ByVal cIsEmriNo As String = "",
                                 Optional ByVal dTarih As Date = #1/1/1950#, Optional cUretFisNo As String = "", Optional ByVal cSiparisNo As String = "")
        Dim cSQL As String = ""
        Dim cDepartmanTipi As String = ""
        Dim oReader As SqlDataReader

        nFiyat = 0
        cDoviz = "TL"

        Try
            ' UTF belliyse planlanandan gerçekleşene doğru fiyat aranıyor
            cSQL = "set dateformat dmy " +
                    " select top 1 b.fiyati, b.doviz " +
                    " from uretimisemri a, uretimisdetayi b " +
                    " where a.isemrino = b.isemrino " +
                    " And b.modelno = '" + cModelNo.Trim + "' " +
                    " and b.departman = '" + cDepartman.Trim + "' " +
                    " and b.fiyati is not null " +
                    " and b.fiyati <> 0 " +
                    IIf(cUTF.Trim = "", "", " and b.uretimtakipno = '" + cUTF.Trim + "' ").ToString +
                    IIf(cFirma.Trim = "", "", " and a.firma = '" + cFirma.Trim + "' ").ToString +
                    IIf(cIsEmriNo.Trim = "", "", " and a.isemrino = '" + cIsEmriNo.Trim + "' ").ToString +
                    IIf(cSiparisNo.Trim = "", "", " and b.uretimtakipno in (select uretimtakipno from sipmodel where modelno = b.modelno and siparisno = '" + cSiparisNo.Trim + "') ").ToString +
                    IIf(dTarih = #1/1/1950#, "", " and a.tarih <= '" + dTarih.ToString + "' ").ToString +
                    " order by a.tarih desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "fiyati")
                cDoviz = SQLReadString(oReader, "doviz")
            End If
            oReader.Close()

            If nFiyat <> 0 Then Exit Sub

            ' daha önce gerçekleşmiş üretim hareketlerinden fiyat alınır
            cSQL = "set dateformat dmy " +
                    " select top 1 b.fiyati, b.fiyatdoviz " +
                    " from uretharfis a, uretharfislines b " +
                    " where a.uretfisno = b.uretfisno " +
                    " and b.modelno = '" + cModelNo.Trim + "' " +
                    " and a.cikisdept = '" + cDepartman.Trim + "' " +
                    IIf(cUTF.Trim = "", "", " and b.uretimtakipno = '" + cUTF.Trim + "' ").ToString +
                    IIf(cFirma.Trim = "", "", " and a.cikisfirm_atl = '" + cFirma.Trim + "' ").ToString +
                    IIf(cIsEmriNo.Trim = "", "", " and b.isemrino = '" + cIsEmriNo.Trim + "' ").ToString +
                    IIf(cSiparisNo.Trim = "", "", " and b.uretimtakipno in (select uretimtakipno from sipmodel where modelno = b.modelno and siparisno = '" + cSiparisNo.Trim + "') ").ToString +
                    IIf(dTarih = #1/1/1950#, "", " and a.fistarihi <= '" + dTarih.ToString + "' ").ToString +
                    IIf(cUretFisNo.Trim = "", "", " and a.uretfisno <> '" + cUretFisNo.Trim + "' ").ToString +
                    " and b.fiyati is not null " +
                    " and b.fiyati <> 0 " +
                    " order by a.fistarihi desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "fiyati")
                cDoviz = SQLReadString(oReader, "fiyatdoviz")
            End If
            oReader.Close()

            If nFiyat <> 0 Then Exit Sub

            ' planlamadaki fiyat alınır, firma kontrol edilMEZ
            cSQL = "select top 1 fiyati, doviz " +
                    " from uretpllines " +
                    " where departman = '" + cDepartman.Trim + "' " +
                    " and uretimtakipno = '" + cUTF.Trim + "' " +
                    " and modelno = '" + cModelNo.Trim + "' " +
                    IIf(cSiparisNo.Trim = "", "", " and uretimtakipno in (select uretimtakipno from sipmodel where modelno = uretpllines.modelno and siparisno = '" + cSiparisNo.Trim + "') ").ToString +
                    " and fiyati is not null " +
                    " and fiyati <> 0 "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "fiyati")
                cDoviz = SQLReadString(oReader, "doviz")
            End If
            oReader.Close()

            If nFiyat <> 0 Then Exit Sub

            ' UTF belli değilse daha global planlamalara bakılır
            ' model fason fiyat tablosundaki fiyat alınır
            cSQL = "select departmantipi " +
                    " from departman " +
                    " where departman = '" + cDepartman.Trim + "' "

            cDepartmanTipi = SQLGetStringConnected(cSQL, ConnYage)

            cSQL = "set dateformat dmy " +
                    " select * " +
                    " from modelfasonfiyat " +
                    " where modelno = '" + cModelNo.Trim + "' " +
                    IIf(cFirma.Trim = "", "", " and firma = '" + cFirma.Trim + "' ").ToString +
                    IIf(dTarih = #1/1/1950#, "", " and onaytarihi <= '" + dTarih.ToString + "' ").ToString +
                    " order by onaytarihi desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                Select Case cDepartman.Trim
                    Case "KESIM+DIKIM", "KESİM                         "
                        nFiyat = SQLReadDouble(oReader, "kesimdikimfiyati")
                        cDoviz = SQLReadString(oReader, "kdfdoviz")
                        Exit Do
                    Case "KESIM", "KESİM                         "
                        nFiyat = SQLReadDouble(oReader, "kesimfiyati")
                        cDoviz = SQLReadString(oReader, "kfdoviz")
                        Exit Do
                    Case "DIKIM", "DİKİM                         "
                        nFiyat = SQLReadDouble(oReader, "dikimfiyati")
                        cDoviz = SQLReadString(oReader, "dfdoviz")
                        If nFiyat = 0 Then
                            nFiyat = SQLReadDouble(oReader, "kesimdikimfiyati")
                            cDoviz = SQLReadString(oReader, "kdfdoviz")
                        End If
                        Exit Do
                    Case "KAL.KONT./ PAKET", "KALİTE KONTROL", "PAKET"
                        nFiyat = SQLReadDouble(oReader, "kalitekontrolfiyati")
                        cDoviz = SQLReadString(oReader, "kkfdoviz")
                        Exit Do
                End Select
            Loop
            oReader.Close()

            If nFiyat <> 0 Then Exit Sub

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                Select Case cDepartmanTipi.Trim
                    Case "KESIM+DIKIM", "KESİM                         "
                        nFiyat = SQLReadDouble(oReader, "kesimdikimfiyati")
                        cDoviz = SQLReadString(oReader, "kdfdoviz")
                        Exit Do
                    Case "KESIM", "KESİM                         "
                        nFiyat = SQLReadDouble(oReader, "kesimfiyati")
                        cDoviz = SQLReadString(oReader, "kfdoviz")
                        Exit Do
                    Case "DIKIM", "DİKİM                         "
                        nFiyat = SQLReadDouble(oReader, "dikimfiyati")
                        cDoviz = SQLReadString(oReader, "dfdoviz")
                        If nFiyat = 0 Then
                            nFiyat = SQLReadDouble(oReader, "kesimdikimfiyati")
                            cDoviz = SQLReadString(oReader, "kdfdoviz")
                        End If
                        Exit Do
                    Case "KAL.KONT./ PAKET", "KALİTE KONTROL", "PAKET"
                        nFiyat = SQLReadDouble(oReader, "kalitekontrolfiyati")
                        cDoviz = SQLReadString(oReader, "kkfdoviz")
                        Exit Do
                End Select
            Loop
            oReader.Close()

            If nFiyat <> 0 Then Exit Sub

            ' model kartındaki üretim planlamasından fiyat alınır
            cSQL = "select top 1 iscilikfiyat, iscilikdoviz " +
                    " from modeluretim " +
                    " where departman = '" + cDepartman.Trim + "' " +
                    " and modelno = '" + cModelNo.Trim + "' " +
                    " and iscilikfiyat is not null " +
                    " and iscilikfiyat <> 0 "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "iscilikfiyat")
                cDoviz = SQLReadString(oReader, "iscilikdoviz")
            End If
            oReader.Close()

            If nFiyat <> 0 Then Exit Sub

            ' ön maliyetteki fiyatları alır
            cSQL = "select top 1 brmaliyet, doviz  " +
                    " from onmaliyetlines " +
                    " where onsipcode = '" + cModelNo.Trim + "' " +
                    " and (mlzadi = '" + cDepartman.Trim + "' or mlzcode = '" + cDepartman.Trim + "') " +
                    " and brmaliyet is not null " +
                    " and brmaliyet > 0 " +
                    " order by brmaliyet desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "brmaliyet")
                cDoviz = SQLReadString(oReader, "doviz")
            End If
            oReader.Close()

            If nFiyat <> 0 Then Exit Sub

            ' departman sabitlerindeki işçilik fiyatı alınır
            cSQL = "select top 1 iscilikfiyat, iscilikdoviz " +
                    " from departman " +
                    " where departman = '" + cDepartman.Trim + "' " +
                    " and iscilikfiyat is not null " +
                    " and iscilikfiyat <> 0 "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "iscilikfiyat")
                cDoviz = SQLReadString(oReader, "iscilikdoviz")
            End If
            oReader.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "GetUretimIscilikFiyat", cSQL)
        End Try
    End Sub

    Public Function STISonMaliyetUretim(ByVal cFilter As String) As Integer

        Dim cSQL As String = ""
        Dim nSiraNo As Double = 0
        Dim nEURTutar As Double = 0
        Dim nFiyat As Double = 0
        Dim cDoviz As String = ""
        Dim dTarih As Date = #1/1/1950#
        Dim nKur As Double = 0
        Dim nEURKur As Double = 0
        Dim nMiktar As Double = 0
        Dim nTLTutar As Double = 0
        Dim nOTutar As Double = 0
        Dim cODoviz As String = ""
        Dim nSonMlytsiraNo As Double = 0
        Dim cUTFFilter As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim aUretFis() As oUretFis
        Dim nCnt As Integer = -1
        Dim nCnt1 As Integer = -1
        Dim nIhtiyac As Double = 0
        Dim aSiparis() As String
        Dim aUIE() As oUIE

        STISonMaliyetUretim = 0

        Try
            cFilter = Replace(cFilter, "||", "'").Trim

            JustForLog("STISonMaliyetUretim start " + cFilter)

            ConnYage = OpenConn()

            cSQL = "select distinct b.uretimtakipno " +
               " from siparis a, sipmodel b, ymodel c " +
               " where a.kullanicisipno = b.siparisno " +
               " and b.modelno = c.modelno " +
               " and b.uretimtakipno is not null " +
               " and b.uretimtakipno <> '' " +
               cFilter +
               " order by b.uretimtakipno "

            If Not CheckExistsConnected(cSQL, ConnYage) Then
                ConnYage.Close()
                Exit Function
            End If

            cUTFFilter = SQLBuildFilterString2(ConnYage, cSQL)

            JustForLog(cUTFFilter)

            nSonMlytsiraNo = 100000

            cSQL = "select x.uretfisno, x.fistarihi, x.belgetarih, x.belgeno, x.faturatarihi, " +
               " x.faturano, x.cikisdept, x.cikisfirm_atl, y.uretimtakipno, y.modelno, " +
               " y.isemrino, y.fiyati, y.fiyatdoviz, " +
               " toplamadet = sum(coalesce(z.adet,0)) " +
               " from  uretharfis x, uretharfislines y, uretharrba z " +
               " where x.uretfisno = y.uretfisno " +
               " and y.uretfisno = z.uretfisno " +
               " and y.ulineno = z.ulineno " +
               " and y.uretimtakipno in (" + cUTFFilter + ") " +
               " group by x.uretfisno, x.fistarihi, x.belgetarih, x.belgeno, x.faturatarihi, " +
               " x.faturano, x.cikisdept, x.cikisfirm_atl, y.uretimtakipno, y.modelno, " +
               " y.isemrino, y.fiyati, y.fiyatdoviz" +
               " order by x.fistarihi, x.cikisdept, x.cikisfirm_atl, y.uretimtakipno "

            If Not CheckExistsConnected(cSQL, ConnYage) Then
                ConnYage.Close()
                Exit Function
            End If

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aUretFis(nCnt)
                aUretFis(nCnt).cBelgeNo = SQLReadString(oReader, "belgeno")
                aUretFis(nCnt).cFaturaNo = SQLReadString(oReader, "faturano")
                aUretFis(nCnt).dFaturaTarihi = SQLReadDate(oReader, "faturatarihi")
                aUretFis(nCnt).cUretFisNo = SQLReadString(oReader, "uretfisno")
                aUretFis(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aUretFis(nCnt).cDepartman = SQLReadString(oReader, "cikisdept")
                aUretFis(nCnt).cFirma = SQLReadString(oReader, "cikisfirm_atl")
                aUretFis(nCnt).cUTF = SQLReadString(oReader, "uretimtakipno")
                aUretFis(nCnt).cIsemriNo = SQLReadString(oReader, "isemrino")
                If SQLReadDate(oReader, "fistarihi") <= CDate("01.01.1950") Then
                    aUretFis(nCnt).dTarih = Today
                Else
                    aUretFis(nCnt).dTarih = SQLReadDate(oReader, "fistarihi")
                End If
                aUretFis(nCnt).nMiktar = SQLReadDouble(oReader, "toplamadet")
                aUretFis(nCnt).nFiyat = SQLReadDouble(oReader, "fiyati")
                If SQLReadString(oReader, "fiyatdoviz") = "" Then
                    aUretFis(nCnt).cDoviz = "TL"
                Else
                    aUretFis(nCnt).cDoviz = SQLReadString(oReader, "fiyatdoviz")
                End If
                aUretFis(nCnt).nKur = 0
                aUretFis(nCnt).nEURKUR = 0
                aUretFis(nCnt).nEURFiyat = 0
                aUretFis(nCnt).nTLFiyat = 0
            Loop
            oReader.Close()

            ' fiyatları tamamla
            For nCnt = 0 To UBound(aUretFis)
                If aUretFis(nCnt).nFiyat = 0 Then
                    GetUretimIscilikFiyat(ConnYage, aUretFis(nCnt).cModelNo, aUretFis(nCnt).cDepartman, aUretFis(nCnt).nFiyat, aUretFis(nCnt).cDoviz,
                                      aUretFis(nCnt).cFirma, aUretFis(nCnt).cUTF, aUretFis(nCnt).cIsemriNo, aUretFis(nCnt).dTarih, aUretFis(nCnt).cUretFisNo)
                End If

                If aUretFis(nCnt).cDoviz = "" Then
                    aUretFis(nCnt).cDoviz = "TL"
                End If

                JustForLog(aUretFis(nCnt).cUTF + " " + aUretFis(nCnt).cDepartman + " " + aUretFis(nCnt).nMiktar.ToString + " " + aUretFis(nCnt).nFiyat.ToString + " " + aUretFis(nCnt).cDoviz)

                If aUretFis(nCnt).nFiyat <> 0 Then

                    aUretFis(nCnt).nTutar = aUretFis(nCnt).nFiyat * aUretFis(nCnt).nMiktar

                    aUretFis(nCnt).nEURKUR = GetKur("EUR", aUretFis(nCnt).dTarih, ConnYage)
                    aUretFis(nCnt).nKur = GetKur(aUretFis(nCnt).cDoviz, aUretFis(nCnt).dTarih, ConnYage)

                    If aUretFis(nCnt).nEURKUR <> 0 Then
                        aUretFis(nCnt).nEURFiyat = aUretFis(nCnt).nKur / aUretFis(nCnt).nEURKUR * aUretFis(nCnt).nFiyat    ' EURO fiyat
                        aUretFis(nCnt).nEURTutar = aUretFis(nCnt).nEURFiyat * aUretFis(nCnt).nMiktar
                    End If

                    If aUretFis(nCnt).nKur <> 0 Then
                        aUretFis(nCnt).nTLFiyat = aUretFis(nCnt).nKur * aUretFis(nCnt).nFiyat    ' TL fiyat
                        aUretFis(nCnt).nTLTutar = aUretFis(nCnt).nTLFiyat * aUretFis(nCnt).nMiktar
                    End If

                    JustForLog(aUretFis(nCnt).cUTF + " " + aUretFis(nCnt).cDepartman + " " + aUretFis(nCnt).nMiktar.ToString + " EUR Tutar " + aUretFis(nCnt).nEURTutar.ToString + " EUR Kur " + aUretFis(nCnt).nEURKUR.ToString + " TL Tutar " + aUretFis(nCnt).nTLTutar.ToString)

                    cSQL = "select sirano " +
                           " from stisonmaliyet " +
                           " where malzemetakipno = '" + aUretFis(nCnt).cUTF + "' " +
                           " and stoktipi = '_" + aUretFis(nCnt).cDepartman + "' "

                    If CheckExistsConnected(cSQL, ConnYage) Then

                        nSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

                        cSQL = "update stisonmaliyet " +
                               " set ihtiyac = 0, " +
                               " karsilanan = coalesce(karsilanan,0) + " + SQLWriteDecimal(aUretFis(nCnt).nMiktar) + ", " +
                               " eurtutar = coalesce(eurtutar,0) + " + SQLWriteDecimal(aUretFis(nCnt).nEURTutar) +
                               " where sirano = " + SQLWriteDecimal(nSiraNo)

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    Else
                        cSQL = "select sira " +
                                " from departman " +
                                " where departman = '" + aUretFis(nCnt).cDepartman + "' "

                        nSonMlytsiraNo = nSonMlytsiraNo + SQLGetDoubleConnected(cSQL, ConnYage)

                        cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, birim, ihtiyac, " +
                                " karsilanan, eurtutar, siralama) "

                        cSQL = cSQL +
                                " values ('" + aUretFis(nCnt).cUTF + "', " +
                                " 'ISCILIK', " +
                                " '_" + aUretFis(nCnt).cDepartman + "', " +
                                " 'AD', " +
                                " 0, "

                        cSQL = cSQL +
                                SQLWriteDecimal(aUretFis(nCnt).nMiktar) + ", " +
                                SQLWriteDecimal(aUretFis(nCnt).nEURTutar) + ", " +
                                SQLWriteDecimal(nSonMlytsiraNo) + " ) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                    If aUretFis(nCnt).nEURKUR <> 0 Then

                        cSQL = "update stisonmaliyet1 " +
                                " set soniscilikkuru = " + SQLWriteDecimal(aUretFis(nCnt).nEURKUR) +
                                " where siparisno = '" + aUretFis(nCnt).cUTF + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                        cSQL = "update stisonmaliyet1 " +
                                " set toplamiscilikmiktarkur = coalesce(toplamiscilikmiktarkur,0) + " + SQLWriteDecimal(aUretFis(nCnt).nEURTutar) +
                                " where siparisno = '" + aUretFis(nCnt).cUTF + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                        cSQL = "update stisonmaliyet1 " +
                                " set toplamiscilikmiktar = coalesce(toplamiscilikmiktar,0) + " + SQLWriteDecimal(aUretFis(nCnt).nMiktar) +
                                " where siparisno = '" + aUretFis(nCnt).cUTF + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                    stisonmaliyet3(ConnYage, aUretFis(nCnt).cUretFisNo, aUretFis(nCnt).cUTF, aUretFis(nCnt).dTarih,
                                   aUretFis(nCnt).cBelgeNo, aUretFis(nCnt).dFaturaTarihi, aUretFis(nCnt).cFaturaNo,
                                   aUretFis(nCnt).cDepartman, aUretFis(nCnt).cFirma, "ISCILIK",
                                   aUretFis(nCnt).nEURKUR, aUretFis(nCnt).nTLTutar, aUretFis(nCnt).nEURTutar, 0, 0, aUretFis(nCnt).nTutar, 0, aUretFis(nCnt).cDoviz, aUretFis(nCnt).nMiktar, 0)

                    stisonmaliyet4(ConnYage, aUretFis(nCnt).cUTF, aUretFis(nCnt).cFirma, aUretFis(nCnt).nTLTutar, aUretFis(nCnt).nEURTutar)

                    stisonmaliyet5(ConnYage, aUretFis(nCnt).cUTF, aUretFis(nCnt).cFirma, aUretFis(nCnt).nTutar, aUretFis(nCnt).cDoviz)
                End If

                JustForLog(aUretFis(nCnt).cUTF + " " + aUretFis(nCnt).cDepartman + " OK ")

            Next

            cSQL = "select distinct a.kullanicisipno " +
                   " from siparis a, sipmodel b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and b.uretimtakipno is not null " +
                   " and b.uretimtakipno <> '' " +
                   cFilter +
                   " order by a.kullanicisipno "

            If Not CheckExistsConnected(cSQL, ConnYage) Then
                ConnYage.Close()
                Exit Function
            End If

            aSiparis = SQLtoStringArrayConnected(cSQL, ConnYage)

            nCnt1 = -1
            For nCnt = 0 To UBound(aSiparis)
                cSQL = "select distinct stoktipi " +
                        " from stisonmaliyet " +
                        " where malzemetakipno = '" + aSiparis(nCnt) + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read
                    If Mid(SQLReadString(oReader, "stoktipi"), 1, 1) = "_" Then
                        nCnt1 = nCnt1 + 1
                        ReDim Preserve aUIE(nCnt1)
                        aUIE(nCnt1).cSiparisNo = aSiparis(nCnt)
                        aUIE(nCnt1).cStokTipi = SQLReadString(oReader, "stoktipi")
                        aUIE(nCnt1).cDepartman = Mid(SQLReadString(oReader, "stoktipi"), 2, 30).Trim
                    End If
                Loop
                oReader.Close()
            Next

            If nCnt1 <> -1 Then
                For nCnt = 0 To UBound(aUIE)
                    cSQL = "select sum(coalesce(c.adet,0)) " +
                            " from uretimisemri a, uretimisdetayi b, uretimisrba c " +
                            " where a.isemrino = b.isemrino " +
                            " and b.isemrino = c.isemrino " +
                            " and b.ulineno = c.ulineno " +
                            " and c.uretimtakipno in (select uretimtakipno " +
                                                    " from sipmodel " +
                                                    " where siparisno = '" + aUIE(nCnt).cSiparisNo + "') " +
                            " and a.departman like '%" + aUIE(nCnt).cDepartman + "%' "

                    nIhtiyac = SQLGetDoubleConnected(cSQL, ConnYage)

                    If nIhtiyac > 0 Then
                        cSQL = "update stisonmaliyet " +
                                " set ihtiyac = " + SQLWriteDecimal(nIhtiyac) +
                                " where malzemetakipno = '" + aUIE(nCnt).cSiparisNo + "' " +
                                " and stoktipi = '" + aUIE(nCnt).cStokTipi + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                Next
            End If

            cSQL = "update stisonmaliyet1 " +
                    " set soniscilikkuru = coalesce(toplamiscilikmiktarkur,0) / coalesce(toplamiscilikmiktar,0) " +
                    " where toplamiscilikmiktar is not null " +
                    " and toplamiscilikmiktar <> 0 "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

            STISonMaliyetUretim = 1

            JustForLog("STISonMaliyetUretim END")

        Catch ex As Exception
            ErrDisp(ex.Message, "STISonMaliyetUretim", cSQL)
        End Try
    End Function

    Public Function STISonMaliyet7Create(cFilter As String) As Integer

        Dim cSQL As String = ""
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim aSiparis() As String
        Dim nKesim As Double = 0
        Dim cStokTipi As String = ""
        Dim nPlAdet As Double = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim aStiSonMaliyet7() As oStiSonMaliyet7

        STISonMaliyet7Create = 0

        Try
            cFilter = Replace(cFilter, "||", "'").Trim

            JustForLog("STISonMaliyet7Create start " + cFilter)

            ConnYage = OpenConn()

            cSQL = "select distinct a.kullanicisipno  " +
                   " from siparis a, sipmodel b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and a.kullanicisipno  is not null " +
                   " and a.kullanicisipno  <> '' " +
                   cFilter +
                   " order by a.kullanicisipno  "

            If Not CheckExistsConnected(cSQL, ConnYage) Then
                ConnYage.Close()
                Exit Function
            End If

            aSiparis = SQLtoStringArrayConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSiparis)

                nCnt1 = -1
                ReDim aStiSonMaliyet7(0)

                ' gerçekleşen
                cSQL = "select stoktipi, birim, karsilanan, eurtutar, " +
                        " anastokgrubu = (select top 1 anastokgrubu " +
                                        " from stoktipi " +
                                        " where kod = stisonmaliyet.stoktipi " +
                                        " and anastokgrubu = 'AKSESUAR') " +
                        " from stisonmaliyet " +
                        " where malzemetakipno = '" + aSiparis(nCnt).Trim + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read

                    cStokTipi = SQLReadString(oReader, "stoktipi")

                    If Mid(cStokTipi, 1, 1) = "_" Then
                        ' işçilik
                        nCnt1 = nCnt1 + 1
                        ReDim Preserve aStiSonMaliyet7(nCnt1)

                        aStiSonMaliyet7(nCnt1).cSiparisNo = aSiparis(nCnt)
                        aStiSonMaliyet7(nCnt1).cTipi = "ISCILIK"
                        aStiSonMaliyet7(nCnt1).cBirim = SQLReadString(oReader, "birim")
                        aStiSonMaliyet7(nCnt1).nMiktar = 0
                        aStiSonMaliyet7(nCnt1).nTutar = SQLReadDouble(oReader, "eurtutar")
                        aStiSonMaliyet7(nCnt1).nPlMiktar = 0
                        aStiSonMaliyet7(nCnt1).nPlTutar = 0
                    ElseIf cStokTipi = "KUMAS" Or
                            cStokTipi = "TELA" Or
                            cStokTipi = "ASTAR" Or
                            cStokTipi = "BIYE" Or
                            cStokTipi = "GARNI" Or
                            cStokTipi = "GENEL GIDER" Then
                        ' kumaş, tela, astar, genel gider
                        If cStokTipi = "GARNI" Then
                            cStokTipi = "KUMAS"
                        End If

                        nCnt1 = nCnt1 + 1
                        ReDim Preserve aStiSonMaliyet7(nCnt1)

                        aStiSonMaliyet7(nCnt1).cSiparisNo = aSiparis(nCnt)
                        aStiSonMaliyet7(nCnt1).cTipi = SQLWriteString(cStokTipi, 30)
                        aStiSonMaliyet7(nCnt1).cBirim = SQLReadString(oReader, "birim")
                        aStiSonMaliyet7(nCnt1).nMiktar = SQLReadDouble(oReader, "karsilanan")
                        aStiSonMaliyet7(nCnt1).nTutar = SQLReadDouble(oReader, "eurtutar")
                        aStiSonMaliyet7(nCnt1).nPlMiktar = 0
                        aStiSonMaliyet7(nCnt1).nPlTutar = 0
                    Else
                        If SQLReadString(oReader, "anastokgrubu") <> "" Then
                            nCnt1 = nCnt1 + 1
                            ReDim Preserve aStiSonMaliyet7(nCnt1)

                            aStiSonMaliyet7(nCnt1).cSiparisNo = aSiparis(nCnt)
                            aStiSonMaliyet7(nCnt1).cTipi = "AKSESUAR"
                            aStiSonMaliyet7(nCnt1).cBirim = ""
                            aStiSonMaliyet7(nCnt1).nMiktar = SQLReadDouble(oReader, "karsilanan")
                            aStiSonMaliyet7(nCnt1).nTutar = SQLReadDouble(oReader, "eurtutar")
                            aStiSonMaliyet7(nCnt1).nPlMiktar = 0
                            aStiSonMaliyet7(nCnt1).nPlTutar = 0
                        Else
                            nCnt1 = nCnt1 + 1
                            ReDim Preserve aStiSonMaliyet7(nCnt1)

                            aStiSonMaliyet7(nCnt1).cSiparisNo = aSiparis(nCnt)
                            aStiSonMaliyet7(nCnt1).cTipi = "DIGER"
                            aStiSonMaliyet7(nCnt1).cBirim = ""
                            aStiSonMaliyet7(nCnt1).nMiktar = SQLReadDouble(oReader, "karsilanan")
                            aStiSonMaliyet7(nCnt1).nTutar = SQLReadDouble(oReader, "eurtutar")
                            aStiSonMaliyet7(nCnt1).nPlMiktar = 0
                            aStiSonMaliyet7(nCnt1).nPlTutar = 0
                        End If
                    End If
                Loop
                oReader.Close()

                ' planlanan
                cSQL = "select mlzcode, birim, miktar, eurtutar " +
                        " from stisonmaliyet2 " +
                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read

                    cStokTipi = SQLReadString(oReader, "mlzcode")

                    If cStokTipi = "AKSESUAR" Or
                        cStokTipi = "ISCILIK" Or
                        cStokTipi = "KUMAS" Or
                        cStokTipi = "TELA" Or
                        cStokTipi = "ASTAR" Or
                        cStokTipi = "BIYE" Or
                        cStokTipi = "GARNI" Or
                        cStokTipi = "GENEL GIDER" Then
                        ' aksesuar, işçilik, kumaş, tela, astar, biye, genel gider
                        If cStokTipi = "GARNI" Then
                            cStokTipi = "KUMAS"
                        End If

                        nCnt1 = nCnt1 + 1
                        ReDim Preserve aStiSonMaliyet7(nCnt1)

                        aStiSonMaliyet7(nCnt1).cSiparisNo = aSiparis(nCnt)
                        aStiSonMaliyet7(nCnt1).cTipi = Mid(cStokTipi, 1, 30)
                        aStiSonMaliyet7(nCnt1).cBirim = SQLReadString(oReader, "birim")
                        aStiSonMaliyet7(nCnt1).nMiktar = 0
                        aStiSonMaliyet7(nCnt1).nTutar = 0
                        aStiSonMaliyet7(nCnt1).nPlMiktar = SQLReadDouble(oReader, "miktar")
                        aStiSonMaliyet7(nCnt1).nPlTutar = SQLReadDouble(oReader, "eurtutar")
                    Else
                        nCnt1 = nCnt1 + 1
                        ReDim Preserve aStiSonMaliyet7(nCnt1)

                        aStiSonMaliyet7(nCnt1).cSiparisNo = aSiparis(nCnt)
                        aStiSonMaliyet7(nCnt1).cTipi = "DIGER"
                        aStiSonMaliyet7(nCnt1).cBirim = SQLReadString(oReader, "birim")
                        aStiSonMaliyet7(nCnt1).nMiktar = 0
                        aStiSonMaliyet7(nCnt1).nTutar = 0
                        aStiSonMaliyet7(nCnt1).nPlMiktar = SQLReadDouble(oReader, "miktar")
                        aStiSonMaliyet7(nCnt1).nPlTutar = SQLReadDouble(oReader, "eurtutar")
                    End If
                Loop
                oReader.Close()

                If nCnt <> -1 Then
                    For nCnt1 = 0 To UBound(aStiSonMaliyet7)
                        stisonmaliyet7(ConnYage, aStiSonMaliyet7(nCnt1).cSiparisNo,
                                       aStiSonMaliyet7(nCnt1).cTipi, aStiSonMaliyet7(nCnt1).cBirim,
                                       aStiSonMaliyet7(nCnt1).nMiktar, aStiSonMaliyet7(nCnt1).nTutar,
                                       aStiSonMaliyet7(nCnt1).nPlMiktar, aStiSonMaliyet7(nCnt1).nPlTutar)
                    Next
                End If

                nPlAdet = GetSipUretIsemri(ConnYage, aSiparis(nCnt).Trim, "KESIM", , , , True)

                cSQL = "select kesimadet " +
                        " from stisonmaliyet1 " +
                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                nKesim = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = "update stisonmaliyet7 " +
                        " set birim = 'AD', " +
                        " plmiktar = " + SQLWriteDecimal(nPlAdet) + ", " +
                        " germiktar = " + SQLWriteDecimal(nKesim) +
                        " where siparisno = '" + aSiparis(nCnt).Trim + "' " +
                        " and mlzcode = 'ISCILIK' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "update stisonmaliyet7 " +
                        " set birim = 'AD', " +
                        " plmiktar = 1, " +
                        " germiktar = 1 " +
                        " where siparisno = '" + aSiparis(nCnt).Trim + "' " +
                        " and mlzcode in ('AKSESUAR','DIGER','GENEL GIDER','GUMRUK MASRAFI','MUSTERI REKLAMASYON','KESTIGIMIZ REKLAMASYON') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next

            ConnYage.Close()

            STISonMaliyet7Create = 1

            JustForLog("STISonMaliyet7Create END")

        Catch ex As Exception
            ErrDisp(ex.Message, "STISonMaliyet7Create", cSQL)
        End Try
    End Function

    Private Sub stisonmaliyet7(ConnYage As SqlConnection, ByRef cSiparisNo As String, ByRef cMlzCode As String, Optional ByRef cBirim As String = "",
                            Optional ByRef nGerMiktar As Double = 0, Optional ByRef nGerTutar As Double = 0,
                            Optional ByRef nPlMiktar As Double = 0, Optional ByRef nPlTutar As Double = 0)
        ' stisonmaliyet7
        ' siparis bazinda
        ' özet planlanan ve gerçeklesen, miktar ve tutarlar
        ' stisonmaliyet, stisonmaliyet2 tablolarindan bilgi cekiyor
        Dim cSQL As String = ""

        Try
            cSQL = "select siparisno " +
                " from stisonmaliyet7 " +
                " where siparisno = '" + cSiparisNo.Trim + "' " +
                " and mlzcode = '" + cMlzCode.Trim + "' "

            If CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "update stisonmaliyet7 " +
                        " set birim = '" + cBirim.Trim + "', " +
                        " germiktar = coalesce(germiktar,0) + " + SQLWriteDecimal(nGerMiktar) + ", " +
                        " gertutar  = coalesce(gertutar,0) + " + SQLWriteDecimal(nGerTutar) + ", " +
                        " plmiktar  = coalesce(plmiktar,0) + " + SQLWriteDecimal(nPlMiktar) + ", " +
                        " pltutar   = coalesce(pltutar,0) + " + SQLWriteDecimal(nPlTutar) +
                        " where siparisno = '" + cSiparisNo.Trim + "' " +
                        " and mlzcode = '" + cMlzCode.Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "insert stisonmaliyet7 (siparisno, mlzcode, birim, germiktar, gertutar, plmiktar, pltutar) " +
                        " values ('" + cSiparisNo.Trim + "', " +
                        " '" + cMlzCode.Trim + "', " +
                        " '" + cBirim.Trim + "', " +
                        SQLWriteDecimal(nGerMiktar) + ", " +
                        SQLWriteDecimal(nGerTutar) + ", " +
                        SQLWriteDecimal(nPlMiktar) + ", " +
                        SQLWriteDecimal(nPlTutar) + ") "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "stisonmaliyet7", cSQL)
        End Try
    End Sub

    Private Sub stisonmaliyet6(ConnYage As SqlConnection, ByRef cMalzemeTakipNo As String, ByRef cStokNo As String, ByRef nUretimeCikan As Double, ByRef nUretimdenIade As Double)
        ' birim harcamalar
        Dim cSQL As String = ""

        Try
            cSQL = "select * " +
                   " from stisonmaliyet6 " +
                   " where stokno = '" + cStokNo.Trim + "' " +
                   " and malzemetakipno = '" + cMalzemeTakipNo.Trim + "' "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "insert stisonmaliyet6 (malzemetakipno, stokno, uretimecikan, uretimdeniade) " +
                       " values ('" + SQLWriteString(cMalzemeTakipNo, 30) + "', " +
                       " '" + SQLWriteString(cStokNo, 30) + "', " +
                       SQLWriteDecimal(nUretimeCikan) + ", " +
                       SQLWriteDecimal(nUretimdenIade) + ") "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "update stisonmaliyet6 " +
                       " set uretimecikan = coalesce(uretimecikan,0) + " + SQLWriteDecimal(nUretimeCikan) + ", " +
                       " uretimdeniade = coalesce(uretimdeniade,0) + " + SQLWriteDecimal(nUretimdenIade) + " " +
                       " where malzemetakipno = '" + cMalzemeTakipNo.Trim + "' " +
                       " and stokno = '" + cStokNo.Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "stisonmaliyet6", cSQL)
        End Try
    End Sub

    Private Sub stisonmaliyet5(ConnYage As SqlConnection, ByVal cMalzemeTakipNo As String, ByVal cFirma As String, ByVal nTutar As Double, ByVal cDoviz As String)
        ' Orjinal Dovizden Hakediş
        Dim cSQL As String = ""

        Try
            cSQL = "select * " +
                   " from stisonmaliyet5 " +
                   " where firma = '" + cFirma.Trim + "' " +
                   " and malzemetakipno = '" + cMalzemeTakipNo.Trim + "' " +
                   " and doviz = '" + cDoviz.Trim + "' "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "insert stisonmaliyet5 (malzemetakipno, firma, tutar, doviz) " +
                       " values ('" + SQLWriteString(cMalzemeTakipNo.Trim, 30) + "', " +
                       " '" + SQLWriteString(cFirma.Trim, 30) + "', " +
                       SQLWriteDecimal(nTutar) + ", " +
                       " '" + SQLWriteString(cDoviz.Trim, 3) + "') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "update stisonmaliyet5 " +
                       " set tutar = coalesce(tutar,0) + " + SQLWriteDecimal(nTutar) + " " +
                       " where firma = '" + cFirma.Trim + "' " +
                       " and malzemetakipno = '" + cMalzemeTakipNo.Trim + "' " +
                       " and doviz = '" + cDoviz.Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "stisonmaliyet5", cSQL)
        End Try
    End Sub

    Private Sub stisonmaliyet4(ConnYage As SqlConnection, ByVal cMalzemeTakipNo As String, ByVal cFirma As String, ByVal nTLTutar As Double, ByVal nEURTutar As Double)
        ' TL ve EUR a çevrilmiş hakediş
        Dim cSQL As String = ""

        Try
            cSQL = "select * " +
               " from stisonmaliyet4 " +
               " where firma = '" + cFirma.Trim + "' " +
               " and malzemetakipno = '" + Trim(cMalzemeTakipNo) + "' "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "insert stisonmaliyet4 (malzemetakipno, firma, tltutar, eurtutar) " +
                       " values ('" + SQLWriteString(cMalzemeTakipNo.Trim, 30) + "', " +
                       " '" + SQLWriteString(cFirma.Trim, 30) + "', " +
                       SQLWriteDecimal(nTLTutar) + ", " +
                       SQLWriteDecimal(nEURTutar) + " ) "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "update stisonmaliyet4 " +
                       " set tltutar = coalesce(tltutar,0) + " + SQLWriteDecimal(nTLTutar) + ", " +
                       " eurtutar = coalesce(eurtutar,0) + " + SQLWriteDecimal(nEURTutar) +
                       " where firma = '" + cFirma.Trim + "' " +
                       " and malzemetakipno = '" + cMalzemeTakipNo.Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "stisonmaliyet4", cSQL)
        End Try
    End Sub


    Private Sub stisonmaliyet3(ConnYage As SqlConnection, cFisNo As String, cMalzemeTakipNo As String, dIrsaliyeTarihi As Date, cIrsaliyeNo As String, dFaturaTarihi As Date,
                           cFaturaNo As String, cDepartman As String, cFirma As String, cAciklama As String, nEURKur As Double,
                           nGirisTLTutar As Double, nGirisEurTutar As Double, nCikisTLTutar As Double, nCikisEurTutar As Double,
                           nOrjinalGirisTutar As Double, nOrjinalCikisTutar As Double, cOrjinalDoviz As String, nGirenMiktar As Double, nCikanMiktar As Double)
        ' Faturalar
        Dim cSQL As String = ""
        Dim cBuffer As String = ""

        Try
            cSQL = "select * " +
                   " from stisonmaliyet3 " +
                   " where fisno = '" + cFisNo.Trim + "' " +
                   " and malzemetakipno = '" + cMalzemeTakipNo.Trim + "' "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "set dateformat dmy " +
                       " insert stisonmaliyet3 (fisno, malzemetakipno, irsaliyetarihi, irsaliyeno, faturatarihi, " +
                       " faturano, departman, firma, aciklama, eurkur, " +
                       " orjinalgiristutar, orjinalcikistutar, orjinaldoviz, giristltutar, giriseurtutar, " +
                       " cikistltutar, cikiseurtutar, GirenMiktar, CikanMiktar) "

                cSQL = cSQL +
                       " values ('" + SQLWriteString(cFisNo) + "', " +
                       " '" + SQLWriteString(cMalzemeTakipNo, 30) + "', " +
                       " '" + SQLWriteDate(dIrsaliyeTarihi) + "', " +
                       " '" + SQLWriteString(cIrsaliyeNo, 30) + "', " +
                       " '" + SQLWriteDate(dFaturaTarihi) + "', "

                cSQL = cSQL +
                       " '" + SQLWriteString(cFaturaNo, 30) + "', " +
                       " '" + SQLWriteString(cDepartman, 30) + "', " +
                       " '" + SQLWriteString(cFirma, 30) + "', " +
                       " '" + SQLWriteString(cAciklama, 500) + "', " +
                       SQLWriteDecimal(nEURKur) + ", "

                cSQL = cSQL +
                       SQLWriteDecimal(nOrjinalGirisTutar) + ", " +
                       SQLWriteDecimal(nOrjinalCikisTutar) + ", " +
                       " '" + SQLWriteString(cOrjinalDoviz, 3) + "', " +
                       SQLWriteDecimal(nGirisTLTutar) + ", " +
                       SQLWriteDecimal(nGirisEurTutar) + ", "

                cSQL = cSQL +
                       SQLWriteDecimal(nCikisTLTutar) + ", " +
                       SQLWriteDecimal(nCikisEurTutar) + ", " +
                       SQLWriteDecimal(nGirenMiktar) + ", " +
                       SQLWriteDecimal(nCikanMiktar) + ") "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "select aciklama " +
                       " from stisonmaliyet3 " +
                       " where fisno = '" + cFisNo.Trim + "' " +
                       " and malzemetakipno = '" + cMalzemeTakipNo.Trim + "' "

                cBuffer = SQLGetStringConnected(cSQL, ConnYage)

                If cBuffer = "" Then
                    cBuffer = cAciklama
                Else
                    If InStr(cBuffer, cAciklama) = 0 Then
                        cBuffer = cBuffer + "," + cAciklama
                    End If
                End If

                cSQL = "update stisonmaliyet3 " +
                       " set giristltutar = coalesce(giristltutar,0) + " + SQLWriteDecimal(nGirisTLTutar) + ", " +
                       " giriseurtutar = coalesce(giriseurtutar,0) + " + SQLWriteDecimal(nGirisEurTutar) + ", " +
                       " cikistltutar = coalesce(cikistltutar,0) + " + SQLWriteDecimal(nCikisTLTutar) + ", " +
                       " cikiseurtutar = coalesce(cikiseurtutar,0) + " + SQLWriteDecimal(nCikisEurTutar) + ", " +
                       " orjinalgiristutar = coalesce(orjinalgiristutar,0) + " + SQLWriteDecimal(nOrjinalGirisTutar) + ", " +
                       " orjinalcikistutar = coalesce(orjinalcikistutar,0) + " + SQLWriteDecimal(nOrjinalCikisTutar) + ", " +
                       " girenmiktar = coalesce(girenmiktar,0) + " + SQLWriteDecimal(nGirenMiktar) + ", " +
                       " cikanmiktar = coalesce(cikanmiktar,0) + " + SQLWriteDecimal(nCikanMiktar) + ", " +
                       " orjinaldoviz = '" + SQLWriteString(cOrjinalDoviz, 3) + "', " +
                       " aciklama = '" + SQLWriteString(cBuffer, 250) + "' " +
                       " where fisno = '" + cFisNo.Trim + "' " +
                       " and malzemetakipno = '" + cMalzemeTakipNo.Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "stisonmaliyet3", cSQL)
        End Try
    End Sub

    Public Sub GetSonStokGiristenFiyat(ConnYage As SqlConnection, ByVal cStokNo As String, ByVal cRenk As String, ByRef nFiyat As Double, ByRef cDoviz As String,
                                       Optional ByRef dTarih As Date = #1/1/1950#, Optional cBeden As String = "", Optional cMTF As String = "")

        Dim cSQL As String = ""
        Dim oReader As SqlDataReader

        Try
            nFiyat = 0
            cDoviz = ""

            cSQL = "set dateformat dmy " +
                " select a.birimfiyat, a.dovizcinsi, b.fistarihi " +
                " from stokfislines a, stokfis b " +
                " where a.stokfisno = b.stokfisno " +
                " and a.stokno = '" + cStokNo.Trim + "' " +
                " and a.renk = '" + cRenk.Trim + "' " +
                " and a.birimfiyat is not null " +
                " and a.birimfiyat > 0 " +
                " and a.stokhareketkodu in ('02 Tedarikten Giris','04 Mlz Uretimden Giris','05 Diger Giris') " +
                IIf(cBeden.Trim = "", "", " and a.beden = '" + cBeden.Trim + "' ").ToString +
                IIf(cMTF.Trim = "", "", " and a.malzemetakipkodu = '" + cMTF.Trim + "' ").ToString +
                IIf(dTarih = #1/1/1950#, "", " and b.fistarihi <= '" + SQLWriteDate(dTarih) + "' ").ToString +
                " order by b.fistarihi desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFiyat = SQLReadDouble(oReader, "birimfiyat")
                cDoviz = SQLReadString(oReader, "dovizcinsi")
                dTarih = SQLReadDate(oReader, "fistarihi")
            End If
            oReader.Close()

            If nFiyat = 0 Then

                cSQL = "set dateformat dmy " +
                " select a.birimfiyat, a.dovizcinsi, b.fistarihi " +
                " from stokfislines a, stokfis b " +
                " where a.stokfisno = b.stokfisno " +
                " and a.stokno = '" + cStokNo.Trim + "' " +
                " and a.renk = '" + cRenk.Trim + "' " +
                " and a.birimfiyat is not null " +
                " and a.birimfiyat > 0 " +
                " and a.stokhareketkodu in ('02 Tedarikten Giris','04 Mlz Uretimden Giris','05 Diger Giris') " +
                IIf(cBeden.Trim = "", "", " and a.beden = '" + cBeden.Trim + "' ").ToString +
                IIf(dTarih = #1/1/1950#, "", " and b.fistarihi <= '" + SQLWriteDate(dTarih) + "' ").ToString +
                " order by b.fistarihi desc "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "birimfiyat")
                    cDoviz = SQLReadString(oReader, "dovizcinsi")
                    dTarih = SQLReadDate(oReader, "fistarihi")
                End If
                oReader.Close()
            End If

            If nFiyat = 0 Then
                cSQL = "set dateformat dmy " +
                        " select top 1 a.fiyat, a.doviz, b.tarih " +
                        " from isemrilines a, isemri b " +
                        " where a.isemrino = b.isemrino " +
                        " and a.stokno = '" + cStokNo.Trim + "' " +
                        " and a.renk = '" + cRenk.Trim + "' " +
                        " and a.fiyat is not null " +
                        " and a.fiyat <> 0 " +
                        IIf(cBeden.Trim = "", "", " and a.beden = '" + cBeden.Trim + "' ").ToString +
                        IIf(cMTF.Trim = "", "", " and a.malzemetakipno = '" + cMTF.Trim + "' ").ToString +
                        IIf(dTarih = #1/1/1950#, "", " and b.tarih <= '" + SQLWriteDate(dTarih) + "' ").ToString +
                        " order by b.tarih desc "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "fiyat")
                    cDoviz = SQLReadString(oReader, "doviz")
                    dTarih = SQLReadDate(oReader, "tarih")
                End If
                oReader.Close()
            End If

            If nFiyat = 0 Then
                cSQL = "set dateformat dmy " +
                        " select top 1 a.fiyat, a.doviz, b.tarih " +
                        " from isemrilines a, isemri b " +
                        " where a.isemrino = b.isemrino " +
                        " and a.stokno = '" + cStokNo.Trim + "' " +
                        " and a.renk = '" + cRenk.Trim + "' " +
                        " and a.fiyat is not null " +
                        " and a.fiyat <> 0 " +
                        IIf(cBeden.Trim = "", "", " and a.beden = '" + cBeden.Trim + "' ").ToString +
                        IIf(dTarih = #1/1/1950#, "", " and b.tarih <= '" + SQLWriteDate(dTarih) + "' ").ToString +
                        " order by b.tarih desc "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "fiyat")
                    cDoviz = SQLReadString(oReader, "doviz")
                    dTarih = SQLReadDate(oReader, "tarih")
                End If
                oReader.Close()
            End If

            If nFiyat <> 0 Then
                If cDoviz.Trim = "" Then
                    cDoviz = "TL"
                End If
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSonStokGiristenFiyat", cSQL)
        End Try
    End Sub

    Public Function STISonMaliyetMalzeme(ByVal cFilter As String, ByVal cFilter2 As String) As Integer

        Dim cSQL As String = ""
        Dim aMTF() As String
        Dim aSiparis() As String
        Dim nCnt As Integer = 0
        Dim nCntSiparis As Integer = 0
        Dim nSiraNo As Double = 0
        Dim nKarsilanan As Double = 0
        Dim nEURTutar As Double = 0
        Dim nFiyat As Double = 0
        Dim cDoviz As String = ""
        Dim dTarih As Date = #1/1/1950#
        Dim nKur As Double = 0
        Dim nEURKur As Double = 0
        Dim nMiktar As Double = 0
        Dim nTLTutar As Double = 0
        Dim cView As String = ""
        Dim nOTutar As Double = 0
        Dim cODoviz As String = ""
        Dim cKesileneGoreIhtiyacTable As String = ""
        Dim nSonMlytsiraNo As Double = 200000
        Dim nMTFOrani As Double = 0
        Dim nMTFSipAdet As Double = 0
        Dim nSipMTFAdet As Double = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim aMalzeme() As oMalzeme
        Dim nMalzeme As Integer = -1
        Dim aStokFis() As oStokFis
        Dim nCnt2 As Integer
        Dim aTransfer() As oTransfer

        STISonMaliyetMalzeme = 0

        Try
            JustForLog("STISonMaliyetMalzeme start")

            If MTFFastGenerateMulti(cFilter) <> 1 Then Exit Function

            JustForLog("STISonMaliyetMalzeme MTFFastGenerateMulti OK")

            cFilter = Replace(cFilter, "||", "'").Trim

            cSQL = "select distinct a.kullanicisipno  " +
                   " from siparis a, sipmodel b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and a.kullanicisipno  is not null " +
                   " and a.kullanicisipno  <> '' " +
                   cFilter +
                   " order by a.kullanicisipno  "

            If Not CheckExists(cSQL) Then Exit Function

            ConnYage = OpenConn()

            aSiparis = SQLtoStringArrayConnected(cSQL, ConnYage)

            For nCntSiparis = 0 To UBound(aSiparis)
                ReDim aMTF(0)

                cSQL = "select distinct malzemetakipno " +
                    " from sipmodel " +
                    " where siparisno = '" + aSiparis(nCntSiparis) + "' " +
                    " and malzemetakipno is not null " +
                    " and malzemetakipno <> '' " +
                    " order by malzemetakipno "

                If CheckExistsConnected(cSQL, ConnYage) Then

                    aMTF = SQLtoStringArrayConnected(cSQL, ConnYage)

                    For nCnt = 0 To UBound(aMTF)
                        ' MTF nin toplam siparis adedi
                        cSQL = "select sum(coalesce(adet,0)) " +
                            " from sipmodel " +
                            " where malzemetakipno = '" + aMTF(nCnt) + "' "

                        nMTFSipAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                        ' MTF nin ilgili siparise bagli toplam siparis adedi
                        cSQL = "select sum(coalesce(adet,0)) " +
                            " from sipmodel " +
                            " where malzemetakipno = '" + aMTF(nCnt) + "' " +
                            " and siparisno = '" + aSiparis(nCntSiparis) + "' "

                        nSipMTFAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                        If nMTFSipAdet = 0 Then
                            nMTFOrani = 1
                        Else
                            nMTFOrani = nSipMTFAdet / nMTFSipAdet
                        End If

                        nMalzeme = -1
                        ReDim aMalzeme(0)

                        cSQL = "select a.stokno, a.renk, a.beden, b.stoktipi, b.anastokgrubu, b.birim1, "

                        If cFilter2 = "yuklenmisler" Then
                            cSQL = cSQL +
                            " ihtiyac = sum(coalesce(ihtiyac,0)), "
                        Else
                            cSQL = cSQL +
                            " ihtiyac = sum(coalesce(kesilenihtiyac,0)), "
                        End If

                        cSQL = cSQL +
                           " uretimecikan = sum(coalesce(uretimicincikis,0)), " +
                           " uretimiade = sum(coalesce(uretimdeniade,0)) " +
                           " from mtkfislines a, stok b " +
                           " where a.stokno = b.stokno " +
                           " and a.malzemetakipno = '" + aMTF(nCnt) + "' " +
                           " group by a.stokno, a.renk, a.beden, b.stoktipi, b.anastokgrubu, b.birim1 " +
                           " order by a.stokno, a.renk, a.beden "

                        oReader = GetSQLReader(cSQL, ConnYage)

                        Do While oReader.Read
                            nMalzeme = nMalzeme + 1
                            ReDim Preserve aMalzeme(nMalzeme)

                            aMalzeme(nMalzeme).cStokNo = SQLReadString(oReader, "stokno")
                            aMalzeme(nMalzeme).cRenk = SQLReadString(oReader, "renk")
                            aMalzeme(nMalzeme).cBeden = SQLReadString(oReader, "beden")
                            aMalzeme(nMalzeme).cStokTipi = SQLReadString(oReader, "stoktipi")
                            aMalzeme(nMalzeme).cAnaStokGrubu = SQLReadString(oReader, "anastokgrubu")
                            aMalzeme(nMalzeme).cBirim = SQLReadString(oReader, "birim1")
                            aMalzeme(nMalzeme).nIhtiyac = SQLReadDouble(oReader, "ihtiyac") * nMTFOrani
                            aMalzeme(nMalzeme).nUretimeCikan = SQLReadDouble(oReader, "uretimecikan") * nMTFOrani
                            aMalzeme(nMalzeme).nUretimIade = SQLReadDouble(oReader, "uretimiade") * nMTFOrani
                        Loop
                        oReader.Close()

                        For nMalzeme = 0 To UBound(aMalzeme)
                            nKarsilanan = 0
                            nEURTutar = 0
                            ' birim harcamalar
                            stisonmaliyet6(ConnYage, aSiparis(nCntSiparis), aMalzeme(nMalzeme).cStokNo, aMalzeme(nMalzeme).nUretimeCikan * nMTFOrani, aMalzeme(nMalzeme).nUretimIade * nMTFOrani)
                            ' stok girişinden gelen
                            nCnt2 = -1
                            cSQL = "select a.stokfisno, a.fistarihi, a.belgeno, a.faturatarihi, a.faturano, " +
                                " a.departman, a.firma, b.netmiktar1, b.birimfiyat, b.dovizcinsi, " +
                                " b.stokhareketkodu, b.kur, b.iscilikfiyat, b.iscilikdoviz, b.iscilikkur  " +
                                " from stokfis a, stokfislines b " +
                                " where a.stokfisno = b.stokfisno " +
                                " And b.malzemetakipkodu = '" + aMTF(nCnt) + "' " +
                                " and b.stokno = '" + aMalzeme(nMalzeme).cStokNo + "' " +
                                " and b.renk = '" + aMalzeme(nMalzeme).cRenk + "' " +
                                " and b.beden = '" + aMalzeme(nMalzeme).cBeden + "' " +
                                " and b.netmiktar1 is not null " +
                                " and b.netmiktar1 <> 0 " +
                                " and b.stokhareketkodu in ('02 Tedarikten Giris','04 Mlz Uretimden Giris','05 Diger Giris','02 Tedarikten iade') "

                            oReader = GetSQLReader(cSQL, ConnYage)

                            Do While oReader.Read
                                nCnt2 = nCnt2 + 1
                                ReDim Preserve aStokFis(nCnt2)

                                aStokFis(nCnt2).cStokHareketKodu = SQLReadString(oReader, "stokhareketkodu")
                                aStokFis(nCnt2).cStokFisNo = SQLReadString(oReader, "stokfisno")
                                aStokFis(nCnt2).dFisTarihi = SQLReadDate(oReader, "fistarihi")
                                aStokFis(nCnt2).cBelgeNo = SQLReadString(oReader, "belgeno")
                                aStokFis(nCnt2).dFaturaTarihi = SQLReadDate(oReader, "faturatarihi")
                                aStokFis(nCnt2).cFaturaNo = SQLReadString(oReader, "faturano")
                                aStokFis(nCnt2).cDepartman = SQLReadString(oReader, "departman")
                                aStokFis(nCnt2).cFirma = SQLReadString(oReader, "firma")
                                aStokFis(nCnt2).nMiktar = SQLReadDouble(oReader, "netmiktar1") * nMTFOrani

                                aStokFis(nCnt2).nFiyat = SQLReadDouble(oReader, "birimfiyat")
                                If SQLReadString(oReader, "dovizcinsi") = "" Or SQLReadString(oReader, "dovizcinsi") = "YTL" Then
                                    aStokFis(nCnt2).cDoviz = "TL"
                                Else
                                    aStokFis(nCnt2).cDoviz = SQLReadString(oReader, "dovizcinsi")
                                End If
                                If aStokFis(nCnt2).cDoviz = "TL" Then
                                    aStokFis(nCnt2).nKur = 1
                                Else
                                    aStokFis(nCnt2).nKur = SQLReadDouble(oReader, "kur")
                                End If

                                aStokFis(nCnt2).nIscilikFiyat = SQLReadDouble(oReader, "iscilikfiyat")
                                If SQLReadString(oReader, "iscilikdoviz") = "" Or SQLReadString(oReader, "iscilikdoviz") = "YTL" Then
                                    aStokFis(nCnt2).cIscilikDoviz = "TL"
                                Else
                                    aStokFis(nCnt2).cIscilikDoviz = SQLReadString(oReader, "iscilikdoviz")
                                End If
                                If aStokFis(nCnt2).cIscilikDoviz = "TL" Then
                                    aStokFis(nCnt2).nIscilikKur = 1
                                Else
                                    aStokFis(nCnt2).nIscilikKur = SQLReadDouble(oReader, "iscilikkur")
                                End If

                                aStokFis(nCnt2).nEURKur = 0
                                aStokFis(nCnt2).nOrjTutar = 0
                                aStokFis(nCnt2).nTLTutar = 0
                                aStokFis(nCnt2).nEURTutar = 0
                            Loop
                            oReader.Close()

                            If nCnt2 <> -1 Then
                                For nCnt2 = 0 To UBound(aStokFis)
                                    If aStokFis(nCnt2).nKur = 0 Then
                                        aStokFis(nCnt2).nKur = GetKurConnected(ConnYage, aStokFis(nCnt2).cDoviz, aStokFis(nCnt2).dFisTarihi,, aStokFis(nCnt2).cFirma)
                                    End If
                                    If aStokFis(nCnt2).nKur = 0 Then aStokFis(nCnt2).nKur = 1

                                    aStokFis(nCnt2).nEURKur = GetKurConnected(ConnYage, "EUR", aStokFis(nCnt2).dFisTarihi,, aStokFis(nCnt2).cFirma)
                                    If aStokFis(nCnt2).nEURKur = 0 Then aStokFis(nCnt2).nEURKur = 1

                                    ' işçilik birim fiyat hesabı
                                    nFiyat = 0
                                    If aStokFis(nCnt2).nIscilikFiyat <> 0 Then
                                        nFiyat = aStokFis(nCnt2).nIscilikFiyat / aStokFis(nCnt2).nMiktar
                                        If aStokFis(nCnt2).cIscilikDoviz = aStokFis(nCnt2).cDoviz Then
                                            nFiyat = aStokFis(nCnt2).nFiyat
                                        Else
                                            If aStokFis(nCnt2).nIscilikKur = 0 Then
                                                aStokFis(nCnt2).nIscilikKur = GetKurConnected(ConnYage, aStokFis(nCnt2).cIscilikDoviz, aStokFis(nCnt2).dFisTarihi,, aStokFis(nCnt2).cFirma)
                                                If aStokFis(nCnt2).nIscilikKur = 0 Then aStokFis(nCnt2).nIscilikKur = 1
                                            End If
                                            nFiyat = nFiyat * aStokFis(nCnt2).nIscilikKur / aStokFis(nCnt2).nKur
                                        End If
                                    End If

                                    aStokFis(nCnt2).nFiyat = aStokFis(nCnt2).nFiyat + nFiyat

                                    aStokFis(nCnt2).nOrjTutar = aStokFis(nCnt2).nMiktar * aStokFis(nCnt2).nFiyat
                                    aStokFis(nCnt2).nTLTutar = aStokFis(nCnt2).nMiktar * aStokFis(nCnt2).nFiyat * aStokFis(nCnt2).nKur
                                    aStokFis(nCnt2).nEURTutar = aStokFis(nCnt2).nMiktar * aStokFis(nCnt2).nFiyat * aStokFis(nCnt2).nKur / aStokFis(nCnt2).nEURKur

                                    If aStokFis(nCnt2).cStokHareketKodu = "02 Tedarikten iade" Then
                                        stisonmaliyet3(ConnYage, aStokFis(nCnt2).cStokFisNo, aSiparis(nCntSiparis), aStokFis(nCnt2).dFisTarihi,
                                                aStokFis(nCnt2).cBelgeNo, aStokFis(nCnt2).dFaturaTarihi, aStokFis(nCnt2).cFaturaNo,
                                                aStokFis(nCnt2).cDepartman, aStokFis(nCnt2).cFirma, aMalzeme(nMalzeme).cStokTipi,
                                                aStokFis(nCnt2).nEURKur, 0, 0, aStokFis(nCnt2).nTLTutar, aStokFis(nCnt2).nEURTutar, 0, aStokFis(nCnt2).nOrjTutar, aStokFis(nCnt2).cDoviz, 0, aStokFis(nCnt2).nMiktar)

                                        stisonmaliyet4(ConnYage, aSiparis(nCntSiparis), aStokFis(nCnt2).cFirma, -1 * aStokFis(nCnt2).nTLTutar, -1 * aStokFis(nCnt2).nEURTutar)

                                        stisonmaliyet5(ConnYage, aSiparis(nCntSiparis), aStokFis(nCnt2).cFirma, -1 * aStokFis(nCnt2).nOrjTutar, aStokFis(nCnt2).cDoviz)

                                        nEURTutar = nEURTutar - aStokFis(nCnt2).nEURTutar
                                        nKarsilanan = nKarsilanan - aStokFis(nCnt2).nMiktar
                                    Else
                                        stisonmaliyet3(ConnYage, aStokFis(nCnt2).cStokFisNo, aSiparis(nCntSiparis), aStokFis(nCnt2).dFisTarihi,
                                                aStokFis(nCnt2).cBelgeNo, aStokFis(nCnt2).dFaturaTarihi, aStokFis(nCnt2).cFaturaNo,
                                                aStokFis(nCnt2).cDepartman, aStokFis(nCnt2).cFirma, aMalzeme(nMalzeme).cStokTipi,
                                                aStokFis(nCnt2).nEURKur, aStokFis(nCnt2).nTLTutar, aStokFis(nCnt2).nEURTutar, 0, 0, aStokFis(nCnt2).nOrjTutar, 0, aStokFis(nCnt2).cDoviz, aStokFis(nCnt2).nMiktar, 0)

                                        stisonmaliyet4(ConnYage, aSiparis(nCntSiparis), aStokFis(nCnt2).cFirma, aStokFis(nCnt2).nTLTutar, aStokFis(nCnt2).nEURTutar)

                                        stisonmaliyet5(ConnYage, aSiparis(nCntSiparis), aStokFis(nCnt2).cFirma, aStokFis(nCnt2).nOrjTutar, aStokFis(nCnt2).cDoviz)

                                        nEURTutar = nEURTutar + aStokFis(nCnt2).nEURTutar
                                        nKarsilanan = nKarsilanan + aStokFis(nCnt2).nMiktar
                                    End If
                                Next
                            End If

                            ' stok transferden gelen
                            nCnt2 = -1
                            cSQL = "select a.transferfisno, a.tarih, a.netmiktar1, a.dovizcinsi, a.birimfiyat, b.stoktipi, a.kaynakmalzemetakipno, a.hedefmalzemetakipno " +
                               " from stoktransfer a, stok b " +
                               " where (a.hedefmalzemetakipno = '" + aMTF(nCnt) + "' or a.kaynakmalzemetakipno = '" + aMTF(nCnt) + "' ) " +
                               " and a.stokno = b.stokno " +
                               " and a.stokno = '" + aMalzeme(nMalzeme).cStokNo + "' " +
                               " and a.renk = '" + aMalzeme(nMalzeme).cRenk + "' " +
                               " and a.beden = '" + aMalzeme(nMalzeme).cBeden + "' " +
                               " and a.netmiktar1 is not null " +
                               " and a.netmiktar1 <> 0 " +
                               " and a.birimfiyat is not null " +
                               " and a.birimfiyat <> 0 "

                            oReader = GetSQLReader(cSQL, ConnYage)

                            Do While oReader.Read
                                nCnt2 = nCnt2 + 1
                                ReDim Preserve aTransfer(nCnt2)

                                aTransfer(nCnt2).cTransferFisNo = SQLReadString(oReader, "transferfisno")
                                aTransfer(nCnt2).dTarih = SQLReadDate(oReader, "tarih")
                                aTransfer(nCnt2).nMiktar = SQLReadDouble(oReader, "netmiktar1") * nMTFOrani
                                If SQLReadString(oReader, "dovizcinsi") = "" Or SQLReadString(oReader, "dovizcinsi") = "YTL" Then
                                    aTransfer(nCnt2).cDoviz = "TL"
                                Else
                                    aTransfer(nCnt2).cDoviz = SQLReadString(oReader, "dovizcinsi")
                                End If
                                aTransfer(nCnt2).nFiyat = SQLReadDouble(oReader, "birimfiyat")
                                aTransfer(nCnt2).cStokTipi = SQLReadString(oReader, "stoktipi")
                                aTransfer(nCnt2).cKaynalMTF = SQLReadString(oReader, "kaynakmalzemetakipno")
                                aTransfer(nCnt2).cHedefMTF = SQLReadString(oReader, "hedefmalzemetakipno")
                                If aTransfer(nCnt2).cDoviz = "TL" Then
                                    aTransfer(nCnt2).nKur = 1
                                Else
                                    aTransfer(nCnt2).nKur = 0
                                End If
                                aTransfer(nCnt2).nEURKur = 0
                                aTransfer(nCnt2).nOrjTutar = 0
                                aTransfer(nCnt2).nTLTutar = 0
                                aTransfer(nCnt2).nEURTutar = 0
                            Loop
                            oReader.Close()

                            If nCnt2 <> -1 Then
                                For nCnt2 = 0 To UBound(aTransfer)
                                    If aTransfer(nCnt2).nKur = 0 Then
                                        aTransfer(nCnt2).nKur = GetKurConnected(ConnYage, aTransfer(nCnt2).cDoviz, aTransfer(nCnt2).dTarih)
                                    End If
                                    If aTransfer(nCnt2).nKur = 0 Then aTransfer(nCnt2).nKur = 1

                                    aTransfer(nCnt2).nEURKur = GetKurConnected(ConnYage, "EUR", aTransfer(nCnt2).dTarih)
                                    If aTransfer(nCnt2).nEURKur = 0 Then aTransfer(nCnt2).nEURKur = 1

                                    aTransfer(nCnt2).nOrjTutar = aTransfer(nCnt2).nMiktar * aTransfer(nCnt2).nFiyat
                                    aTransfer(nCnt2).nTLTutar = aTransfer(nCnt2).nMiktar * aTransfer(nCnt2).nFiyat * aTransfer(nCnt2).nKur
                                    aTransfer(nCnt2).nEURTutar = aTransfer(nCnt2).nMiktar * aTransfer(nCnt2).nFiyat * aTransfer(nCnt2).nKur / aTransfer(nCnt2).nEURKur

                                    If aMTF(nCnt) = aTransfer(nCnt2).cHedefMTF Then
                                        stisonmaliyet3(ConnYage, aTransfer(nCnt2).cTransferFisNo, aSiparis(nCntSiparis), aTransfer(nCnt2).dTarih,
                                                   "TRANSFER", #1/1/1950#, "", "DEPO", "DAHILI", aMalzeme(nMalzeme).cStokTipi, aTransfer(nCnt2).nEURKur,
                                                   aTransfer(nCnt2).nTLTutar, aTransfer(nCnt2).nEURTutar, 0, 0, aTransfer(nCnt2).nOrjTutar, 0,
                                                   aTransfer(nCnt2).cDoviz, aTransfer(nCnt2).nMiktar, 0)

                                        nEURTutar = nEURTutar + aTransfer(nCnt2).nEURTutar
                                        nKarsilanan = nKarsilanan + aTransfer(nCnt2).nMiktar
                                    Else
                                        stisonmaliyet3(ConnYage, aTransfer(nCnt2).cTransferFisNo, aSiparis(nCntSiparis), aTransfer(nCnt2).dTarih,
                                                   "TRANSFER", #1/1/1950#, "", "DEPO", "DAHILI", aMalzeme(nMalzeme).cStokTipi, aTransfer(nCnt2).nEURKur,
                                                   0, 0, aTransfer(nCnt2).nTLTutar, aTransfer(nCnt2).nEURTutar, 0, aTransfer(nCnt2).nOrjTutar,
                                                   aTransfer(nCnt2).cDoviz, 0, aTransfer(nCnt2).nMiktar)

                                        nEURTutar = nEURTutar - aTransfer(nCnt2).nEURTutar
                                        nKarsilanan = nKarsilanan - aTransfer(nCnt2).nMiktar
                                    End If
                                Next
                            End If

                            cSQL = "select sirano " +
                               " from stisonmaliyet " +
                               " where malzemetakipno = '" + aSiparis(nCntSiparis) + "' " +
                               " and stoktipi = '" + aMalzeme(nMalzeme).cStokTipi + "' "

                            If CheckExistsConnected(cSQL, ConnYage) Then
                                nSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

                                cSQL = "update stisonmaliyet " +
                                   " set ihtiyac = coalesce(ihtiyac,0) + " + SQLWriteDecimal(aMalzeme(nMalzeme).nIhtiyac) + ", " +
                                   " karsilanan = coalesce(karsilanan,0) + " + SQLWriteDecimal(nKarsilanan) + ", " +
                                   " eurtutar = coalesce(eurtutar,0) + " + SQLWriteDecimal(nEURTutar) +
                                   " where sirano = " + SQLWriteDecimal(nSiraNo)

                                ExecuteSQLCommandConnected(cSQL, ConnYage)
                            Else
                                cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, birim, ihtiyac, " +
                                    " karsilanan, eurtutar, siralama) "

                                cSQL = cSQL +
                                    " values ('" + aSiparis(nCntSiparis) + "', " +
                                    " '" + aMalzeme(nMalzeme).cAnaStokGrubu + "', " +
                                    " '" + aMalzeme(nMalzeme).cStokTipi + "', " +
                                    " '" + aMalzeme(nMalzeme).cBirim + "', " +
                                   SQLWriteDecimal(aMalzeme(nMalzeme).nIhtiyac) + ", "

                                cSQL = cSQL +
                                    SQLWriteDecimal(nKarsilanan) + ", " +
                                    SQLWriteDecimal(nEURTutar) + ", " +
                                    SQLWriteDecimal(nSonMlytsiraNo) + " ) "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                                nSonMlytsiraNo = nSonMlytsiraNo + 10
                            End If
                        Next
                    Next
                End If
            Next

            ConnYage.Close()

            STISonMaliyetMalzeme = 1

            JustForLog("STISonMaliyetMalzeme END")

        Catch ex As Exception
            ErrDisp(ex.Message, "STISonMaliyetMalzeme", cSQL)
        End Try
    End Function

    Public Function STISonMaliyetOnMaliyet(ByVal cFilter As String) As Integer

        Dim cSQL As String = ""
        Dim nCnt As Integer = 0
        Dim nCnt2 As Integer = 0
        Dim nCnt3 As Integer = 0
        Dim nSiraNo As Double = 0
        Dim nEURTutar As Double = 0
        Dim cDoviz As String = ""
        Dim nKur As Double = 0
        Dim nEURKur As Double = 0
        Dim nMiktar As Double = 0
        Dim cOnSipCode As String = ""
        Dim nKesimAdet As Double = 0
        Dim nBirimMiktar As Double = 0
        Dim cStokTipi As String = ""
        Dim nSonMlytsiraNo As Double = 0
        Dim cMlzCode As String = ""
        Dim nGenelGider As Double = 0
        Dim nEURFiyat As Double = 0
        Dim cAnaStokGrubu As String = ""
        Dim dTarih As Date = #1/1/1950#
        Dim nModelKur As Double = 0
        Dim cMlzAdi As String = ""
        Dim cBirim As String = ""
        Dim nFiyat As Double = 0

        Dim aSiparis() As String
        Dim aAksesuarStokTipi() As String
        Dim aKumasStokTipi() As String
        Dim adept() As String
        Dim aOnMaliyet() As oOnMaliyet

        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        STISonMaliyetOnMaliyet = 0

        Try
            JustForLog("STISonMaliyetOnMaliyet START")

            cFilter = Replace(cFilter, "||", "'").Trim

            cSQL = "select distinct a.kullanicisipno  " +
                   " from siparis a, sipmodel b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and a.kullanicisipno  is not null " +
                   " and a.kullanicisipno  <> '' " +
                   " and exists (select onsipcode " +
                                " from onmaliyetmodel " +
                                " where onsipcode in (select onmaliyetmodelno " +
                                                    " from sipfiyat " +
                                                    " where siparisno = a.kullanicisipno) " +
                                " and onsipcode is not null " +
                                " and onsipcode <> '') " +
                   cFilter +
                   " order by a.kullanicisipno  "

            If Not CheckExists(cSQL) Then Exit Function

            ConnYage = OpenConn()

            aSiparis = SQLtoStringArrayConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSiparis)

                JustForLog("STISonMaliyetOnMaliyet : " + aSiparis(nCnt))

                cSQL = "select top 1 onmaliyetmodelno " +
                       " from sipfiyat " +
                       " where siparisno = '" + aSiparis(nCnt) + "' " +
                       " and onmaliyetmodelno is not null " +
                       " and onmaliyetmodelno <> '' " +
                       " order by satisfiyati "

                cOnSipCode = SQLGetStringConnected(cSQL, ConnYage)

                cSQL = "select distinct stoktipi " +
                        " from stisonmaliyet " +
                        " where malzemetakipno = '" + aSiparis(nCnt) + "' " +
                        " and anastokgrubu = 'AKSESUAR' " +
                        " and stoktipi is not null " +
                        " and stoktipi <> '' " +
                        " order by stoktipi "

                If CheckExistsConnected(cSQL, ConnYage) Then
                    aAksesuarStokTipi = SQLtoStringArrayConnected(cSQL, ConnYage)
                Else
                    cSQL = "select distinct kod " +
                            " from stoktipi " +
                            " where anastokgrubu = 'AKSESUAR' " +
                            " and kod is not null " +
                            " and kod <> '' " +
                            " order by kod "

                    aAksesuarStokTipi = SQLtoStringArrayConnected(cSQL, ConnYage)
                End If

                cSQL = "select distinct stoktipi " +
                        " from stisonmaliyet " +
                        " where malzemetakipno = '" + aSiparis(nCnt) + "' " +
                        " and anastokgrubu = 'KUMAS' " +
                        " and stoktipi is not null " +
                        " and stoktipi <> '' " +
                        " order by stoktipi "

                If CheckExistsConnected(cSQL, ConnYage) Then
                    aKumasStokTipi = SQLtoStringArrayConnected(cSQL, ConnYage)
                Else
                    cSQL = "select distinct kod " +
                            " from stoktipi " +
                            " where anastokgrubu = 'KUMAS' " +
                            " and kod is not null " +
                            " and kod <> '' " +
                            " order by kod "

                    aKumasStokTipi = SQLtoStringArrayConnected(cSQL, ConnYage)
                End If

                cSQL = "select distinct stoktipi " +
                        " from stisonmaliyet " +
                        " where malzemetakipno = '" + aSiparis(nCnt) + "' " +
                        " and anastokgrubu = 'ISCILIK' " +
                        " and stoktipi is not null " +
                        " and stoktipi <> '' " +
                        " order by stoktipi "

                If CheckExistsConnected(cSQL, ConnYage) Then
                    adept = SQLtoStringArrayConnected(cSQL, ConnYage)
                Else
                    cSQL = "select distinct departman, sira " +
                            " from departman " +
                            " where departman is not null " +
                            " and departman <> '' " +
                            " order by sira, departman "

                    adept = SQLtoStringArrayConnected(cSQL, ConnYage)
                End If

                ' kesilen adede göre ön maliyet hesaplanacak

                cSQL = "select kesimadet " +
                        " from stisonmaliyet1 " +
                        " where siparisno = '" + aSiparis(nCnt) + "' "

                nKesimAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = "update stisonmaliyet1 " +
                        " set onsipcode = '" + cOnSipCode + "' " +
                        " where siparisno = '" + aSiparis(nCnt) + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "select urettarih, modelkur, genelgidertutar, genelgiderdoviz " +
                        " from onmaliyetmodel " +
                        " where onsipcode = '" + cOnSipCode + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nGenelGider = SQLReadDouble(oReader, "genelgidertutar")
                    cDoviz = SQLReadString(oReader, "genelgiderdoviz")
                    dTarih = SQLReadDate(oReader, "urettarih")
                    nModelKur = SQLReadDouble(oReader, "modelkur")
                    nBirimMiktar = 1
                    nMiktar = nKesimAdet
                    nEURFiyat = 0
                    nEURTutar = 0
                    nEURKur = 0
                End If
                oReader.Close()

                nEURKur = GetKur("EUR", dTarih, ConnYage, "On Maliyet Kuru")

                If cDoviz = "EUR" Then
                    nEURFiyat = nGenelGider
                Else
                    nKur = GetKur(cDoviz, dTarih, ConnYage, "On Maliyet Kuru")
                    If nEURKur <> 0 Then
                        nEURFiyat = nGenelGider * nKur / nEURKur
                    End If
                End If

                nEURTutar = nMiktar * nEURFiyat

                If nEURKur <> 0 Then
                    cSQL = "update stisonmaliyet1 " +
                        " set onmaliyetaliskuru = " + SQLWriteDecimal(nEURKur) +
                        " where siparisno = '" + aSiparis(nCnt) + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                If nModelKur <> 0 Then
                    cSQL = "update stisonmaliyet1 " +
                            " set onmaliyetsatiskuru = " + SQLWriteDecimal(nModelKur) +
                            " where siparisno = '" + aSiparis(nCnt) + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                cMlzCode = "GENEL GIDER"

                cSQL = "select sirano " +
                       " from stisonmaliyet2 " +
                       " where siparisno = '" + aSiparis(nCnt) + "' " +
                       " and mlzcode = '" + cMlzCode + "' "

                If CheckExistsConnected(cSQL, ConnYage) Then

                    nSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

                    cSQL = "update stisonmaliyet2 " +
                           " set miktar = coalesce(miktar,0) + " + SQLWriteDecimal(nMiktar) + ", " +
                           " birimharcama = coalesce(birimharcama,0) + " + SQLWriteDecimal(nBirimMiktar) + ", " +
                           " eurtutar = coalesce(eurtutar,0) + " + SQLWriteDecimal(nEURTutar) +
                           " where sirano = " + SQLWriteDecimal(nSiraNo)

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    cSQL = "insert stisonmaliyet2 (siparisno, mlzcode, birim, birimharcama, miktar, eurtutar, kesim) " +
                           " values ('" + aSiparis(nCnt) + "', " +
                           " '" + cMlzCode + "', " +
                           " 'AD', " +
                           SQLWriteDecimal(nBirimMiktar) + ", " +
                           SQLWriteDecimal(nMiktar) + ", " +
                           SQLWriteDecimal(nEURTutar) + ", " +
                           SQLWriteDecimal(nKesimAdet) + " ) "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                cSQL = "select sirano " +
                       " from stisonmaliyet " +
                       " where malzemetakipno = '" + aSiparis(nCnt) + "' " +
                       " and stoktipi = 'GENEL GIDER' "

                If CheckExistsConnected(cSQL, ConnYage) Then

                    nSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

                    cSQL = "update stisonmaliyet " +
                           " set onmaliyetbirimeur = coalesce(onmaliyetbirimeur,0) + " + SQLWriteDecimal(nEURFiyat) + ", " +
                           " onmaliyettoplamharcama = coalesce(onmaliyettoplamharcama,0) + " + SQLWriteDecimal(nMiktar) + ", " +
                           " onmaliyettoplameur = coalesce(onmaliyettoplameur,0) + " + SQLWriteDecimal(nEURTutar) +
                           " where sirano = " + SQLWriteDecimal(nSiraNo)

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    nSonMlytsiraNo = 300000

                    cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, birim, ihtiyac, " +
                            " karsilanan, eurtutar, siralama, onmaliyetbirimharcama, onmaliyetbirimeur, " +
                            " onmaliyettoplamharcama, onmaliyettoplameur) "

                    cSQL = cSQL +
                            " values ('" + aSiparis(nCnt) + "','DIGER','GENEL GIDER','AD',0," +
                            " 0,0," + SQLWriteDecimal(nSonMlytsiraNo) + ",1," + SQLWriteDecimal(nEURFiyat) + "," +
                            SQLWriteDecimal(nMiktar) + "," + SQLWriteDecimal(nEURTutar) + ") "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                nCnt3 = -1

                cSQL = "select a.imalatfire, a.modeldoviz, a.urettarih, a.modeltarih, a.modelkur, " +
                       " b.mlzcode, b.mlzadi, b.brmaliyet, b.doviz, b.miktar, b.birim, b.fire, b.kullanimyuzde " +
                       " from onmaliyetmodel a , onmaliyetlines b " +
                       " where a.onsipcode = b.onsipcode " +
                       " and a.onsipcode = '" + cOnSipCode + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read

                    cMlzCode = SQLReadString(oReader, "mlzcode")
                    cMlzAdi = SQLReadString(oReader, "mlzadi")
                    cDoviz = SQLReadString(oReader, "doviz")
                    cBirim = SQLReadString(oReader, "birim")
                    nBirimMiktar = SQLReadDouble(oReader, "miktar")
                    nMiktar = nBirimMiktar * nKesimAdet
                    nFiyat = SQLReadDouble(oReader, "brmaliyet")
                    cAnaStokGrubu = ""
                    cStokTipi = ""

                    Select Case cMlzCode
                        Case "KUMAS"
                            cMlzCode = "KUMAS"
                            cAnaStokGrubu = "KUMAS"

                            cStokTipi = ""
                            For nCnt2 = 0 To UBound(aKumasStokTipi)
                                If InStr(cMlzAdi, aKumasStokTipi(nCnt2)) > 0 Then
                                    cStokTipi = aKumasStokTipi(nCnt2)
                                    Exit For
                                End If
                            Next

                            If cStokTipi.Trim = "" Then
                                If InStr(cMlzAdi, "TELA") > 0 Then
                                    cAnaStokGrubu = "AKSESUAR"
                                    cStokTipi = "TELA"
                                Else
                                    cStokTipi = "KUMAS"
                                End If
                            End If

                        Case "GARNI", "GARNİ"
                            cMlzCode = "GARNI"
                            cAnaStokGrubu = "KUMAS"

                            cStokTipi = ""
                            For nCnt2 = 0 To UBound(aKumasStokTipi)
                                If InStr(cMlzAdi, aKumasStokTipi(nCnt2)) > 0 Then
                                    cStokTipi = aKumasStokTipi(nCnt2)
                                    Exit For
                                End If
                            Next

                            If cStokTipi.Trim = "" Then
                                If InStr(cMlzAdi, "TELA") > 0 Then
                                    cAnaStokGrubu = "AKSESUAR"
                                    cStokTipi = "TELA"
                                Else
                                    cStokTipi = "GARNI"
                                End If
                            End If

                        Case "ASTAR"
                            cMlzCode = "ASTAR"
                            cAnaStokGrubu = "KUMAS"

                            cStokTipi = ""
                            For nCnt2 = 0 To UBound(aKumasStokTipi)
                                If InStr(cMlzAdi, aKumasStokTipi(nCnt2)) > 0 Then
                                    cStokTipi = aKumasStokTipi(nCnt2)
                                    Exit For
                                End If
                            Next

                            If cStokTipi.Trim = "" Then
                                If InStr(cMlzAdi, "TELA") > 0 Then
                                    cAnaStokGrubu = "AKSESUAR"
                                    cStokTipi = "TELA"
                                Else
                                    cStokTipi = "ASTAR"
                                End If
                            End If

                        Case "TELA"
                            cMlzCode = "TELA"
                            cAnaStokGrubu = "AKSESUAR"
                            cStokTipi = "TELA"

                        Case "BIYE"
                            cMlzCode = "BIYE"
                            cAnaStokGrubu = "KUMAS"

                            cStokTipi = ""
                            For nCnt2 = 0 To UBound(aKumasStokTipi)
                                If InStr(cMlzAdi, aKumasStokTipi(nCnt2)) > 0 Then
                                    cStokTipi = aKumasStokTipi(nCnt2)
                                    Exit For
                                End If
                            Next

                            If cStokTipi.Trim = "" Then
                                If InStr(cMlzAdi, "TELA") > 0 Then
                                    cAnaStokGrubu = "AKSESUAR"
                                    cStokTipi = "TELA"
                                Else
                                    cStokTipi = "BIYE"
                                End If
                            End If

                        Case "DIKIM AKSESUAR", "PAKET AKSESUAR", "AKSESUAR", "DERI"
                            cMlzCode = "AKSESUAR"
                            cAnaStokGrubu = "AKSESUAR"

                            cStokTipi = ""
                            For nCnt2 = 0 To UBound(aAksesuarStokTipi)
                                If InStr(cMlzAdi, aAksesuarStokTipi(nCnt2)) > 0 Then
                                    cStokTipi = aAksesuarStokTipi(nCnt2)
                                    Exit For
                                End If
                            Next

                            If cStokTipi.Trim = "" Then
                                cStokTipi = "DIGER AKSESUAR"
                            End If

                        Case "YAN ISLEM", "ISCILIK"
                            cMlzCode = "ISCILIK"
                            cAnaStokGrubu = "ISCILIK"

                            cStokTipi = ""
                            For nCnt2 = 0 To UBound(adept)
                                If InStr(LCase(Conv_Tr_Char(cMlzAdi)), LCase(Conv_Tr_Char(Replace(adept(nCnt2), "_", "")))) > 0 Then
                                    cStokTipi = adept(nCnt2)
                                    Exit For
                                End If
                            Next

                            If cStokTipi.Trim = "" Then
                                cStokTipi = cMlzAdi
                            End If

                        Case "SEVKIYAT", ""
                            cMlzCode = "DIGER"
                            cAnaStokGrubu = "DIGER"
                            cStokTipi = "DIGER"

                        Case Else
                            cAnaStokGrubu = "DIGER"
                            cStokTipi = "DIGER"
                    End Select

                    nCnt3 = nCnt3 + 1
                    ReDim Preserve aOnMaliyet(nCnt3)

                    aOnMaliyet(nCnt3).cMlzAdi = cMlzAdi
                    aOnMaliyet(nCnt3).cMlzCode = cMlzCode
                    aOnMaliyet(nCnt3).cAnaStokGrubu = cAnaStokGrubu
                    aOnMaliyet(nCnt3).cStokTipi = cStokTipi
                    aOnMaliyet(nCnt3).cBirim = cBirim
                    aOnMaliyet(nCnt3).cDoviz = cDoviz
                    aOnMaliyet(nCnt3).nBirimMiktar = nBirimMiktar
                    aOnMaliyet(nCnt3).nMiktar = nMiktar
                    aOnMaliyet(nCnt3).nFiyat = nFiyat
                Loop
                oReader.Close()

                If nCnt3 <> -1 Then
                    For nCnt3 = 0 To UBound(aOnMaliyet)
                        If aOnMaliyet(nCnt3).cDoviz = "EUR" Then
                            aOnMaliyet(nCnt3).nEURFiyat = aOnMaliyet(nCnt3).nFiyat
                        Else
                            aOnMaliyet(nCnt3).nKur = GetKur(aOnMaliyet(nCnt3).cDoviz, dTarih, ConnYage, "On Maliyet Kuru")
                            If nEURKur <> 0 Then
                                aOnMaliyet(nCnt3).nEURFiyat = aOnMaliyet(nCnt3).nFiyat * aOnMaliyet(nCnt3).nKur / nEURKur
                            End If
                        End If

                        aOnMaliyet(nCnt3).nEURTutar = aOnMaliyet(nCnt3).nMiktar * aOnMaliyet(nCnt3).nEURFiyat

                        cSQL = "select sirano " +
                               " from stisonmaliyet2 " +
                               " where siparisno = '" + aSiparis(nCnt) + "' " +
                               " and mlzcode = '" + aOnMaliyet(nCnt3).cMlzCode + "' "

                        If CheckExistsConnected(cSQL, ConnYage) Then

                            nSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

                            cSQL = "update stisonmaliyet2 " +
                                   " set miktar = coalesce(miktar,0) + " + SQLWriteDecimal(aOnMaliyet(nCnt3).nMiktar) + ", " +
                                   " birimharcama = coalesce(birimharcama,0) + " + SQLWriteDecimal(aOnMaliyet(nCnt3).nBirimMiktar) + ", " +
                                   " eurtutar = coalesce(eurtutar,0) + " + SQLWriteDecimal(aOnMaliyet(nCnt3).nEURTutar) +
                                   " where sirano = " + SQLWriteDecimal(nSiraNo)

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        Else
                            cSQL = "insert stisonmaliyet2 (siparisno, mlzcode, birim, birimharcama, miktar, eurtutar, kesim) " +
                                   " values ('" + aSiparis(nCnt) + "', " +
                                   " '" + aOnMaliyet(nCnt3).cMlzCode + "', " +
                                   " '" + aOnMaliyet(nCnt3).cBirim + "', " +
                                   SQLWriteDecimal(aOnMaliyet(nCnt3).nBirimMiktar) + ", " +
                                   SQLWriteDecimal(aOnMaliyet(nCnt3).nMiktar) + ", " +
                                   SQLWriteDecimal(aOnMaliyet(nCnt3).nEURTutar) + ", " +
                                   SQLWriteDecimal(nKesimAdet) + " ) "

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End If

                        cSQL = "select sirano " +
                               " from stisonmaliyet " +
                               " where malzemetakipno = '" + aSiparis(nCnt) + "' " +
                               " and anastokgrubu = '" + aOnMaliyet(nCnt3).cAnaStokGrubu + "' " +
                               " and stoktipi = '" + aOnMaliyet(nCnt3).cStokTipi + "' "

                        If CheckExistsConnected(cSQL, ConnYage) Then

                            nSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

                            cSQL = "update stisonmaliyet " +
                                   " set onmaliyetbirimharcama  = coalesce(onmaliyetbirimharcama,0) + " + SQLWriteDecimal(aOnMaliyet(nCnt3).nBirimMiktar) + ", " +
                                   " onmaliyetbirimeur = coalesce(onmaliyetbirimeur,0) + " + SQLWriteDecimal(aOnMaliyet(nCnt3).nEURFiyat * aOnMaliyet(nCnt3).nBirimMiktar) + ", " +
                                   " onmaliyettoplamharcama = coalesce(onmaliyettoplamharcama,0) + " + SQLWriteDecimal(aOnMaliyet(nCnt3).nMiktar) + ", " +
                                   " onmaliyettoplameur = coalesce(onmaliyettoplameur,0) + " + SQLWriteDecimal(aOnMaliyet(nCnt3).nEURTutar) +
                                   " where sirano = " + SQLWriteDecimal(nSiraNo)

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        Else
                            nSonMlytsiraNo = 300000

                            cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, birim, ihtiyac, " +
                                    " karsilanan, eurtutar, siralama, onmaliyetbirimharcama, onmaliyetbirimeur, " +
                                    " onmaliyettoplamharcama, onmaliyettoplameur) "

                            cSQL = cSQL +
                                    " values ('" + aSiparis(nCnt) + "', " +
                                    " '" + SQLWriteString(aOnMaliyet(nCnt3).cAnaStokGrubu, 30) + "', " +
                                    " '" + SQLWriteString(aOnMaliyet(nCnt3).cStokTipi, 30) + "', " +
                                    " '" + SQLWriteString(aOnMaliyet(nCnt3).cBirim, 30) + "', " +
                                    " 0,0,0," +
                                    SQLWriteDecimal(nSonMlytsiraNo) + "," +
                                    SQLWriteDecimal(aOnMaliyet(nCnt3).nBirimMiktar) + "," +
                                    SQLWriteDecimal(aOnMaliyet(nCnt3).nEURFiyat * aOnMaliyet(nCnt3).nBirimMiktar) + "," +
                                    SQLWriteDecimal(aOnMaliyet(nCnt3).nMiktar) + "," +
                                    SQLWriteDecimal(aOnMaliyet(nCnt3).nEURTutar) + ") "

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End If
                    Next
                End If
            Next

            ConnYage.Close()
            STISonMaliyetOnMaliyet = 1
            JustForLog("STISonMaliyetOnMaliyet STOP")

        Catch ex As Exception
            ErrDisp(ex.Message, "STISonMaliyetOnMaliyet", cSQL)
        End Try
    End Function

    Public Function STISonMaliyetDigerMasraf(cFilter As String) As Integer

        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim cDoviz As String = ""
        Dim dTarih As Date = #1/1/1950#
        Dim nKur As Double = 0
        Dim nEURKur As Double = 0
        Dim dSipTarih As Date = #1/1/1950#
        Dim nSipKur As Double = 0
        Dim nSipEURKur As Double = 0
        Dim nSipAdet As Double = 0
        Dim nSevkAdet As Double = 0
        Dim nKesim As Double = 0
        Dim nTutar As Double = 0
        Dim nSonMlytsiraNo As Double = 0
        Dim cMusteri As String = ""
        Dim nindirim As Double = 0
        Dim nSevkTutar As Double = 0
        Dim aSiparis() As String
        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim aTanimlamaData() As oTanimlamaData
        Dim aReklamasyon() As oReklamasyon

        STISonMaliyetDigerMasraf = 0

        Try
            JustForLog("STISonMaliyetDigerMasraf START")

            cFilter = Replace(cFilter, "||", "'").Trim

            cSQL = "select distinct a.kullanicisipno  " +
                   " from siparis a, sipmodel b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and a.kullanicisipno  is not null " +
                   " and a.kullanicisipno  <> '' " +
                   cFilter +
                   " order by a.kullanicisipno  "

            If Not CheckExists(cSQL) Then Exit Function

            ConnYage = OpenConn()

            aSiparis = SQLtoStringArrayConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSiparis)

                cSQL = "select siparistarihi " +
                        " from siparis " +
                        " where kullanicisipno = '" + aSiparis(nCnt).Trim + "' "

                dSipTarih = SQLGetDateConnected(cSQL, ConnYage)

                If dSipTarih = #1/1/1950# Then
                    dSipTarih = Today.Date
                End If

                nSipEURKur = GetKur("EUR", dSipTarih, ConnYage)

                cSQL = "select siparisadet " +
                        " from stisonmaliyet1 " +
                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                nSipAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = "select sevkiyatadet " +
                        " from stisonmaliyet1 " +
                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                nSevkAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                If nSipAdet <> 0 And nSevkAdet <> 0 Then
                    If nSevkAdet >= nSipAdet Then
                        cSQL = "update stisonmaliyet5 " +
                               " set hakedis = tutar " +
                               " where malzemetakipno = '" + aSiparis(nCnt).Trim + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    Else
                        cSQL = "update stisonmaliyet5 " +
                               " set hakedis = tutar * " + SQLWriteDecimal(nSevkAdet / nSipAdet) +
                               " where malzemetakipno = '" + aSiparis(nCnt).Trim + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                End If

                ' ortalama malzeme harcamalari

                cSQL = "select sum(coalesce(a.adet,0)) " +
                        " from uretharrba a , uretharfis b " +
                        " where a.uretfisno = b.uretfisno " +
                        " and a.uretimtakipno in (select uretimtakipno " +
                                                " from sipmodel " +
                                                " where siparisno = '" + aSiparis(nCnt).Trim + "') " +
                        " and b.cikisdept like '%KES%' "

                nKesim = SQLGetDoubleConnected(cSQL, ConnYage)

                If nKesim <> 0 Then
                    cSQL = "update stisonmaliyet6 " +
                           " set kesim = " + SQLWriteDecimal(nKesim) + ", " +
                           " ortalamaharcama = (coalesce(uretimecikan,0) - coalesce(uretimdeniade,0)) / " + SQLWriteDecimal(nKesim) +
                           " where malzemetakipno = '" + aSiparis(nCnt).Trim + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                ' son maliyete sipariş kartından gelen masraflar

                nSonMlytsiraNo = 300000

                nCnt1 = -1
                ReDim aTanimlamaData(0)

                cSQL = "select a.alan,a.s_numeric, a.s_doviz, b.alanetiketi " +
                        " from tanimlamadata a, tanimlama b " +
                        " where a.karttipi = b.karttipi " +
                        " and a.alan = b.alan " +
                        " and a.karttipi = 'siparismaliyet' " +
                        " and a.kayitno = '" + aSiparis(nCnt).Trim + "' " +
                        " and a.s_numeric is not null " +
                        " and a.s_numeric <> 0 "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read
                    nCnt1 = nCnt1 + 1
                    ReDim Preserve aTanimlamaData(nCnt1)

                    aTanimlamaData(nCnt1).cMasraf = SQLReadString(oReader, "alanetiketi")
                    If SQLReadString(oReader, "s_doviz") = "" Then
                        aTanimlamaData(nCnt1).cDoviz = "EUR"
                    Else
                        aTanimlamaData(nCnt1).cDoviz = SQLReadString(oReader, "s_doviz")
                    End If
                    aTanimlamaData(nCnt1).nTutar = SQLReadDouble(oReader, "s_numeric")
                Loop
                oReader.Close()

                If nCnt1 <> -1 Then
                    For nCnt1 = 0 To UBound(aTanimlamaData)

                        nTutar = aTanimlamaData(nCnt1).nTutar
                        cDoviz = aTanimlamaData(nCnt1).cDoviz

                        If cDoviz = "EUR" Then
                            nKur = nSipEURKur
                        Else
                            GetKur(cDoviz, dSipTarih, ConnYage)
                        End If

                        If nSipEURKur <> 0 Then
                            nTutar = nTutar * nKur / nSipEURKur
                        End If

                        Select Case aTanimlamaData(nCnt1).cMasraf
                            Case "GENEL GIDER"
                                cSQL = "update stisonmaliyet1 " +
                                        " set genelgidertutar = " + SQLWriteDecimal(nTutar) + ", " +
                                        " genelgiderdoviz = 'EUR' " +
                                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "GUMRUK MASRAFI"
                                cSQL = "update stisonmaliyet1 " +
                                        " set gumruktutar = " + SQLWriteDecimal(nTutar) + ", " +
                                        " gumrukdoviz = 'EUR' " +
                                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "MUSTERI REKLAMASYON"
                                ' siparis kartindan gelen reklamasyon
                                cSQL = "update stisonmaliyet1 " +
                                        " set musterireklamasyontutar = " + SQLWriteDecimal(nTutar) + ", " +
                                        " musterireklamasyondoviz = 'EUR' " +
                                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "KESTIGIMIZ REKLAMASYON"
                                ' siparis kartindan gelen reklamasyon
                                cSQL = "update stisonmaliyet1 " +
                                        " set reklamasyontutar = coalesce(reklamasyontutar,0) + " + SQLWriteDecimal(nTutar) + ", " +
                                        " reklamasyondoviz = 'EUR' " +
                                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End Select

                        If aTanimlamaData(nCnt1).cMasraf = "KESTIGIMIZ REKLAMASYON" Then
                            ' siparis kartindan gelen reklamasyon
                            cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, siralama, eurtutar, birim, ihtiyac, karsilanan) " +
                                    " values ('" + aSiparis(nCnt).Trim + "', " +
                                    " 'KESTIGIMIZ REKLAMASYON', " +
                                    " 'KESTIGIMIZ REKLAMASYON', " +
                                    SQLWriteDecimal(nSonMlytsiraNo) + ", " +
                                    SQLWriteDecimal(-1 * nTutar) + ", " +
                                    " 'AD', " +
                                    " 0, " +
                                    " 0) "

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        Else
                            cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, siralama, eurtutar, birim, ihtiyac, karsilanan) " +
                                    " values ('" + aSiparis(nCnt).Trim + "', " +
                                    " 'DIGER', " +
                                    " '" + SQLWriteString(aTanimlamaData(nCnt1).cMasraf, 30) + "', " +
                                    SQLWriteDecimal(nSonMlytsiraNo) + ", " +
                                    SQLWriteDecimal(nTutar) + ", " +
                                    " 'AD', " +
                                    " 1, " +
                                    " 1) "

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End If

                        nSonMlytsiraNo = nSonMlytsiraNo + 10
                    Next
                End If

                ' reklamasyon fisleriyle kestigimiz reklamasyon tutari

                nCnt1 = -1
                ReDim aReklamasyon(0)

                cSQL = "select a.reklamasyontutar, a.doviz, b.tarih " +
                        " from sipkapanis a, sipreklamasyon b " +
                        " where a.siprekfisno = b.siprekfisno " +
                        " and a.siparisno = '" + aSiparis(nCnt).Trim + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read
                    nCnt1 = nCnt1 + 1
                    ReDim Preserve aReklamasyon(nCnt1)

                    aReklamasyon(nCnt1).dTarih = SQLReadDate(oReader, "tarih")
                    If SQLReadString(oReader, "doviz") = "" Then
                        aReklamasyon(nCnt1).cDoviz = "TL"
                    Else
                        aReklamasyon(nCnt1).cDoviz = SQLReadString(oReader, "doviz")
                    End If
                    aReklamasyon(nCnt1).nTutar = SQLReadDouble(oReader, "reklamasyontutar")
                Loop
                oReader.Close()

                If nCnt1 <> -1 Then
                    nTutar = 0
                    For nCnt1 = 0 To UBound(aReklamasyon)
                        If aReklamasyon(nCnt1).cDoviz = "EUR" Then
                            nTutar = nTutar + aReklamasyon(nCnt1).nTutar
                        Else
                            nEURKur = GetKur("EUR", aReklamasyon(nCnt1).dTarih, ConnYage)
                            nKur = GetKur(aReklamasyon(nCnt1).cDoviz, aReklamasyon(nCnt1).dTarih, ConnYage)

                            If nEURKur <> 0 Then
                                nTutar = nTutar + (aReklamasyon(nCnt1).nTutar * nKur / nEURKur)
                            End If
                        End If
                    Next

                    If nTutar <> 0 Then
                        cSQL = "update stisonmaliyet1 " +
                                " set reklamasyontutar = coalesce(reklamasyontutar,0) + " + SQLWriteDecimal(nTutar) + ", " +
                                " reklamasyondoviz = 'EUR' " +
                                " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                        nSonMlytsiraNo = 400000

                        cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, siralama, eurtutar, " +
                                " birim, ihtiyac, karsilanan) "

                        cSQL = cSQL +
                                " values ('" + aSiparis(nCnt).Trim + "', " +
                                " 'KESTIGIMIZ REKLAMASYON', " +
                                " 'KESTIGIMIZ REKLAMASYON', " +
                                SQLWriteDecimal(nSonMlytsiraNo) + ", " +
                                SQLWriteDecimal(-1 * nTutar) + ", "

                        cSQL = cSQL +
                                " 'AD', " +
                                " 0, " +
                                " 0) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                End If

                ' Müşteri indirimi
                cSQL = "select a.indirim " +
                        " from firma a, siparis b " +
                        " where a.firma = b.musterino " +
                        " and b.kullanicisipno = '" + aSiparis(nCnt).Trim + "' "

                nindirim = SQLGetDoubleConnected(cSQL, ConnYage)

                If nindirim <> 0 Then

                    nSevkTutar = GetSipSevkDVZTutarConnected(ConnYage, aSiparis(nCnt).Trim, , "EUR")
                    nindirim = nSevkTutar * nindirim / 100

                    nSonMlytsiraNo = 400001

                    cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, siralama, eurtutar, birim, ihtiyac, karsilanan) " +
                            " values ('" + aSiparis(nCnt).Trim + "', " +
                            " 'DIGER', " +
                            " 'INDIRIM', " +
                            SQLWriteDecimal(nSonMlytsiraNo) + ", " +
                            SQLWriteDecimal(nindirim) + ", " +
                            " 'AD', " +
                            " 1, " +
                            " 1) "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                ' reçeteden birim harcamalar

                ' ana receteden ara/bul

                cSQL = "update stisonmaliyet " +
                        " set receteharcama = (select top 1 kullanimmiktari " +
                                            " from modelhammadde a, stok b " +
                                            " where a.hammaddekodu = b.stokno " +
                                            " and b.stoktipi = stisonmaliyet.stoktipi " +
                                            " and a.modelno in (select modelno " +
                                                                " from sipmodel " +
                                                                " where siparisno = stisonmaliyet.malzemetakipno)) " +
                        " where malzemetakipno = '" + aSiparis(nCnt).Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                ' alternatif receteden ara/bul

                cSQL = "update stisonmaliyet " +
                        " set receteharcama = (select top 1 kullanimmiktari " +
                                            " from modelhammadde2 a, stok b " +
                                            " where a.hammaddekodu = b.stokno " +
                                            " and b.stoktipi = stisonmaliyet.stoktipi " +
                                            " and a.modelno in (select modelno " +
                                                                " from sipmodel " +
                                                                " where siparisno = stisonmaliyet.malzemetakipno)) " +
                        " where malzemetakipno = '" + aSiparis(nCnt).Trim + "' " +
                        " and (receteharcama is null or receteharcama = 0) "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                ' siparis bazinda gerçeklesen maliyet tutar toplami **** EN SON HESAPLANIR
                cSQL = "update stisonmaliyet1 " +
                        " set maliyettutar = (select sum(coalesce(eurtutar,0)) " +
                                            " from stisonmaliyet " +
                                            " where malzemetakipno = stisonmaliyet1.siparisno), " +
                        " maliyetdoviz = 'EUR' " +
                        " where siparisno = '" + aSiparis(nCnt).Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next

            ConnYage.Close()
            STISonMaliyetDigerMasraf = 1
            JustForLog("STISonMaliyetDigerMasraf STOP")

        Catch ex As Exception
            ErrDisp(ex.Message, "STISonMaliyetDigerMasraf", cSQL)
        End Try
    End Function

    Public Function StokFisKurTamamla(cFilter As String) As Integer

        Dim cSQL As String = ""
        Dim aSHKur() As oSHKur
        Dim nCnt As Integer = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nKur As Double = 0

        StokFisKurTamamla = 0

        Try
            JustForLog("StokFisKurTamamla START")

            cFilter = Replace(cFilter, "||", "'").Trim

            ConnYage = OpenConn()

            cSQL = "update stokfislines " +
                    " set kur = 1 " +
                    " where dovizcinsi in ('','TL','YTL','TRL') " +
                    " and (kur = 0 or kur is null) " +
                    " and stokfisno in (select a.stokfisno " +
                                        " from stokfis a, stokfislines b, firma c " +
                                        " where a.stokfisno = b.stokfisno " +
                                        " and a.firma = c.firma " +
                                        cFilter.Trim + " ) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update stokfislines " +
                    " set iscilikkur = 1 " +
                    " where iscilikdoviz in ('','TL','YTL','TRL') " +
                    " and (iscilikkur = 0 or iscilikkur is null) " +
                    " and stokfisno in (select a.stokfisno " +
                                        " from stokfis a, stokfislines b, firma c " +
                                        " where a.stokfisno = b.stokfisno " +
                                        " and a.firma = c.firma " +
                                        cFilter.Trim + " ) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = -1
            ReDim aSHKur(0)

            cSQL = "select a.fistarihi, a.stokfisno, b.dovizcinsi, b.stokhareketno, c.kurcinsi " +
                    " from stokfis a, stokfislines b, firma c " +
                    " where a.stokfisno = b.stokfisno " +
                    " and a.firma = c.firma " +
                    " and a.fistarihi is not null " +
                    " and a.fistarihi <> '01.01.1950' " +
                    " and b.birimfiyat is not null " +
                    " and b.birimfiyat <> 0 " +
                    " and b.dovizcinsi not in ('','TL','YTL','TRL') " +
                    " and (b.kur = 0 or b.kur is null) " +
                    cFilter.Trim +
                    " order by a.fistarihi "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aSHKur(nCnt)

                aSHKur(nCnt).nStokHareketNo = SQLReadDouble(oReader, "stokhareketno")
                aSHKur(nCnt).dTarih = SQLReadDate(oReader, "fistarihi")
                aSHKur(nCnt).cDoviz = SQLReadString(oReader, "dovizcinsi")
                If SQLReadString(oReader, "kurcinsi") = "" Then
                    aSHKur(nCnt).cKurCinsi = "Kur"
                Else
                    aSHKur(nCnt).cKurCinsi = SQLReadString(oReader, "kurcinsi")
                End If
            Loop
            oReader.Close()

            If nCnt <> -1 Then
                For nCnt = 0 To UBound(aSHKur)
                    nKur = GetKurConnected(ConnYage, aSHKur(nCnt).cDoviz, aSHKur(nCnt).dTarih, aSHKur(nCnt).cKurCinsi)
                    If nKur <> 0 Then
                        cSQL = "update stokfislines " +
                                " set kur = " + SQLWriteDecimal(nKur) +
                                " where stokhareketno = " + SQLWriteDecimal(aSHKur(nCnt).nStokHareketNo)

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                        JustForLog("StokFisKurTamamla / stok hareket no : " + CStr(aSHKur(nCnt).nStokHareketNo))
                    End If
                Next
            End If

            nCnt = -1
            ReDim aSHKur(0)

            cSQL = "select a.fistarihi, a.stokfisno, b.iscilikdoviz, b.stokhareketno, c.kurcinsi " +
                    " from stokfis a, stokfislines b, firma c " +
                    " where a.stokfisno = b.stokfisno " +
                    " and a.firma = c.firma " +
                    " and a.fistarihi is not null " +
                    " and a.fistarihi <> '01.01.1950' " +
                    " and b.iscilikfiyat is not null " +
                    " and b.iscilikfiyat <> 0 " +
                    " and b.iscilikdoviz not in ('','TL','YTL','TRL') " +
                    " and (b.iscilikkur = 0 or b.iscilikkur is null) " +
                    cFilter.Trim +
                    " order by a.fistarihi "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aSHKur(nCnt)

                aSHKur(nCnt).nStokHareketNo = SQLReadDouble(oReader, "stokhareketno")
                aSHKur(nCnt).dTarih = SQLReadDate(oReader, "fistarihi")
                aSHKur(nCnt).cDoviz = SQLReadString(oReader, "iscilikdoviz")
                If SQLReadString(oReader, "kurcinsi") = "" Then
                    aSHKur(nCnt).cKurCinsi = "Kur"
                Else
                    aSHKur(nCnt).cKurCinsi = SQLReadString(oReader, "kurcinsi")
                End If
            Loop
            oReader.Close()

            If nCnt <> -1 Then
                For nCnt = 0 To UBound(aSHKur)
                    nKur = GetKurConnected(ConnYage, aSHKur(nCnt).cDoviz, aSHKur(nCnt).dTarih, aSHKur(nCnt).cKurCinsi)
                    If nKur <> 0 Then
                        cSQL = "update stokfislines " +
                                " set iscilikkur = " + SQLWriteDecimal(nKur) +
                                " where stokhareketno = " + SQLWriteDecimal(aSHKur(nCnt).nStokHareketNo)

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                        JustForLog("StokFisKurTamamla / stok hareket no : " + CStr(aSHKur(nCnt).nStokHareketNo))
                    End If
                Next
            End If

            ConnYage.Close()
            StokFisKurTamamla = 1
            JustForLog("StokFisKurTamamla STOP")

        Catch ex As Exception
            ErrDisp(ex.Message, "StokFisKurTamamla", cSQL)
        End Try
    End Function

    'Public Function STISonMaliyet8(cFilter As String) As Integer

    '    Dim ConnYage As SqlConnection
    '    Dim oReader As SqlDataReader
    '    Dim cSQL As String = ""
    '    Dim nSipAdet As Double = 0
    '    Dim nSipTutar As Double = 0
    '    Dim nSevkiyatAdet As Double = 0
    '    Dim nSevkTutar As Double = 0
    '    Dim nMaliyetTutar As Double = 0
    '    Dim cModelNo As String = ""
    '    Dim cAciklama As String = ""
    '    Dim cDosyaKapandi As String = ""
    '    Dim nFiyat As Double = 0
    '    Dim cDoviz As String = ""
    '    Dim dSevkTarih As Date = #1/1/1950#
    '    Dim nKumasTutar As Double = 0
    '    Dim nAksesuarTutar As Double = 0
    '    Dim nIscilikYutar As Double = 0
    '    Dim nDigerTutar As Double = 0
    '    Dim nGenelGiderTutar As Double = 0
    '    Dim dPSevkTarih As Date
    '    Dim nPKumasTutar As Double = 0
    '    Dim nPAksesuarTutar As Double = 0
    '    Dim nPIscilikYutar As Double = 0
    '    Dim nPDigerTutar As Double = 0
    '    Dim nPGenelGiderTutar As Double = 0
    '    Dim cMusteriNo As String = ""
    '    Dim nEURFiyat As Double = 0
    '    Dim cEURDoviz As String = ""
    '    Dim cSipFiyatDoviz As String = ""
    '    Dim aSiparis() As String
    '    Dim cSipFilter As String = ""
    '    Dim nCnt As Integer = 0
    '    Dim nCnt2 As Integer = 0
    '    Dim aSTI8() As ostisonmaliyet8
    '    Dim nRatio As Double = 0

    '    STISonMaliyet8 = 0

    '    Try
    '        JustForLog("STISonMaliyet8 START")

    '        cFilter = Replace(cFilter, "||", "'").Trim

    '        cSQL = "select distinct a.kullanicisipno  " +
    '               " from siparis a, sipmodel b, ymodel c " +
    '               " where a.kullanicisipno = b.siparisno " +
    '               " and b.modelno = c.modelno " +
    '               " and a.kullanicisipno  is not null " +
    '               " and a.kullanicisipno  <> '' " +
    '               cFilter +
    '               " order by a.kullanicisipno  "

    '        If Not CheckExists(cSQL) Then Exit Function

    '        ReDim aSTI8(0)
    '        nCnt2 = -1

    '        cSQL = "delete stisonmaliyet8 " +
    '                " where siparisno in (select a.kullanicisipno " +
    '                                   " from siparis a, sipmodel b, ymodel c " +
    '                                   " where a.kullanicisipno = b.siparisno " +
    '                                   " and b.modelno = c.modelno " +
    '                                   " and a.kullanicisipno  is not null " +
    '                                   " and a.kullanicisipno  <> '' " +
    '                                   cFilter + ") "

    '        ConnYage = OpenConn()

    '        aSiparis = SQLtoStringArrayConnected(cSQL, ConnYage)

    '        For nCnt = 0 To aSiparis.GetUpperBound(0)

    '            nSipAdet = 0
    '            nSipTutar = 0
    '            nSevkiyatAdet = 0
    '            nSevkTutar = 0
    '            nMaliyetTutar = 0
    '            cModelNo = ""
    '            cAciklama = ""
    '            cDosyaKapandi = ""
    '            nFiyat = 0
    '            cDoviz = ""
    '            dSevkTarih = #1/1/1950#
    '            nKumasTutar = 0
    '            nAksesuarTutar = 0
    '            nIscilikYutar = 0
    '            nDigerTutar = 0
    '            nGenelGiderTutar = 0
    '            dPSevkTarih = #1/1/1950#
    '            nPKumasTutar = 0
    '            nPAksesuarTutar = 0
    '            nPIscilikYutar = 0
    '            nPDigerTutar = 0
    '            nPGenelGiderTutar = 0
    '            cSipFiyatDoviz = ""

    '            JustForLog("STISonMaliyet8 Siparis : " + aSiparis(nCnt))

    '            cSQL = "select a.siparisno, a.siparisadet, a.siptutar, a.sevkiyatadet, a.sevktutar, b.musterino, " +
    '                    " a.maliyettutar, a.genelgidertutar, a.sevktarih, b.dosyakapandi, b.ilksevktarihi " +
    '                    " from stisonmaliyet1 a, siparis b " +
    '                    " where a.siparisno = b.kullanicisipno " +
    '                    " and a.siparisno = '" + aSiparis(nCnt) + "' "

    '            oReader = GetSQLReader(cSQL, ConnYage)

    '            If oReader.Read Then
    '                nSipAdet = SQLReadDouble(oReader, "siparisadet")
    '                nSipTutar = SQLReadDouble(oReader, "siptutar")
    '                nSevkiyatAdet = SQLReadDouble(oReader, "sevkiyatadet")
    '                nSevkTutar = SQLReadDouble(oReader, "sevktutar")
    '                nMaliyetTutar = SQLReadDouble(oReader, "maliyettutar")
    '                nGenelGiderTutar = SQLReadDouble(oReader, "genelgidertutar")
    '                cDosyaKapandi = SQLReadString(oReader, "dosyakapandi")
    '                dSevkTarih = SQLReadDate(oReader, "sevktarih")
    '                cMusteriNo = SQLReadString(oReader, "musterino")
    '            End If
    '            oReader.Close()

    '            cModelNo = ""
    '            cAciklama = ""

    '            cSQL = "select b.modelno, b.aciklama " +
    '                    " from sipmodel a, ymodel b " +
    '                    " where a.modelno = b.modelno " +
    '                    " and a.siparisno = '" + aSiparis(nCnt) + "' "

    '            oReader = GetSQLReader(cSQL, ConnYage)

    '            If oReader.Read Then
    '                cModelNo = SQLReadString(oReader, "modelno")
    '                cAciklama = SQLReadString(oReader, "aciklama")
    '            End If
    '            oReader.Close()

    '            nFiyat = 0
    '            cSipFiyatDoviz = "TL"

    '            cSQL = "select satisfiyati, satisdoviz " +
    '                    " from sipfiyat " +
    '                    " where siparisno = '" + aSiparis(nCnt) + "' " +
    '                    " and satisfiyati is not null " +
    '                    " and satisfiyati <> 0 "

    '            oReader = GetSQLReader(cSQL, ConnYage)

    '            If oReader.Read Then
    '                nFiyat = SQLReadDouble(oReader, "satisfiyati")
    '                cSipFiyatDoviz = SQLReadString(oReader, "satisdoviz")
    '            End If
    '            oReader.Close()

    '            nKumasTutar = 0
    '            nPKumasTutar = 0
    '            nAksesuarTutar = 0
    '            nPAksesuarTutar = 0
    '            nIscilikYutar = 0
    '            nPIscilikYutar = 0
    '            nGenelGiderTutar = 0
    '            nPGenelGiderTutar = 0
    '            nDigerTutar = 0
    '            nPDigerTutar = 0

    '            cSQL = "select mlzcode, " +
    '                    " gertutar = sum(coalesce(gertutar,0)), " +
    '                    " pltutar = sum(coalesce(pltutar,0)) " +
    '                    " from stisonmaliyet7 " +
    '                    " where siparisno = '" + aSiparis(nCnt) + "' " +
    '                    " group by mlzcode "

    '            oReader = GetSQLReader(cSQL, ConnYage)

    '            Do While oReader.Read
    '                Select Case SQLReadString(oReader, "mlzcode")
    '                    Case "KUMAS", "ASTAR"
    '                        nKumasTutar = nKumasTutar + SQLReadDouble(oReader, "gertutar")
    '                        nPKumasTutar = nPKumasTutar + SQLReadDouble(oReader, "pltutar")
    '                    Case "AKSESUAR", "TELA"
    '                        nAksesuarTutar = nAksesuarTutar + SQLReadDouble(oReader, "gertutar")
    '                        nPAksesuarTutar = nPAksesuarTutar + SQLReadDouble(oReader, "pltutar")
    '                    Case "ISCILIK"
    '                        nIscilikYutar = nIscilikYutar + SQLReadDouble(oReader, "gertutar")
    '                        nPIscilikYutar = nPIscilikYutar + SQLReadDouble(oReader, "pltutar")
    '                    Case "GENEL GIDER"
    '                        nGenelGiderTutar = nGenelGiderTutar + SQLReadDouble(oReader, "gertutar")
    '                        nPGenelGiderTutar = nPGenelGiderTutar + SQLReadDouble(oReader, "pltutar")
    '                    Case "DIGER"
    '                        nDigerTutar = nDigerTutar + SQLReadDouble(oReader, "gertutar")
    '                        nPDigerTutar = nPDigerTutar + SQLReadDouble(oReader, "pltutar")
    '                End Select
    '            Loop
    '            oReader.Close()

    '            cSQL = "select top 1 ilksevktar " +
    '                    " from sevkplfislines " +
    '                    " where siparisno = '" + aSiparis(nCnt) + "' " +
    '                    " order by ilksevktar "

    '            dPSevkTarih = SQLGetDateConnected(cSQL, ConnYage)

    '            ' gerçekleşen adet ve tutarlar gerçekleşen sevkiyatlara,
    '            ' planlanan adet ve tutarlar planlanan sevkiyatlara bölünür

    '            cSQL = "select w.siparisno, w.modelno, " +
    '                    " ptarih = w.ilksevktar, " +
    '                    " padet = sum(coalesce(w.toplam,0)), " +
    '                    " gtarih = min(w.gtarih), " +
    '                    " gadet = sum(coalesce(w.gadet,0)), "

    '            ' siparişin toplam giden adedi

    '            cSQL = cSQL +
    '                    " tgadet = (select sum((y.koliend - y.kolibeg + 1) * z.adet) " +
    '                            " from sevkform x, sevkformlines y, sevkformlinesrba z " +
    '                            " Where x.sevkformno = Y.sevkformno " +
    '                            " and y.sevkformno = z.sevkformno " +
    '                            " and y.ulineno = z.ulineno " +
    '                            " and x.ok = 'E' " +
    '                            " and y.siparisno = w.siparisno " +
    '                            " and y.modelno = w.modelno), "

    '            ' siparişin toplam planlanan adedi

    '            cSQL = cSQL +
    '                    " tpadet = (select sum(coalesce(toplam,0)) " +
    '                            " From sevkplfislines " +
    '                            " where siparisno = w.siparisno " +
    '                            " and modelno = w.modelno " +
    '                            " and exists (select siparisno " +
    '                                        " from sipmodel " +
    '                                        " where siparisno = sevkplfislines.siparisno " +
    '                                        " and modelno = sevkplfislines.modelno " +
    '                                        " and sevkiyattakipno = sevkplfislines.sevkiyattakipno) ) "

    '            cSQL = cSQL +
    '                    " from (select a.sevkiyattakipno, a.sevkemrino, a.siparisno, a.modelno, a.toplam, a.ilksevktar, " +
    '                            " gadet = (select sum((y.koliend - y.kolibeg + 1) * z.adet) " +
    '                                    " from sevkform x, sevkformlines y, sevkformlinesrba z " +
    '                                    " Where x.sevkformno = Y.sevkformno " +
    '                                    " and y.sevkformno = z.sevkformno " +
    '                                    " and y.ulineno = z.ulineno " +
    '                                    " and x.ok = 'E' " +
    '                                    " and y.sevkiyattakipno = a.sevkiyattakipno " +
    '                                    " and y.sevkemrino = a.sevkemrino " +
    '                                    " and y.siparisno = a.siparisno " +
    '                                    " and y.modelno = a.modelno), " +
    '                            " gtarih = (select top 1 x.sevktar " +
    '                                    " from sevkform x, sevkformlines y, sevkformlinesrba z " +
    '                                    " Where x.sevkformno = Y.sevkformno " +
    '                                    " and y.sevkformno = z.sevkformno " +
    '                                    " and y.ulineno = z.ulineno " +
    '                                    " and x.ok = 'E' " +
    '                                    " and y.sevkiyattakipno = a.sevkiyattakipno " +
    '                                    " and y.sevkemrino = a.sevkemrino " +
    '                                    " and y.siparisno = a.siparisno " +
    '                                    " and y.modelno = a.modelno " +
    '                                    " order by x.sevktar) "

    '            cSQL = cSQL +
    '                " from sevkplfislines a " +
    '                " where a.siparisno = '" + aSiparis(nCnt) + "' " +
    '                " and a.modelno = '" + cModelNo + "' " +
    '                " and a.toplam is not null " +
    '                " and a.toplam <> 0 " +
    '                " and exists (select siparisno " +
    '                                " from sipmodel " +
    '                                " where siparisno = a.siparisno " +
    '                                " and modelno = a.modelno " +
    '                                " and sevkiyattakipno = a.sevkiyattakipno)) w "

    '            cSQL = cSQL +
    '                " group by w.siparisno, w.modelno, w.ilksevktar " +
    '                " order by w.siparisno, w.modelno, w.ilksevktar "

    '            oReader = GetSQLReader(cSQL, ConnYage)

    '            Do While oReader.Read
    '                nEURFiyat = nFiyat
    '                cEURDoviz = cSipFiyatDoviz
    '                If cSipFiyatDoviz <> "EUR" Then
    '                    GetSipFiyat(ConnYage, aSiparis(nCnt), nEURFiyat, cEURDoviz, "EUR")
    '                End If
    '                ' sipariş / sevkiyat adet ve tutarları
    '                nSipAdet = SQLReadDouble(oReader, "padet")
    '                nSipTutar = nEURFiyat * SQLReadDouble(oReader, "padet")
    '                nSevkiyatAdet = SQLReadDouble(oReader, "gadet")
    '                nSevkTutar = nEURFiyat * SQLReadDouble(oReader, "gadet")
    '                ' gerçekleşen kısım planlanan sevkiyat adedine göre paylaştırılıyor
    '                dSevkTarih = SQLReadDate(oReader, "gtarih")
    '                ' planlanan kısım yine planlanan sevkiyat adedine göre paylaştırılıyor
    '                dPSevkTarih = SQLReadDate(oReader, "ptarih")
    '                If SQLReadDouble(oReader, "tpadet") = 0 Then
    '                    nRatio = 1
    '                Else
    '                    nRatio = SQLReadDouble(oReader, "padet") / SQLReadDouble(oReader, "tpadet")
    '                End If

    '                nCnt2 = nCnt2 + 1
    '                ReDim Preserve aSTI8(nCnt2)

    '                aSTI8(nCnt2).csiparisno = aSiparis(nCnt)
    '                aSTI8(nCnt2).cmodelno = cModelNo
    '                aSTI8(nCnt2).caciklama = cAciklama
    '                aSTI8(nCnt2).nfiyat = nFiyat
    '                aSTI8(nCnt2).cdoviz = cSipFiyatDoviz
    '                aSTI8(nCnt2).nsiparisadet = nSipAdet
    '                aSTI8(nCnt2).nsiparistutar = nSipTutar
    '                aSTI8(nCnt2).nsevkadet = nSevkiyatAdet
    '                aSTI8(nCnt2).nsevktutar = nSevkTutar
    '                aSTI8(nCnt2).g_sevktarihi = dSevkTarih
    '                aSTI8(nCnt2).g_kumastutar = nKumasTutar * nRatio
    '                aSTI8(nCnt2).g_aksesuartutar = nAksesuarTutar * nRatio
    '                aSTI8(nCnt2).g_isciliktutar = nIscilikYutar * nRatio
    '                aSTI8(nCnt2).g_genelgider = nGenelGiderTutar * nRatio
    '                aSTI8(nCnt2).g_digertutar = nDigerTutar * nRatio
    '                aSTI8(nCnt2).p_sevktarihi = dPSevkTarih
    '                aSTI8(nCnt2).p_kumastutar = nPKumasTutar * nRatio
    '                aSTI8(nCnt2).p_aksesuartutar = nPAksesuarTutar * nRatio
    '                aSTI8(nCnt2).p_isciliktutar = nPIscilikYutar * nRatio
    '                aSTI8(nCnt2).p_genelgider = nPGenelGiderTutar * nRatio
    '                aSTI8(nCnt2).p_digertutar = nPDigerTutar * nRatio
    '                aSTI8(nCnt2).cmusterino = cMusteriNo
    '                aSTI8(nCnt2).cdosyakapandi = cDosyaKapandi
    '            Loop
    '            oReader.Close()
    '        Next

    '        If nCnt2 <> -1 Then
    '            For nCnt = 0 To aSTI8.GetUpperBound(0)
    '                cSQL = "set dateformat dmy " +
    '                        " insert stisonmaliyet8 (siparisno, modelno, aciklama, fiyat, doviz, " +
    '                                                " siparisadet, siparistutar, sevkadet, sevktutar, g_sevktarihi, " +
    '                                                " g_kumastutar, g_aksesuartutar, g_isciliktutar, g_genelgider, g_digertutar, " +
    '                                                " p_sevktarihi, p_kumastutar, p_aksesuartutar, p_isciliktutar, p_genelgider, " +
    '                                                " p_digertutar, musterino, dosyakapandi) "
    '                cSQL = cSQL +
    '                        " values ('" + SQLWriteString(aSTI8(nCnt).csiparisno, 30) + "', " +
    '                        " '" + SQLWriteString(aSTI8(nCnt).cmodelno, 30) + "', " +
    '                        " '" + SQLWriteString(aSTI8(nCnt).caciklama, 250) + "', " +
    '                        SQLWriteDecimal(aSTI8(nCnt).nfiyat) + ", " +
    '                        " '" + SQLWriteString(aSTI8(nCnt).cdoviz, 3) + "', "

    '                cSQL = cSQL +
    '                        SQLWriteDecimal(aSTI8(nCnt).nsiparisadet) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).nsiparistutar) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).nsevkadet) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).nsevktutar) + ", " +
    '                        " '" + SQLWriteDate(aSTI8(nCnt).g_sevktarihi) + "', "

    '                cSQL = cSQL +
    '                        SQLWriteDecimal(aSTI8(nCnt).g_kumastutar) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).g_aksesuartutar) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).g_isciliktutar) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).g_genelgider) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).g_digertutar) + ", "

    '                cSQL = cSQL +
    '                        " '" + SQLWriteDate(aSTI8(nCnt).p_sevktarihi) + "', " +
    '                        SQLWriteDecimal(aSTI8(nCnt).p_kumastutar) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).p_aksesuartutar) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).p_isciliktutar) + ", " +
    '                        SQLWriteDecimal(aSTI8(nCnt).p_genelgider) + ", "

    '                cSQL = cSQL +
    '                        SQLWriteDecimal(aSTI8(nCnt).p_digertutar) + ", " +
    '                        " '" + SQLWriteString(aSTI8(nCnt).cmusterino, 30) + "', " +
    '                        " '" + SQLWriteString(aSTI8(nCnt).cdosyakapandi, 1) + "' ) "

    '                ExecuteSQLCommandConnected(cSQL, ConnYage)
    '            Next
    '        End If

    '        ConnYage.Close()
    '        STISonMaliyet8 = 1
    '        JustForLog("STISonMaliyet8 STOP")
    '    Catch ex As Exception
    '        ErrDisp(ex.Message, "STISonMaliyet8", cSQL)
    '    End Try
    'End Function
    Public Function STISonMaliyetIhracat(cFilter As String) As Integer

        Dim cSQL As String = ""
        Dim aSHKur() As oSHKur
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nKur As Double = 0
        Dim aSiparis() As String
        Dim nTutar As Double = 0
        Dim cDoviz As String = ""
        Dim aMasraf() As oMasraf

        STISonMaliyetIhracat = 0

        Try
            JustForLog("STISonMaliyetIhracat START")

            cFilter = Replace(cFilter, "||", "'").Trim

                    cSQL = "select distinct a.kullanicisipno  " +
                   " from siparis a, sipmodel b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and a.kullanicisipno  is not null " +
                   " and a.kullanicisipno  <> '' " +
                   cFilter +
                   " order by a.kullanicisipno  "

            If Not CheckExists(cSQL) Then Exit Function

            ConnYage = OpenConn()

            aSiparis = SQLtoStringArrayConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSiparis)
                nCnt1 = -1
                ReDim aMasraf(0)

                cSQL = "select siparisno, tutar, doviz, kur, tarih  " +
                        " from masraf " +
                        " where siparisno = '" + aSiparis(nCnt) + "' " +
                        " and tutar is not null " +
                        " and tutar <> 0 "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read
                    nCnt1 = nCnt1 + 1
                    ReDim Preserve aMasraf(nCnt1)

                    aMasraf(nCnt1).nKur = SQLReadDouble(oReader, "kur")
                    If aMasraf(nCnt1).dTarih = CDate("01.01.1950") Then
                        aMasraf(nCnt1).dTarih = Today
                    Else
                        aMasraf(nCnt1).dTarih = SQLReadDate(oReader, "tarih")
                    End If
                    aMasraf(nCnt1).nTutar = SQLReadDouble(oReader, "tutar")
                    If SQLReadString(oReader, "doviz") = "" Then
                        aMasraf(nCnt1).cDoviz = "TL"
                    Else
                        aMasraf(nCnt1).cDoviz = SQLReadString(oReader, "doviz")
                    End If
                Loop
                oReader.Close()

                If nCnt1 > -1 Then
                    nTutar = 0
                    For nCnt1 = 0 To UBound(aMasraf)
                        If aMasraf(nCnt1).nKur = 0 Then
                            aMasraf(nCnt1).nKur = GetKur(aMasraf(nCnt1).cDoviz, aMasraf(nCnt1).dTarih, ConnYage)
                        End If

                        If aMasraf(nCnt1).cDoviz = "EUR" Then
                            aMasraf(nCnt1).nEURKur = aMasraf(nCnt1).nKur
                        Else
                            aMasraf(nCnt1).nEURKur = GetKur("EUR", aMasraf(nCnt1).dTarih, ConnYage)
                        End If

                        If aMasraf(nCnt1).nEURKur <> 0 Then
                            nTutar = nTutar + (aMasraf(nCnt1).nTutar * aMasraf(nCnt1).nKur / aMasraf(nCnt1).nEURKur)
                        End If
                    Next

                    If nTutar > 0 Then
                        cSQL = "insert stisonmaliyet (malzemetakipno, anastokgrubu, stoktipi, siralama, eurtutar, birim, ihtiyac, karsilanan) " +
                            " values ('" + aSiparis(nCnt) + "', " +
                            " 'DIGER', " +
                            " 'IHRACAT', " +
                            " 400000, " +
                            SQLWriteDecimal(nTutar) + ", " +
                            " 'AD', " +
                            " 1, " +
                            " 1) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                End If
            Next

            ConnYage.Close()
            STISonMaliyetIhracat = 1
            JustForLog("STISonMaliyetIhracat STOP")

        Catch ex As Exception
            ErrDisp(ex.Message, "STISonMaliyetIhracat", cSQL)
        End Try
    End Function


End Module

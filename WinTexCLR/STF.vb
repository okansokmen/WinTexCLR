Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module STF

    Private Structure oMUSR
        Dim cModelNo As String
        Dim cUTf As String
        Dim cSiparisNo As String
        Dim cRenk As String
    End Structure

    Private Structure oSevk
        Dim cSiparisNo As String
        Dim cModelNo As String
        Dim cBedenSeti As String
        Dim dIlkSevkTarihi As Date
        Dim dSonSevkTarihi As Date
        Dim dEkSevkTarihi1 As Date
        Dim dEkSevkTarihi2 As Date
        Dim cAcenta As String
        Dim nKomisyon As Double
        Dim cTeslimat As String
        Dim cOdemesi As String
        Dim cMusteriNo As String
        Dim cAtolye As String
        Dim cAmbalaj As String
        Dim cMusteriSiparisNo As String
        Dim cGidecegiUlke As String
        Dim cTasimaSekli As String
    End Structure

    Private Structure oSevkPlFisLines
        Dim cSevkEmriNo As String
        Dim cSiparisNo As String
        Dim cModelNo As String
        Dim cBedenSeti As String
        Dim dIlkSevkTarihi As Date
        Dim dSonSevkTarihi As Date
        Dim dEkSevkTarihi1 As Date
        Dim dEkSevkTarihi2 As Date
        Dim cAtolye As String
        Dim cAmbalaj As String
        Dim cMusteriSiparisNo As String
        Dim cGidecegiUlke As String
        Dim cTasimaSekli As String
        Dim nSevkPlFisSiraNo As Double
    End Structure

    Public Function STFFastGenerateAll(cFilter As String) As SqlInt32

        Dim cSQL As String = ""
        Dim aSTF() As String = Nothing
        Dim nCnt As Integer = 0
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""

        STFFastGenerateAll = 0

        Try
            cFilter = Replace(cFilter, "||", "'")
            lAltModelDetay = (GetSysPar("altmodeltakibi") = "1")

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            If Trim(cFilter) = "" Then
                cFilter = " and (a.dosyakapandi = 'H' or a.dosyakapandi = '' or a.dosyakapandi is null) "
            End If

            cSQL = "select distinct b.SevkiyatTakipNo " +
                   " from siparis a, " + cSipModelTableName + " b, ymodel c " +
                   " where a.kullanicisipno = b.siparisno " +
                   " and b.modelno = c.modelno " +
                   " and b.SevkiyatTakipNo is not null " +
                   " and b.SevkiyatTakipNo <> '' " +
                   cFilter +
                   " order by b.sevkiyattakipno "

            If CheckExists(cSQL) Then
                aSTF = SQLtoStringArray(cSQL)
                For nCnt = 0 To UBound(aSTF)
                    If STFGenerate(aSTF(nCnt)) = 0 Then
                        ' hata var
                        Exit Function
                    End If
                Next
            End If

            STFFastGenerateAll = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "STFFastGenerateAll", cSQL)
        End Try
    End Function

    Public Function STFGenerate(cSevkiyatTakipNo As String) As Integer

        Dim cSQL As String = ""
        Dim nSevkPlFisSiraNo As Double = 0
        Dim cSevkEmriNo As String = ""
        Dim nSevkEmriNo As Double = 0
        Dim nSipAdet As Double = 0
        Dim nFiyat As Double = 0
        Dim nSevkAdet As Double = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim dTarih As Date = #1/1/1950#
        Dim aSevk() As oSevk = Nothing
        Dim aSevkPlFisLines() As oSevkPlFisLines = Nothing
        Dim nCnt As Integer = 0
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""
        Dim aMUSR() As oMUSR
        Dim cAtolye As String = ""
        Dim cOK As String = "H"

        STFGenerate = 0

        Try
            ConnYage = OpenConn()

            lAltModelDetay = (GetSysParConnected("altmodeltakibi", ConnYage) = "1")

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            dTarih = GetNowFromServer(ConnYage)

            ' Çeki Listesi Olmayan Sevkiyat Planlamalarını Sil 

            cSQL = "delete sevkplfis " +
                    " where sevkiyattakipno not in (select sevkiyattakipno from sevkform) " +
                    " and sevkiyattakipno not in (select sevkiyattakipno from sevkformlines) " +
                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete sevkplfislines " +
                    " where sevkiyattakipno not in (select sevkiyattakipno from sevkform) " +
                    " and sevkiyattakipno not in (select sevkiyattakipno from sevkformlines) " +
                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete sevkplfisrba " +
                    " where sevkiyattakipno not in (select sevkiyattakipno from sevkform) " +
                    " and sevkiyattakipno not in (select sevkiyattakipno from sevkformlines) " +
                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete sevkplfisgadet " +
                    " where sevkiyattakipno not in (select sevkiyattakipno from sevkform) " +
                    " and sevkiyattakipno not in (select sevkiyattakipno from sevkformlines) " +
                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete sevkplfisfiyat " +
                    " where sevkiyattakipno not in (select sevkiyattakipno from sevkform) " +
                    " and sevkiyattakipno not in (select sevkiyattakipno from sevkformlines) " +
                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "select a.dosyakapandi " +
                    " from siparis a, " + cSipModelTableName + " b " +
                    " where a.kullanicisipno = b.siparisno " +
                    " and b.sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            cOK = SQLGetStringConnected(cSQL, ConnYage)

            If cOK <> "E" Then cOK = "H"

            ' SevkPlFis tablosuna her sevkiyat takip föyü için bir kayıt ekleniyor
            cSQL = "select sevkiyattakipno " +
                    " from sevkplfis " +
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "set dateformat dmy " +
                        " insert into sevkplfis (sevkiyattakipno, verildigitarih, ok, notlar) " +
                        " values ('" + SQLWriteString(cSevkiyatTakipNo) + "', " +
                        " '" + SQLWriteDate(dTarih) + "', " +
                        " '" + cOK.Trim + "', " +
                        " 'CLR-otomatik STF' ) "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            ' Bu durumda sevkplfissirano alanı gereksiz oluyor fakat eskiye uyumluluk için bu alanı update etmeye devam ediyoruz
            cSQL = "select sevkplfissirano " +
                    " from sevkplfis " +
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            nSevkPlFisSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

            If nSevkPlFisSiraNo = 0 Then
                nSevkPlFisSiraNo = 1
            End If

            ' Sevkiyat İşemirleri sevkiyattakipno + siparisno + modelno + bedenseti bazında
            cSQL = "select distinct a.siparisno, a.modelno, a.bedenseti, a.ilksevktar, a.sonsevktar, " +
                   " a.ektermin1, a.ektermin2, a.firma, a.ulke, a.ambalaj, a.musterisiparisno, " +
                   " b.ilksevktarihi, b.sonsevktarihi, b.eksevktarihi1, b.eksevktarihi2, b.acenta, " +
                   " b.komisyon, b.teslimat, b.odemesi, b.musterino, a.tasimasekli " +
                   " from " + cSipModelTableName + " a, siparis b " +
                   " where a.siparisno = b.kullanicisipno " +
                   " and a.sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                   " order by a.siparisno, a.modelno, a.bedenseti "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                ReDim Preserve aSevk(nCnt)

                aSevk(nCnt).cSiparisNo = SQLReadString(oReader, "siparisno")
                aSevk(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aSevk(nCnt).cBedenSeti = SQLReadString(oReader, "bedenseti")
                aSevk(nCnt).cAtolye = SQLReadString(oReader, "firma")
                aSevk(nCnt).cGidecegiUlke = SQLReadString(oReader, "ulke")
                aSevk(nCnt).cAmbalaj = SQLReadString(oReader, "ambalaj")
                aSevk(nCnt).cMusteriSiparisNo = SQLReadString(oReader, "musterisiparisno")
                aSevk(nCnt).cTasimaSekli = SQLReadString(oReader, "tasimasekli")
                aSevk(nCnt).dIlkSevkTarihi = SQLReadDate(oReader, "ilksevktar")
                aSevk(nCnt).dSonSevkTarihi = SQLReadDate(oReader, "sonsevktar")
                aSevk(nCnt).dEkSevkTarihi1 = SQLReadDate(oReader, "ektermin1")
                aSevk(nCnt).dEkSevkTarihi2 = SQLReadDate(oReader, "ektermin2")
                aSevk(nCnt).cAcenta = SQLReadString(oReader, "acenta")
                aSevk(nCnt).nKomisyon = SQLReadDouble(oReader, "komisyon")
                aSevk(nCnt).cTeslimat = SQLReadString(oReader, "teslimat")
                aSevk(nCnt).cOdemesi = SQLReadString(oReader, "odemesi")
                aSevk(nCnt).cMusteriNo = SQLReadString(oReader, "musterino")

                nCnt = nCnt + 1
            Loop
            oReader.Close()

            ' her STF altına en az bir işemri olacak şekilde açılmamış sevk emirlerini aç
            For nCnt = 0 To UBound(aSevk)

                cSQL = "select sevkemrino " +
                        " from sevkplfislines " +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevk(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevk(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevk(nCnt).cBedenSeti + "' " +
                        " and sevkemrino is not null " +
                        " and sevkemrino <> '' "

                If CheckExistsConnected(cSQL, ConnYage) Then

                    cSevkEmriNo = SQLGetStringConnected(cSQL, ConnYage)

                    If cSevkEmriNo.Trim <> "" Then
                        cSQL = "set dateformat dmy " +
                            " update sevkplfislines set " +
                            " atolye = '" + aSevk(nCnt).cAtolye + "', " +
                            " gidecegiulke = '" + aSevk(nCnt).cGidecegiUlke + "', " +
                            " ambalaj = '" + aSevk(nCnt).cAmbalaj + "', " +
                            " musterisiparisno = '" + aSevk(nCnt).cMusteriSiparisNo + "', " +
                            " tasima = '" + aSevk(nCnt).cTasimaSekli + "', " +
                            " ilksevktar = '" + SQLWriteDate(aSevk(nCnt).dIlkSevkTarihi) + "', " +
                            " sonsevktar = '" + SQLWriteDate(aSevk(nCnt).dSonSevkTarihi) + "', " +
                            " ektermin1 = '" + SQLWriteDate(aSevk(nCnt).dEkSevkTarihi1) + "', " +
                            " ektermin2 = '" + SQLWriteDate(aSevk(nCnt).dEkSevkTarihi2) + "', " +
                            " altmusteri = '" + aSevk(nCnt).cMusteriNo + "', " +
                            " komfirma = '" + aSevk(nCnt).cAcenta + "', " +
                            " komisyon = " + SQLWriteDecimal(aSevk(nCnt).nKomisyon)

                        cSQL = cSQL +
                            " where sevkemrino = '" + cSevkEmriNo.Trim + "' " +
                            " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                            " and siparisno = '" + aSevk(nCnt).cSiparisNo + "' " +
                            " and modelno = '" + aSevk(nCnt).cModelNo + "' " +
                            " and bedenseti = '" + aSevk(nCnt).cModelNo + "' "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                Else
                    cSQL = "select count(*) " +
                            " from sevkplfislines " +
                            " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

                    nSevkEmriNo = CDbl(SQLGetIntegerConnected(cSQL, ConnYage) + 1)

                    cSevkEmriNo = Trim(Mid(cSevkiyatTakipNo, 1, 27)) + "_" + Microsoft.VisualBasic.Format(nSevkEmriNo, "00")

                    Do While True
                        cSQL = "select sevkemrino " +
                            " from sevkplfislines " +
                            " where sevkemrino = '" + cSevkEmriNo.Trim + "' "

                        If CheckExistsConnected(cSQL, ConnYage) Then
                            nSevkEmriNo = nSevkEmriNo + 1
                            cSevkEmriNo = Trim(Mid(cSevkiyatTakipNo, 1, 27)) + "_" + Microsoft.VisualBasic.Format(nSevkEmriNo, "00")
                        Else
                            Exit Do
                        End If
                    Loop

                    cSQL = "select max(sevkplfissirano) " +
                            " from sevkplfislines "

                    nSevkPlFisSiraNo = SQLGetDoubleConnected(cSQL, ConnYage) + 1

                    cSQL = "set dateformat dmy " +
                           " insert sevkplfislines " +
                           " (sevkplfissirano, sevkemrino, sevkiyattakipno, siparisno, modelno, " +
                           " bedenseti, altmusteri, ilksevktar, sonsevktar, komfirma, " +
                           " komisyon, ektermin1, ektermin2, gidecegiulke, ok, " +
                           " ambalaj, musterisiparisno, atolye, tasima) "

                    cSQL = cSQL +
                            " values (" + SQLWriteDecimal(nSevkPlFisSiraNo) + ", " +
                            " '" + cSevkEmriNo + "', " +
                            " '" + cSevkiyatTakipNo + "', " +
                            " '" + aSevk(nCnt).cSiparisNo + "', " +
                            " '" + aSevk(nCnt).cModelNo + "', "

                    cSQL = cSQL +
                            " '" + aSevk(nCnt).cBedenSeti + "', " +
                            " '" + aSevk(nCnt).cMusteriNo + "', " +
                            " '" + SQLWriteDate(aSevk(nCnt).dIlkSevkTarihi) + "', " +
                            " '" + SQLWriteDate(aSevk(nCnt).dSonSevkTarihi) + "', " +
                            " '" + aSevk(nCnt).cAcenta + "', "

                    cSQL = cSQL +
                            SQLWriteDecimal(aSevk(nCnt).nKomisyon) + ", " +
                            " '" + SQLWriteDate(aSevk(nCnt).dEkSevkTarihi1) + "', " +
                            " '" + SQLWriteDate(aSevk(nCnt).dEkSevkTarihi2) + "', " +
                            " '" + aSevk(nCnt).cGidecegiUlke + "', " +
                            " 'H', "

                    cSQL = cSQL +
                            " '" + aSevk(nCnt).cAmbalaj + "', " +
                            " '" + aSevk(nCnt).cMusteriSiparisNo + "', " +
                            " '" + aSevk(nCnt).cAtolye + "', " +
                            " '" + aSevk(nCnt).cTasimaSekli + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            nCnt = 0

            cSQL = "select sevkemrino, siparisno, modelno, bedenseti, ilksevktar, sonsevktar, " +
                    " ektermin1, ektermin2, atolye, ambalaj, gidecegiulke, " +
                    " musterisiparisno, tasima, sevkplfissirano " +
                    " from sevkplfislines " +
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                    " order by sevkemrino, siparisno, modelno, bedenseti "

            If CheckExistsConnected(cSQL, ConnYage) Then

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read

                    ReDim Preserve aSevkPlFisLines(nCnt)

                    aSevkPlFisLines(nCnt).cSevkEmriNo = SQLReadString(oReader, "sevkemrino")
                    aSevkPlFisLines(nCnt).cSiparisNo = SQLReadString(oReader, "siparisno")
                    aSevkPlFisLines(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                    aSevkPlFisLines(nCnt).cBedenSeti = SQLReadString(oReader, "bedenseti")
                    aSevkPlFisLines(nCnt).dIlkSevkTarihi = SQLReadDate(oReader, "ilksevktar")
                    aSevkPlFisLines(nCnt).dSonSevkTarihi = SQLReadDate(oReader, "sonsevktar")
                    aSevkPlFisLines(nCnt).dEkSevkTarihi1 = SQLReadDate(oReader, "ektermin1")
                    aSevkPlFisLines(nCnt).dEkSevkTarihi2 = SQLReadDate(oReader, "ektermin2")
                    aSevkPlFisLines(nCnt).cAtolye = SQLReadString(oReader, "atolye")
                    aSevkPlFisLines(nCnt).cAmbalaj = SQLReadString(oReader, "ambalaj")
                    aSevkPlFisLines(nCnt).cGidecegiUlke = SQLReadString(oReader, "gidecegiulke")
                    aSevkPlFisLines(nCnt).cMusteriSiparisNo = SQLReadString(oReader, "musterisiparisno")
                    aSevkPlFisLines(nCnt).cTasimaSekli = SQLReadString(oReader, "tasima")
                    aSevkPlFisLines(nCnt).nSevkPlFisSiraNo = SQLReadDouble(oReader, "sevkplfissirano")

                    nCnt = nCnt + 1
                Loop
                oReader.Close()

                For nCnt = 0 To UBound(aSevkPlFisLines)
                    cSQL = "select sum(coalesce(adet,0)) " +
                        " from " + cSipModelTableName +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' "

                    nSipAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                    cSQL = "select sum((b.koliend - b.kolibeg + 1) * c.adet) " +
                        " from sevkform a, sevkformlines b, sevkformlinesrba c " +
                        " where a.sevkformno = b.sevkformno " +
                        " and b.sevkformno = c.sevkformno " +
                        " and b.ulineno = c.ulineno " +
                        " and a.ok = 'E' " +
                        " and (a.sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' or b.sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "') " +
                        " and b.siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and b.modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and b.bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " and b.sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                    nSevkAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                    cSQL = "update sevkplfislines " +
                        " set toplam = " + SQLWriteDecimal(nSipAdet) + " , " +
                        " planlanan = " + SQLWriteDecimal(nSipAdet) + " , " +
                        " giden = " + SQLWriteDecimal(nSevkAdet) +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete sevkplfisrba " +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "insert sevkplfisrba " +
                        " (sevkiyattakipno, sevkplfissirano, sevkemrino, siparisno, modelno, " +
                        " bedenseti, renk, beden, adet) "

                    cSQL = cSQL +
                        " select sevkiyattakipno, " +
                        " sevkplfissirano = " + SQLWriteDecimal(aSevkPlFisLines(nCnt).nSevkPlFisSiraNo) + ", " +
                        " sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "', " +
                        " siparisno, modelno, bedenseti, renk, beden, " +
                        " adet = sum(coalesce(adet,0)) " +
                        " from " + cSipModelTableName +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " group by sevkiyattakipno, siparisno, modelno, bedenseti, renk, beden"

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete sevkplfisgadet " +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "insert sevkplfisgadet " +
                        " (sevkiyattakipno, sevkplfissirano, sevkemrino, siparisno, modelno, " +
                        " bedenseti, renk, beden, adet) "

                    cSQL = cSQL +
                        " select sevkiyattakipno, " +
                        " sevkplfissirano = " + SQLWriteDecimal(aSevkPlFisLines(nCnt).nSevkPlFisSiraNo) + ", " +
                        " sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "', " +
                        " siparisno, modelno, bedenseti, renk, beden, " +
                        " adet = sum(coalesce(adet,0)) " +
                        " from " + cSipModelTableName +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " group by sevkiyattakipno, siparisno, modelno, bedenseti, renk, beden"

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete sevkplfisfiyat " +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "insert sevkplfisfiyat " +
                        " (sevkiyattakipno, sevkplfissirano, sevkemrino, siparisno, modelno, " +
                        " bedenseti, renk, beden, fiyat) "

                    cSQL = cSQL +
                        " select sevkiyattakipno, " +
                        " sevkplfissirano = " + SQLWriteDecimal(aSevkPlFisLines(nCnt).nSevkPlFisSiraNo) + ", " +
                        " sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "', " +
                        " siparisno, modelno, bedenseti, renk, beden, " +
                        " fiyat = (select top 1 satisfiyati " +
                                    " from sipfiyat " +
                                    " where siparisno = " + cSipModelTableName + ".siparisno " +
                                    " and modelkodu = " + cSipModelTableName + ".modelno " +
                                    " and (sevkiyattakipno = " + cSipModelTableName + ".sevkiyattakipno or sevkiyattakipno = 'HEPSI') " +
                                    " and (renk = " + cSipModelTableName + ".renk or renk = 'HEPSI') " +
                                    " and (beden = " + cSipModelTableName + ".beden or beden = 'HEPSI') ) "

                    cSQL = cSQL +
                        " from " + cSipModelTableName +
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " +
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " +
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " +
                        " group by sevkiyattakipno, siparisno, modelno, bedenseti, renk, beden"

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Next
            End If

            ' Sevkiyat İşemirleri otomatik kapatma
            cSQL = "update sevkplfislines " +
                    " set ok = 'E' " +
                    " where coalesce(planlanan,0) <= coalesce(giden,0) " +
                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update sevkplfislines " +
                    " set oktarihi = (select top 1 a.sevktar " +
                                    " from sevkform a, sevkformlines b " +
                                    " where a.sevkformno = b.sevkformno " +
                                    " and b.sevkemrino = sevkplfislines.sevkemrino " +
                                    " and b.siparisno = sevkplfislines.siparisno " +
                                    " and b.modelno = sevkplfislines.modelno " +
                                    " and b.bedenseti = sevkplfislines.bedenseti " +
                                    " and a.sevktar is not null " +
                                    " and a.sevktar <> '01.01.1950' " +
                                    " and a.ok = 'E' " +
                                    " order by a.sevktar desc) " +
                    " where ok = 'E' " +
                    " and ok is not null " +
                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                    " and (oktarihi is null or oktarihi = '01.01.1950') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' Sevkiyat Takip Föyü Otomatik Kapatma
            cSQL = "update sevkplfis " +
                    " set ok = 'E' " +
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                    " and not exists (select sevkiyattakipno " +
                                    " from sevkplfislines " +
                                    " where sevkiyattakipno = sevkplfis.sevkiyattakipno " +
                                    " and (ok is null or ok = '' or ok = 'H')) " +
                    " and coalesce((select sum(coalesce(adet,0)) from sipmodel where sevkiyattakipno = sevkplfis.sevkiyattakipno),0) <= coalesce((select sum(coalesce(giden,0)) from sevkplfislines where sevkiyattakipno = sevkplfis.sevkiyattakipno),0)"

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' atolyeleri tamamla
            cSQL = "select atolye " +
                    " from sevkplfislines " +
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                    " and (atolye is null or atolye = '') "

            If CheckExistsConnected(cSQL, ConnYage) Then
                nCnt = -1
                ReDim aMUSR(0)

                ' boş atölyeleri doldur
                cSQL = "select distinct modelno, uretimtakipno, siparisno, renk " +
                   " from sipmodel " +
                   " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read
                    nCnt = nCnt + 1
                    ReDim Preserve aMUSR(nCnt)

                    aMUSR(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                    aMUSR(nCnt).cUTf = SQLReadString(oReader, "uretimtakipno")
                    aMUSR(nCnt).cSiparisNo = SQLReadString(oReader, "siparisno")
                    aMUSR(nCnt).cRenk = SQLReadString(oReader, "renk")
                Loop
                oReader.Close()

                If nCnt <> -1 Then
                    For nCnt = 0 To UBound(aMUSR)

                        cSQL = "select top 1 firma " +
                                " from sipmodel " +
                                " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                                " and uretimtakipno = '" + aMUSR(nCnt).cUTf + "' " +
                                " and siparisno = '" + aMUSR(nCnt).cSiparisNo + "' " +
                                " and modelno = '" + aMUSR(nCnt).cModelNo + "' " +
                                " and renk = '" + aMUSR(nCnt).cRenk + "' " +
                                " and firma is not null " +
                                " and firma <> '' "

                        cAtolye = SQLGetStringConnected(cSQL, ConnYage)

                        If cAtolye.Trim = "" Then
                            cSQL = "select top 1 atolye " +
                                    " from sevkformlines " +
                                    " where siparisno = '" + aMUSR(nCnt).cSiparisNo + "' " +
                                    " and modelno = '" + aMUSR(nCnt).cModelNo + "' " +
                                    " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                                    " and atolye is not null  " +
                                    " and atolye <> '' "

                            cAtolye = SQLGetStringConnected(cSQL, ConnYage)
                        End If

                        If cAtolye.Trim = "" Then
                            cSQL = "select top 1 a.girisfirm_atl  " +
                                    " from uretharfis a, uretharfislines b " +
                                    " where a.uretfisno = b.uretfisno " +
                                    " and b.modelno = '" + aMUSR(nCnt).cModelNo + "' " +
                                    " and b.uretimtakipno = '" + aMUSR(nCnt).cUTf + "' " +
                                    " and a.girisdept like 'SEVK%' " +
                                    " and a.girisfirm_atl is not null " +
                                    " and a.girisfirm_atl <> '' "

                            cAtolye = SQLGetStringConnected(cSQL, ConnYage)
                        End If

                        If cAtolye.Trim = "" Then
                            cSQL = "select top 1 a.firma  " +
                                    " from uretimisemri a, uretimisdetayi b " +
                                    " where a.uretimtakipno = b.uretimtakipno " +
                                    " and a.isemrino = b.isemrino " +
                                    " and b.modelno = '" + aMUSR(nCnt).cModelNo + "' " +
                                    " and a.uretimtakipno = '" + aMUSR(nCnt).cUTf + "' " +
                                    " and a.departman like 'SEVK%' " +
                                    " and a.firma is not null " +
                                    " and a.firma <> '' "

                            cAtolye = SQLGetStringConnected(cSQL, ConnYage)
                        End If

                        If cAtolye.Trim = "" Then
                            cSQL = "select top 1 plfirma  " +
                                    " from uretpllines " +
                                    " where modelno = '" + aMUSR(nCnt).cModelNo + "' " +
                                    " and uretimtakipno = '" + aMUSR(nCnt).cUTf + "' " +
                                    " and departman like 'SEVK%' " +
                                    " and plfirma is not null " +
                                    " and plfirma <> '' "

                            cAtolye = SQLGetStringConnected(cSQL, ConnYage)
                        End If

                        If cAtolye.Trim <> "" Then
                            cSQL = "update sevkplfislines " +
                                   " set atolye = '" + cAtolye.Trim + "' " +
                                   " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                                   " and siparisno = '" + aMUSR(nCnt).cSiparisNo + "' " +
                                   " and modelno = '" + aMUSR(nCnt).cModelNo + "' " +
                                   " and (atolye is null or atolye = '') "

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End If
                    Next
                End If

                ' en son boş kalmış atolyeleri DAHILI olarak doldur
                cSQL = "update sevkplfislines " +
                       " set atolye = 'DAHILI' " +
                       " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " +
                       " and (atolye is null or atolye = '') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

            End If

            ConnYage.Close()

            STFGenerate = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "STFGenerate", cSQL)
        End Try
    End Function

End Module

Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module STF

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
    End Structure

    Public Sub STFFastGenerateAll()

        Dim cSQL As String = ""
        Dim aSTF() As String = Nothing
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

            cSQL = "select distinct a.sevkiyattakipno " + _
                    " from " + cSipModelTableName + " a, siparis b  " + _
                    " where a.siparisno = b.kullanicisipno " + _
                    " and a.sevkiyattakipno is not null " + _
                    " and a.sevkiyattakipno <> '' " + _
                    " and (b.dosyakapandi = 'H' or b.dosyakapandi = '' or b.dosyakapandi is null) " + _
                    " order by a.sevkiyattakipno "

            If CheckExists(cSQL) Then
                aSTF = SQLtoStringArray(cSQL)
                For nCnt = 0 To UBound(aSTF)
                    STFGenerate(aSTF(nCnt))
                Next
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "STFFastGenerateAll", cSQL)
        End Try
    End Sub

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

            cSQL = "select sevkiyattakipno " + _
                    " from sevkiyattakipno " + _
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "set dateformat dmy " + _
                        " insert into sevkplfis (sevkiyattakipno, verildigitarih, ok, notlar) " + _
                        " values ('" + SQLWriteString(cSevkiyatTakipNo) + "', " + _
                        " '" + SQLWriteDate(dTarih) + "', " + _
                        " 'H', " + _
                        " 'CLR-otomatik STF' ) "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            cSQL = "select sevkplfissirano " + _
                    " from sevkplfis " + _
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

            nSevkPlFisSiraNo = SQLGetDoubleConnected(cSQL, ConnYage)

            ' Sevkiyat İşemirleri sevkiyattakipno + siparisno + modelno + bedenseti bazında
            cSQL = "select distinct a.siparisno, a.modelno, a.bedenseti, " + _
                    " b.ilksevktarihi, b.sonsevktarihi, b.eksevktarihi1, b.eksevktarihi2, b.acenta, " + _
                    " b.komisyon, b.teslimat, b.odemesi, b.musterino " + _
                    " from " + cSipModelTableName + " a, siparis b " + _
                    " where a.siparisno = b.kullanicisipno " + _
                    " and a.sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                    " order by a.siparisno, a.modelno, a.bedenseti "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aSevk(nCnt)
                aSevk(nCnt).cSiparisNo = SQLReadString(oReader, "siparisno")
                aSevk(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aSevk(nCnt).cBedenSeti = SQLReadString(oReader, "bedenseti")
                aSevk(nCnt).dIlkSevkTarihi = SQLReadDate(oReader, "ilksevktarihi")
                aSevk(nCnt).dSonSevkTarihi = SQLReadDate(oReader, "sonsevktarihi")
                aSevk(nCnt).dEkSevkTarihi1 = SQLReadDate(oReader, "eksevktarihi1")
                aSevk(nCnt).dEkSevkTarihi2 = SQLReadDate(oReader, "eksevktarihi2")
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
                cSQL = "select sevkiyattakipno " + _
                        " from sevkplfislines " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

                If Not CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "select count(*) " + _
                            " from sevkplfislines " + _
                            " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' "

                    nSevkEmriNo = SQLGetDoubleConnected(cSQL, ConnYage) + 1

                    cSevkEmriNo = Trim(Mid(cSevkiyatTakipNo, 1, 25)) + "-" + Microsoft.VisualBasic.Format(nSevkEmriNo, "0000")

                    cSQL = "set dateformat dmy " + _
                            " insert into sevkplfislines " + _
                            " (sevkplfissirano, sevkemrino, sevkiyattakipno, siparisno, modelno, " + _
                            " bedenseti, altmusteri, ilksevktar, sonsevktar, komfirma, " + _
                            " komisyon, ektermin1, ektermin2, ok) "

                    cSQL = cSQL + _
                            " values (" + SQLWriteDecimal(nSevkPlFisSiraNo) + ", " + _
                            " '" + cSevkEmriNo + "', " + _
                            " '" + cSevkiyatTakipNo + "', " + _
                            " '" + aSevk(nCnt).cSiparisNo + "', " + _
                            " '" + aSevk(nCnt).cModelNo + "', "

                    cSQL = cSQL + _
                            " '" + aSevk(nCnt).cBedenSeti + "', " + _
                            " '" + aSevk(nCnt).cMusteriNo + "', " + _
                            " '" + SQLWriteDate(aSevk(nCnt).dIlkSevkTarihi) + "', " + _
                            " '" + SQLWriteDate(aSevk(nCnt).dSonSevkTarihi) + "', " + _
                            " '" + aSevk(nCnt).cAcenta + "', "

                    cSQL = cSQL + _
                            SQLWriteDecimal(aSevk(nCnt).nKomisyon) + ", " + _
                            " '" + SQLWriteDate(aSevk(nCnt).dEkSevkTarihi1) + "', " + _
                            " '" + SQLWriteDate(aSevk(nCnt).dEkSevkTarihi2) + "', " + _
                            " 'H') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            nCnt = 0

            cSQL = "select sevkemrino, siparisno, modelno, bedenseti, ilksevktar, sonsevktar, ektermin1, ektermin2  " + _
                    " from sevkplfislines " + _
                    " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                    " order by sevkemrino, siparisno, modelno, bedenseti "

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
                nCnt = nCnt + 1
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aSevkPlFisLines)
                cSQL = "select sum(coalesce(adet,0)) " + _
                        " from " + cSipModelTableName + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' "

                nSipAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = "select sum((b.koliend - b.kolibeg + 1) * c.adet) " + _
                        " from sevkform a, sevkformlines b, sevkformlinesrba c " + _
                        " where a.sevkformno = b.sevkformno " + _
                        " and b.sevkformno = c.sevkformno " + _
                        " and b.ulineno = c.ulineno " + _
                        " and (a.sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' or b.sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "') " + _
                        " and b.siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and b.modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and b.bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and b.sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                nSevkAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                cSQL = "update sevkplfislines " + _
                        " set sevkplfissirano = " + SQLWriteDecimal(nSevkPlFisSiraNo) + ", " + _
                        " toplam = " + SQLWriteDecimal(nSipAdet) + " , " + _
                        " planlanan = " + SQLWriteDecimal(nSipAdet) + " , " + _
                        " giden = " + SQLWriteDecimal(nSevkAdet) + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "set dateformat dmy " + _
                        " update sevkplfislines " + _
                        " set ilksevktar = '" + SQLWriteDate(aSevkPlFisLines(nCnt).dIlkSevkTarihi) + "' " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' " + _
                        " and (ilksevktar is null or ilksevktar = '01.01.1950') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "set dateformat dmy " + _
                        " update sevkplfislines " + _
                        " set sonsevktar = '" + SQLWriteDate(aSevkPlFisLines(nCnt).dSonSevkTarihi) + "' " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' " + _
                        " and (sonsevktar is null or sonsevktar = '01.01.1950') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "set dateformat dmy " + _
                        " update sevkplfislines " + _
                        " set ektermin1 = '" + SQLWriteDate(aSevkPlFisLines(nCnt).dEkSevkTarihi1) + "' " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' " + _
                        " and (ektermin1 is null or ektermin1 = '01.01.1950') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "set dateformat dmy " + _
                        " update sevkplfislines " + _
                        " set ektermin2 = '" + SQLWriteDate(aSevkPlFisLines(nCnt).dEkSevkTarihi2) + "' " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' " + _
                        " and (ektermin2 is null or ektermin2 = '01.01.1950') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                ' detayları silerken sevkplfissirano kontrol edilmiyor
                ' Böylece NULL olarak yanlışlıkla açılmış kayıtlar otomatikman siliniyor
                cSQL = "delete sevkplfisrba " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "insert sevkplfisrba " + _
                        " (sevkiyattakipno, sevkplfissirano, sevkemrino, siparisno, modelno, " + _
                        " bedenseti, renk, beden, adet) "

                cSQL = cSQL + _
                        " select sevkiyattakipno, " + _
                        " sevkplfissirano = " + SQLWriteDecimal(nSevkPlFisSiraNo) + ", " + _
                        " sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "', " + _
                        " siparisno, modelno, bedenseti, renk, beden, " + _
                        " adet = sum(coalesce(adet,0)) " + _
                        " from " + cSipModelTableName + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " group by sevkiyattakipno, siparisno, modelno, bedenseti, renk, beden"

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "delete sevkplfisgadet " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "insert sevkplfisgadet " + _
                        " (sevkiyattakipno, sevkplfissirano, sevkemrino, siparisno, modelno, " + _
                        " bedenseti, renk, beden, adet) "

                cSQL = cSQL + _
                        " select sevkiyattakipno, " + _
                        " sevkplfissirano = " + SQLWriteDecimal(nSevkPlFisSiraNo) + ", " + _
                        " sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "', " + _
                        " siparisno, modelno, bedenseti, renk, beden, " + _
                        " adet = sum(coalesce(adet,0)) " + _
                        " from " + cSipModelTableName + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " group by sevkiyattakipno, siparisno, modelno, bedenseti, renk, beden"

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "delete sevkplfisfiyat " + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " and sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = "insert sevkplfisfiyat " + _
                        " (sevkiyattakipno, sevkplfissirano, sevkemrino, siparisno, modelno, " + _
                        " bedenseti, renk, beden, fiyat) "

                cSQL = cSQL + _
                        " select sevkiyattakipno, " + _
                        " sevkplfissirano = " + SQLWriteDecimal(nSevkPlFisSiraNo) + ", " + _
                        " sevkemrino = '" + aSevkPlFisLines(nCnt).cSevkEmriNo + "', " + _
                        " siparisno, modelno, bedenseti, renk, beden, " + _
                        " fiyat = (select satisfiyati " + _
                                    " from sipfiyat " + _
                                    " where siparisno = " + cSipModelTableName + ".siparisno " + _
                                    " and modelkodu = " + cSipModelTableName + ".modelno " + _
                                    " and (sevkiyattakipno = " + cSipModelTableName + ".sevkiyattakipno or sevkiyattakipno = 'HEPSI') " + _
                                    " and (renk = " + cSipModelTableName + ".renk or renk = 'HEPSI') " + _
                                    " and (beden = " + cSipModelTableName + ".beden or beden = 'HEPSI') ) "

                cSQL = cSQL + _
                        " from " + cSipModelTableName + _
                        " where sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' " + _
                        " and siparisno = '" + aSevkPlFisLines(nCnt).cSiparisNo + "' " + _
                        " and modelno = '" + aSevkPlFisLines(nCnt).cModelNo + "' " + _
                        " and bedenseti = '" + aSevkPlFisLines(nCnt).cBedenSeti + "' " + _
                        " group by sevkiyattakipno, siparisno, modelno, bedenseti, renk, beden"

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next

            ConnYage.Close()

            STFGenerate = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "STFGenerate", cSQL)
        End Try
    End Function

End Module

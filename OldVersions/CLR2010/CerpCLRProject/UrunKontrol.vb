Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module UrunKontrol

    Public Function BarkodluUrunKontrol(ByVal cukFisNo As String, ByVal cDepo As String, ByVal cFirma As String, ByVal cTarih As String) As String

        Dim dTarih As Date
        Dim cSQL As String = ""
        Dim aUKFisNo() As String
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nCnt As Integer = 0
        Dim cSiparisNo As String = ""
        Dim cModelNo As String = ""
        Dim cAciklama As String = ""
        Dim cRenk As String = ""
        Dim cBeden As String = ""
        Dim cAnamodeltipi As String = ""
        Dim cGTIP As String = ""
        Dim cAsortino As String = ""
        Dim cicerik As String = ""
        Dim cStokNo As String = ""
        Dim nAgirlik As Double = 0
        Dim cKoliFisno As String = ""
        Dim nAdet As Double = 0
        Dim nStokMiktari As Double = 0
        Dim cAsortiBedenSeti As String = ""
        Dim nAsortidekiAdet As Double = 0
        Dim nAsortiBedenCount As Double = 0
        Dim nFiyat As Double = 0
        Dim cFiyat As String = ""
        Dim nPoz As Integer = 0
        Dim cBoy As String = ""

        BarkodluUrunKontrol = "OK"

        Try

            ConnYage = OpenConn()

            dTarih = GetNowFromServer(ConnYage)

            ReDim aUKFisNo(0)

            cSQL = "select distinct barcode " + _
                    " from barkodluurunkontrol " + _
                    " where ukfisno = '" + cukFisNo.Trim + "' " + _
                    " order by barcode "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aUKFisNo(nCnt)
                aUKFisNo(nCnt) = SQLReadString(oReader, "barcode")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            cSQL = "delete from urunkontrol where ukfisno = '" + cukFisNo + "' "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from urunkontrollines where ukfisno = '" + cukFisNo + "' "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = " insert into urunkontrol (ukfisno,fistarihi,firma,depo,creationdate,modificationdate,username) " + _
                   " values ('" + cukFisNo + "', " + _
                   " '" + cTarih + "', " + _
                   " '" + cFirma + "', " + _
                   " '" + cDepo + "', " + _
                   " '" + SQLWriteDate(dTarih) + "', " + _
                   " '" + SQLWriteDate(dTarih) + "', " + _
                   " 'CLR')"

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

            For nCnt = 0 To UBound(aUKFisNo)

                cSiparisNo = ""
                cModelNo = ""
                cRenk = ""
                cBeden = ""
                cAciklama = ""
                cAnamodeltipi = ""
                cGTIP = ""

                cSQL = "select top 1 a.siparisno,a.stokno,a.renk,a.beden,b.aciklama,b.modelno,b.anamodeltipi,b.gtip  " + _
                        " from stokbarkod a, ymodel b " + _
                        " where (a.barcode = '" + aUKFisNo(nCnt) + "' or a.barcode2 = '" + aUKFisNo(nCnt) + "') " + _
                        " and a.stokno = b.entegrekodu "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    cSiparisNo = SQLReadString(oReader, "SiparisNo")
                    cModelNo = SQLReadString(oReader, "ModelNo")
                    cRenk = SQLReadString(oReader, "renk")
                    cBeden = SQLReadString(oReader, "beden")
                    cAciklama = SQLReadString(oReader, "Aciklama")
                    cAnamodeltipi = SQLReadString(oReader, "anamodeltipi")
                    cGTIP = SQLReadString(oReader, "gtip")
                End If
                oReader.Close()
                oReader = Nothing
                ' asorti

                cAsortino = ""

                cSQL = "select asortino from sipasorti where siparisno = '" + cSiparisNo + "' "
                cAsortino = ReadSingleValueConnected(cSQL, ConnYage)

                If cAsortino = "" Then
                    cSQL = "select asortino " + _
                            " from modelasorti " + _
                            " where modelno = '" + cModelNo + "' " + _
                            " and asortino is not null " + _
                            " and asortino <> '' " + _
                            " and exists (select adet from firmaasorti where asortino = modelasorti.asortino and adet > 0) " + _
                            " order by sayac desc "
                    cAsortino = ReadSingleValueConnected(cSQL, ConnYage)
                End If

                ' içerik 

                cicerik = ""

                cSQL = "select hammaddekodu " + _
                        " from modelhammadde " + _
                        " where modelno = '" + cModelNo + "' " + _
                        " and anakumas = 'E' " + _
                        " union all " + _
                        " select hammaddekodu " + _
                        " from modelhammadde2 " + _
                        " where modelno = '" + cModelNo + "' " + _
                        " and anakumas = 'E' "
                cStokNo = ReadSingleValueConnected(cSQL, ConnYage)

                cSQL = "select icerik from stokdokuma where stokno = '" + cStokNo + "' "
                cicerik = ReadSingleValueConnected(cSQL, ConnYage)

                If cicerik = "" Then
                    cSQL = "select icerik from stokorme where stokno = '" + cStokNo + "' "
                    cicerik = ReadSingleValueConnected(cSQL, ConnYage)
                End If

                If cicerik = "" Then
                    cicerik = "Problem, reçete"
                End If

                ' ağırlık

                nAgirlik = 0
                cKoliFisno = ""

                cSQL = "select top 1 b.kolifisno, b.netagirlik " + _
                        " from kolilines a, kolileme b  " + _
                        " where modelno = '" + cModelNo + "' " + _
                        " and a.kolifisno = b.kolifisno " + _
                        " and b.netagirlik is not null " + _
                        " and b.netagirlik <> 0 "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nAgirlik = SQLReadDouble(oReader, "netagirlik")
                    cKoliFisno = SQLReadString(oReader, "kolifisno")
                End If
                oReader.Close()
                oReader = Nothing
                nAdet = 0

                cSQL = "select adet = sum(coalesce(adet,0)) " + _
                        " from kolilines " + _
                        " where kolifisno = '" + cKoliFisno + "' "
                nAdet = ReadSingleDoubleValueConnected(cSQL, ConnYage)

                If nAdet <> 0 Then
                    nAgirlik = nAgirlik / nAdet
                End If

                ' stok miktarı

                nStokMiktari = 0

                If cAsortino = "TEKLEME" Then

                    cSQL = "select adet = coalesce(sum(coalesce(donemgiris1,0)) - sum(coalesce(donemcikis1,0)),0) " + _
                            " from stokrb " + _
                            " where stokno = '" + cModelNo + "' " + _
                            " and renk = '" + cRenk + "' " + _
                            " and beden = '" + cBeden + "' " + _
                            " and depo = 'KIRIK DEPO' "
                    nStokMiktari = ReadSingleDoubleValueConnected(cSQL, ConnYage)
                Else
                    cAsortiBedenSeti = ""
                    nAsortidekiAdet = 0

                    cSQL = "select sum(coalesce(adet,0)) " + _
                            " from firmaasorti " + _
                            " where asortino = '" + cAsortino + "' " + _
                            " and bedenseti is not null " + _
                            " and beden is not null " + _
                            " and adet is not null "
                    nAsortidekiAdet = ReadSingleDoubleValueConnected(cSQL, ConnYage)

                    cSQL = "select bedenseti " + _
                            " from firmaasorti " + _
                            " where asortino = '" + cAsortino + "' " + _
                            " and bedenseti is not null " + _
                            " and beden is not null " + _
                            " and adet is not null "
                    cAsortiBedenSeti = ReadSingleValueConnected(cSQL, ConnYage)

                    cSQL = "select coalesce(sum(coalesce(donemgiris1,0)) - sum(coalesce(donemcikis1,0)),0) " + _
                            " from stokrb " + _
                            " where stokno = '" + cModelNo + "' " + _
                            " and renk = '" + cRenk + "' " + _
                            " and depo = 'MAMUL DEPO' " + _
                            " and malzemetakipkodu is not null " + _
                            " and exists (select siparisno from sipmodel " + _
                                            " where malzemetakipno = stokrb.malzemetakipkodu " + _
                                            " and bedenseti = '" + cAsortiBedenSeti + "')"
                    nStokMiktari = ReadSingleDoubleValueConnected(cSQL, ConnYage)

                    If nAsortidekiAdet <> 0 Then
                        nStokMiktari = nStokMiktari / nAsortidekiAdet
                    End If
                End If

                ' fiyat
                nFiyat = 0
                cFiyat = ""

                If cAsortino <> "TEKLEME" Then

                    cSQL = "select top 1 coalesce(maliyet,0)  " + _
                        " from ModelAstFiyat " + _
                        " where ModelNo = '" + cModelNo + "' " + _
                        " and asortino = '" + cAsortino + "' " + _
                        " and maliyet is not null " + _
                        " and maliyet <> 0 "
                Else
                    nPoz = InStr(cAsortino, "/")

                    cBoy = Mid(cAsortino, nPoz + 1, 2)

                    cSQL = "select top 1 coalesce(maliyet,0)  " + _
                        " from ModelAstFiyat " + _
                        " where ModelNo = '" + cModelNo + "' " + _
                        " " + IIf(IsNumeric(cBoy), " and BedenSeti Like '%" + cBoy + "%' ", "").ToString + _
                        " and maliyet is not null " + _
                        " and maliyet <> 0 "
                End If

                nFiyat = ReadSingleDoubleValueConnected(cSQL, ConnYage)
                If nFiyat = 0 Then
                    cFiyat = "Eksik"
                Else
                    cFiyat = "OK"
                End If

                ' update stok

                cSQL = "select stokno from stok where stokno = '" + cModelNo + "' "
                If Not CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "insert into stok " + _
                            "(stokno,cinsaciklamasi,stoktipi,birimseti,entegrekodu,temindepartmani,uretimecikisdepo, " + _
                            " anastokgrubu,paratakipesasi,maltakipesasi,birim1) " + _
                            " values " + _
                            "('" + cModelNo + "'," + _
                            " '" + cAciklama + "'," + _
                            " '" + cAnamodeltipi + "'," + _
                            " '001'," + _
                            " '" + cModelNo + "'," + _
                            " 'MAMUL','MAMUL DEPO','MAMUL','4','4','AD')"
                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                ' update details
                cSQL = " insert into urunkontrollines (ukfisno, siparisno,modelno, renk, beden, asortino, aciklama," + _
                        " anamodeltipi, gtip, icerik, agirlik, miktar, fiyatok, barcode, mail) " + _
                        " values ('" + cukFisNo + "', " + _
                        " '" + cSiparisNo + "', " + _
                        " '" + cModelNo + "', " + _
                        " '" + cRenk + "', " + _
                        " '" + cBeden + "', " + _
                        " '" + cAsortino + "', " + _
                        " '" + cAciklama + "', " + _
                        " '" + cAnamodeltipi + "', " + _
                        " '" + cGTIP + "', " + _
                        " '" + cicerik + "', " + _
                        SQLWriteDecimal(nAgirlik) + ", " + _
                        SQLWriteDecimal(nStokMiktari) + ", " + _
                        " '" + cFiyat + "', " + _
                        " '" + aUKFisNo(nCnt) + "', 'H') "
                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next

            CloseConn(ConnYage)

            BarkodluUrunKontrol = "OK"

        Catch Err As Exception
            BarkodluUrunKontrol = "Hata"
            ErrDisp("BarkodluUrunKontrol : " + Err.Message.Trim + vbCrLf + cSQL)
        End Try
    End Function
End Module

Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module Paketleme

    Private Structure FisLines
        Dim cSiparisNo As String
        Dim cStokNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cBedenSeti As String
        Dim cMTF As String
        Dim nMiktar As Double
        Dim nFiyat As Double
        Dim cDoviz As String
        Dim cDepo As String
    End Structure

    Public Sub Cuvalla(ByVal cCuvalFisNo As String)

        Dim cSQL As String
        Dim ConnYage As SqlConnection
        Dim aKoliFisNo() As String
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim nKolinumarasi As Integer
        Dim cSiparisNo As String
        Dim cModelNo As String
        Dim cFirma As String
        Dim cAltFirma As String
        Dim nNetAgirlik As Double
        Dim nBrutAgirlik As Double
        Dim nToplamAdet As Double
        Dim lOK As Boolean
        Dim oSysFlags As SysFlags = Nothing

        Try
            ReadSysFlagsMain(oSysFlags)

            ConnYage = OpenConn()

            nCnt = 0
            ReDim aKoliFisNo(0)

            cSQL = "Select kolifisno " + _
                    " from barkodlucuval " + _
                    " where cuvalfisno = '" + cCuvalFisNo + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aKoliFisNo(nCnt)
                aKoliFisNo(nCnt) = SQLReadString(oReader, "kolifisno")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

            cSQL = "select cuvalfisno " + _
                    " from cuvalfis " + _
                    " where cuvalfisno = '" + cCuvalFisNo + "' "

            If Not CheckExists(cSQL) Then

                cSQL = "insert into cuvalfis (cuvalfisno,tarih,ambalaj,gonderenfirma) " + _
                        " values ('" + cCuvalFisNo + "', " + _
                        " '" + SQLWriteDate(Today) + "', " + _
                        " 'CVL','DAHiLi') "

                ExecuteSQLCommand(cSQL, True)
            End If

            cSQL = "delete cuvalfislines " + _
                    " where cuvalfisno = '" + cCuvalFisNo + "' "

            ExecuteSQLCommand(cSQL)

            For nCnt = 0 To UBound(aKoliFisNo)

                ConnYage = OpenConn()

                lOK = False
                nKolinumarasi = 0
                cSiparisNo = ""
                cModelNo = ""
                cFirma = ""
                cAltFirma = ""
                nNetAgirlik = 0
                nBrutAgirlik = 0
                nToplamAdet = 0

                If Mid(aKoliFisNo(nCnt), 1, 3) = "869" Then
                    ' model no
                    cSQL = "select stokno " + _
                            " from stokbarkod " + _
                            " where barcode2 = '" + aKoliFisNo(nCnt) + "' "
                    cModelNo = ReadSingleValueConnected(cSQL, ConnYage)
                    ' sipariş no
                    cSQL = " select b.malzemetakipkodu " + _
                           " from stokbarkod a, stokrb b " + _
                           " Where a.stokno = b.stokno " + _
                           " and a.beden = b.beden " + _
                           " and a.renk = b.renk " + _
                           " and (b.depo = 'KIRIK DEPO' or b.depo = '" + oSysFlags.G_SevkStokDeposu + "') " + _
                           " and a.barcode2 = '" + aKoliFisNo(nCnt) + "' " + _
                           " and b.malzemetakipkodu is not null " + _
                           " and b.malzemetakipkodu <> '' "
                    cSiparisNo = ReadSingleValueConnected(cSQL, ConnYage)
                    ' müşteri firma 
                    cSQL = "select musterino from siparis where kullanicisipno = '" + cSiparisNo + "' "
                    cFirma = ReadSingleValueConnected(cSQL, ConnYage)
                    ' adet
                    nToplamAdet = 1
                    lOK = True
                Else
                    cSQL = "SELECT a.kolinumarasi, b.siparisno, b.modelno, a.netagirlik, a.brutagirlik, a.sevkformno, adet = sum(coalesce(b.adet,0)), " + _
                            " firma = (select top 1 musterino from siparis where kullanicisipno = b.siparisno), " + _
                            " altfirma = (select top 1 d.altmusteri " + _
                                            " from sipmodel c, sevkplfislines d " + _
                                            " where d.sevkiyattakipno = c.sevkiyattakipno " + _
                                            " and c.siparisno = b.siparisno " + _
                                            " and c.modelno = b.modelno) " + _
                            " FROM kolileme a, kolilines b " + _
                            " WHERE a.kolifisno = b.kolifisno " + _
                            " AND a.kolifisno = '" + aKoliFisNo(nCnt) + " ' " + _
                            " GROUP BY a.kolinumarasi, b.siparisno, b.modelno, a.netagirlik, a.brutagirlik, a.sevkformno " + _
                            " ORDER BY a.kolinumarasi, b.siparisno, b.modelno, a.netagirlik, a.brutagirlik, a.sevkformno "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        nKolinumarasi = SQLReadInteger(oReader, "kolinumarasi")
                        cSiparisNo = SQLReadString(oReader, "siparisno")
                        cModelNo = SQLReadString(oReader, "modelno")
                        cFirma = SQLReadString(oReader, "firma")
                        cAltFirma = SQLReadString(oReader, "altfirma")
                        nNetAgirlik = SQLReadDouble(oReader, "netagirlik")
                        nBrutAgirlik = SQLReadDouble(oReader, "brutagirlik")
                        nToplamAdet = SQLReadDouble(oReader, "adet")

                        lOK = True
                    End If
                    oReader.Close()
                    oReader = Nothing
                End If
                CloseConn(ConnYage)

                If lOK Then
                    cSQL = "insert into cuvalfislines (cuvalfisno, kolifisno, kolinumarasi, siparisno, modelno, " + _
                                                        " firma, altfirma, netagirlik, brutagirlik, toplamadet) " + _
                            " values ('" + cCuvalFisNo + "', " + _
                            " '" + aKoliFisNo(nCnt) + "', " + _
                            SQLWriteInteger(nKolinumarasi) + ", " + _
                            " '" + cSiparisNo + "', " + _
                            " '" + cModelNo + "', " + _
                            " '" + cFirma + "', " + _
                            " '" + cAltFirma + "', " + _
                            SQLWriteDecimal(nNetAgirlik) + ", " + _
                            SQLWriteDecimal(nBrutAgirlik) + ", " + _
                            SQLWriteDecimal(nToplamAdet) + ") "

                    ExecuteSQLCommand(cSQL)
                End If
            Next

            cSQL = "delete barkodlucuval where cuvalfisno = '" + cCuvalFisNo + "' "
            ExecuteSQLCommand(cSQL)

        Catch Err As Exception
            ErrDisp("Cuvalla : " + Err.Message)
        End Try
    End Sub

    Public Function BarkodluYukleme(ByVal cKoliFisNo As String, ByVal cTSiparisNo As String) As Integer

        Dim cSQL As String
        Dim ConnYage As SqlConnection
        Dim aKoliGirisFisNo() As String
        Dim aCuvalFisNo() As String
        Dim oReader As SqlDataReader
        Dim nCnt As Integer

        Dim nKolinumarasi As Integer
        Dim cSiparisNo As String
        Dim cModelNo As String
        Dim cFirma As String
        Dim cAltFirma As String
        Dim nNetAgirlik As Double
        Dim nBrutAgirlik As Double
        Dim nToplamAdet As Double
        Dim cBedenSeti As String
        Dim cSevkFormNo As String
        Dim cStokFisNo As String
        Dim cCuvalFisNo As String
        Dim lOK As Boolean
        Dim oSysFlags As SysFlags = Nothing
        Dim nSFCnt As Integer

        BarkodluYukleme = 0
        cSQL = ""

        Try
            ReadSysFlagsMain(oSysFlags)

            ConnYage = OpenConn()

            cSQL = "delete kolifisi where kolifisno = '" + cKoliFisNo + "' "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete kolifislines where kolifisno = '" + cKoliFisNo + "' "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = 0
            nSFCnt = 0
            ReDim aKoliGirisFisNo(0)
            ReDim aCuvalFisNo(0)

            cSQL = "Select koligirisfisno, cuvalfisno " + _
                    " from barkodluyukleme " + _
                    " where kolifisno = '" + cKoliFisNo + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aKoliGirisFisNo(nCnt)
                aKoliGirisFisNo(nCnt) = SQLReadString(oReader, "koligirisfisno")

                ReDim Preserve aCuvalFisNo(nCnt)
                aCuvalFisNo(nCnt) = SQLReadString(oReader, "cuvalfisno")
                nCnt = nCnt + 1
            Loop

            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

            cSQL = " insert into kolifisi (kolifisno,fistipi,fistarihi,faturatarihi,departman,firma,harekettipi,siparisno) " + _
                        " values ('" + cKoliFisNo + "', " + _
                        " 'Cikis', " + _
                        " '" + SQLWriteDate(Today) + "', " + _
                        " '" + SQLWriteDate(Today) + "', " + _
                        " 'MAMUL','DAHiLi','07 Satis', " + _
                        " '" + cTSiparisNo + "') "

            ExecuteSQLCommand(cSQL, True)

            For nCnt = 0 To UBound(aKoliGirisFisNo)

                ConnYage = OpenConn()

                lOK = False
                nKolinumarasi = 0
                cSiparisNo = ""
                cModelNo = ""
                cFirma = ""
                cAltFirma = ""
                nNetAgirlik = 0
                nBrutAgirlik = 0
                nToplamAdet = 0
                cBedenSeti = ""
                cSevkFormNo = ""
                cStokFisNo = ""
                cCuvalFisNo = ""

                If Mid(aKoliGirisFisNo(nCnt), 1, 3) = "869" Then
                    ' model no
                    cSQL = "select stokno " + _
                            " from stokbarkod " + _
                            " where barcode2 = '" + aKoliGirisFisNo(nCnt) + "' "
                    cModelNo = ReadSingleValueConnected(cSQL, ConnYage)
                    ' sipariş no
                    cSQL = " select b.malzemetakipkodu " + _
                           " from stokbarkod a, stokrb b " + _
                           " Where a.stokno = b.stokno " + _
                           " and a.beden = b.beden " + _
                           " and a.renk = b.renk " + _
                           " and (b.depo = 'KIRIK DEPO' or b.depo = 'T.KIRIK DEPO' or b.depo = '" + oSysFlags.G_SevkStokDeposu + "' or b.depo = '" + oSysFlags.G_TSevkStokDeposu + "') " + _
                           " and a.barcode2 = '" + aKoliGirisFisNo(nCnt) + "' " + _
                           " and b.malzemetakipkodu is not null " + _
                           " and b.malzemetakipkodu <> '' "
                    cSiparisNo = ReadSingleValueConnected(cSQL, ConnYage)
                    ' müşteri firma 
                    cSQL = "select musterino from siparis where kullanicisipno = '" + cSiparisNo + "' "
                    cFirma = ReadSingleValueConnected(cSQL, ConnYage)
                    ' adet
                    nToplamAdet = 1
                    lOK = True
                Else
                    cSQL = "SELECT a.kolinumarasi, b.siparisno, b.modelno, a.netagirlik, a.brutagirlik, a.sevkformno, adet = sum(coalesce(b.adet,0)), " + _
                            " firma = (select top 1 musterino from siparis where kullanicisipno = b.siparisno), " + _
                            " altfirma = (select top 1 d.altmusteri " + _
                                            " from sipmodel c, sevkplfislines d " + _
                                            " where d.sevkiyattakipno = c.sevkiyattakipno " + _
                                            " and c.siparisno = b.siparisno " + _
                                            " and c.modelno = b.modelno) " + _
                            " FROM kolileme a, kolilines b " + _
                            " WHERE a.kolifisno = b.kolifisno " + _
                            " AND a.kolifisno = '" + aKoliGirisFisNo(nCnt) + " ' " + _
                            " GROUP BY a.kolinumarasi, b.siparisno, b.modelno, a.netagirlik, a.brutagirlik, a.sevkformno " + _
                            " ORDER BY a.kolinumarasi, b.siparisno, b.modelno, a.netagirlik, a.brutagirlik, a.sevkformno "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        nKolinumarasi = SQLReadInteger(oReader, "kolinumarasi")
                        cSiparisNo = SQLReadString(oReader, "siparisno")
                        cModelNo = SQLReadString(oReader, "modelno")
                        cFirma = SQLReadString(oReader, "firma")
                        cAltFirma = SQLReadString(oReader, "altfirma")
                        nNetAgirlik = SQLReadDouble(oReader, "netagirlik")
                        nBrutAgirlik = SQLReadDouble(oReader, "brutagirlik")
                        nToplamAdet = SQLReadDouble(oReader, "adet")
                        cSevkFormNo = SQLReadString(oReader, "sevkformno")
                        lOK = True
                    End If
                    oReader.Close()
                    oReader = Nothing
                End If

                cSQL = "select bedenseti " + _
                        " from sipasorti a, firmaasorti b " + _
                        " where a.asortino = b.asortino " + _
                        " and a.siparisno = '" + cSiparisNo + "' " + _
                        " and a.modelno = '" + cModelNo + "' "

                cBedenSeti = ReadSingleValueConnected(cSQL, ConnYage)

                CloseConn(ConnYage)

                If lOK Then
                    cSQL = "insert into kolifislines (kolifisno, koligirisfisno, kolinumarasi, siparisno, modelno, " + _
                            " firma, altfirma, netagirlik, brutagirlik, toplamadet, sevkformno, cuvalfisno, bedenseti, " + _
                            " stokfisno,  grup )"

                    cSQL = cSQL + _
                            " values ('" + cKoliFisNo + "', " + _
                            " '" + aKoliGirisFisNo(nCnt) + "', " + _
                            SQLWriteInteger(nKolinumarasi) + ", " + _
                            " '" + cSiparisNo + "', " + _
                            " '" + cModelNo + "', "

                    cSQL = cSQL + _
                            " '" + cFirma + "', " + _
                            " '" + cAltFirma + "', " + _
                            SQLWriteDecimal(nNetAgirlik) + ", " + _
                            SQLWriteDecimal(nBrutAgirlik) + ", " + _
                            SQLWriteDecimal(nToplamAdet) + ", "

                    cSQL = cSQL + _
                            " '" + cSevkFormNo + "', " + _
                            " '" + aCuvalFisNo(nCnt) + "', " + _
                            " '" + cBedenSeti + "', " + _
                            " '', '' )"

                    ExecuteSQLCommand(cSQL)
                End If
            Next

            cSQL = "delete barkodluyukleme where kolifisno = '" + cKoliFisNo + "' "
            ExecuteSQLCommand(cSQL)

            BarkodluYukleme = 1

        Catch Err As Exception
            BarkodluYukleme = 0
            ErrDisp("BarkodluYukleme : " + cKoliFisNo + vbCrLf + Err.Message + vbCrLf + cSQL)
        End Try
    End Function

    Public Function YuklemedenStokCikisFisi(ByVal cKoliFisNo As String) As Integer

        Dim cSQL As String
        Dim ConnYage As SqlConnection
        Dim aModelNo() As String
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim cAciklama As String
        Dim cStokTipi As String
        Dim cStokFisNo As String
        Dim cFirma As String
        Dim nOK As SqlInt32
        Dim cSiparisNo As String
        Dim cModelNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim nAdet As Double
        Dim cBedenSeti As String
        Dim lOK As Boolean
        Dim oSysFlags As SysFlags = Nothing
        Dim aSatir() As FisLines
        Dim nSCnt As Integer
        Dim cMTF As String
        Dim nFiyat As Double
        Dim cDoviz As String
        Dim cTSiparisNo As String
        Dim aSevkFormNo() As String
        Dim aKoliGirisFisNo() As String
        Dim nSFCnt As Integer
        Dim nCnt2 As Integer
        Dim nFound As Integer
        Dim lFirst As Boolean
        Dim cDepo As String
        Dim cUrtGrubu As String

        YuklemedenStokCikisFisi = 0
        cSQL = ""

        Try
            lFirst = True
            ReDim aSatir(0)
            nSCnt = 0
            nSFCnt = 0
            ReadSysFlagsMain(oSysFlags)

            ConnYage = OpenConn()

            cSQL = "select siparisno from kolifisi where kolifisno = '" + cKoliFisNo.Trim + "' "
            cTSiparisNo = ReadSingleValueConnected(cSQL, ConnYage)

            nCnt = 0
            ReDim aKoliGirisFisNo(0)
            ReDim aSevkFormNo(0)

            cSQL = "Select koligirisfisno " + _
                    " from kolifislines " + _
                    " where kolifisno = '" + cKoliFisNo + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aKoliGirisFisNo(nCnt)
                aKoliGirisFisNo(nCnt) = SQLReadString(oReader, "koligirisfisno")
                nCnt = nCnt + 1
            Loop

            oReader.Close()
            oReader = Nothing

            cSQL = "select kod from anastokgrubu where kod = 'MAMUL' "
            If Not CheckExistsConnected(cSQL, ConnYage) Then
                cSQL = "insert into anastokgrubu (kod,aciklama) values ('MAMUL','MAMUL') "
                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            nCnt = 0
            ReDim aModelNo(0)

            cSQL = "Select distinct modelno " + _
                    " from kolifislines " + _
                    " where kolifisno = '" + cKoliFisNo + "' " + _
                    " and (stokfisno = '' or stokfisno is NULL) "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aModelNo(nCnt)
                aModelNo(nCnt) = SQLReadString(oReader, "modelno")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aModelNo)
                cSQL = "select stokno from stok where stokno = '" + aModelNo(nCnt) + "' "
                If Not CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "select aciklama from ymodel where modelno = '" + aModelNo(nCnt) + "' "
                    cAciklama = ReadSingleValueConnected(cSQL, ConnYage)

                    cSQL = "select anamodeltipi from ymodel where modelno = '" + aModelNo(nCnt) + "' "
                    cStokTipi = ReadSingleValueConnected(cSQL, ConnYage)

                    cSQL = "insert into stok (stokno, cinsaciklamasi, stoktipi, paratakipesasi, maltakipesasi, Birim1, anastokgrubu) " + _
                            " values ('" + aModelNo(nCnt) + "', " + _
                            " '" + cAciklama + "', " + _
                            " '" + cStokTipi + "', " + _
                            " '1','4','AD','MAMUL' ) "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "select kod from stoktipi where kod = '" + cStokTipi + "' "
                    If Not CheckExistsConnected(cSQL, ConnYage) Then
                        cSQL = "insert into stoktipi (kod,aciklama) values ('" + cStokTipi + "','" + cStokTipi + "')"
                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                End If
            Next

            ' tek firmaya sevkiyat yapıyoruz varsayılıyor

            cSQL = "select firma from kolifislines " + _
                    " where kolifisno = '" + cKoliFisNo + "' " + _
                    " and firma is not null " + _
                    " and firma <> '' "

            cFirma = ReadSingleValueConnected(cSQL, ConnYage)

            CloseConn(ConnYage)

            ' add stokfis 

            DeleteStokFis(cKoliFisNo)

            cStokFisNo = GetStokFisNo()

            cSQL = " insert into stokfis (StokFisNo, fistarihi, faturatarihi, StokFisTipi, departman, " + _
                    " firma, tsiparisno, modificationdate, CreationDate, username, notlar) " + _
                    " values ('" + cStokFisNo + "', " + _
                    " '" + SQLWriteDate(Today) + "', " + _
                    " '" + SQLWriteDate(Today) + "', " + _
                    " 'Cikis', 'MAMUL', " + _
                    " '" + cFirma + "', " + _
                    " '" + cTSiparisNo + "', " + _
                    " '" + SQLWriteDate(Today) + "', " + _
                    " '" + SQLWriteDate(Today) + "', " + _
                    " 'WinTexCLR', 'WinTexCLR Yükleme Fişi : " + cKoliFisNo + "') "

            ExecuteSQLCommand(cSQL, True)

            ' add StokFisLines-read

            ConnYage = OpenConn()

            For nCnt = 0 To UBound(aKoliGirisFisNo)

                lOK = False
                cSiparisNo = ""
                cModelNo = ""
                cRenk = ""
                cBeden = ""
                nAdet = 0

                If Mid(aKoliGirisFisNo(nCnt), 1, 3) = "869" Then
                    ' model no
                    cSQL = "select stokno, renk, beden " + _
                            " from stokbarkod " + _
                            " where barcode2 = '" + aKoliGirisFisNo(nCnt) + "' "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        cModelNo = SQLReadString(oReader, "stokno")
                        cRenk = SQLReadString(oReader, "renk")
                        cBeden = SQLReadString(oReader, "beden")
                    End If
                    oReader.Close()
                    oReader = Nothing

                    ' sipariş no
                    cSQL = " select b.malzemetakipkodu " + _
                           " from stokbarkod a, stokrb b " + _
                           " Where a.stokno = b.stokno " + _
                           " and a.beden = b.beden " + _
                           " and a.renk = b.renk " + _
                           " and (b.depo = 'KIRIK DEPO' or b.depo = 'T.KIRIK DEPO' or b.depo = '" + oSysFlags.G_SevkStokDeposu + "' or b.depo = '" + oSysFlags.G_TSevkStokDeposu + "') " + _
                           " and a.barcode2 = '" + aKoliGirisFisNo(nCnt) + "' " + _
                           " and b.malzemetakipkodu is not null " + _
                           " and b.malzemetakipkodu <> '' "

                    cSiparisNo = ReadSingleValueConnected(cSQL, ConnYage)

                    cSQL = "select urtgrubu from siparis where siparisno = '" + cSiparisNo + "' "

                    cUrtGrubu = ReadSingleValueConnected(cSQL, ConnYage)

                    cDepo = "KIRIK DEPO"
                    If cUrtGrubu <> "Üretim" Then
                        cDepo = "T.KIRIK DEPO"
                    End If

                    cSQL = "select bedenseti " + _
                            " from sipasorti a, firmaasorti b " + _
                            " where a.asortino = b.asortino " + _
                            " and a.siparisno = '" + cSiparisNo + "' " + _
                            " and a.modelno = '" + cModelNo + "' "

                    cBedenSeti = ReadSingleValueConnected(cSQL, ConnYage)

                    If lFirst Then
                        ReDim Preserve aSatir(nSCnt)
                        aSatir(nSCnt).cSiparisNo = cSiparisNo
                        aSatir(nSCnt).cStokNo = cModelNo
                        aSatir(nSCnt).cRenk = cRenk
                        aSatir(nSCnt).cBeden = cBeden
                        aSatir(nSCnt).cBedenSeti = cBedenSeti
                        aSatir(nSCnt).cDepo = cDepo
                        aSatir(nSCnt).nMiktar = 1
                        nSCnt = nSCnt + 1
                        lFirst = False
                    Else
                        nFound = -1
                        For nCnt2 = 0 To UBound(aSatir)
                            If aSatir(nCnt2).cSiparisNo = cSiparisNo And _
                                aSatir(nCnt2).cStokNo = cModelNo And _
                                aSatir(nCnt2).cRenk = cRenk And _
                                aSatir(nCnt2).cBeden = cBeden And _
                                aSatir(nCnt2).cBedenSeti = cBedenSeti And _
                                aSatir(nSCnt).cDepo = cDepo Then
                                nFound = nCnt2
                                Exit For
                            End If
                        Next
                        If nFound = -1 Then
                            ReDim Preserve aSatir(nSCnt)
                            aSatir(nSCnt).cSiparisNo = cSiparisNo
                            aSatir(nSCnt).cStokNo = cModelNo
                            aSatir(nSCnt).cRenk = cRenk
                            aSatir(nSCnt).cBeden = cBeden
                            aSatir(nSCnt).cBedenSeti = cBedenSeti
                            aSatir(nSCnt).cDepo = cDepo
                            aSatir(nSCnt).nMiktar = 1
                            nSCnt = nSCnt + 1
                        Else
                            aSatir(nFound).nMiktar = aSatir(nFound).nMiktar + 1
                        End If
                    End If

                Else
                    cSQL = "SELECT b.siparisno, b.modelno, b.renk, b.beden, b.bedenseti, c.urtgrubu, adet = sum(coalesce(b.adet,0)) " + _
                            " FROM kolileme a, kolilines b, siparis c " + _
                            " WHERE a.kolifisno = b.kolifisno " + _
                            " and a.kolifisno = '" + aKoliGirisFisNo(nCnt) + " ' " + _
                            " and b.siparisno = c.kullanicisipno " + _
                            " GROUP BY b.siparisno, b.modelno, b.renk, b.beden, b.bedenseti, c.urtgrubu " + _
                            " ORDER BY b.siparisno, b.modelno, b.renk, b.beden, b.bedenseti, c.urtgrubu "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    Do While oReader.Read
                        cSiparisNo = SQLReadString(oReader, "siparisno")
                        cModelNo = SQLReadString(oReader, "modelno")
                        cRenk = SQLReadString(oReader, "renk")
                        cBeden = SQLReadString(oReader, "beden")
                        cBedenSeti = SQLReadString(oReader, "bedenseti")
                        nAdet = SQLReadDouble(oReader, "adet")

                        cUrtGrubu = SQLReadString(oReader, "urtgrubu")

                        cDepo = oSysFlags.G_SevkStokDeposu
                        If cUrtGrubu <> "Üretim" Then
                            cDepo = oSysFlags.G_TSevkStokDeposu
                        End If

                        If lFirst Then
                            ReDim Preserve aSatir(nSCnt)
                            aSatir(nSCnt).cSiparisNo = cSiparisNo
                            aSatir(nSCnt).cStokNo = cModelNo
                            aSatir(nSCnt).cRenk = cRenk
                            aSatir(nSCnt).cBeden = cBeden
                            aSatir(nSCnt).cBedenSeti = cBedenSeti
                            aSatir(nSCnt).cDepo = cDepo
                            aSatir(nSCnt).nMiktar = nAdet
                            nSCnt = nSCnt + 1
                            lFirst = False
                        Else
                            nFound = -1
                            For nCnt2 = 0 To UBound(aSatir)
                                If aSatir(nCnt2).cSiparisNo = cSiparisNo And _
                                    aSatir(nCnt2).cStokNo = cModelNo And _
                                    aSatir(nCnt2).cRenk = cRenk And _
                                    aSatir(nCnt2).cBeden = cBeden And _
                                    aSatir(nCnt2).cBedenSeti = cBedenSeti And _
                                    aSatir(nCnt2).cDepo = cDepo Then
                                    nFound = nCnt2
                                    Exit For
                                End If
                            Next
                            If nFound = -1 Then
                                ReDim Preserve aSatir(nSCnt)
                                aSatir(nSCnt).cSiparisNo = cSiparisNo
                                aSatir(nSCnt).cStokNo = cModelNo
                                aSatir(nSCnt).cRenk = cRenk
                                aSatir(nSCnt).cBeden = cBeden
                                aSatir(nSCnt).cBedenSeti = cBedenSeti
                                aSatir(nSCnt).cDepo = cDepo
                                aSatir(nSCnt).nMiktar = nAdet
                                nSCnt = nSCnt + 1
                            Else
                                aSatir(nFound).nMiktar = aSatir(nFound).nMiktar + nAdet
                            End If
                        End If
                    Loop
                    oReader.Close()
                    oReader = Nothing

                    cSQL = "SELECT sevkformno " + _
                            " FROM kolileme " + _
                            " WHERE kolifisno = '" + aKoliGirisFisNo(nCnt) + "' "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        If SQLReadString(oReader, "sevkformno") <> "" Then
                            ReDim Preserve aSevkFormNo(nSFCnt)
                            aSevkFormNo(nSFCnt) = SQLReadString(oReader, "sevkformno")
                            nSFCnt = nSFCnt + 1
                        End If
                    End If
                    oReader.Close()
                    oReader = Nothing
                End If
            Next

            ' add stokfislines-write
            For nCnt = 0 To UBound(aSatir)

                cSQL = "SELECT malzemetakipno " + _
                        " from sipmodel " + _
                        " where siparisno = '" + aSatir(nCnt).cSiparisNo + " ' " + _
                        " and modelno = '" + aSatir(nCnt).cStokNo + " ' " + _
                        " and malzemetakipno is not null " + _
                        " and malzemetakipno <> '' " + _
                        IIf(aSatir(nCnt).cBedenSeti = "", "", " and bedenseti = '" + aSatir(nCnt).cBedenSeti + "' ").ToString()

                cMTF = ReadSingleValueConnected(cSQL, ConnYage)

                nFiyat = 0
                cDoviz = "USD"

                cSQL = "SELECT maliyet " + _
                        " from modelastfiyat " + _
                        " where modelno = '" + aSatir(nCnt).cStokNo + "' " + _
                        " and bedenseti = '" + aSatir(nCnt).cBedenSeti + "' " + _
                        " and maliyet is not null " + _
                        " and maliyet <> 0 "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "maliyet")
                End If
                oReader.Close()
                oReader = Nothing

                cSQL = "insert into stokfislines " + _
                        " (stokfisno, stokhareketkodu, malzemetakipkodu, Stokno, renk, beden, " + _
                        " depo, netmiktar1, brutmiktar1, birim1, fissirano, birimfiyat, dovizcinsi, notlar) " + _
                        " values ('" + cStokFisNo + "', " + _
                        " '07 Satis', " + _
                        " '" + cMTF + "', " + _
                        " '" + aSatir(nCnt).cStokNo + "', " + _
                        " '" + aSatir(nCnt).cRenk + "', " + _
                        " '" + aSatir(nCnt).cBeden + "', " + _
                        " '" + aSatir(nCnt).cDepo + "', " + _
                        SQLWriteDecimal(aSatir(nCnt).nMiktar) + ", " + _
                        SQLWriteDecimal(aSatir(nCnt).nMiktar) + ", " + _
                        " 'AD',0, " + _
                        SQLWriteDecimal(nFiyat) + ", " + _
                        " '" + cDoviz + "', " + _
                        " '" + aSatir(nCnt).cBedenSeti + "') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next
            CloseConn(ConnYage)

            ' validate
            nOK = SingleStokFisValidate("validate", cStokFisNo, "", "", "")
            If nOK = 0 Then
                ErrDisp("Stok fis validate err : " + cStokFisNo)
                YuklemedenStokCikisFisi = 0
                Exit Function
            End If

            ' sevkformları düzelt
            For nCnt = 0 To UBound(aSevkFormNo)
                cSQL = "UPDATE sevkform " + _
                        " SET ok = 'E', " + _
                        " stokfisno = '" + cStokFisNo + "' " + _
                        " WHERE sevkformno = '" + aSevkFormNo(nCnt) + "' "

                ExecuteSQLCommand(cSQL)
            Next

            ' stok fisi nosunu kolifişlines yaz
            cSQL = "update kolifislines " + _
                    " set stokfisno = '" + cStokFisNo + "' " + _
                    " where kolifisno = '" + cKoliFisNo + "' "

            ExecuteSQLCommand(cSQL)

            YuklemedenStokCikisFisi = 1

        Catch Err As Exception
            YuklemedenStokCikisFisi = 0
            ErrDisp("YuklemedenStokCikisFisi : " + cKoliFisNo + vbCrLf + Err.Message + vbCrLf + cSQL)
        End Try

    End Function

    Private Sub DeleteStokFis(ByVal cKoliFisNo As String)

        Dim cStokFisNo As String
        Dim cSQL As String
        Dim nOK As SqlInt32

        ' 1 adet stok fişi bağlı kabul ediliyor
        cSQL = "select distinct stokfisno " + _
                " from kolifislines " + _
                " where kolifisno = '" + cKoliFisNo + "'"

        cStokFisNo = ReadSingleValue(cSQL)

        If cStokFisNo = "" Then Exit Sub

        nOK = SingleStokFisValidate("revert", cStokFisNo, "", "", "")

        If nOK = 1 Then
            cSQL = "delete from stokfis where stokfisno = '" + cStokFisNo + "' "
            ExecuteSQLCommand(cSQL)

            cSQL = "delete from stokfislines where stokfisno = '" + cStokFisNo + "' "
            ExecuteSQLCommand(cSQL)
        End If

        cSQL = "update kolifislines set stokfisno = '' where kolifisno = '" + cKoliFisNo + "' "
        ExecuteSQLCommand(cSQL)
    End Sub

    Public Function KontrolYukleme(ByVal cKoliFisNo As String) As String

        Dim aSatir() As FisLines
        Dim cSQL As String
        Dim ConnYage As SqlConnection
        Dim nCnt As Integer
        Dim nDepoMiktar As Double
        Dim aKoliGirisFisNo() As String

        cSQL = ""
        KontrolYukleme = "Hata"

        Try
            KontrolYukleme = ""
            ' Fill aSatir
            ReDim aSatir(0)
            ReDim aKoliGirisFisNo(0)

            BreakYukleme(cKoliFisNo, aSatir, aKoliGirisFisNo)

            ConnYage = OpenConn()

            For nCnt = 0 To UBound(aSatir)
                nDepoMiktar = 0
                cSQL = "select miktar = coalesce(donemgiris1,0) + coalesce(devirgiris1,0) - coalesce(donemcikis1,0) - coalesce(devircikis1,0) " + _
                        " from stokrb " + _
                        " where stokno = '" + aSatir(nCnt).cStokNo + "' " + _
                        " and renk = '" + aSatir(nCnt).cRenk + "' " + _
                        " and beden = '" + aSatir(nCnt).cBeden + "' " + _
                        " and depo = 'MAMUL DEPO' " + _
                        " and malzemetakipkodu = '" + aSatir(nCnt).cMTF + "' "

                nDepoMiktar = ReadSingleDoubleValueConnected(cSQL, ConnYage)

                If nDepoMiktar < aSatir(nCnt).nMiktar Then
                    KontrolYukleme = KontrolYukleme + _
                                    aSatir(nCnt).cStokNo + " " + aSatir(nCnt).cRenk + " " + aSatir(nCnt).cBeden + _
                                    " depo:" + nDepoMiktar.ToString + " cikis:" + aSatir(nCnt).nMiktar.ToString + vbCrLf
                End If
            Next

            CloseConn(ConnYage)

            If KontrolYukleme = "" Then
                KontrolYukleme = "OK"
            Else
                KontrolYukleme = "Hata : " + KontrolYukleme
            End If
        Catch Err As Exception
            KontrolYukleme = "Hata"
            ErrDisp("KontrolYukleme : " + cKoliFisNo + vbCrLf + Err.Message + vbCrLf + cSQL)
        End Try
    End Function

    Private Sub BreakYukleme(ByVal cKoliFisNo As String, ByRef aSatir() As FisLines, ByRef aKoliGirisFisNo() As String)

        Dim cSQL As String
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim nSCnt As Integer
        Dim cSiparisNo As String
        Dim cModelNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim nAdet As Double
        Dim cBedenSeti As String
        Dim lOK As Boolean
        Dim oSysFlags As SysFlags = Nothing
        Dim lFirst As Boolean
        Dim nFound As Integer
        Dim nCnt2 As Integer

        cSQL = ""

        Try
            ReadSysFlagsMain(oSysFlags)

            ConnYage = OpenConn()

            lFirst = True
            nCnt = 0

            ReDim aSatir(0)
            ReDim aKoliGirisFisNo(0)

            cSQL = "Select koligirisfisno, cuvalfisno " + _
                    " from barkodluyukleme " + _
                    " where kolifisno = '" + cKoliFisNo + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aKoliGirisFisNo(nCnt)
                aKoliGirisFisNo(nCnt) = SQLReadString(oReader, "koligirisfisno")
                nCnt = nCnt + 1
            Loop

            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aKoliGirisFisNo)

                lOK = False
                cSiparisNo = ""
                cModelNo = ""
                cRenk = ""
                cBeden = ""
                nAdet = 0

                If Mid(aKoliGirisFisNo(nCnt), 1, 3) = "869" Then
                    ' model no
                    cSQL = "select stokno, renk, beden " + _
                            " from stokbarkod " + _
                            " where barcode2 = '" + aKoliGirisFisNo(nCnt) + "' "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        cModelNo = SQLReadString(oReader, "stokno")
                        cRenk = SQLReadString(oReader, "renk")
                        cBeden = SQLReadString(oReader, "beden")
                    End If
                    oReader.Close()
                    oReader = Nothing

                    ' sipariş no
                    cSQL = " select b.malzemetakipkodu " + _
                           " from stokbarkod a, stokrb b " + _
                           " Where a.stokno = b.stokno " + _
                           " and a.beden = b.beden " + _
                           " and a.renk = b.renk " + _
                           " and (b.depo = 'KIRIK DEPO' or b.depo = '" + oSysFlags.G_SevkStokDeposu + "') " + _
                           " and a.barcode2 = '" + aKoliGirisFisNo(nCnt) + "' " + _
                           " and b.malzemetakipkodu is not null " + _
                           " and b.malzemetakipkodu <> '' "

                    cSiparisNo = ReadSingleValueConnected(cSQL, ConnYage)

                    cSQL = "select bedenseti " + _
                            " from sipasorti a, firmaasorti b " + _
                            " where a.asortino = b.asortino " + _
                            " and a.siparisno = '" + cSiparisNo + "' " + _
                            " and a.modelno = '" + cModelNo + "' "

                    cBedenSeti = ReadSingleValueConnected(cSQL, ConnYage)

                    If lFirst Then
                        ReDim Preserve aSatir(nSCnt)
                        aSatir(nSCnt).cSiparisNo = cSiparisNo
                        aSatir(nSCnt).cStokNo = cModelNo
                        aSatir(nSCnt).cRenk = cRenk
                        aSatir(nSCnt).cBeden = cBeden
                        aSatir(nSCnt).cBedenSeti = cBedenSeti
                        aSatir(nSCnt).nMiktar = 1
                        nSCnt = nSCnt + 1
                        lFirst = False
                    Else
                        nFound = -1
                        For nCnt2 = 0 To UBound(aSatir)
                            If aSatir(nCnt2).cSiparisNo = cSiparisNo And _
                                aSatir(nCnt2).cStokNo = cModelNo And _
                                aSatir(nCnt2).cRenk = cRenk And _
                                aSatir(nCnt2).cBeden = cBeden And _
                                aSatir(nCnt2).cBedenSeti = cBedenSeti Then
                                nFound = nCnt2
                                Exit For
                            End If
                        Next
                        If nFound = -1 Then
                            ReDim Preserve aSatir(nSCnt)
                            aSatir(nSCnt).cSiparisNo = cSiparisNo
                            aSatir(nSCnt).cStokNo = cModelNo
                            aSatir(nSCnt).cRenk = cRenk
                            aSatir(nSCnt).cBeden = cBeden
                            aSatir(nSCnt).cBedenSeti = cBedenSeti
                            aSatir(nSCnt).nMiktar = 1
                            nSCnt = nSCnt + 1
                        Else
                            aSatir(nFound).nMiktar = aSatir(nFound).nMiktar + 1
                        End If
                    End If

                Else
                    cSQL = "SELECT b.siparisno, b.modelno, b.renk, b.beden, b.bedenseti, adet = sum(coalesce(b.adet,0)) " + _
                            " FROM kolileme a, kolilines b " + _
                            " WHERE a.kolifisno = b.kolifisno " + _
                            " AND a.kolifisno = '" + aKoliGirisFisNo(nCnt) + " ' " + _
                            " GROUP BY b.siparisno, b.modelno, b.renk, b.beden, b.bedenseti " + _
                            " ORDER BY b.siparisno, b.modelno, b.renk, b.beden, b.bedenseti "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    Do While oReader.Read
                        cSiparisNo = SQLReadString(oReader, "siparisno")
                        cModelNo = SQLReadString(oReader, "modelno")
                        cRenk = SQLReadString(oReader, "renk")
                        cBeden = SQLReadString(oReader, "beden")
                        cBedenSeti = SQLReadString(oReader, "bedenseti")
                        nAdet = SQLReadDouble(oReader, "adet")

                        If lFirst Then
                            ReDim Preserve aSatir(nSCnt)
                            aSatir(nSCnt).cSiparisNo = cSiparisNo
                            aSatir(nSCnt).cStokNo = cModelNo
                            aSatir(nSCnt).cRenk = cRenk
                            aSatir(nSCnt).cBeden = cBeden
                            aSatir(nSCnt).cBedenSeti = cBedenSeti
                            aSatir(nSCnt).nMiktar = nAdet
                            nSCnt = nSCnt + 1
                            lFirst = False
                        Else
                            nFound = -1
                            For nCnt2 = 0 To UBound(aSatir)
                                If aSatir(nCnt2).cSiparisNo = cSiparisNo And _
                                    aSatir(nCnt2).cStokNo = cModelNo And _
                                    aSatir(nCnt2).cRenk = cRenk And _
                                    aSatir(nCnt2).cBeden = cBeden And _
                                    aSatir(nCnt2).cBedenSeti = cBedenSeti Then
                                    nFound = nCnt2
                                    Exit For
                                End If
                            Next
                            If nFound = -1 Then
                                ReDim Preserve aSatir(nSCnt)
                                aSatir(nSCnt).cSiparisNo = cSiparisNo
                                aSatir(nSCnt).cStokNo = cModelNo
                                aSatir(nSCnt).cRenk = cRenk
                                aSatir(nSCnt).cBeden = cBeden
                                aSatir(nSCnt).cBedenSeti = cBedenSeti
                                aSatir(nSCnt).nMiktar = nAdet
                                nSCnt = nSCnt + 1
                            Else
                                aSatir(nFound).nMiktar = aSatir(nFound).nMiktar + nAdet
                            End If
                        End If
                    Loop
                    oReader.Close()
                    oReader = Nothing
                End If
            Next

            For nCnt = 0 To UBound(aSatir)
                cSQL = "SELECT malzemetakipno " + _
                        " from sipmodel " + _
                        " where siparisno = '" + aSatir(nCnt).cSiparisNo + " ' " + _
                        " and modelno = '" + aSatir(nCnt).cStokNo + " ' " + _
                        " and malzemetakipno is not null " + _
                        " and malzemetakipno <> '' " + _
                        IIf(aSatir(nCnt).cBedenSeti = "", "", " and bedenseti = '" + aSatir(nCnt).cBedenSeti + "' ").ToString()

                aSatir(nCnt).cMTF = ReadSingleValueConnected(cSQL, ConnYage)
            Next
            CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp("BreakYukleme : " + cKoliFisNo + vbCrLf + Err.Message + vbCrLf + cSQL)
        End Try
    End Sub

    Public Function CheckPaketStatus(ByVal cKoliFisNo As String, ByVal cOriginalBarcode As String) As String

        Dim cSQL As String
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cBarcode As String
        Dim aKoliGirisFisNo() As String
        Dim nCnt As Integer

        cSQL = ""
        CheckPaketStatus = "OK"

        Try
            CheckPaketStatus = ""
            ConnYage = OpenConn()

            Select Case Mid(cOriginalBarcode, 1, 3)
                Case "914"

                    ' cuval fişi
                    cBarcode = Mid(cOriginalBarcode, 1, 12)

                    cSQL = "select cuvalfisno " + _
                            " from cuvalfis " + _
                            " where cuvalfisno = '" + cBarcode.Trim + "' "

                    If Not CheckExistsConnected(cSQL, ConnYage) Then
                        CheckPaketStatus = CheckPaketStatus + "Çuval bulunamadı : " + cBarcode + vbCrLf
                    End If

                    cSQL = "select cuvalfisno " + _
                            " from barkodluyukleme " + _
                            " where cuvalfisno = '" + cBarcode.Trim + "' " + _
                            " and kolifisno = '" + cKoliFisNo + "' "

                    If CheckExistsConnected(cSQL, ConnYage) Then
                        CheckPaketStatus = CheckPaketStatus + "Çuval bu yüklemede var : " + cBarcode + vbCrLf
                    End If

                    ' çuvalın içindeki paketler

                    nCnt = 0
                    ReDim aKoliGirisFisNo(0)

                    cSQL = "select kolifisno " + _
                            " from cuvalfislines " + _
                            " where cuvalfisno = '" + cBarcode.Trim + "' " + _
                            " and substring (kolifisno,1,3) = '000' "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    Do While oReader.Read
                        ReDim Preserve aKoliGirisFisNo(nCnt)
                        aKoliGirisFisNo(nCnt) = SQLReadString(oReader, "kolifisno")
                        nCnt = nCnt + 1
                    Loop
                    oReader.Close()
                    oReader = Nothing

                    For nCnt = 0 To UBound(aKoliGirisFisNo)
                        CheckPaketStatus = CheckPaketStatus + CheckPaket(cKoliFisNo, aKoliGirisFisNo(nCnt), ConnYage)
                    Next

                Case "000"
                    CheckPaketStatus = CheckPaketStatus + CheckPaket(cKoliFisNo, cOriginalBarcode, ConnYage)
            End Select

            CloseConn(ConnYage)

            If CheckPaketStatus = "" Then
                CheckPaketStatus = "OK"
            Else
                CheckPaketStatus = "Hata : " + CheckPaketStatus
            End If

        Catch Err As Exception
            CheckPaketStatus = "Hata"
            ErrDisp("CheckPaketStatus : " + cKoliFisNo + " " + cOriginalBarcode + vbCrLf + Err.Message + vbCrLf + cSQL)
        End Try
    End Function

    Private Function CheckPaket(ByVal cKoliFisNo As String, ByVal cKoliGirisFisNo As String, ByVal ConnYage As SqlConnection) As String

        Dim cSQL As String = ""
        Dim oReader As SqlDataReader
        Dim nGiris As Integer
        Dim nCikis As Integer

        CheckPaket = ""
        nGiris = 1 ' depoya ilk girisi zaten var
        nCikis = 0

        Try

            cSQL = "select a.fistipi " + _
                    " from kolifisi a, kolifislines b " + _
                    " where a.kolifisno = b.kolifisno " + _
                    " and koligirisfisno = '" + cKoliGirisFisNo + "' " + _
                    " and a.kolifisno <> '" + cKoliFisNo + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                If SQLReadString(oReader, "fistipi") = "Giris" Then
                    nGiris = nGiris + 1
                Else
                    nCikis = nCikis + 1
                End If
            Loop
            oReader.Close()
            oReader = Nothing

            If nGiris <= nCikis Then
                CheckPaket = CheckPaket + "Paket cikilmis : " + cKoliGirisFisNo + vbCrLf
            End If

        Catch Err As Exception
            ErrDisp("CheckPaket : " + cKoliFisNo + " " + cKoliGirisFisNo + vbCrLf + Err.Message + vbCrLf + cSQL)
        End Try
    End Function
End Module

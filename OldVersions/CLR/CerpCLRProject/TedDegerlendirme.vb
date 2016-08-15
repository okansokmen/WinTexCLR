Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Module TedDegerlendirme

    Private Structure TDParam
        Dim PuanTipi As String
        Dim Oran As Double
        Dim nMin As Double
        Dim nMax As Double
        Dim Ceza As Double
    End Structure

    Private Function TDReadParam(ByVal ConnYage As SqlConnection, ByVal Rapor As String) As TDParam()
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim cSQL As String
        Dim aTDParam() As TDParam
        TDReadParam = Nothing

        Try
            ReDim aTDParam(0)
            nCnt = 0
            cSQL = "SELECT * FROM Tdparametre WHERE raporadi='" + Rapor.Trim + "' ORDER BY raporadi,puantipi"

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aTDParam(nCnt)
                aTDParam(nCnt).PuanTipi = SQLReadString(oReader, "PuanTipi").Trim
                aTDParam(nCnt).Oran = SQLReadDouble(oReader, "Oran")
                aTDParam(nCnt).nMin = SQLReadDouble(oReader, "Deger1")
                aTDParam(nCnt).nMax = SQLReadDouble(oReader, "Deger2")
                aTDParam(nCnt).Ceza = SQLReadDouble(oReader, "Ceza")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            TDReadParam = aTDParam

        Catch Err As Exception
            ErrDisp("Error TDReadParam" + Err.Message)
        End Try
    End Function
    Private Function GetTDCezaPuan(ByRef aTDParam() As TDParam, ByVal PuanTipi As String, Optional ByVal Fark As Double = 0, Optional ByVal Oran As Boolean = False) As Double
        Dim nCnt As Integer

        GetTDCezaPuan = 0
        If Not Oran Then
            For nCnt = 0 To UBound(aTDParam)
                If aTDParam(nCnt).PuanTipi = PuanTipi And _
                   aTDParam(nCnt).nMin <= Math.Abs(Fark) And _
                   aTDParam(nCnt).nMax >= Math.Abs(Fark) Then
                    GetTDCezaPuan = aTDParam(nCnt).Ceza
                    Exit For
                End If
            Next
        Else
            For nCnt = 0 To UBound(aTDParam)
                If aTDParam(nCnt).PuanTipi = PuanTipi Then
                    GetTDCezaPuan = aTDParam(nCnt).Oran
                    Exit For
                End If
            Next
        End If
    End Function
    Public Function TDKumas(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal FirmaFilter As String) As SqlString

        Dim ConnYage As SqlConnection
        Dim oSysFlags As SysFlags = Nothing
        Dim oReader As SqlDataReader
        Dim aTDParam() As TDParam
        Dim aSQL() As String
        Dim cSQL As String
        Dim cTableName As String
        Dim nMiktarPuan As Double
        Dim nTerminPuan As Double
        Dim nKalitePuan As Double
        Dim nGenelPuan As Double
        Dim nFark As Double
        Dim nTestPuan As Double
        Dim nIadePuan As Double
        Dim nKabulPuan As Double
        Dim nSipMiktar As Double
        Dim nSipGelen As Double
        Dim nTeslimOran As Double
        Dim nTeslimsure As Double
        Dim nSipSayisi As Double
        Dim nCnt As Integer

        Try
            ReadSysFlagsMain(oSysFlags)
            ConnYage = OpenConn()
            cTableName = CreateTempTable(ConnYage)

            TDKumas = ""
            ReDim aSQL(0)
            aTDParam = TDReadParam(ConnYage, "Kumaş")

            cSQL = " SELECT Sno=CAST(ROW_NUMBER()OVER(ORDER BY b.isemrino DESC) as decimal (4,0)), " + _
                   " a.tarih,a.firma,b.isemrino,b.stokno,c.cinsaciklamasi,b.renk,b.termintarihi, " + _
                   " MGelen=COALESCE(SUM(b.tedarikgelen),0), Miktar=COALESCE(SUM(b.miktar1),0), " + _
                   " Miade=(SELECT CAST(COALESCE(SUM(x.netmiktar1),0) as decimal (18,12)) FROM Stokfislines x " + _
                           "WHERE x.isemrino=b.isemrino AND x.stokhareketkodu='02 Tedarikten iade' " + _
                           "AND x.stokno=b.stokno AND x.renk=b.renk AND x.sakatkodu='Miktariade'), " + _
                   " MiktarPuan= NULL, TerminPuan= NULL, KalitePuan= NULL, " + _
                   " Songelis=(SELECT COALESCE(MAX(x.fistarihi),'" + BtsTarihi + "') FROM stokfis x,stokfislines x1  " + _
                              "WHERE x.stokfisno=x1.stokfisno " + _
                              "AND x1.isemrino=b.isemrino AND x1.stokhareketkodu='02 Tedarikten Giris' " + _
                              "AND x1.stokno=b.stokno AND x1.renk=b.renk), " + _
                   " Tiade=(SELECT CAST(COALESCE(SUM(netmiktar1),0) as decimal (18,12)) FROM Stokfislines x " + _
                           "WHERE x.isemrino=b.isemrino AND x.stokhareketkodu='02 Tedarikten iade' " + _
                           "AND x.stokno=b.stokno AND x.renk=b.renk AND x.sakatkodu='Terminiade'), " + _
                   " En= COALESCE(b.miktar2,'0'), " + _
                   " GelenEn=(SELECT COALESCE(AVG(x1.en),0) FROM kesimteyid x,kesimteyidlines x1 " + _
                             "WHERE x.fisno=x1.fisno AND x.isemrino=b.isemrino " + _
                             "AND x.stokno=b.stokno AND x.renk=b.renk), " + _
                   " Gramaj= COALESCE(b.miktar3,'0'), " + _
                   " GelenGr=(SELECT COALESCE(AVG(x1.grm2),0) FROM kesimteyid x,kesimteyidlines x1 " + _
                             "WHERE x.fisno=x1.fisno AND x.isemrino=b.isemrino " + _
                             "AND x.stokno=b.stokno AND x.renk=b.renk), " + _
                   " CekmeEn=(SELECT COALESCE(AVG(x1.encekme),0) FROM kesimteyid x,kesimteyidlines x1 " + _
                             "WHERE x.fisno=x1.fisno AND x.isemrino=b.isemrino " + _
                             "AND x.stokno=b.stokno AND x.renk=b.renk), " + _
                   " CekmeBoy=(SELECT COALESCE(AVG(x1.boycekme),0) FROM kesimteyid x,kesimteyidlines x1 " + _
                              "WHERE x.fisno=x1.fisno AND x.isemrino=b.isemrino  " + _
                              "AND x.stokno=b.stokno AND x.renk=b.renk), " + _
                   " Kiade=(SELECT CAST(COALESCE(SUM(netmiktar1),0) as decimal (18,12)) FROM Stokfislines x  " + _
                           "WHERE x.isemrino=b.isemrino AND x.stokhareketkodu='02 Tedarikten iade'  " + _
                           "AND x.stokno=b.stokno AND x.renk=b.renk AND x.sakatkodu='Kaliteiade'), " + _
                   " PartiPuan=(SELECT CAST(COALESCE(AVG(x2.puan),0) as decimal (18,12)) FROM topongiris x,topongirislines x1,wfctopkontrolfis x2 " + _
                               "WHERE x.toponfisno=x1.toponfisno AND x1.topno=x2.topno AND x1.isemrino=b.isemrino " + _
                               "AND x1.stokno=b.stokno AND x1.renk=b.renk), " + _
                   " OnayR=(SELECT COALESCE(COUNT(x.sonuc),'0') FROM topongiris x,topongirislines x1  " + _
                           "WHERE x.toponfisno=x1.toponfisno AND x.belgeno=b.isemrino " + _
                           "AND x.belgeno=b.isemrino AND x1.stokno=b.stokno AND x1.renk=b.renk AND x.sonuc='Red'), " + _
                   " OnaySK=(SELECT COALESCE(COUNT(x.sonuc),'0') FROM topongiris x,topongirislines x1  " + _
                            "WHERE x.toponfisno=x1.toponfisno AND x.belgeno=b.isemrino " + _
                            "AND x.belgeno=b.isemrino AND x1.stokno=b.stokno AND x1.renk=b.renk AND x.sonuc='Sartlı Kabul'), " + _
                   " Hesap=(SELECT COALESCE(MAX(x.fistarihi),'" + oSysFlags.G_Date + "') FROM stokfis x,stokfislines x1  " + _
                           "WHERE x.stokfisno=x1.stokfisno AND x1.isemrino=b.isemrino AND x1.stokhareketkodu='02 Tedarikten Giris' " + _
                           "AND x1.stokno=b.stokno AND x1.renk=b.renk) " + _
                   " INTO " + cTableName + " " + _
                   " FROM isemri a,isemrilines b, stok c " + _
                   " WHERE a.isemrino=b.isemrino " + _
                   " AND b.stokno=c.stokno " + _
                   " AND a.departman='KUMAS SATINALMA' " + _
                   " AND b.termintarihi>='" + BslTarihi + "' " + _
                   " AND b.termintarihi<='" + BtsTarihi + "' " + _
                   IIf(FirmaFilter = "", "", "AND " + FirmaFilter).ToString + _
                   " GROUP BY a.tarih,a.firma,b.isemrino,b.stokno,c.cinsaciklamasi,b.renk,b.termintarihi,b.miktar2,b.miktar3 "

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

            nCnt = 0
            cSQL = "SELECT * FROM " + cTableName + " ORDER BY sno"

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nMiktarPuan = 100
                nTerminPuan = 100
                nKalitePuan = 100
                nTestPuan = 0
                nIadePuan = 0
                nKabulPuan = 0

                If SQLReadDate(oReader, "Hesap") <> CDate(oSysFlags.G_Date) Then
                    'Miktar Puan=Eğer Miktar iade varsa 100 puan yoksa  Sip. Miktar ile Gelen Miktar arasındaki farka göre Ceza Puanı ver
                    nFark = 0
                    If SQLReadDouble(oReader, "Miade") > 0 Then
                        nMiktarPuan = 0
                    Else
                        nFark = fPercent(SQLReadDouble(oReader, "Miktar"), SQLReadDouble(oReader, "MGelen"))
                        nMiktarPuan = GetTDCezaPuan(aTDParam, "miktar", nFark)
                        nMiktarPuan = 100 - nMiktarPuan
                    End If

                    'Termin Puan=Eğer Termin iade varsa 100 puan yoksa  Sip.Tarihi ile Gelen Tarih arasındaki farka göre Ceza Puanı ver
                    nFark = 0
                    If SQLReadDouble(oReader, "Tiade") > 0 Then
                        nTerminPuan = 0
                    Else
                        nFark = (SQLReadDate(oReader, "Songelis") - SQLReadDate(oReader, "termintarihi")).Days
                        If nFark > 0 Then
                            nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                            nTerminPuan = 100 - nTerminPuan
                        End If
                    End If

                    'Test Puan=Eğer En Fark ±2 üstü ise 20 Puan, Eğer Gramaj Fark ±%5 üstü ise 20 Puan, Eğer En,Boy çekme Fark ±5 üstü ise 20 Puan
                    nFark = 0
                    nFark = (SQLReadDouble(oReader, "En") - SQLReadDouble(oReader, "GelenEn"))
                    If Math.Abs(nFark) > 2 Then nTestPuan = 20

                    nFark = fPercent(SQLReadDouble(oReader, "Gramaj"), SQLReadDouble(oReader, "GelenGr"))
                    If Math.Abs(nFark) > 5 Then nTestPuan = 20

                    If Math.Abs(SQLReadDouble(oReader, "CekmeEn")) > 5 Then
                        nTestPuan = 20
                    ElseIf Math.Abs(SQLReadDouble(oReader, "CekmeBoy")) > 5 Then
                        nTestPuan = 20
                    End If

                    'İade Puan=Eğer İade varsa iade oranına göre ceza ver
                    nFark = 0
                    If SQLReadDouble(oReader, "Kiade") > 0 Then
                        nFark = fPercent(SQLReadDouble(oReader, "Miktar"), SQLReadDouble(oReader, "Kiade"), 1)
                        nIadePuan = GetTDCezaPuan(aTDParam, "kalite", nFark)
                    End If

                    'KabulPuan=Eğer Parti Red ise 100, Şartlı Kabul ise 50
                    If SQLReadInteger(oReader, "OnayR") > 0 Then
                        nKabulPuan = 100
                    ElseIf SQLReadInteger(oReader, "OnaySK") > 0 Then
                        nKabulPuan = 50
                    End If
                    nKalitePuan = fMax(nTestPuan, nIadePuan, nKabulPuan)
                    nKalitePuan = 100 - nKalitePuan

                    If nKalitePuan = 100 Then
                        nKalitePuan = SQLReadDouble(oReader, "PartiPuan") / 4
                        nKalitePuan = 100 - nKalitePuan
                    End If

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " MiktarPuan= " + SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                                 " KalitePuan= " + SQLWriteDecimal(nKalitePuan).ToString + " " + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString

                Else
                    nFark = (SQLReadDate(oReader, "Songelis") - SQLReadDate(oReader, "termintarihi")).Days
                    If nFark > 0 Then
                        nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                        nTerminPuan = 100 - nTerminPuan
                    End If

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString
                End If
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next


            cSQL = " SELECT Firma,Miktar=CAST(COALESCE(AVG(MiktarPuan),0) as decimal (18,12)), " + _
                   " Termin=CAST(COALESCE(AVG(TerminPuan),0)as decimal (18,12)), " + _
                   " Kalite=CAST(COALESCE(AVG(KalitePuan),0)as decimal (18,12)), " + _
                   " SipSayisi=CAST(COALESCE(COUNT(isemrino),0)as decimal (18,12)), " + _
                   " Sipmiktar=CAST(COALESCE(SUM(Miktar),0)as decimal (18,12)), " + _
                   " Gelen=CAST(COALESCE(SUM(MGelen),0)as decimal (18,12)), " + _
                   " Teslimsure=CAST(COALESCE(AVG(CAST((Songelis-tarih)as float (8))),0)as decimal (18,12)) " + _
                   " FROM " + cTableName + _
                   " GROUP BY Firma ORDER BY Firma"

            nCnt = 0

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nMiktarPuan = SQLReadDouble(oReader, "Miktar")
                nTerminPuan = SQLReadDouble(oReader, "Termin")
                nKalitePuan = SQLReadDouble(oReader, "Kalite")
                nGenelPuan = (nMiktarPuan * GetTDCezaPuan(aTDParam, "miktarorani", , True) +
                              nTerminPuan * GetTDCezaPuan(aTDParam, "terminorani", , True) +
                              nKalitePuan * GetTDCezaPuan(aTDParam, "kaliteorani", , True)) / 100
                nSipSayisi = SQLReadDouble(oReader, "SipSayisi")
                nSipMiktar = SQLReadDouble(oReader, "Sipmiktar")
                nSipGelen = SQLReadDouble(oReader, "Gelen")
                nTeslimOran = Math.Round(nSipGelen / nSipMiktar, 12)
                nTeslimsure = Math.Round(SQLReadDouble(oReader, "Teslimsure"), 12)

                ReDim Preserve aSQL(nCnt)
                aSQL(nCnt) = " INSERT INTO " + cTableName + "_Ozet (Firma,Miktar,Termin,Kalite,Genel," + _
                             " SipSayisi,SipMiktar,SipGelen,TeslimOrani,TeslimSuresi) " + _
                             " VALUES ('" + SQLReadString(oReader, "Firma") + "', " + _
                              SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                              SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                              SQLWriteDecimal(nKalitePuan).ToString + ", " + _
                              SQLWriteDecimal(nGenelPuan).ToString + ", " + _
                              SQLWriteDecimal(nSipSayisi).ToString + ", " + _
                              SQLWriteDecimal(nSipMiktar).ToString + ", " + _
                              SQLWriteDecimal(nSipGelen).ToString + ", " + _
                              SQLWriteDecimal(nTeslimOran).ToString + ", " + _
                              SQLWriteDecimal(nTeslimsure).ToString + ") "
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            cSQL = "CREATE TABLE " + cTableName + "_Ozet (Firma char(30) NULL,Miktar decimal(18, 12) NULL,Termin decimal(18, 12) NULL," + _
                   "Kalite decimal(18, 12) NULL,Genel decimal(18, 12) NULL,SipSayisi decimal(18, 12) NULL,SipMiktar decimal(18, 12) NULL," + _
                   "SipGelen decimal(18, 12) NULL,TeslimOrani decimal(18, 12) NULL,TeslimSuresi decimal(18, 12) NULL) ON [PRIMARY]"
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next

            CloseConn(ConnYage)
            TDKumas = cTableName


        Catch Err As Exception
            TDKumas = "Hata"
            ErrDisp("Error Tedarikçi Değerlendirme (Kumaş) " + Err.Message)
        End Try
    End Function
    Public Function TDAksesuar(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal FirmaFilter As String) As SqlString
        Dim ConnYage As SqlConnection
        Dim oSysFlags As SysFlags = Nothing
        Dim oReader As SqlDataReader
        Dim aTDParam() As TDParam
        Dim aSQL() As String
        Dim cSQL As String
        Dim cTableName As String
        Dim nMiktarPuan As Double
        Dim nTerminPuan As Double
        Dim nKalitePuan As Double
        Dim nGenelPuan As Double
        Dim nFark As Double
        Dim nTestPuan As Double
        Dim nATestPuan As Double
        Dim nIadePuan As Double
        Dim nKabulPuan As Double
        Dim nSipMiktar As Double
        Dim nSipGelen As Double
        Dim nTeslimOran As Double
        Dim nTeslimsure As Double
        Dim nSipSayisi As Double
        Dim nCnt As Integer

        Try
            ReadSysFlagsMain(oSysFlags)
            ConnYage = OpenConn()
            cTableName = CreateTempTable(ConnYage)

            TDAksesuar = ""
            ReDim aSQL(0)
            aTDParam = TDReadParam(ConnYage, "Aksesuar")

            cSQL = " SELECT Sno=CAST(ROW_NUMBER()OVER(ORDER BY b.isemrino DESC) as decimal (4,0)), " + _
                   " a.tarih,a.firma,b.isemrino,b.stokno,c.cinsaciklamasi,b.renk,b.termintarihi, " + _
                   " Miktar=COALESCE(SUM(b.miktar1),0), MGelen=COALESCE(SUM(b.tedarikgelen),0), " + _
                   " Miade=(SELECT CAST(COALESCE(SUM(x.netmiktar1),0) as decimal (18,12)) FROM Stokfislines x " + _
                           "WHERE x.isemrino=b.isemrino AND x.stokhareketkodu='02 Tedarikten iade' " + _
                           "AND x.stokno=b.stokno AND x.renk=b.renk AND x.sakatkodu='Miktariade'), " + _
                   " MiktarPuan= NULL, TerminPuan= NULL, KalitePuan= NULL,ATestPuan= CAST(NULL as decimal(18,12)), " + _
                   " TestPuan=CAST(NULL as decimal(18,12)), " + _
                   " Songelis=(SELECT COALESCE(MAX(x.fistarihi),'" + BtsTarihi + "') FROM stokfis x,stokfislines x1  " + _
                              " WHERE x.stokfisno=x1.stokfisno " + _
                              " AND x1.isemrino=b.isemrino AND x1.stokhareketkodu='02 Tedarikten Giris' " + _
                              " AND x1.stokno=b.stokno AND x1.renk=b.renk), " + _
                   " Tiade=(SELECT CAST(COALESCE(SUM(netmiktar1),0) as decimal (18,12)) FROM Stokfislines x " + _
                           " WHERE x.isemrino=b.isemrino AND x.stokhareketkodu='02 Tedarikten iade' " + _
                           " AND x.stokno=b.stokno AND x.renk=b.renk AND x.sakatkodu='Terminiade'), " + _
                   " TestK=(SELECT COALESCE(COUNT(x.testsonuc),0) FROM Aksesuartestlines x " + _
                           " WHERE x.testsonuc='Kabul' AND x.isemrino=b.isemrino " + _
                           " AND x.stokno=b.stokno and x.renk=b.renk), " + _
                   " TestSK=(SELECT COALESCE(COUNT(x.testsonuc),0) FROM Aksesuartestlines x " + _
                            " WHERE x.testsonuc='Sartlı Kabul' AND x.isemrino=b.isemrino " + _
                            " AND x.stokno=b.stokno and x.renk=b.renk), " + _
                   " TestR=(SELECT COALESCE(COUNT(x.testsonuc),0) FROM Aksesuartestlines x " + _
                            " WHERE x.testsonuc='Red' AND x.isemrino=b.isemrino " + _
                            " AND x.stokno=b.stokno and x.renk=b.renk), " + _
                   " TestS=(SELECT COALESCE(COUNT(x.testsonuc),0) FROM Aksesuartestlines x " + _
                            " WHERE x.isemrino=b.isemrino AND x.stokno=b.stokno and x.renk=b.renk), " + _
                   " Kiade=(SELECT CAST(COALESCE(SUM(netmiktar1),0)as decimal (18,12)) FROM Stokfislines x " + _
                           " WHERE x.isemrino=b.isemrino AND x.stokhareketkodu='02 Tedarikten iade' " + _
                           " AND x.stokno=b.stokno AND x.renk=b.renk AND x.sakatkodu='Kaliteiade'), " + _
                   " OnayR=(SELECT COALESCE(COUNT(x.sonuc),'0') FROM aksesuarongiris x,aksesuarongirislines x1  " + _
                          " WHERE x.toponfisno=x1.toponfisno AND x.isEmriNo=b.isemrino " + _
                          " AND x1.stokno=b.stokno AND x1.renk=b.renk AND x.sonuc='Red'), " + _
                   " OnaySK=(SELECT COALESCE(COUNT(x.sonuc),'0') FROM aksesuarongiris x,aksesuarongirislines x1  " + _
                          " WHERE x.toponfisno=x1.toponfisno AND x.isEmriNo=b.isemrino " + _
                          " AND x1.stokno=b.stokno AND x1.renk=b.renk AND x.sonuc='Sartlı Kabul'), " + _
                   " Hesap=(SELECT COALESCE(MAX(x.fistarihi),'" + oSysFlags.G_Date + "') FROM stokfis x,stokfislines x1  " + _
                           "WHERE x.stokfisno=x1.stokfisno AND x1.isemrino=b.isemrino AND x1.stokhareketkodu='02 Tedarikten Giris' " + _
                           "AND x1.stokno=b.stokno AND x1.renk=b.renk) " + _
                   " INTO " + cTableName + " " + _
                   " FROM isemri a,isemrilines b, stok c " + _
                   " WHERE a.isemrino=b.isemrino " + _
                   " AND b.stokno=c.stokno " + _
                   " AND a.departman='AKSESUAR SATINALMA' " + _
                   " AND b.termintarihi>='" + BslTarihi + "' " + _
                   " AND b.termintarihi<='" + BtsTarihi + "' " + _
                   IIf(FirmaFilter = "", "", "AND " + FirmaFilter).ToString + _
                   " GROUP BY a.tarih,a.firma,b.isemrino,b.stokno,c.cinsaciklamasi,b.renk,b.termintarihi "

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

            nCnt = 0
            cSQL = "SELECT * FROM " + cTableName + " ORDER BY sno"

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nMiktarPuan = 100
                nTerminPuan = 100
                nKalitePuan = 100
                nKabulPuan = 0
                nATestPuan = 0
                nTestPuan = 0
                nIadePuan = 0

                If SQLReadDate(oReader, "Hesap") <> CDate(oSysFlags.G_Date) Then
                    'Miktar Puan=Eğer Miktar iade varsa 100 puan yoksa  Sip. Miktar ile Gelen Miktar arasındaki farka göre Ceza Puanı ver
                    nFark = 0
                    If SQLReadDouble(oReader, "Miade") > 0 Then
                        nMiktarPuan = 0
                    Else
                        nFark = fPercent(SQLReadDouble(oReader, "Miktar"), SQLReadDouble(oReader, "MGelen"))
                        nMiktarPuan = GetTDCezaPuan(aTDParam, "miktar", nFark)
                        nMiktarPuan = 100 - nMiktarPuan
                    End If

                    'Termin Puan=Eğer Termin iade varsa 100 puan yoksa  Sip.Tarihi ile Gelen Tarih arasındaki farka göre Ceza Puanı ver
                    nFark = 0
                    If SQLReadDouble(oReader, "Tiade") > 0 Then
                        nTerminPuan = 0
                    Else
                        nFark = (SQLReadDate(oReader, "Songelis") - SQLReadDate(oReader, "termintarihi")).Days
                        If nFark > 0 Then
                            nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                            nTerminPuan = 100 - nTerminPuan
                        End If
                    End If

                    'Ağırlıklı Test Puanı 
                    nTestPuan = 0
                    nATestPuan = CDbl(SQLReadInteger(oReader, "TestS"))
                    If nATestPuan > 0 Then
                        nATestPuan = Math.Round(100 / nATestPuan, 2)
                        nTestPuan = nATestPuan * CDbl(SQLReadInteger(oReader, "TestR"))
                        nTestPuan = Math.Round(nTestPuan + ((nATestPuan / 2) * CDbl(SQLReadInteger(oReader, "TestSK"))), 2)
                    End If

                    'İade Puan=Eğer İade varsa iade oranına göre ceza ver
                    nFark = 0
                    If SQLReadDouble(oReader, "Kiade") > 0 Then
                        nFark = fPercent(SQLReadDouble(oReader, "Miktar"), SQLReadDouble(oReader, "Kiade"), 1)
                        nIadePuan = GetTDCezaPuan(aTDParam, "kalite", nFark)
                    End If

                    'KabulPuan=Eğer Parti Red ise 100, Şartlı Kabul ise 50
                    If SQLReadInteger(oReader, "OnayR") > 0 Then
                        nKabulPuan = 100
                    ElseIf SQLReadInteger(oReader, "OnaySK") > 0 Then
                        nKabulPuan = 50
                    End If
                    nKalitePuan = fMax(nTestPuan, nIadePuan, nKabulPuan)
                    nKalitePuan = 100 - nKalitePuan

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " MiktarPuan= " + SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                                 " ATestPuan= " + SQLWriteDecimal(nATestPuan).ToString + ", " + _
                                 " TestPuan= " + SQLWriteDecimal(nTestPuan).ToString + ", " + _
                                 " KalitePuan= " + SQLWriteDecimal(nKalitePuan).ToString + " " + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString

                Else
                    nFark = (SQLReadDate(oReader, "Songelis") - SQLReadDate(oReader, "termintarihi")).Days
                    If nFark > 0 Then
                        nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                        nTerminPuan = 100 - nTerminPuan
                    End If

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString
                End If
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next


            cSQL = " SELECT Firma,Miktar=CAST(COALESCE(AVG(MiktarPuan),0) as decimal (18,12)), " + _
                   " Termin=CAST(COALESCE(AVG(TerminPuan),0)as decimal (18,12)), " + _
                   " Kalite=CAST(COALESCE(AVG(KalitePuan),0)as decimal (18,12)), " + _
                   " SipSayisi=CAST(COALESCE(COUNT(isemrino),0)as decimal (18,12)), " + _
                   " Sipmiktar=CAST(COALESCE(SUM(Miktar),0)as decimal (18,12)), " + _
                   " Gelen=CAST(COALESCE(SUM(MGelen),0)as decimal (18,12)), " + _
                   " Teslimsure=CAST(COALESCE(AVG(CAST((Songelis-tarih)as float (8))),0)as decimal (18,12)) " + _
                   " FROM " + cTableName + _
                   " GROUP BY Firma ORDER BY Firma"

            nCnt = 0

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                nMiktarPuan = SQLReadDouble(oReader, "Miktar")
                nTerminPuan = SQLReadDouble(oReader, "Termin")
                nKalitePuan = SQLReadDouble(oReader, "Kalite")
                nGenelPuan = (nMiktarPuan * GetTDCezaPuan(aTDParam, "miktarorani", , True) +
                              nTerminPuan * GetTDCezaPuan(aTDParam, "terminorani", , True) +
                              nKalitePuan * GetTDCezaPuan(aTDParam, "kaliteorani", , True)) / 100
                nSipSayisi = SQLReadDouble(oReader, "SipSayisi")
                nSipMiktar = SQLReadDouble(oReader, "Sipmiktar")
                nSipGelen = SQLReadDouble(oReader, "Gelen")
                nTeslimOran = Math.Round(nSipGelen / nSipMiktar, 12)
                nTeslimsure = Math.Round(SQLReadDouble(oReader, "Teslimsure"), 12)

                ReDim Preserve aSQL(nCnt)
                aSQL(nCnt) = " INSERT INTO " + cTableName + "_Ozet (Firma,Miktar,Termin,Kalite,Genel, " + _
                             " SipSayisi,SipMiktar,SipGelen,TeslimOrani,TeslimSuresi,Ihtisas) " + _
                             " VALUES ('" + SQLReadString(oReader, "Firma") + "', " + _
                              SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                              SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                              SQLWriteDecimal(nKalitePuan).ToString + ", " + _
                              SQLWriteDecimal(nGenelPuan).ToString + ", " + _
                              SQLWriteDecimal(nSipSayisi).ToString + ", " + _
                              SQLWriteDecimal(nSipMiktar).ToString + ", " + _
                              SQLWriteDecimal(nSipGelen).ToString + ", " + _
                              SQLWriteDecimal(nTeslimOran).ToString + ", " + _
                              SQLWriteDecimal(nTeslimsure).ToString + ", Null ) "
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            cSQL = " CREATE TABLE " + cTableName + "_Ozet (Firma char(30) NULL,Miktar decimal(18, 12) NULL,Termin decimal(18, 12) NULL, " + _
                   " Kalite decimal(18, 12) NULL,Genel decimal(18, 12) NULL,SipSayisi decimal(18, 12) NULL,SipMiktar decimal(18, 12) NULL, " + _
                   " SipGelen decimal(18, 12) NULL,TeslimOrani decimal(18, 12) NULL,TeslimSuresi decimal(18, 12) NULL,Ihtisas char(30) NULL) ON [PRIMARY] "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next

            cSQL = " UPDATE " + cTableName + "_Ozet SET Ihtisas=x.ihtisas " + _
                   " FROM " + cTableName + "_Ozet a, Firma x " + _
                   " WHERE x.Firma=a.Firma"

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            CloseConn(ConnYage)
            TDAksesuar = cTableName


        Catch Err As Exception
            TDAksesuar = "Hata"
            ErrDisp("Error Tedarikçi Değerlendirme (Aksesuar) " + Err.Message)
        End Try

    End Function
    Public Function TDUretim(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal FirmaFilter As String) As SqlString

        Dim ConnYage As SqlConnection
        Dim oSysFlags As SysFlags = Nothing
        Dim oReader As SqlDataReader
        Dim aTDParam() As TDParam
        Dim aSQL() As String
        Dim cSQL As String
        Dim cTableName As String
        Dim cInspection As String
        Dim nMiktarPuan As Double
        Dim nTerminPuan As Double
        Dim nKalitePuan As Double
        Dim nGenelPuan As Double
        Dim nTamirPuan As Double
        Dim nKabulPuan As Double
        Dim nSipMiktar As Double
        Dim nSipGelen As Double
        Dim nTeslimOran As Double
        Dim nTeslimsure As Double
        Dim nSipSayisi As Double
        Dim nPlnSure As Double
        Dim nGrcSure As Double
        Dim nOlcuPuan As Double
        Dim nCnt As Integer
        Dim nFark As Double

        Try
            ReadSysFlagsMain(oSysFlags)
            ConnYage = OpenConn()
            cTableName = CreateTempTable(ConnYage)

            TDUretim = ""
            ReDim aSQL(0)
            aTDParam = TDReadParam(ConnYage, "Üretim")

            cSQL = " SELECT Sno=CAST(ROW_NUMBER()OVER(ORDER BY a.uretimtakipno DESC) as decimal (4,0)), " + _
                   " a.firma,a.departman,a.uretimtakipno,b.modelno,Plbsltarihi=b.baslama_tar,Plbtstarihi=b.bitis_tar, " + _
                   " SipMiktar=CAST(COALESCE(SUM(b.toplamadet),0) as decimal (6,0)), " + _
                   " Renk=(SELECT DISTINCT x.renk FROM sipmodel x WHERE x.siparisno=a.uretimtakipno), " + _
                   " Bsltarihi=(SELECT MIN(x.fistarihi) FROM uretharfis x,uretharfislines x1 " + _
                               "WHERE x.uretfisno=x1.uretfisno AND x.girisdept=a.departman AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " Btstarihi=(SELECT COALESCE(MAX(x.fistarihi),'" + BtsTarihi + "') FROM uretharfis x,uretharfislines x1 " + _
                               "WHERE x.uretfisno=x1.uretfisno AND x.cikisdept=a.departman AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " CikanAdet=(SELECT CAST(COALESCE(SUM(x1.toplamadet),0) as decimal (6,0)) FROM uretharfis x,uretharfislines x1 " + _
                               "WHERE x.uretfisno=x1.uretfisno AND x.girisdept=a.departman AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " GelenAdet=(SELECT CAST(COALESCE(SUM(x1.toplamadet),0) as decimal (6,0)) FROM uretharfis x,uretharfislines x1 " + _
                               "WHERE x.uretfisno=x1.uretfisno AND x.cikisdept=a.departman AND x.girisdept<>'FIRE' AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " FireAdet=(SELECT CAST(COALESCE(SUM(x1.toplamadet),0) as decimal (6,0)) FROM uretharfis x,uretharfislines x1 " + _
                              "WHERE x.uretfisno=x1.uretfisno AND x.cikisdept=a.departman AND x.girisdept='FIRE' AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " Tamir=CASE a.departman " + _
                         " WHEN 'BASKI' THEN (SELECT CAST(COALESCE(SUM(x.baskiad),0) as decimal (6,0)) FROM utupaket x WHERE x.siparisno=a.uretimtakipno) " + _
                         " WHEN 'NAKIS' THEN (SELECT CAST(COALESCE(SUM(x.nakisad),0) as decimal (6,0)) FROM utupaket x WHERE x.siparisno=a.uretimtakipno) " + _
                         " WHEN 'DIKIM' THEN (SELECT CAST(COALESCE(SUM(x.uretimad),0) as decimal (6,0)) FROM utupaket x WHERE x.siparisno=a.uretimtakipno) " + _
                         " WHEN 'YIKAMA' THEN (SELECT CAST(COALESCE(SUM(x.yikamaad),0) as decimal (6,0)) FROM utupaket x WHERE x.siparisno=a.uretimtakipno) " + _
                         " WHEN 'UTU&PAKET' THEN (SELECT CAST(COALESCE(SUM(x.Drivetad),0) as decimal (6,0)) FROM utupaket x WHERE x.siparisno=a.uretimtakipno) END, " + _
                   " OnayR=CASE a.departman " + _
                         " WHEN 'DIKIM' THEN (SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x " + _
                                            "WHERE x.siparisno=a.uretimtakipno AND x.oksafhasi='1.SAFHA' AND x.oktipi='Red') " + _
                         " WHEN 'YIKAMA' THEN (SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x " + _
                                              "WHERE x.siparisno=a.uretimtakipno AND x.oksafhasi='3.SAFHA' AND x.oktipi='Red') " + _
                         " WHEN 'UTU&PAKET' THEN (SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x " + _
                                               " WHERE x.siparisno=a.uretimtakipno AND x.oksafhasi='2.SAFHA' AND x.oktipi='Red') END, " + _
                  " OnaySK=CASE a.departman " + _
                         " WHEN 'DIKIM' THEN (SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x " + _
                                            "WHERE x.siparisno=a.uretimtakipno AND x.oksafhasi='1.SAFHA' AND x.oktipi='Sartlı Kabul') " + _
                         " WHEN 'YIKAMA' THEN (SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x " + _
                                             "WHERE x.siparisno=a.uretimtakipno AND x.oksafhasi='3.SAFHA' AND x.oktipi='Sartlı Kabul') " + _
                         " WHEN 'UTU&PAKET' THEN (SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x " + _
                                                 "WHERE x.siparisno=a.uretimtakipno AND x.oksafhasi='2.SAFHA' AND x.oktipi='Sartlı Kabul') END, " + _
                   " Hesap=(SELECT COALESCE(MAX(x.fistarihi),'" + oSysFlags.G_Date + "') FROM uretharfis x,uretharfislines x1 " + _
                           "WHERE x.uretfisno=x1.uretfisno AND x.girisdept=a.departman AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " OlcuPuan=CAST(NULL as decimal(3,0)), PlanlananSure=NULL, Gerceklesensure=NULL, Inspection=CAST(NULL as char(30)), MiktarPuan=NULL, TerminPuan=NULL, KalitePuan=NULL " + _
                   " INTO " + cTableName + " " + _
                   " FROM uretimisemri a, uretimisdetayi b " + _
                   " WHERE b.uretimtakipno=a.uretimtakipno " + _
                   " AND a.isemrino=b.isemrino " + _
                   " AND b.bitis_tar>='" + BslTarihi + "' " + _
                   " AND b.bitis_tar<='" + BtsTarihi + "' " + _
                   " AND a.departman not in ('URETIM PLANLAMA','KESIM','MAMUL','TASNIF','TASNIF1','FIRE') " + _
                   IIf(FirmaFilter = "", "", "AND " + FirmaFilter).ToString + _
                   " GROUP BY a.firma,a.departman,a.uretimtakipno,b.modelno,b.baslama_tar,b.bitis_tar "

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

            cSQL = "DELETE FROM " + cTableName + " WHERE uretimtakipno " + _
                   "IN (SELECT kullanicisipno FROM siparis WHERE UrtGrubu IN('Ticari Ürün','Ticari Aksesuar'))"
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = 0
            cSQL = "SELECT * FROM " + cTableName + " ORDER BY sno"

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nMiktarPuan = 100
                nTerminPuan = 100
                nKalitePuan = 100
                nOlcuPuan = 0
                nTamirPuan = 0
                nKabulPuan = 0

                nPlnSure = (SQLReadDate(oReader, "Plbtstarihi") - SQLReadDate(oReader, "Plbsltarihi")).Days
                nGrcSure = (SQLReadDate(oReader, "Btstarihi") - SQLReadDate(oReader, "Bsltarihi")).Days
                If nPlnSure = 0 Then nPlnSure = 1
                If nGrcSure = 0 Then nGrcSure = 1

                If SQLReadDate(oReader, "Hesap") <> CDate(oSysFlags.G_Date) Then
                    'Miktar Puan= Çıkan Miktar ile Fire Miktarı arasındaki orana göre Ceza Puanı ver
                    nFark = 0
                    nFark = fPercent(SQLReadDouble(oReader, "CikanAdet"), SQLReadDouble(oReader, "FireAdet"), 1)
                    nMiktarPuan = GetTDCezaPuan(aTDParam, "miktar", nFark)
                    nMiktarPuan = 100 - nMiktarPuan

                    'Termin Puan=Planlanan Süre ile Gercekleşen Süre arasındaki farka göre Ceza Puanı ver
                    nFark = 0
                    If nGrcSure > nPlnSure Then nFark = nPlnSure - nGrcSure
                    nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                    nTerminPuan = 100 - nTerminPuan

                    'Tamir Puanı 
                    nFark = 0
                    nFark = fPercent(SQLReadDouble(oReader, "CikanAdet"), SQLReadDouble(oReader, "Tamir"), 1)
                    nTamirPuan = GetTDCezaPuan(aTDParam, "kalite", nFark)

                    'Ölçü Puanı 
                    If SQLReadString(oReader, "departman") = "DIKIM" Then nOlcuPuan = GetOlcuPuan(SQLReadString(oReader, "uretimtakipno"))

                    'KabulPuan=Eğer Parti Red ise 100, Şartlı Kabul ise 50
                    cInspection = "Kabul"
                    If SQLReadInteger(oReader, "OnayR") > 0 Then
                        nKabulPuan = 100
                        cInspection = "Red"
                    ElseIf SQLReadInteger(oReader, "OnaySK") > 0 Then
                        nKabulPuan = 50
                        cInspection = "Sartlı Kabul"
                    End If

                    nKalitePuan = fMax(nTamirPuan, nOlcuPuan, nKabulPuan)
                    nKalitePuan = 100 - nKalitePuan

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " MiktarPuan= " + SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                                 " PlanlananSure= " + SQLWriteDecimal(nPlnSure).ToString + ", " + _
                                 " Gerceklesensure= " + SQLWriteDecimal(nGrcSure).ToString + ", " + _
                                 " OlcuPuan=" + SQLWriteDecimal(nOlcuPuan).ToString + ", " + _
                                 " KalitePuan= " + SQLWriteDecimal(nKalitePuan).ToString + ", " + _
                                 " Inspection= '" + cInspection + "' " + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString

                Else
                    nFark = 0
                    If nGrcSure > nPlnSure Then nFark = nPlnSure - nGrcSure
                    nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                    nTerminPuan = 100 - nTerminPuan

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " PlanlananSure= " + SQLWriteDecimal(nPlnSure).ToString + ", " + _
                                 " Gerceklesensure= " + SQLWriteDecimal(nGrcSure).ToString + ", " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString
                End If
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next


            cSQL = " SELECT Firma,Miktar=CAST(COALESCE(AVG(MiktarPuan),0) as decimal (18,12)), " + _
                   " Termin=CAST(COALESCE(AVG(TerminPuan),0)as decimal (18,12)), " + _
                   " Kalite=CAST(COALESCE(AVG(KalitePuan),0)as decimal (18,12)), " + _
                   " SipSayisi=CAST(COALESCE(COUNT(uretimtakipno),0)as decimal (18,12)), " + _
                   " Sipmiktar=CAST(COALESCE(SUM(Sipmiktar),0)as decimal (18,12)), " + _
                   " Gelen=CAST(COALESCE(SUM(GelenAdet),0)as decimal (18,12)), " + _
                   " Teslimsure=CAST(COALESCE(AVG(Gerceklesensure),0)as decimal (18,12)) " + _
                   " FROM " + cTableName + _
                   " GROUP BY Firma ORDER BY Firma"

            nCnt = 0

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                nMiktarPuan = SQLReadDouble(oReader, "Miktar")
                nTerminPuan = SQLReadDouble(oReader, "Termin")
                nKalitePuan = SQLReadDouble(oReader, "Kalite")
                nGenelPuan = (nMiktarPuan * GetTDCezaPuan(aTDParam, "miktarorani", , True) +
                              nTerminPuan * GetTDCezaPuan(aTDParam, "terminorani", , True) +
                              nKalitePuan * GetTDCezaPuan(aTDParam, "kaliteorani", , True)) / 100
                nSipSayisi = SQLReadDouble(oReader, "SipSayisi")
                nSipMiktar = SQLReadDouble(oReader, "Sipmiktar")
                nSipGelen = SQLReadDouble(oReader, "Gelen")
                nTeslimOran = Math.Round(nSipGelen / nSipMiktar, 12)
                nTeslimsure = Math.Round(SQLReadDouble(oReader, "Teslimsure"), 12)

                ReDim Preserve aSQL(nCnt)
                aSQL(nCnt) = " INSERT INTO " + cTableName + "_Ozet (Firma,Miktar,Termin,Kalite,Genel, " + _
                             " SipSayisi,SipMiktar,SipGelen,TeslimOrani,TeslimSuresi) " + _
                             " VALUES ('" + SQLReadString(oReader, "Firma") + "', " + _
                              SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                              SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                              SQLWriteDecimal(nKalitePuan).ToString + ", " + _
                              SQLWriteDecimal(nGenelPuan).ToString + ", " + _
                              SQLWriteDecimal(nSipSayisi).ToString + ", " + _
                              SQLWriteDecimal(nSipMiktar).ToString + ", " + _
                              SQLWriteDecimal(nSipGelen).ToString + ", " + _
                              SQLWriteDecimal(nTeslimOran).ToString + ", " + _
                              SQLWriteDecimal(nTeslimsure).ToString + ") "
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            cSQL = " CREATE TABLE " + cTableName + "_Ozet (Firma char(30) NULL,Miktar decimal(18, 12) NULL,Termin decimal(18, 12) NULL, " + _
                   " Kalite decimal(18, 12) NULL,Genel decimal(18, 12) NULL,SipSayisi decimal(18, 12) NULL,SipMiktar decimal(18, 12) NULL, " + _
                   " SipGelen decimal(18, 12) NULL,TeslimOrani decimal(18, 12) NULL,TeslimSuresi decimal(18, 12) NULL) ON [PRIMARY] "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next


            CloseConn(ConnYage)
            TDUretim = cTableName

        Catch Err As Exception
            TDUretim = "Hata"
            ErrDisp("Error Tedarikçi Değerlendirme (Üretim) " + Err.Message)
        End Try

    End Function
    Public Function TDTUretim(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal FirmaFilter As String) As SqlString

        Dim ConnYage As SqlConnection
        Dim oSysFlags As SysFlags = Nothing
        Dim oReader As SqlDataReader
        Dim aTDParam() As TDParam
        Dim aSQL() As String
        Dim cSQL As String
        Dim cTableName As String
        Dim cInspection As String
        Dim nMiktarPuan As Double
        Dim nTerminPuan As Double
        Dim nKalitePuan As Double
        Dim nGenelPuan As Double
        Dim nKabulPuan As Double
        Dim nSipMiktar As Double
        Dim nSipGelen As Double
        Dim nTeslimOran As Double
        Dim nTeslimsure As Double
        Dim nSipSayisi As Double
        Dim nOlcuPuan As Double
        Dim nCnt As Integer
        Dim nFark As Double
        Dim nPlnSure As Double
        Dim nGrcSure As Double

        Try
            ReadSysFlagsMain(oSysFlags)
            ConnYage = OpenConn()
            cTableName = CreateTempTable(ConnYage)

            TDTUretim = ""
            ReDim aSQL(0)
            aTDParam = TDReadParam(ConnYage, "Ticari Üretim")

            cSQL = " SELECT Sno=CAST(ROW_NUMBER()OVER(ORDER BY a.uretimtakipno DESC) as decimal (4,0)), " + _
                   " a.firma,a.uretimtakipno,b.modelno,Plbtstarihi=b.bitis_tar, " + _
                   " Renk=(SELECT DISTINCT x.renk FROM sipmodel x WHERE x.siparisno=a.uretimtakipno), " + _
                   " Sipmiktar=CAST(ISNULL((SELECT x1.toplamadet FROM uretharfis x,uretharfislines x1 " + _
                              "WHERE x.uretfisno=x1.uretfisno AND x1.uretimtakipno=a.uretimtakipno AND x.cikisdept='KESIM')," + _
                              "(SELECT SUM(x.adet) FROM sipmodel x WHERE x.siparisno=a.uretimtakipno))as decimal (6,0)), " + _
                   " GelenAdet=CAST((SELECT COALESCE(SUM(netmiktar1),0) FROM stokfislines x " + _
                              "WHERE x.malzemetakipkodu=a.uretimtakipno AND x.stokhareketkodu='02 Tedarikten Giris' " + _
                              "AND x.stokno=b.modelno AND x.renk=(SELECT DISTINCT x1.renk FROM sipmodel x1 " + _
                              "WHERE x1.siparisno= a.uretimtakipno))as decimal (6,0)), " + _
                   " Plbsltarihi=(SELECT MIN(x1.baslama_tar) FROM uretimisemri x,uretimisdetayi x1 " + _
                                 "WHERE x.isemrino=x1.isemrino AND x.departman='KESIM' AND x.uretimtakipno=a.uretimtakipno), " + _
                   " Bsltarihi=(SELECT MIN(x.fistarihi) FROM uretharfis x,uretharfislines x1 " + _
                               "WHERE x.uretfisno=x1.uretfisno AND x.girisdept='KESIM' AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " Btstarihi=(SELECT COALESCE(MAX(x.fistarihi),'" + BtsTarihi + "') FROM uretharfis x,uretharfislines x1 " + _
                               "WHERE x.uretfisno=x1.uretfisno AND x.cikisdept=a.departman AND x.girisdept<>'FIRE' AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " OnayR=(SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x WHERE x.siparisno=a.uretimtakipno " + _
                           "AND x.oksafhasi='2.SAFHA' AND x.oktipi='Red'), " + _
                   " OnaySK=(SELECT COALESCE(COUNT(x.oktipi),'0') FROM sipok x WHERE x.siparisno=a.uretimtakipno " + _
                            "AND x.oksafhasi='2.SAFHA' AND x.oktipi='Sartlı Kabul'), " + _
                   " Hesap=(SELECT COALESCE(MAX(x.fistarihi),'" + oSysFlags.G_Date + "') FROM uretharfis x,uretharfislines x1 " + _
                           "WHERE x.uretfisno=x1.uretfisno AND x.cikisdept=a.departman AND x.girisdept<>'FIRE' AND x1.uretimtakipno=a.uretimtakipno), " + _
                   " OlcuPuan=CAST(NULL as decimal(3,0)),MiktarPuan=NULL, TerminPuan=NULL, KalitePuan=NULL, PlanlananSure=NULL, " + _
                   " Gerceklesensure=NULL, Inspection=CAST(NULL as char(30)) " + _
                   " INTO " + cTableName + " " + _
                   " FROM uretimisemri a, uretimisdetayi b " + _
                   " WHERE b.uretimtakipno=a.uretimtakipno " + _
                   " AND a.isemrino=b.isemrino " + _
                   " AND b.bitis_tar>='" + BslTarihi + "' " + _
                   " AND b.bitis_tar<='" + BtsTarihi + "' " + _
                   " AND a.departman='FASON' " + _
                   IIf(FirmaFilter = "", "", "AND " + FirmaFilter).ToString + _
                   " GROUP BY a.firma,a.departman,a.uretimtakipno,b.modelno,b.bitis_tar "

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

            cSQL = "DELETE FROM " + cTableName + " WHERE uretimtakipno " + _
                   "NOT IN (SELECT kullanicisipno FROM siparis WHERE UrtGrubu IN('Ticari Ürün','Ticari Aksesuar'))"
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = 0
            cSQL = "SELECT * FROM " + cTableName + " ORDER BY sno"

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nMiktarPuan = 100
                nTerminPuan = 100
                nKalitePuan = 100
                nOlcuPuan = 0
                nKabulPuan = 0

                nPlnSure = (SQLReadDate(oReader, "Plbtstarihi") - SQLReadDate(oReader, "Plbsltarihi")).Days
                nGrcSure = (SQLReadDate(oReader, "Btstarihi") - SQLReadDate(oReader, "Bsltarihi")).Days
                If nPlnSure = 0 Then nPlnSure = 1
                If nGrcSure = 0 Then nGrcSure = 1

                If SQLReadDate(oReader, "Hesap") <> CDate(oSysFlags.G_Date) Then
                    'Miktar Puan= Çıkan Miktar ile Fire Miktarı arasındaki orana göre Ceza Puanı ver
                    nFark = 0
                    nFark = fPercent(SQLReadDouble(oReader, "Sipmiktar"), SQLReadDouble(oReader, "GelenAdet"))
                    nMiktarPuan = GetTDCezaPuan(aTDParam, "miktar", nFark)
                    nMiktarPuan = 100 - nMiktarPuan

                    'Termin Puan=Planlanan Süre ile Gercekleşen Süre arasındaki farka göre Ceza Puanı ver
                    nFark = 0
                    If nGrcSure > nPlnSure Then nFark = nPlnSure - nGrcSure
                    nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                    nTerminPuan = 100 - nTerminPuan

                    'Ölçü Puanı 
                    nOlcuPuan = GetOlcuPuan(SQLReadString(oReader, "uretimtakipno"))

                    'KabulPuan=Eğer Parti Red ise 100, Şartlı Kabul ise 50
                    cInspection = "Kabul"
                    If SQLReadInteger(oReader, "OnayR") > 0 Then
                        nKabulPuan = 100
                        cInspection = "Red"
                    ElseIf SQLReadInteger(oReader, "OnaySK") > 0 Then
                        nKabulPuan = 50
                        cInspection = "Sartlı Kabul"
                    End If

                    nKalitePuan = fMax(nOlcuPuan, nKabulPuan)
                    nKalitePuan = 100 - nKalitePuan

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " MiktarPuan= " + SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                                 " PlanlananSure= " + SQLWriteDecimal(nPlnSure).ToString + ", " + _
                                 " Gerceklesensure= " + SQLWriteDecimal(nGrcSure).ToString + ", " + _
                                 " OlcuPuan=" + SQLWriteDecimal(nOlcuPuan).ToString + ", " + _
                                 " KalitePuan= " + SQLWriteDecimal(nKalitePuan).ToString + ", " + _
                                 " Inspection= '" + cInspection + "' " + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString

                Else
                    nFark = 0
                    If nGrcSure > nPlnSure Then nFark = nPlnSure - nGrcSure
                    nTerminPuan = GetTDCezaPuan(aTDParam, "termin", nFark)
                    nTerminPuan = 100 - nTerminPuan

                    ReDim Preserve aSQL(nCnt)
                    aSQL(nCnt) = " UPDATE " + cTableName + " SET " + _
                                 " PlanlananSure= " + SQLWriteDecimal(nPlnSure).ToString + ", " + _
                                 " Gerceklesensure= " + SQLWriteDecimal(nGrcSure).ToString + ", " + _
                                 " TerminPuan= " + SQLWriteDecimal(nTerminPuan).ToString + _
                                 " WHERE sno=  " + SQLWriteDecimal(SQLReadDouble(oReader, "Sno")).ToString
                End If
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next


            cSQL = " SELECT Firma,Miktar=CAST(COALESCE(AVG(MiktarPuan),0) as decimal (18,12)), " + _
                   " Termin=CAST(COALESCE(AVG(TerminPuan),0)as decimal (18,12)), " + _
                   " Kalite=CAST(COALESCE(AVG(KalitePuan),0)as decimal (18,12)), " + _
                   " SipSayisi=CAST(COALESCE(COUNT(uretimtakipno),0)as decimal (18,12)), " + _
                   " Sipmiktar=CAST(COALESCE(SUM(Sipmiktar),0)as decimal (18,12)), " + _
                   " Gelen=CAST(COALESCE(SUM(GelenAdet),0)as decimal (18,12)), " + _
                   " Teslimsure=CAST(COALESCE(AVG(Gerceklesensure),0)as decimal (18,12)) " + _
                   " FROM " + cTableName + _
                   " GROUP BY Firma ORDER BY Firma"

            nCnt = 0

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                nMiktarPuan = SQLReadDouble(oReader, "Miktar")
                nTerminPuan = SQLReadDouble(oReader, "Termin")
                nKalitePuan = SQLReadDouble(oReader, "Kalite")
                nGenelPuan = (nMiktarPuan * GetTDCezaPuan(aTDParam, "miktarorani", , True) +
                              nTerminPuan * GetTDCezaPuan(aTDParam, "terminorani", , True) +
                              nKalitePuan * GetTDCezaPuan(aTDParam, "kaliteorani", , True)) / 100
                nSipSayisi = SQLReadDouble(oReader, "SipSayisi")
                nSipMiktar = SQLReadDouble(oReader, "Sipmiktar")
                nSipGelen = SQLReadDouble(oReader, "Gelen")
                nTeslimOran = Math.Round(nSipGelen / nSipMiktar, 12)
                nTeslimsure = Math.Round(SQLReadDouble(oReader, "Teslimsure"), 12)

                ReDim Preserve aSQL(nCnt)
                aSQL(nCnt) = " INSERT INTO " + cTableName + "_Ozet (Firma,Miktar,Termin,Kalite,Genel, " + _
                             " SipSayisi,SipMiktar,SipGelen,TeslimOrani,TeslimSuresi) " + _
                             " VALUES ('" + SQLReadString(oReader, "Firma") + "', " + _
                              SQLWriteDecimal(nMiktarPuan).ToString + ", " + _
                              SQLWriteDecimal(nTerminPuan).ToString + ", " + _
                              SQLWriteDecimal(nKalitePuan).ToString + ", " + _
                              SQLWriteDecimal(nGenelPuan).ToString + ", " + _
                              SQLWriteDecimal(nSipSayisi).ToString + ", " + _
                              SQLWriteDecimal(nSipMiktar).ToString + ", " + _
                              SQLWriteDecimal(nSipGelen).ToString + ", " + _
                              SQLWriteDecimal(nTeslimOran).ToString + ", " + _
                              SQLWriteDecimal(nTeslimsure).ToString + ") "
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            cSQL = " CREATE TABLE " + cTableName + "_Ozet (Firma char(30) NULL,Miktar decimal(18, 12) NULL,Termin decimal(18, 12) NULL, " + _
                   " Kalite decimal(18, 12) NULL,Genel decimal(18, 12) NULL,SipSayisi decimal(18, 12) NULL,SipMiktar decimal(18, 12) NULL, " + _
                   " SipGelen decimal(18, 12) NULL,TeslimOrani decimal(18, 12) NULL,TeslimSuresi decimal(18, 12) NULL) ON [PRIMARY] "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            For nCnt = 0 To UBound(aSQL)
                ExecuteSQLCommandConnected(aSQL(nCnt), ConnYage)
            Next


            CloseConn(ConnYage)
            TDTUretim = cTableName

        Catch Err As Exception
            TDTUretim = "Hata"
            ErrDisp("Error Tedarikçi Değerlendirme (Ticari Üretim) " + Err.Message)
        End Try


    End Function
    Public Function TDGHizmet(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal Firma As String) As SqlString

        Dim ConnYage As SqlConnection
        ' Dim cSQL As String
        Dim cTableName As String

        Try
            ConnYage = OpenConn()
            cTableName = CreateTempTable(ConnYage)
            TDGHizmet = ""

        Catch Err As Exception
            TDGHizmet = "Hata"
            ErrDisp("Error Tedarikçi Değerlendirme (Genel Hizmet) " + Err.Message)
        End Try

    End Function
    Public Function GetOlcuPuan(ByVal SiparisNo As String) As Double
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cSQL As String
        Dim nOlcuPuan As Double = 0

        Try
            ConnYage = OpenConn()
            GetOlcuPuan = 0

            cSQL = " SELECT DISTINCT Bolum, Olcu=CAST(COALESCE(AVG(ABS(notlar)),'0') as decimal (4,2)) " + _
                   " FROM sipolcu where siparisno='" + SiparisNo + "' AND olcutablosuno='ARAKONTROL-YS' GROUP BY bolum "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                Select Case LCase(SQLReadString(oReader, "Bolum"))
                    Case "bel", "basen", "ıc boy 1", "arka orta boy", "gogus", "etek"
                        If SQLReadDouble(oReader, "Olcu") > 0.5 Then nOlcuPuan = nOlcuPuan + 5
                    Case "kol boyu"
                        If SQLReadDouble(oReader, "Olcu") > 0.5 Then nOlcuPuan = nOlcuPuan + 3
                    Case Else
                        If SQLReadDouble(oReader, "Olcu") > 0.5 Then nOlcuPuan = nOlcuPuan + 1
                End Select
            Loop
            oReader.Close()
            oReader = Nothing

            GetOlcuPuan = nOlcuPuan
            CloseConn(ConnYage)

        Catch Err As Exception
            GetOlcuPuan = 0
            ErrDisp("Error Tedarikçi Değerlendirme (Ölçü Puanı) " + Err.Message)
        End Try

    End Function

End Module

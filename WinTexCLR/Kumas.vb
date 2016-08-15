Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module Kumas

    Private Structure KumasTop
        Dim cTopNo As String
        Dim cMTK As String
        Dim cIsemriNo As String
        Dim cStokNo As String
        Dim cRenk As String
        Dim cDepo As String
        Dim cPartiNo As String
        Dim nNet As Double
        Dim nBrut As Double
        Dim cBirim As String
        Dim nFiyat As Double
        Dim cDoviz As String
        Dim nIFiyat As Double
        Dim cIDoviz As String
        Dim cSakatKodu As String
    End Structure

    Public Function BarkodluKumasCikis(ByVal cStokFisNo As String, ByVal cStokHareketKodu As String, ByVal cDepartman As String, ByVal cFirma As String, ByVal cNotlar As String) As String

        Dim cSQL As String

        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim aTop() As KumasTop
        Dim nCnt As Integer = 0
        Dim dTarih As Date
        Dim nOK As SqlInt32

        BarkodluKumasCikis = "OK"
        cSQL = ""

        Try
            ReDim aTop(0)

            ConnYage = OpenConn()

            dTarih = GetNowFromServer(ConnYage)

            cSQL = "select topno " + _
                    " from barkodlukumas " + _
                    " where stokfisno = '" + cStokFisNo + "' " + _
                    " order by topno "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read()
                ReDim Preserve aTop(nCnt)
                aTop(nCnt).cTopNo = oReader.GetString(oReader.GetOrdinal("topno")).Trim
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aTop)

                cSQL = "select stokno, renk, malzemetakipkodu, depo, partino, " + _
                        " miktar = coalesce(donemgiris1,0) - coalesce(donemcikis1,0) " + _
                        " from stoktoprb " + _
                        " where topno = '" + aTop(nCnt).cTopNo + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    aTop(nCnt).cStokNo = SQLReadString(oReader, "stokno")
                    aTop(nCnt).cRenk = SQLReadString(oReader, "renk")
                    aTop(nCnt).cDepo = SQLReadString(oReader, "depo")
                    aTop(nCnt).cPartiNo = SQLReadString(oReader, "partino")
                    aTop(nCnt).cMTK = SQLReadString(oReader, "malzemetakipkodu")
                    aTop(nCnt).nNet = SQLReadDouble(oReader, "miktar")
                    aTop(nCnt).nBrut = SQLReadDouble(oReader, "miktar")
                End If
                oReader.Close()
                oReader = Nothing

                cSQL = "select birim, birimfiyat, doviz, toplamiscilik, iscilikdoviz, sakatkodu " + _
                        " from topongirislines " + _
                        " where topno = '" + aTop(nCnt).cTopNo + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    aTop(nCnt).cIsemriNo = "" ' SQLReadString(dr, "isemrino")
                    aTop(nCnt).cBirim = SQLReadString(oReader, "birim")
                    aTop(nCnt).nFiyat = SQLReadDouble(oReader, "birimfiyat")
                    aTop(nCnt).cDoviz = SQLReadString(oReader, "doviz")
                    aTop(nCnt).nIFiyat = SQLReadDouble(oReader, "toplamiscilik")
                    aTop(nCnt).cIDoviz = SQLReadString(oReader, "iscilikdoviz")
                    aTop(nCnt).cSakatKodu = SQLReadString(oReader, "sakatkodu")
                End If
                oReader.Close()
                oReader = Nothing
            Next

            ' fiş kafasını yaz

            cSQL = " insert into stokfis (StokFisNo,StokFisTipi,fistarihi,departman,firma,notlar) " + _
                    " values ('" + cStokFisNo + "', " + _
                    " 'Cikis', " + _
                    " '" + SQLWriteDate(dTarih) + "', " + _
                    " '" + cDepartman + "', " + _
                    " '" + cFirma + "', " + _
                    " '" + cNotlar + "') "

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

            ' satirlarini yaz

            For nCnt = 0 To UBound(aTop)
                If aTop(nCnt).cStokNo <> "" And aTop(nCnt).nNet <> 0 Then
                    cSQL = "insert into stokfislines " + _
                            " (stokfisno, stokhareketkodu, malzemetakipkodu, isemrino, stokno, renk, beden, depo, sakatkodu, partino, " + _
                            " netmiktar1, brutmiktar1, birim1, fissirano, topno, iscilikfiyat, iscilikdoviz, birimfiyat, dovizcinsi) " + _
                            " values " + _
                            "('" + cStokFisNo + "', " + _
                            " '" + cStokHareketKodu + "', " + _
                            " '" + aTop(nCnt).cMTK + "', " + _
                            " '" + aTop(nCnt).cIsemriNo + "', " + _
                            " '" + aTop(nCnt).cStokNo + "', " + _
                            " '" + aTop(nCnt).cRenk + "', " + _
                            " 'HEPSI', " + _
                            " '" + aTop(nCnt).cDepo + "', " + _
                            " '" + aTop(nCnt).cSakatKodu + "', " + _
                            " '" + aTop(nCnt).cPartiNo + "', " + _
                            SQLWriteDecimal(aTop(nCnt).nNet) + ", " + _
                            SQLWriteDecimal(aTop(nCnt).nBrut) + ", " + _
                            " '" + aTop(nCnt).cBirim + "', " + _
                            " 0, " + _
                            " '" + aTop(nCnt).cTopNo + "', " + _
                            SQLWriteDecimal(aTop(nCnt).nIFiyat) + ", " + _
                            " '" + aTop(nCnt).cIDoviz + "', " + _
                            SQLWriteDecimal(aTop(nCnt).nFiyat) + ", " + _
                            " '" + aTop(nCnt).cDoviz + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            ' validate
            nOK = SingleStokFisValidateConnected(ConnYage, "validate", cStokFisNo, "", "", "")
            If nOK = 0 Then
                ErrDisp("Stok fis validate err : " + cStokFisNo)
                BarkodluKumasCikis = "Hata"
                Exit Function
            End If

            Call CloseConn(ConnYage)
            BarkodluKumasCikis = "OK"

        Catch Err As Exception
            BarkodluKumasCikis = "Hata"
            ErrDisp("BarkodluKumasCikis : " + Err.Message.Trim + vbCrLf + cSQL)
        End Try

    End Function
End Module

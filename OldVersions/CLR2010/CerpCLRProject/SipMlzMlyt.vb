Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module SipMlzMlyt

    Public Function SipMlzMlytList(ByVal cSiparisNo As String) As String

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cReceteNo As String = ""
        Dim cModelNo As String = ""
        Dim cView As String

        SipMlzMlytList = ""

        Try
            ConnYage = OpenConn()

            SipMlzMlytList = ""
            cSQL = "select receteno, modelno from sipmodel where siparisno = '" + cSiparisNo + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                cReceteNo = SQLReadString(oReader, "receteno")
                cModelNo = SQLReadString(oReader, "modelno")
            End If
            oReader.Close()
            oReader = Nothing

            cView = CreateTempView(ConnYage)

            cSQL = "create view " + cView + " as " + _
                    " select " + _
                    " Tablo = 'MTF', " + _
                    " a.departman, a.stokno, a.renk, " + _
                    " satis = 0, " + _
                    " Miktar = (sum(coalesce(a.uretimicincikis,0))-sum(coalesce(a.uretimdeniade,0))), " + _
                    " a.birim, " + _
                    " birimfiyat = (select top 1 fiyat2 from stokfiyat where stokno = a.stokno  and (renk = a.renk or renk = 'HEPSI') order by tarih desc)," + _
                    " dovizcinsi = (select top 1 doviz2 from stokfiyat where stokno = a.stokno  and (renk = a.renk or renk = 'HEPSI') order by tarih desc) " + _
                    " from mtkfislines a " + _
                    " where a.malzemetakipno = '" + cSiparisNo + "' " + _
                    " and a.stokno not in(SELECT x1.stokno " + _
                                        " from stokfis x, stokfislines x1 " + _
                                        " where x.stokfisno = x1.stokfisno " + _
                                        " and x1.malzemetakipkodu = a.malzemetakipno " + _
                                        " and x1.stokhareketkodu in ('01 Uretime Cikis','07 Satis')) " + _
                    " group by a.malzemetakipno,a.departman,a.stokno,a.renk,a.birim,a.departman"

            cSQL = cSQL + " union all  " + _
                    " SELECT Tablo = 'SHF', " + _
                    " Departman = coalesce((select top 1 uretimdepartmani " + _
                                         " from " + IIf(cReceteNo = "", "modelhammadde where modelno = '" + cModelNo + "' ", _
                                         " modelhammadde2 where modelno = '" + cModelNo + "' " + _
                                         " and receteno = '" + cReceteNo + "'").ToString + " " + _
                                         " and hammaddekodu = b.stokno and (hammadderenk = b.renk or hammadderenk = 'HEPSI')),a.departman), " + _
                    " b.stokno, b.renk , " + _
                    " satis = 0, " + _
                    " Miktar = sum(coalesce(b.netmiktar1,0)) - " + _
                            " (SELECT coalesce(sum(coalesce(b1.netmiktar1,0)),0)" + _
                            " from stokfis a1, stokfislines b1 " + _
                            " where a1.stokfisno = b1.stokfisno " + _
                            " and b1.malzemetakipkodu = b.malzemetakipkodu " + _
                            " and b1.stokhareketkodu = '01 Uretimden iade' " + _
                            " and b1.stokno = b.stokno " + _
                            " and b1.renk = b.renk ) + " + _
                            " (SELECT coalesce(sum(coalesce(b1.netmiktar1,0)),0) " + _
                            " from stokfis a1, stokfislines b1 " + _
                            " where a1.stokfisno = b1.stokfisno " + _
                            " and b1.malzemetakipkodu = b.malzemetakipkodu " + _
                            " and a1.departman = 'KUMAS LABTEST ' " + _
                            " and b1.stokhareketkodu = '05 Diger Cikis' " + _
                            " and b1.stokno = b.stokno " + _
                            " and b1.renk = b.renk) , " + _
                    " birim1, birimfiyat=max(coalesce(b.birimfiyat,0)),  b.dovizcinsi " + _
                    " from stokfis a, stokfislines b " + _
                    " where a.stokfisno = b.stokfisno " + _
                    " and b.malzemetakipkodu = '" + cSiparisNo + "'  " + _
                    " and b.stokhareketkodu = '01 Uretime Cikis' " + _
                    " group by b.malzemetakipkodu,b.stokno,b.renk,b.birim1,b.dovizcinsi,a.departman  "

            cSQL = cSQL + " union all  " + _
                    " SELECT Tablo = 'SHF', " + _
                    " Departman = coalesce((select top 1 uretimdepartmani " + _
                                         " from " + IIf(cReceteNo = "", "modelhammadde where modelno = '" + cModelNo + "' ", _
                                         " modelhammadde2 where modelno = '" + cModelNo + "' " + _
                                         " and receteno = '" + cReceteNo + "'").ToString + " " + _
                                         " and hammaddekodu = b.stokno and (hammadderenk = b.renk or hammadderenk = 'HEPSI')),a.departman), " + _
                    " b.stokno, b.renk , " + _
                    " satis = sum(coalesce(b.netmiktar1,0)), " + _
                    " Miktar = 0, " + _
                    " birim1, birimfiyat=max(coalesce(b.birimfiyat,0)),  b.dovizcinsi " + _
                    " from stokfis a, stokfislines b " + _
                    " where a.stokfisno = b.stokfisno " + _
                    " and b.malzemetakipkodu = '" + cSiparisNo + "'  " + _
                    " and (b.stokno like '0%' or b.stokno like '1%' or b.stokno like '2%')" + _
                    " and b.stokhareketkodu = '07 Satis' " + _
                    " group by b.malzemetakipkodu,a.departman,b.stokno,b.renk,b.birim1,b.dovizcinsi,a.departman  "

            cSQL = cSQL + " union all  " + _
                    " select Tablo = 'SHF', " + _
                    " Departman = coalesce((select top 1 uretimdepartmani " + _
                                         " from " + IIf(cReceteNo = "", "modelhammadde where modelno = '" + cModelNo + "' ", _
                                         " modelhammadde2 where modelno = '" + cModelNo + "' " + _
                                         " and receteno = '" + cReceteNo + "'").ToString + " " + _
                                         " and hammaddekodu = b.stokno and (hammadderenk = b.renk or hammadderenk = 'HEPSI')),a.departman), " + _
                    " b.stokno, b.renk , " + _
                    " satis = 0 - sum(coalesce(b.netmiktar1,0)), " + _
                    " Miktar = 0, " + _
                    " birim1, birimfiyat=max(coalesce(b.birimfiyat,0)),  b.dovizcinsi " + _
                    " from stokfis a, stokfislines b " + _
                    " where a.stokfisno = b.stokfisno " + _
                    " and b.malzemetakipkodu = '" + cSiparisNo + "'  " + _
                    " and b.stokhareketkodu = '07 Satis Iade' " + _
                    " and (b.stokno like '0%' or b.stokno like '1%' or b.stokno like '2%')" + _
                    " group by b.malzemetakipkodu,a.departman,b.stokno,b.renk,b.birim1,b.dovizcinsi,a.departman  "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = " select Tablo, a.stokno, a.renk, a.Departman, " + _
                   " stoktipi = (select top 1 b.stoktipi from stok b where b.stokno = a.stokno), " + _
                   " Brihtiyac = (select top 1 coalesce(kullanimmiktari,0) " + _
                                 " from " + IIf(cReceteNo = "", "modelhammadde  where modelno = '" + cModelNo + "' ", _
                                            " modelhammadde2 where modelno = '" + cModelNo + "' " + _
                                            " and receteno = '" + cReceteNo + "'").ToString + " " + _
                                            " and hammaddekodu = stokno and (hammadderenk = a.renk or hammadderenk = 'HEPSI')), " + _
                   " miktar = sum(coalesce(a.miktar,0)), " + _
                   " satis = sum(coalesce(a.satis,0)), " + _
                   " a.birim, a.birimfiyat, a.dovizcinsi  " + _
                   " from " + cView + " a " + _
                   " group by a.Tablo, a.stokno, a.renk, a.birim, a.birimfiyat, a.dovizcinsi, a.departman " + _
                   " order by a.stokno, a.renk, a.birim, a.birimfiyat, a.dovizcinsi, a.departman "

            oReader = GetSQLReader(cSQL, ConnYage)

            SqlContext.Pipe.Send(oReader)
            oReader.Close()
            oReader = Nothing

            DropView(cView, ConnYage)

            CloseConn(ConnYage)

        Catch Err As Exception
            SipMlzMlytList = "Hata"
            ErrDisp("SipMlzMlytList : " + Err.Message.Trim + vbCrLf + cSQL)
        End Try
    End Function
End Module

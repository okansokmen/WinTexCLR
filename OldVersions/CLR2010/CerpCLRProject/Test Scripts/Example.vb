Imports System
Imports System.Data
Imports Microsoft.SqlServer.Server
Imports System.Data.SqlTypes
Imports System.Runtime.InteropServices
Imports System.Data.SqlClient

Partial Public Class StoredProcedures

    <SqlProcedure()> _
    Public Shared Sub okan(ByVal cSiparisGrubu As String)

        Dim cSQL As String
        Dim dr As SqlDataReader
        Dim cm As SqlCommand
        Dim cBuffer As String
        Dim ConnYage As SqlConnection

        ConnYage = OpenConn()

        cSQL = "select musterino from siparis where siparisgrubu = '" + cSiparisGrubu + "' "

        cm = New SqlCommand

        cm.Connection = ConnYage
        cm.CommandText = cSQL
        dr = cm.ExecuteReader(CommandBehavior.SingleResult)
        Do While dr.Read()
            cBuffer = dr.GetString(dr.GetOrdinal("musterino")).Trim
            SqlContext.Pipe.Send("sonuc-1 : " & cBuffer & Environment.NewLine)
        Loop
        dr.Close()
        dr = Nothing
        cm = Nothing

        SqlContext.Pipe.Send("Parameter : " & cSiparisGrubu & Environment.NewLine)

        CloseConn(ConnYage)

    End Sub
End Class

Public Class demo2
    Public Shared Sub VBproc1()
        Using conn As New SqlConnection("context connection=true")
            conn.Open()
            Dim cmd As SqlCommand = New SqlCommand("SELECT col1 FROM dbo.table1", conn)
            SqlContext.Pipe.ExecuteAndSend(cmd)
            conn.Close()
        End Using
    End Sub
End Class


Public Class HelloWorldProc
    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub AA1(<Out()> ByRef text As String)

        Dim dr As SqlDataReader
        Dim cm As SqlCommand
        Dim cBuffer As String
        Dim cSQL As String
        Dim ConnYage As SqlClient.SqlConnection

        Try
            ConnYage = OpenConn()

            cSQL = "select musterino from siparis"

            cm = New SqlCommand
            cm.Connection = ConnYage
            cm.CommandText = cSQL
            dr = cm.ExecuteReader(CommandBehavior.SingleResult)
            Do While dr.Read()
                cBuffer = dr.GetString(dr.GetOrdinal("musterino")).Trim
                SqlContext.Pipe.Send("sonuc-1 : " & cBuffer & Environment.NewLine)
                text = "sonuc-2 : " & cBuffer
            Loop
            dr.Close()
            dr = Nothing
            cm = Nothing

            Call CloseConn(ConnYage)

        Catch ex As Exception

        End Try
    End Sub
End Class

Public Class HelloWorldProc2
    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub AA2(<Out()> ByRef text As String)
        SqlContext.Pipe.Send("Hello okan-x2!" & Environment.NewLine)
        text = "Hello okan-x22!"
    End Sub
End Class

Partial Public Class StoredProcedures

    <SqlProcedure()> _
    Public Shared Sub StokrbTarih(ByVal nKumasAks As Integer, Optional ByVal cTarih As String = "", Optional ByVal lSayim As Boolean = False, Optional ByVal topno As String = "", _
                                Optional ByVal cSayimDepo As String = "", Optional ByVal nBakim As Integer = 0, Optional ByVal cStokno As String = "", _
                                Optional ByVal nWANoMTK As Integer = 1, Optional ByVal nWFNoMTK As Integer = 1)

        ' Barkod numarasi bazinda giris-cikis hareketlerinden tarih itibariyla kalan miktari hesapliyor
        ' topno deneme filtresi olarak kullanılıyor
        ' nKumasAks = 0 barkodsuz sistemdir, nKumasAks = 1 barkodlu kumas , nKumasAks = 2 barkodlu aksesuar, nKumasAks = 3 hem kumas hem aksesuar

        Dim cTopFilter As String
        Dim cOrgTopFilter As String
        Dim cNewTopFilter As String
        Dim cDepoFilter As String
        Dim cHDepoFilter As String
        Dim cKDepoFilter As String
        Dim cBDepoFilter As String
        Dim cSQL As String
        Dim cBarkodNull As String
        Dim cBarkodBolmeDosyasi As String
        Dim cBarkodData As String
        Dim cTarihFilter1 As String
        Dim cTarihFilter2 As String
        Dim cTarihFilter3 As String
        Dim cStokNoFilter As String
        Dim Tarih As Date
        Dim lBakim As Boolean
        Dim G_WANoMTK As Boolean
        Dim G_WFNoMTK As Boolean

        G_WANoMTK = False
        If nWANoMTK = 1 Then G_WANoMTK = True

        G_WFNoMTK = False
        If nWFNoMTK = 1 Then G_WFNoMTK = True

        lBakim = False
        If nBakim = 1 Then lBakim = True

        Tarih = Today
        If cTarih <> "" Then Tarih = CDate(cTarih)

        cStokNoFilter = ""
        If cStokno <> "" Then cStokNoFilter = " and stokno = '" + cStokno + "' "

        cTarihFilter1 = ""
        cTarihFilter2 = ""
        cTarihFilter3 = ""
        If Not lBakim Then
            cTarihFilter1 = " and fistarihi <= '" + CStr(Tarih) + "' "
            cTarihFilter2 = " and tarih <= '" + CStr(Tarih) + "' "
            cTarihFilter3 = " and tarih <= '" + CStr(Tarih) + "' "
        End If

        cDepoFilter = ""
        cHDepoFilter = ""
        cKDepoFilter = ""
        cBDepoFilter = ""
        If Trim(cSayimDepo) <> "" Then
            cDepoFilter = " and b.depo='" & cSayimDepo & "' "
            cHDepoFilter = " and hedefdepo='" & cSayimDepo & "' "
            cKDepoFilter = " and kaynakdepo='" & cSayimDepo & "' "
            cBDepoFilter = " and a.depo='" + cSayimDepo + "' "
        End If

        cBarkodNull = ""
        If nKumasAks = 1 Or nKumasAks = 2 Or nKumasAks = 3 Then
            cBarkodNull = " and topno is not null and topno <> '' "
        End If

        cTopFilter = ""
        cOrgTopFilter = ""
        cNewTopFilter = ""
        If Trim(topno) <> "" Then
            cTopFilter = " and topno = '" + Trim(topno) + "'"
            cOrgTopFilter = " and orgtopno = '" + Trim(topno) + "'"
            cNewTopFilter = " and newtopno = '" + Trim(topno) + "'"
        End If

        If nKumasAks = 1 Then
            cBarkodBolmeDosyasi = "TopBolmeNew"
            cBarkodData = "stoktoprb"
        End If

        If nKumasAks = 2 Then
            cBarkodBolmeDosyasi = "AksesuarBolmeNew"
            cBarkodData = "stokaksesuarrb"
        End If

        ExecuteSQLCommandConnected("set dateformat 'dmy'")

        ' stok fislerinden hesapla

        ExecuteSQLCommandConnected("if exists (select * from sysobjects where id = object_id('dbo.stokfis_yedek')) drop view stokfis_yedek")

        cSQL = "create view stokfis_yedek as " + _
            " SELECT stokno = coalesce(stokno,''), " + _
                " renk = coalesce(renk,''), " + _
                " beden = coalesce(beden,''), " + _
                " mtkno = coalesce(malzemetakipkodu,''), " + _
                " partino = coalesce(partino,''), " + _
                " depo = coalesce(depo,''), " + _
                " topno = coalesce(topno, ''), " + _
                " giris = sum(coalesce(netmiktar1,0)), " + _
                " cikis = 0 , " + _
                " gir_agirlik = sum(coalesce(agirlik,0)), " + _
                " cik_agirlik = 0 " + _
            " FROM stokfis a, stokfislines b " + _
            " WHERE a.stokfisno = b.stokfisno " + _
            " and (stokfistipi = 'Giris') " + _
            cTarihFilter1 + cTopFilter + cBarkodNull + cDepoFilter + cStokNoFilter + _
            " GROUP BY stokno, renk, beden, depo, partino, malzemetakipkodu, topno "

        cSQL = cSQL + _
            " Union All " + _
            " SELECT stokno = coalesce(stokno,''), " + _
                " renk = coalesce(renk,''), " + _
                " beden = coalesce(beden,''), " + _
                " mtkno = coalesce(malzemetakipkodu,''), " + _
                " partino = coalesce(partino,''), " + _
                " depo = coalesce(depo,''), " + _
                " topno = coalesce(topno, '')," + _
                " giris = 0, " + _
                " cikis = sum(coalesce(netmiktar1,0)), " + _
                " gir_agirlik = 0, " + _
                " cik_agirlik = sum(coalesce(agirlik,0))" + _
            " FROM stokfis a, stokfislines b " + _
            " WHERE a.stokfisno = b.stokfisno " + _
            " and (stokfistipi = 'Cikis') " + _
            cTarihFilter1 + cTopFilter + cBarkodNull + cDepoFilter + cStokNoFilter + _
            " GROUP BY stokno, renk, beden, depo, partino, malzemetakipkodu, topno "

        ExecuteSQLCommandConnected(cSQL)

        ' rezervasyonlardan hesapla

        ExecuteSQLCommandConnected("if exists (select * from sysobjects where id = object_id('dbo.stokrez_yedek')) drop view stokrez_yedek")

        cSQL = "create view stokrez_yedek as " + _
            " SELECT stokno = coalesce(stokno,''), " + _
                " renk = coalesce(renk,''), " + _
                " beden = coalesce(beden,''), " + _
                " mtkno = coalesce(hedefmalzemetakipno,''), " + _
                " partino = coalesce(hedefpartino,''), " + _
                " depo = coalesce(hedefdepo,''), " + _
                " topno = coalesce(topno, ''), " + _
                " giris = sum(coalesce(netmiktar1,0)), " + _
                " cikis = 0, " + _
                " gir_agirlik = 0, " + _
                " cik_agirlik = 0 " + _
            " FROM StokTransfer " + _
            " WHERE rtrim(stokno) <> '' " + _
            cTarihFilter2 + _
            cTopFilter + _
            cBarkodNull + _
            cHDepoFilter + _
            cStokNoFilter + _
            " GROUP BY stokno, renk, beden, hedefpartino, hedefdepo, hedefmalzemetakipno, topno "

        cSQL = cSQL + _
            " Union All " + _
            " SELECT stokno = coalesce(stokno,''), " + _
                " renk = coalesce(renk,''), " + _
                " beden = coalesce(beden,''), " + _
                " mtkno = coalesce(kaynakmalzemetakipno,''), " + _
                " partino = coalesce(kaynakpartino,''), " + _
                " depo = coalesce(kaynakdepo,''), " + _
                " topno = coalesce(topno, ''), " + _
                " giris = 0, " + _
                " cikis = sum(coalesce(netmiktar1,0)), " + _
                " gir_agirlik = 0, " + _
                " cik_agirlik = 0 " + _
            " FROM StokTransfer " + _
            " WHERE rtrim(stokno) <> '' " + _
            cTarihFilter2 + _
            cTopFilter + _
            cBarkodNull + _
            cKDepoFilter + _
            cStokNoFilter + _
            " GROUP BY stokno, renk, beden, kaynakpartino, kaynakdepo, kaynakmalzemetakipno, topno"

        ExecuteSQLCommandConnected(cSQL)

        ExecuteSQLCommandConnected("if exists (select * from sysobjects where id = object_id('dbo.stokrbeskidurum')) drop view stokrbeskidurum")

        If lBakim Then
            ' stok ve rezervasyon fislerini al
            cSQL = "create view stokrbeskidurum as " + _
                           "SELECT stokno, renk, beden, mtkno, partino, depo, topno = coalesce(topno, ''), " + _
                           "giris = sum(coalesce(giris,0)), cikis=sum(coalesce(cikis,0)) " + _
                           "FROM stokfis_yedek " + _
                           "GROUP BY stokno, renk, beden, partino, depo, mtkno, topno " + _
                           "Union All " + _
                           "SELECT stokno, renk, beden, mtkno, partino, depo, topno = coalesce(topno, ''), " + _
                           "giris = sum(coalesce(giris,0)), cikis=sum(coalesce(cikis,0)) " + _
                           "FROM stokrez_yedek " + _
                           "GROUP BY stokno, renk, beden, partino, depo, mtkno, topno "
        Else
            ' sayim buradan bakiyor
            ' stok ve rezervasyon fislerini al
            cSQL = "create view stokrbeskidurum as " + _
                           "SELECT stokno, renk, beden, mtkno, partino, depo, topno , " + _
                           "rbaadet = sum(coalesce(giris,0) - coalesce(cikis,0)), " + _
                           "rbaagirlik = sum(coalesce(gir_agirlik,0)-coalesce(cik_agirlik,0)) " + _
                           "FROM stokfis_yedek " + _
                           "GROUP BY stokno, renk, beden, partino, depo, mtkno, topno " + _
                           "Union All " + _
                           "SELECT stokno, renk, beden, mtkno, partino, depo, topno, " + _
                           "rbaadet = sum(coalesce(giris,0) - coalesce(cikis,0)), " + _
                           "rbaagirlik = sum(coalesce(gir_agirlik,0)-coalesce(cik_agirlik,0))  " + _
                           "FROM stokrez_yedek " + _
                           "GROUP BY stokno, renk, beden, partino, depo, mtkno, topno "
        End If

        ExecuteSQLCommandConnected(cSQL)

        If lBakim Then
            ' stokrb toplamlari isini yap
            If nKumasAks = 0 Then
                ExecuteSQLCommandConnected("delete from stokrb")
                ExecuteSQLCommandConnected("insert stokrb (stokno,renk,beden,depo,partino,malzemetakipkodu,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                " select  stokno,renk,beden,depo,partino,mtkno, sum(coalesce(giris,0)), sum(coalesce(cikis,0)), 0,0,0,0 " + _
                 " From stokrbeskidurum " + _
                 " group by stokno,renk,beden,depo,partino,mtkno ")
            End If
            ' StokTopRB (Kumas)
            If nKumasAks = 1 Then
                ExecuteSQLCommandConnected("if exists (select * from sysobjects where id = object_id('dbo.StokTopRB_yedek') and type in('U')) drop table dbo.StokTopRB_yedek")
                ExecuteSQLCommandConnected("CREATE TABLE StokTopRB_yedek (topno char(30),stokno char(30),renk char(30),beden char(30),partino char(30), " + _
                                  "malzemetakipkodu char(30),depo char(30),songiristarihi datetime,songirisfiyati decimal(18,6),songirisdovizi char(3), " + _
                                  "songirisdovizfiyati decimal(18,6),songirisdept char(30),songirisfirmasi char(30),donemgiris1 decimal(18,3),donemcikis1 decimal(18,3), " + _
                                  "donemgiris2 decimal(18,3),donemcikis2 decimal(18,3),donemgiris3 decimal(18,3),donemcikis3 decimal(18,3), " + _
                                  "devirgiris1 decimal(18,3),devircikis1 decimal(18,3),devircikis2 decimal(18,3),devirgiris2 decimal(18,3), " + _
                                  "devircikis3 decimal(18,3),devirgiris3 decimal(18,3),alismiktari1 decimal(18,3),alistutari1 decimal(18,3), " + _
                                  "magazakodu char(30),agirlik decimal(9,3),bolundu char(1),topsirano decimal(18,3), " + _
                                  "encekme decimal(15,3),boycekme decimal(15,3),pacatest decimal(15,3),grm2 decimal(15,3),netkullanimeni decimal(15,3),tedarikcitopno char(30) ) ")
                ' sadece kumas barkodlarini getirsin
                If G_WFNoMTK Then
                    ExecuteSQLCommandConnected("insert StokTopRB_yedek (topno,stokno,renk,beden,depo,partino,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,sum(coalesce(giris,0)), sum(coalesce(cikis,0)), 0,0,0,0 " + _
                     " From stokrbeskidurum " + _
                     " where exists (select stokno from stok where stokrbeskidurum.stokno = stok.stokno and toptakibi='E') " + _
                     " group by topno,stokno,renk,beden,depo,partino ")
                Else
                    ExecuteSQLCommandConnected("insert StokTopRB_yedek (topno,stokno,renk,beden,depo,partino,malzemetakipkodu,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,mtkno , sum(coalesce(giris,0)), sum(coalesce(cikis,0)), 0,0,0,0 " + _
                     " From stokrbeskidurum " + _
                     " where exists (select stokno from stok where stokrbeskidurum.stokno = stok.stokno and toptakibi='E') " + _
                     " group by topno,stokno,renk,beden,depo,partino,mtkno ")
                End If
                ' StokTopRB yaratiliyor
                ExecuteSQLCommandConnected("delete from StokTopRB")
                If G_WFNoMTK Then
                    ExecuteSQLCommandConnected("insert StokTopRB (topno,stokno,renk,beden,depo,partino,malzemetakipkodu,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,malzemetakipkodu='' , sum(coalesce(donemgiris1,0)), sum(coalesce(donemcikis1,0)), 0,0,0,0 " + _
                     " From StokTopRB_yedek " + _
                     " group by topno,stokno,renk,beden,depo,partino,malzemetakipkodu ")
                Else
                    ExecuteSQLCommandConnected("insert StokTopRB (topno,stokno,renk,beden,depo,partino,malzemetakipkodu,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,malzemetakipkodu , sum(coalesce(donemgiris1,0)), sum(coalesce(donemcikis1,0)), 0,0,0,0 " + _
                     " From StokTopRB_yedek " + _
                     " group by topno,stokno,renk,beden,depo,partino,malzemetakipkodu ")
                End If
            End If
            ' StokAksesuarRB (Aksesuar)
            If nKumasAks = 2 Then
                ExecuteSQLCommandConnected("if exists (select * from sysobjects where id = object_id('dbo.StokAksesuarRB_yedek') and type in('U')) drop table dbo.StokAksesuarRB_yedek")
                ExecuteSQLCommandConnected("CREATE TABLE StokAksesuarRB_yedek (topno char(30),stokno char(30),renk char(30),beden char(30),partino char(30), " + _
                                  "malzemetakipkodu char(30),depo char(30),songiristarihi datetime,songirisfiyati decimal(18,6),songirisdovizi char(3), " + _
                                  "songirisdovizfiyati decimal(18,6),songirisdept char(30),songirisfirmasi char(30),donemgiris1 decimal(18,3),donemcikis1 decimal(18,3), " + _
                                  "donemgiris2 decimal(18,3),donemcikis2 decimal(18,3),donemgiris3 decimal(18,3),donemcikis3 decimal(18,3), " + _
                                  "devirgiris1 decimal(18,3),devircikis1 decimal(18,3),devircikis2 decimal(18,3),devirgiris2 decimal(18,3), " + _
                                  "devircikis3 decimal(18,3),devirgiris3 decimal(18,3),alismiktari1 decimal(18,3),alistutari1 decimal(18,3), " + _
                                  "magazakodu char(30),agirlik decimal(9,3),bolundu char(1),topsirano decimal(18,3), " + _
                                  "encekme decimal(15,3),boycekme decimal(15,3),pacatest decimal(15,3),grm2 decimal(15,3),netkullanimeni decimal(15,3),tedarikcitopno char(30) ) ")
                ' sadece aksesuar barkodlarini getirsin
                If G_WANoMTK Then
                    ExecuteSQLCommandConnected("insert StokAksesuarRB_yedek (topno,stokno,renk,beden,depo,partino,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,sum(coalesce(giris,0)), sum(coalesce(cikis,0)), 0,0,0,0 " + _
                     " From stokrbeskidurum " + _
                     " where exists (select stokno from stok where stokrbeskidurum.stokno = stok.stokno and aksesuartakibi='E') " + _
                     " group by topno,stokno,renk,beden,depo,partino ")
                Else
                    ExecuteSQLCommandConnected("insert StokAksesuarRB_yedek (topno,stokno,renk,beden,depo,partino,malzemetakipkodu,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,mtkno , sum(coalesce(giris,0)), sum(coalesce(cikis,0)), 0,0,0,0 " + _
                     " From stokrbeskidurum " + _
                     " where exists (select stokno from stok where stokrbeskidurum.stokno = stok.stokno and aksesuartakibi='E') " + _
                     " group by topno,stokno,renk,beden,depo,partino,mtkno ")
                End If
                ' StokAksesuarRB yaratiliyor
                ExecuteSQLCommandConnected("delete from StokAksesuarRB")
                If G_WANoMTK Then
                    ExecuteSQLCommandConnected("insert StokAksesuarRB (topno,stokno,renk,beden,depo,partino,malzemetakipkodu,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,malzemetakipkodu='', sum(coalesce(donemgiris1,0)), sum(coalesce(donemcikis1,0)), 0,0,0,0 " + _
                     " From StokAksesuarRB_yedek " + _
                     " group by topno,stokno,renk,beden,depo,partino,malzemetakipkodu ")
                Else
                    ExecuteSQLCommandConnected("insert StokAksesuarRB (topno,stokno,renk,beden,depo,partino,malzemetakipkodu,donemgiris1,donemcikis1,devirgiris1,devircikis1,alismiktari1,alistutari1)  " + _
                    " select  topno,stokno,renk,beden,depo,partino,malzemetakipkodu , sum(coalesce(donemgiris1,0)), sum(coalesce(donemcikis1,0)), 0,0,0,0 " + _
                     " From StokAksesuarRB_yedek " + _
                     " group by topno,stokno,renk,beden,depo,partino,malzemetakipkodu ")
                End If
            End If
        Else
            ' sayim isleri ara dosyalari
            If lSayim Then
                ' sayim buraya bakiyor
                ' butun stoklarin eski durumlarini veriyor

                ExecuteSQLCommandConnected("if exists (select * from sysobjects where id = object_id('dbo.winstorestokrba') and sysstat & 0xf = 2) drop view winstorestokrba")

                ExecuteSQLCommandConnected("create view winstorestokrba as " + _
                                  " SELECT stokno, cinsaciklama = (select cinsaciklamasi from stok where stokno = st.stokno), " + _
                                  " depo, renk, beden, partino, mtkno, topno = coalesce(topno, ''), sayimadet = 0, " + _
                                  " rbaadet = sum(coalesce(rbaadet,0)), rbaagirlik = sum(coalesce(rbaagirlik,0)) " + _
                                  " FROM stokrbeskidurum st " + _
                                  " GROUP BY stokno, renk, beden, depo, partino, mtkno, topno ")
            Else
                ' sadece sayimda olan stoklarin eski durumlarini veriyor

                ExecuteSQLCommandConnected("if exists (select * from sysobjects where id = object_id('dbo.wintexstokrba') and sysstat & 0xf = 2) drop view wintexstokrba")

                ExecuteSQLCommandConnected("create view wintexstokrba as " + _
                                  " SELECT a.stokno, a.depo, a.renk, a.beden, a.partino, a.mtkno, a.topno, " + _
                                  " sayimadet = 0, rbaadet = sum(coalesce(a.rbaadet,0)), topagirlik = 0 " + _
                                  " FROM stokrbeskidurum a, wintexsayimrba b " + _
                                  " WHERE a.stokno = b.stokno " + _
                                  " and a.topno = b.topno " + _
                                  " and a.renk = b.renk " + _
                                  " and a.beden = b.beden " + _
                                  " and a.partino = b.partino " + _
                                  " and a.mtkno = b.mtkno " + _
                                  " GROUP BY a.stokno, a.renk, a.beden, a.depo, a.partino, a.mtkno, a.topno")
            End If
        End If
    End Sub
End Class
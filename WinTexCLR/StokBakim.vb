Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module StokBakim

    Public Function GetViewStokDurumu(ByVal nCase As Integer, Optional ByVal cFilterFis As String = "", Optional ByVal cFilterTransfer1 As String = "", _
                                      Optional ByVal cFilterTransfer2 As String = "", Optional ByVal dTarih As Date = #1/1/1950#) As String
        ' nCase = 1 stokRB
        ' nCase = 2 stokTopRB
        ' nCase = 3 stokAksesuarRB
        Dim cSQL As String = ""
        Dim cStkHarMTFFilter As String = ""
        Dim cTranMTFFilter As String = ""
        Dim cStokFisTarihFilter As String = ""
        Dim cTransferTarihFilter As String = ""
        Dim oSysFlags As New General.SysFlags
        Dim ConnYage As SqlConnection

        GetViewStokDurumu = ""

        Try
            ConnYage = OpenConn()

            ReadSysFlags(oSysFlags, ConnYage)

            If dTarih <> CDate("01.01.1950") Then
                cStokFisTarihFilter = " and a.fistarihi <= '" + SQLWriteDate(dTarih) + "' "
                cTransferTarihFilter = " and a.tarih <= '" + SQLWriteDate(dTarih) + "' "
            End If

            Select Case nCase
                Case 1
                    cStkHarMTFFilter = " malzemetakipkodu = coalesce(b.malzemetakipkodu,''), "
                Case 2
                    If oSysFlags.G_WFNoMTK Then
                        cStkHarMTFFilter = " malzemetakipkodu = '', "
                    Else
                        cStkHarMTFFilter = " malzemetakipkodu = coalesce(b.malzemetakipkodu,''), "
                    End If
                Case 3
                    If oSysFlags.G_WANoMTK Then
                        cStkHarMTFFilter = " malzemetakipkodu = '', "
                    Else
                        cStkHarMTFFilter = " malzemetakipkodu = coalesce(b.malzemetakipkodu,''), "
                    End If
            End Select

            cSQL = "select stokno = coalesce(b.stokno,''), " + _
                    IIf(nCase = 1, "", " topno = coalesce(b.topno,''), ").ToString + _
                    " renk = coalesce(b.renk,''), " + _
                    " beden = coalesce(b.beden,''), " + _
                    cStkHarMTFFilter + _
                    " partino = coalesce(b.partino,''), " + _
                    " depo = coalesce(b.depo,''), " + _
                    " c.anastokgrubu, " + _
                    " c.stoktipi, " + _
                    " c.cinsaciklamasi, " + _
                    " giris = sum(coalesce(b.netmiktar1,0)), " + _
                    " cikis = 0 " + _
                    " from stokfis a , stokfislines b, stok c " + _
                    " where a.stokfisno = b.stokfisno " + _
                    " and a.stokfistipi in ('Giris','02 Satis Iade','03 Defolu iade') " + _
                    " and (a.iptal <> 'E' or a.iptal is null) " + _
                    " and b.stokno = c.stokno " + _
                    " and (c.kapandi is null or c.kapandi = 'H') " + _
                    cFilterFis + cStokFisTarihFilter

            Select Case nCase
                Case 1
                    cSQL = cSQL + _
                        " group by b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
                Case 2
                    cSQL = cSQL + _
                        " and c.toptakibi = 'E' " + _
                        " and b.topno is not null " + _
                        " and b.topno <> '' " + _
                        " group by b.topno, b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
                Case 3
                    cSQL = cSQL + _
                        " and c.aksesuartakibi = 'E' " + _
                        " and b.topno is not null " + _
                        " and b.topno <> '' " + _
                        " group by b.topno, b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
                Case 4
                    cSQL = cSQL + _
                        " and (c.toptakibi = 'E' or c.aksesuartakibi = 'E') " + _
                        " and b.topno is not null " + _
                        " and b.topno <> '' " + _
                        " group by b.topno, b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
            End Select

            cSQL = cSQL + _
                    " select stokno = coalesce(b.stokno,''), " + _
                    IIf(nCase = 1, "", " topno = coalesce(b.topno,''), ").ToString + _
                    " renk = coalesce(b.renk,''), " + _
                    " beden = coalesce(b.beden,''), " + _
                    cStkHarMTFFilter + _
                    " partino = coalesce(b.partino,''), " + _
                    " depo = coalesce(b.depo,''), " + _
                    " c.anastokgrubu, " + _
                    " c.stoktipi, " + _
                    " c.cinsaciklamasi, " + _
                    " giris = 0, " + _
                    " cikis = sum(coalesce(b.netmiktar1,0)) " + _
                    " from stokfis a , stokfislines b, stok c " + _
                    " where a.stokfisno = b.stokfisno  " + _
                    " and a.stokfistipi in ('Cikis','01 Satis') " + _
                    " and (a.iptal <> 'E' or a.iptal is null) " + _
                    " and b.stokno = c.stokno " + _
                    " and (c.kapandi is null or c.kapandi = 'H') " + _
                    cFilterFis + cStokFisTarihFilter

            Select Case nCase
                Case 1
                    cSQL = cSQL + _
                        " group by b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
                Case 2
                    cSQL = cSQL + _
                        " and c.toptakibi = 'E' " + _
                        " and b.topno is not null " + _
                        " and b.topno <> '' " + _
                        " group by b.topno, b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
                Case 3
                    cSQL = cSQL + _
                        " and c.aksesuartakibi = 'E' " + _
                        " and b.topno is not null " + _
                        " and b.topno <> '' " + _
                        " group by b.topno, b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
                Case 4
                    cSQL = cSQL + _
                        " and (c.toptakibi = 'E' or c.aksesuartakibi = 'E') " + _
                        " and b.topno is not null " + _
                        " and b.topno <> '' " + _
                        " group by b.topno, b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.partino, b.depo, c.anastokgrubu, c.stoktipi, c.cinsaciklamasi " + _
                        " Union All "
            End Select

            Select Case nCase
                Case 1
                    cTranMTFFilter = " malzemetakipkodu = coalesce(a.hedefmalzemetakipno ,''), "
                Case 2
                    If oSysFlags.G_WFNoMTK Then
                        cTranMTFFilter = " malzemetakipkodu = '', "
                    Else
                        cTranMTFFilter = " malzemetakipkodu = coalesce(a.hedefmalzemetakipno ,''), "
                    End If
                Case 3
                    If oSysFlags.G_WANoMTK Then
                        cTranMTFFilter = " malzemetakipkodu = '', "
                    Else
                        cTranMTFFilter = " malzemetakipkodu = coalesce(a.hedefmalzemetakipno ,''), "
                    End If
            End Select

            cSQL = cSQL + _
                    "select stokno = coalesce(a.stokno,''), " + _
                    IIf(nCase = 1, "", " topno = coalesce(a.topno,''), ").ToString + _
                    " renk = coalesce(a.renk ,''), " + _
                    " beden = coalesce(a.beden ,''), " + _
                    cTranMTFFilter + _
                    " partino = coalesce(a.hedefpartino,''), " + _
                    " depo = coalesce(a.hedefdepo,''), " + _
                    " b.anastokgrubu, " + _
                    " b.stoktipi, " + _
                    " b.cinsaciklamasi, " + _
                    " giris = sum(coalesce(a.netmiktar1,0)), " + _
                    " cikis = 0 " + _
                    " from StokTransfer a, stok b " + _
                    " where a.stokno = b.stokno " + _
                    " and (b.kapandi is null or b.kapandi = 'H') " + _
                    cFilterTransfer1 + cTransferTarihFilter

            Select Case nCase
                Case 1
                    cSQL = cSQL + _
                        " group by a.stokno, a.renk, a.beden, a.hedefmalzemetakipno, a.hedefpartino, a.hedefdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi  " + _
                        " Union All "
                Case 2
                    cSQL = cSQL + _
                        " and b.toptakibi = 'E' " + _
                        " and a.topno is not null " + _
                        " and a.topno <> '' " + _
                        " group by a.topno, a.stokno, a.renk, a.beden, a.hedefmalzemetakipno, a.hedefpartino, a.hedefdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi   " + _
                        " Union All "
                Case 3
                    cSQL = cSQL + _
                        " and b.aksesuartakibi = 'E' " + _
                        " and a.topno is not null " + _
                        " and a.topno <> '' " + _
                        " group by a.topno, a.stokno, a.renk, a.beden, a.hedefmalzemetakipno, a.hedefpartino, a.hedefdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi   " + _
                        " Union All "
                Case 4
                    cSQL = cSQL + _
                        " and (b.toptakibi = 'E' or b.aksesuartakibi = 'E') " + _
                        " and a.topno is not null " + _
                        " and a.topno <> '' " + _
                        " group by a.topno, a.stokno, a.renk, a.beden, a.hedefmalzemetakipno, a.hedefpartino, a.hedefdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi   " + _
                        " Union All "
            End Select

            Select Case nCase
                Case 1
                    cTranMTFFilter = " malzemetakipkodu = coalesce(a.kaynakmalzemetakipno ,''), "
                Case 2
                    If oSysFlags.G_WFNoMTK Then
                        cTranMTFFilter = " malzemetakipkodu = '', "
                    Else
                        cTranMTFFilter = " malzemetakipkodu = coalesce(a.kaynakmalzemetakipno ,''), "
                    End If
                Case 3
                    If oSysFlags.G_WANoMTK Then
                        cTranMTFFilter = " malzemetakipkodu = '', "
                    Else
                        cTranMTFFilter = " malzemetakipkodu = coalesce(a.kaynakmalzemetakipno ,''), "
                    End If
            End Select

            cSQL = cSQL + _
                    " select stokno = coalesce(a.stokno,''), " + _
                    IIf(nCase = 1, "", " topno = coalesce(a.topno,''), ").ToString + _
                    " renk = coalesce(a.renk ,''), " + _
                    " beden = coalesce(a.beden,''), " + _
                    cTranMTFFilter + _
                    " partino = coalesce(a.kaynakpartino,''), " + _
                    " depo = coalesce(a.kaynakdepo,''), " + _
                    " b.anastokgrubu, " + _
                    " b.stoktipi, " + _
                    " b.cinsaciklamasi, " + _
                    " giris = 0, " + _
                    " cikis = sum(coalesce(a.netmiktar1,0)) " + _
                    " from StokTransfer a, stok b " + _
                    " where a.stokno = b.stokno " + _
                    " and (b.kapandi is null or b.kapandi = 'H') " + _
                    cFilterTransfer2 + cTransferTarihFilter

            Select Case nCase
                Case 1
                    cSQL = cSQL + _
                        " group by a.stokno, a.renk, a.beden, a.kaynakmalzemetakipno, a.kaynakpartino, a.kaynakdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi  "
                Case 2
                    cSQL = cSQL + _
                        " and b.toptakibi = 'E' " + _
                        " and a.topno is not null " + _
                        " and a.topno <> '' " + _
                        " group by a.topno, a.stokno, a.renk, a.beden, a.kaynakmalzemetakipno, a.kaynakpartino, a.kaynakdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi   "
                Case 3
                    cSQL = cSQL + _
                        " and b.aksesuartakibi = 'E' " + _
                        " and a.topno is not null " + _
                        " and a.topno <> '' " + _
                        " group by a.topno, a.stokno, a.renk, a.beden, a.kaynakmalzemetakipno, a.kaynakpartino, a.kaynakdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi   "
                Case 4
                    cSQL = cSQL + _
                        " and (b.toptakibi = 'E' or b.aksesuartakibi = 'E') " + _
                        " and a.topno is not null " + _
                        " and a.topno <> '' " + _
                        " group by a.topno, a.stokno, a.renk, a.beden, a.kaynakmalzemetakipno, a.kaynakpartino, a.kaynakdepo, b.anastokgrubu, b.stoktipi, b.cinsaciklamasi  "
            End Select

            GetViewStokDurumu = CreateTempView(ConnYage, cSQL)

            CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "GetViewStokDurumu", cSQL)
        End Try
    End Function

    Public Sub HizliStokRBBakimi()

        Dim cSQL As String = ""
        Dim cView As String = ""
        Dim ConnYage As SqlConnection

        Try
            JustForLog("HizliStokRBBakimi basladi")

            cView = GetViewStokDurumu(1)

            ConnYage = OpenConn()

            cSQL = "delete from stokrb"
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "insert stokrb (StokNo, renk, beden, depo, partino, malzemetakipkodu, donemgiris1, donemcikis1, devirgiris1, devircikis1, alismiktari1, alistutari1) " + _
                    " select stokno, renk, beden, depo, partino, malzemetakipkodu, " + _
                    " donemgiris1 = sum(coalesce(giris,0)), " + _
                    " donemcikis1 = sum(coalesce(cikis,0)), " + _
                    " devirgiris1 = 0, " + _
                    " devircikis1 = 0, " + _
                    " alismiktari1 = 0, " + _
                    " alistutari1 = 0 " + _
                    " from " + cView + _
                    " group by stokno, renk, beden, depo, partino, malzemetakipkodu "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropView(cView, ConnYage)

            cSQL = "update stok set " + _
                    " donemgiris1 = coalesce((select sum(coalesce(donemgiris1,0)) from stokrb where stokno = stok.stokno),0), " + _
                    " donemcikis1 = coalesce((select sum(coalesce(donemcikis1,0)) from stokrb where stokno = stok.stokno),0), " + _
                    " devirgiris1 = 0, " + _
                    " devircikis1 = 0 "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' fiyat son girişten okunsun
            cSQL = "update stokrb " + _
                    " set songiristarihi = (select top 1 a.fistarihi " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stokrb.stokno = b.stokno " + _
                                            " and stokrb.renk = b.renk " + _
                                            " and stokrb.beden = b.beden " + _
                                            " and stokrb.malzemetakipkodu = b.malzemetakipkodu " + _
                                            " and stokrb.depo = b.depo " + _
                                            " and stokrb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisfiyati = (select top 1 b.maliyetbirimfiyati " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stokrb.stokno = b.stokno " + _
                                            " and stokrb.renk = b.renk " + _
                                            " and stokrb.beden = b.beden " + _
                                            " and stokrb.malzemetakipkodu = b.malzemetakipkodu " + _
                                            " and stokrb.depo = b.depo " + _
                                            " and stokrb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisdovizi = (select top 1 b.maliyetdovizi " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stokrb.stokno = b.stokno " + _
                                            " and stokrb.renk = b.renk " + _
                                            " and stokrb.beden = b.beden " + _
                                            " and stokrb.malzemetakipkodu = b.malzemetakipkodu " + _
                                            " and stokrb.depo = b.depo " + _
                                            " and stokrb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisdept = (select top 1 a.departman " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stokrb.stokno = b.stokno " + _
                                            " and stokrb.renk = b.renk " + _
                                            " and stokrb.beden = b.beden " + _
                                            " and stokrb.malzemetakipkodu = b.malzemetakipkodu " + _
                                            " and stokrb.depo = b.depo " + _
                                            " and stokrb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisfirmasi = (select top 1 a.firma " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stokrb.stokno = b.stokno " + _
                                            " and stokrb.renk = b.renk " + _
                                            " and stokrb.beden = b.beden " + _
                                            " and stokrb.malzemetakipkodu = b.malzemetakipkodu " + _
                                            " and stokrb.depo = b.depo " + _
                                            " and stokrb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            CloseConn(ConnYage)

            JustForLog("HizliStokRBBakimi bitti")

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "HizliStokRBBakimi", cSQL)
        End Try
    End Sub

    Public Sub HizliStokTopRBBakimi()

        Dim cSQL As String = ""
        Dim cView As String = ""
        Dim cTableName As String = ""
        Dim ConnYage As SqlConnection
        Dim oSysFlags As New General.SysFlags

        Try
            JustForLog("HizliStokTopRBBakimi basladi")

            ConnYage = OpenConn()

            ReadSysFlags(oSysFlags, ConnYage)

            If Not oSysFlags.G_WinFabric Then
                ConnYage.Close()
                Exit Sub
            End If

            cView = GetViewStokDurumu(2)

            cSQL = "(topno char(30) null, " + _
                    " stokno char(30) null, " + _
                    " renk char(30) null, " + _
                    " beden char(30) null, " + _
                    " depo char(30) null, " + _
                    " partino char(30) null, " + _
                    " malzemetakipkodu char(30) null, " + _
                    " donemgiris1 decimal(18,3) null, " + _
                    " donemcikis1 decimal(18,3) null)"

            cTableName = CreateTempTable(ConnYage, cSQL)

            ' hız ve hafıza için stok durumunu tablolaştır
            cSQL = "insert " + cTableName + " (topno, stokno, renk, beden, depo, partino, malzemetakipkodu, donemgiris1, donemcikis1) " + _
                    " select topno, stokno, renk, beden, depo, partino, " + _
                    IIf(oSysFlags.G_WFNoMTK, " malzemetakipkodu = '', ", " malzemetakipkodu, ").ToString + _
                    " donemgiris1 = sum(coalesce(giris,0)), " + _
                    " donemcikis1 = sum(coalesce(cikis,0)) " + _
                    " from " + cView + _
                    " group by topno, stokno, renk, beden, depo, partino " + _
                    IIf(oSysFlags.G_WFNoMTK, "", ", malzemetakipkodu ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' delete all
            cSQL = "delete  stoktoprb"

            ExecuteSQLCommandConnected(cSQL, ConnYage)
            ' rebuild all
            cSQL = "insert stoktoprb (topno, stokno, renk, beden, depo, partino, malzemetakipkodu, donemgiris1, donemcikis1) " + _
                    " select topno, stokno, renk, beden, depo, partino, malzemetakipkodu, donemgiris1, donemcikis1 " + _
                    " from " + cTableName

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            'cSQL = "update stoktoprb set donemgiris1 = 0, donemcikis1 = 0 "
            'ExecuteSQLCommandConnected(cSQL, ConnYage)

            '' stok fiş satırlarında OLMAYAN kayıtlar silinir
            'cSQL = "delete  stoktoprb" + _
            '        " where not exists (select stokfisno  " + _
            '                            " from stokfislines " + _
            '                            " where topno = stoktoprb.topno " + _
            '                            " and stokno = stoktoprb.stokno " + _
            '                            " and renk = stoktoprb.renk " + _
            '                            " and beden = stoktoprb.beden ) "

            'ExecuteSQLCommandConnected(cSQL, ConnYage)

            'cSQL = "update stoktoprb " + _
            '        " set donemgiris1 = (select top 1 coalesce(donemgiris1,0) " + _
            '                            " from " + cTableName + _
            '                            " where topno = stoktoprb.topno " + _
            '                            " and stokno = stoktoprb.stokno " + _
            '                            " and renk = stoktoprb.renk " + _
            '                            " and beden = stoktoprb.beden " + _
            '                            " and depo = stoktoprb.depo " + _
            '                            " and partino = stoktoprb.partino " + _
            '                            IIf(oSysFlags.G_WFNoMTK, "", " and malzemetakipkodu = stoktoprb.malzemetakipkodu ").ToString + _
            '                            " ), "

            'cSQL = cSQL + _
            '        " donemcikis1 = (select top 1 coalesce(donemcikis1,0) " + _
            '                            " from " + cTableName + _
            '                            " where topno = stoktoprb.topno " + _
            '                            " and stokno = stoktoprb.stokno " + _
            '                            " and renk = stoktoprb.renk " + _
            '                            " and beden = stoktoprb.beden " + _
            '                            " and depo = stoktoprb.depo " + _
            '                            " and partino = stoktoprb.partino " + _
            '                            IIf(oSysFlags.G_WFNoMTK, "", " and malzemetakipkodu = stoktoprb.malzemetakipkodu ").ToString + _
            '                            " ) "
            'ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update stoktoprb " + _
                        " set songiristarihi = (select top 1 a.fistarihi " + _
                                                " from stokfis a, stokfislines b " + _
                                                " where a.stokfisno = b.stokfisno " + _
                                                " and stoktoprb.topno = b.topno " + _
                                                " and stoktoprb.stokno = b.stokno " + _
                                                " and stoktoprb.renk = b.renk " + _
                                                " and stoktoprb.beden = b.beden " + _
                                                IIf(oSysFlags.G_WFNoMTK, "", " and stoktoprb.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                                " and stoktoprb.depo = b.depo " + _
                                                " and stoktoprb.partino = b.partino " + _
                                                " and a.stokfistipi = 'Giris' " + _
                                                " and b.maliyetbirimfiyati is not null " + _
                                                " and b.maliyetbirimfiyati <> 0 " + _
                                                " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisfiyati = (select top 1 b.maliyetbirimfiyati " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stoktoprb.topno = b.topno " + _
                                            " and stoktoprb.stokno = b.stokno " + _
                                            " and stoktoprb.renk = b.renk " + _
                                            " and stoktoprb.beden = b.beden " + _
                                            IIf(oSysFlags.G_WFNoMTK, "", " and stoktoprb.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and stoktoprb.depo = b.depo " + _
                                            " and stoktoprb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisdovizi = (select top 1 b.maliyetdovizi " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stoktoprb.topno = b.topno " + _
                                            " and stoktoprb.stokno = b.stokno " + _
                                            " and stoktoprb.renk = b.renk " + _
                                            " and stoktoprb.beden = b.beden " + _
                                            IIf(oSysFlags.G_WFNoMTK, "", " and stoktoprb.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and stoktoprb.depo = b.depo " + _
                                            " and stoktoprb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisdept = (select top 1 a.departman " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stoktoprb.topno = b.topno " + _
                                            " and stoktoprb.stokno = b.stokno " + _
                                            " and stoktoprb.renk = b.renk " + _
                                            " and stoktoprb.beden = b.beden " + _
                                            IIf(oSysFlags.G_WFNoMTK, "", " and stoktoprb.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and stoktoprb.depo = b.depo " + _
                                            " and stoktoprb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisfirmasi = (select top 1 a.firma " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and stoktoprb.topno = b.topno " + _
                                            " and stoktoprb.stokno = b.stokno " + _
                                            " and stoktoprb.renk = b.renk " + _
                                            " and stoktoprb.beden = b.beden " + _
                                            IIf(oSysFlags.G_WFNoMTK, "", " and stoktoprb.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and stoktoprb.depo = b.depo " + _
                                            " and stoktoprb.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete stoktoprb " + _
                    " where coalesce(donemgiris1, 0) - coalesce(donemcikis1, 0) = 0 " + _
                    " and not exists (select * from stokfislines where topno = stoktoprb.topno) " + _
                    " and not exists (select * from stoktransfer where topno = stoktoprb.topno) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropView(cView, ConnYage)
            DropTable(cTableName, ConnYage)

            CloseConn(ConnYage)

            JustForLog("HizliStokTopRBBakimi bitti")

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "HizliStokTopRBBakimi", cSQL)
        End Try
    End Sub

    Public Sub HizliStokAksesuarRBBakimi()

        Dim cSQL As String = ""
        Dim cView As String = ""
        Dim cTableName As String = ""
        Dim ConnYage As SqlConnection
        Dim oSysFlags As New General.SysFlags

        Try
            JustForLog("HizliStokAksesuarRBBakimi basladi")

            ConnYage = OpenConn()

            ReadSysFlags(oSysFlags, ConnYage)

            If Not oSysFlags.G_WinAccessory Then
                ConnYage.Close()
                Exit Sub
            End If

            cView = GetViewStokDurumu(3)

            cSQL = "(topno char(30) null, " + _
                    " stokno char(30) null, " + _
                    " renk char(30) null, " + _
                    " beden char(30) null, " + _
                    " depo char(30) null, " + _
                    " partino char(30) null, " + _
                    " malzemetakipkodu char(30) null, " + _
                    " donemgiris1 decimal(18,3) null, " + _
                    " donemcikis1 decimal(18,3) null)"

            cTableName = CreateTempTable(ConnYage, cSQL)

            ' hız ve hafıza için stok durumunu tablolaştır
            cSQL = "insert " + cTableName + " (topno, stokno, renk, beden, depo, partino, malzemetakipkodu, donemgiris1, donemcikis1) " + _
                    " select topno, stokno, renk, beden, depo, partino, " + _
                    IIf(oSysFlags.G_WANoMTK, " malzemetakipkodu = '', ", " malzemetakipkodu, ").ToString + _
                    " donemgiris1 = sum(coalesce(giris,0)), " + _
                    " donemcikis1 = sum(coalesce(cikis,0)) " + _
                    " from " + cView + _
                    " group by topno, stokno, renk, beden, depo, partino " + _
                    IIf(oSysFlags.G_WANoMTK, "", ", malzemetakipkodu ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' delete all
            cSQL = "delete stokaksesuarrb "

            ExecuteSQLCommandConnected(cSQL, ConnYage)
            ' rebuild all
            cSQL = "insert stokaksesuarrb (topno, stokno, renk, beden, depo, partino, malzemetakipkodu, donemgiris1, donemcikis1) " + _
                    " select topno, stokno, renk, beden, depo, partino, malzemetakipkodu, donemgiris1, donemcikis1 " + _
                    " from " + cTableName

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            '' stok fiş satırlarında OLMAYAN kayıtlar silinir
            'cSQL = "delete stokaksesuarrb " + _
            '        " where not exists (select stokfisno  " + _
            '                            " from stokfislines " + _
            '                            " where topno = stokaksesuarrb.topno " + _
            '                            " and stokno = stokaksesuarrb.stokno " + _
            '                            " and renk = stokaksesuarrb.renk " + _
            '                            " and beden = stokaksesuarrb.beden ) "

            'ExecuteSQLCommandConnected(cSQL, ConnYage)
            '' miktarlar sıfırlanır
            'cSQL = "update StokAksesuarRB set donemgiris1 = 0, donemcikis1 = 0 "
            'ExecuteSQLCommandConnected(cSQL, ConnYage)
            '' miktarlar güncellenir
            'cSQL = "update StokAksesuarRB " + _
            '        " set donemgiris1 = (select top 1 coalesce(donemgiris1,0) " + _
            '                            " from " + cTableName + _
            '                            " where topno = StokAksesuarRB.topno " + _
            '                            " and stokno = StokAksesuarRB.stokno " + _
            '                            " and renk = StokAksesuarRB.renk " + _
            '                            " and beden = StokAksesuarRB.beden " + _
            '                            " and depo = StokAksesuarRB.depo " + _
            '                            " and partino = StokAksesuarRB.partino " + _
            '                            IIf(oSysFlags.G_WANoMTK, "", " and malzemetakipkodu = StokAksesuarRB.malzemetakipkodu ").ToString + _
            '                            " ), " + _
            '        " donemcikis1 = (select top 1 coalesce(donemcikis1,0) " + _
            '                            " from " + cTableName + _
            '                            " where topno = StokAksesuarRB.topno " + _
            '                            " and stokno = StokAksesuarRB.stokno " + _
            '                            " and renk = StokAksesuarRB.renk " + _
            '                            " and beden = StokAksesuarRB.beden " + _
            '                            " and depo = StokAksesuarRB.depo " + _
            '                            " and partino = StokAksesuarRB.partino " + _
            '                            IIf(oSysFlags.G_WANoMTK, "", " and malzemetakipkodu = StokAksesuarRB.malzemetakipkodu ").ToString + _
            '                            " ) "

            'ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update StokAksesuarRB " + _
                            " set songiristarihi = (select top 1 a.fistarihi " + _
                                                    " from stokfis a, stokfislines b " + _
                                                    " where a.stokfisno = b.stokfisno " + _
                                                    " and StokAksesuarRB.topno = b.topno " + _
                                                    " and StokAksesuarRB.stokno = b.stokno " + _
                                                    " and StokAksesuarRB.renk = b.renk " + _
                                                    " and StokAksesuarRB.beden = b.beden " + _
                                                    IIf(oSysFlags.G_WANoMTK, "", " and StokAksesuarRB.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                                    " and StokAksesuarRB.depo = b.depo " + _
                                                    " and StokAksesuarRB.partino = b.partino " + _
                                                    " and a.stokfistipi = 'Giris' " + _
                                                    " and b.maliyetbirimfiyati is not null " + _
                                                    " and b.maliyetbirimfiyati <> 0 " + _
                                                    " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisfiyati = (select top 1 b.maliyetbirimfiyati " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and StokAksesuarRB.topno = b.topno " + _
                                            " and StokAksesuarRB.stokno = b.stokno " + _
                                            " and StokAksesuarRB.renk = b.renk " + _
                                            " and StokAksesuarRB.beden = b.beden " + _
                                            IIf(oSysFlags.G_WANoMTK, "", " and StokAksesuarRB.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and StokAksesuarRB.depo = b.depo " + _
                                            " and StokAksesuarRB.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisdovizi = (select top 1 b.maliyetdovizi " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and StokAksesuarRB.topno = b.topno " + _
                                            " and StokAksesuarRB.stokno = b.stokno " + _
                                            " and StokAksesuarRB.renk = b.renk " + _
                                            " and StokAksesuarRB.beden = b.beden " + _
                                            IIf(oSysFlags.G_WANoMTK, "", " and StokAksesuarRB.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and StokAksesuarRB.depo = b.depo " + _
                                            " and StokAksesuarRB.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisdept = (select top 1 a.departman " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and StokAksesuarRB.topno = b.topno " + _
                                            " and StokAksesuarRB.stokno = b.stokno " + _
                                            " and StokAksesuarRB.renk = b.renk " + _
                                            " and StokAksesuarRB.beden = b.beden " + _
                                            IIf(oSysFlags.G_WANoMTK, "", " and StokAksesuarRB.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and StokAksesuarRB.depo = b.depo " + _
                                            " and StokAksesuarRB.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc), "
            cSQL = cSQL + _
                    " songirisfirmasi = (select top 1 a.firma " + _
                                            " from stokfis a, stokfislines b " + _
                                            " where a.stokfisno = b.stokfisno " + _
                                            " and StokAksesuarRB.topno = b.topno " + _
                                            " and StokAksesuarRB.stokno = b.stokno " + _
                                            " and StokAksesuarRB.renk = b.renk " + _
                                            " and StokAksesuarRB.beden = b.beden " + _
                                            IIf(oSysFlags.G_WANoMTK, "", " and StokAksesuarRB.malzemetakipkodu = b.malzemetakipkodu ").ToString + _
                                            " and StokAksesuarRB.depo = b.depo " + _
                                            " and StokAksesuarRB.partino = b.partino " + _
                                            " and a.stokfistipi = 'Giris' " + _
                                            " and b.maliyetbirimfiyati is not null " + _
                                            " and b.maliyetbirimfiyati <> 0 " + _
                                            " order by a.fistarihi desc) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete stokaksesuarrb " + _
                    " where coalesce(donemgiris1, 0) - coalesce(donemcikis1, 0) = 0 " + _
                    " and not exists (select * from stokfislines where topno = stokaksesuarrb.topno) " + _
                    " and not exists (select * from stoktransfer where topno = stokaksesuarrb.topno) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropView(cView, ConnYage)
            DropTable(cTableName, ConnYage)

            ConnYage.Close()

            JustForLog("HizliStokAksesuarRBBakimi bitti")

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "HizliStokAksesuarRBBakimi", cSQL)
        End Try
    End Sub
End Module

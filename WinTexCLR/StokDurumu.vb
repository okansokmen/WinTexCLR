Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module StokDurumu

    Structure SipStok
        Dim cSiparisNo As String
        Dim nAdet As Double
    End Structure

    Structure Asorti
        Dim cBeden As String
        Dim nAdet As Double
    End Structure

    Public Function BarkodStokDurumu(ByVal cTarih As String, ByVal nKumasAks As Integer) As String

        Dim ConnYage As SqlConnection
        Dim cTarihFilter1 As String
        Dim cTarihFilter2 As String
        Dim cTarihFilter3 As String
        Dim cSQL As String
        Dim cBarkodNull As String
        Dim oSysFlags As New General.SysFlags
        Dim cViewName As String
        Dim cTableName As String
        Dim cRandom As String

        BarkodStokDurumu = ""

        Try
            cRandom = Rnd().ToString
            cViewName = "os_v_sed_" + cRandom
            cTableName = "os_t_sed_" + cRandom

            cTarihFilter1 = " and fistarihi <= '" + cTarih + "' "
            cTarihFilter2 = " and tarih <= '" + cTarih + "' "
            cTarihFilter3 = " and tarih <= '" + cTarih + "' "

            cBarkodNull = ""
            Select Case nKumasAks
                Case 1 ' kumas
                    cBarkodNull = " and SUBSTRING (topno,1,3) = '000' "
                Case 2 ' aksesuar
                    cBarkodNull = " and SUBSTRING (topno,1,3) = '001' "
            End Select

            ConnYage = OpenConn()

            ReadSysFlags(oSysFlags, ConnYage)

            DropView(cViewName, ConnYage)

            cSQL = "create view " + cViewName + " as " + _
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
                    " cik_agirlik = 0" + _
                " FROM stokfis a, stokfislines b " + _
                " WHERE a.stokfisno = b.stokfisno " + _
                " and (stokfistipi = 'Giris') " + _
                cTarihFilter1 + _
                cBarkodNull + _
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
                cTarihFilter1 + _
                cBarkodNull + _
                " GROUP BY stokno, renk, beden, depo, partino, malzemetakipkodu, topno "

            cSQL = cSQL + _
                " Union All " + _
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
                " WHERE stokno is not null " + _
                cTarihFilter2 + _
                cBarkodNull + _
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
                " WHERE stokno is not null " + _
                cTarihFilter2 + _
                cBarkodNull + _
                " GROUP BY stokno, renk, beden, kaynakpartino, kaynakdepo, kaynakmalzemetakipno, topno"

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropTable(cTableName, ConnYage)

            cSQL = "CREATE TABLE " + cTableName + _
                    " (stokno char(30)," + _
                    " renk char(30), " + _
                    " beden char(30), " + _
                    " mtkno char(30), " + _
                    " partino char(30), " + _
                    " depo char(30), " + _
                    " topno char(30), " + _
                    " rbaadet decimal(18,3), " + _
                    " rbaagirlik decimal(18,3))  "

            ExecuteSQLCommandConnected(cSQL, ConnYage)


            If (nKumasAks = 1 And oSysFlags.G_WFNoMTK) Or (nKumasAks = 2 And oSysFlags.G_WANoMTK) Then
                cSQL = "insert into " + cTableName + " (stokno, renk, beden, mtkno, partino, depo, topno, rbaadet, rbaagirlik) " + _
                        " select stokno, renk, beden, mtkno = '', partino, depo, topno, " + _
                        " rbaadet = sum(coalesce(giris,0) - coalesce(cikis,0)), " + _
                        " rbaagirlik = sum(coalesce(gir_agirlik,0) - coalesce(cik_agirlik,0)) " + _
                        " from " + cViewName + _
                        " group by stokno, renk, beden, partino, depo, topno "
            Else
                cSQL = "insert into " + cTableName + " (stokno, renk, beden, mtkno, partino, depo, topno, rbaadet, rbaagirlik) " + _
                        " select stokno, renk, beden, mtkno, partino, depo, topno, " + _
                        " rbaadet = sum(coalesce(giris,0) - coalesce(cikis,0)), " + _
                        " rbaagirlik = sum(coalesce(gir_agirlik,0) - coalesce(cik_agirlik,0)) " + _
                        " from " + cViewName + _
                        " group by stokno, renk, beden, mtkno, partino, depo, topno "
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            'Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM " + cTableName, ConnYage)
            'SqlContext.Pipe.ExecuteAndSend(cmd)

            DropView(cViewName, ConnYage)
            'DropTable(cTableName, ConnYage)

            CloseConn(ConnYage)

            BarkodStokDurumu = cTableName
        Catch
            BarkodStokDurumu = ""
            ErrDisp("Error BarkodStokDurumu " + Err.Description.Trim)
        End Try
    End Function

    Public Function StokDurumuTarih(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal cStokNo As String, ByVal nCase As Integer) As SqlString

        Dim ConnYage As SqlConnection
        Dim cSQL As String
        Dim cTableName As String

        Try
            ConnYage = OpenConn()
            cTableName = CreateTempTable(ConnYage)
            StokDurumuTarih = ""
            cSQL = " Select x.entegrekodu, x.cinsaciklamasi "
            If CBool(CInt(nCase = 0) Or CInt(nCase = 3)) Then
                cSQL = cSQL + ", Kumas=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='KUMAS DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='KUMAS DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.kumas),0) From Devir2009 b Where b.stokno = x.stokno )"
            End If
            If CBool(CInt(nCase = 1) Or CInt(nCase = 3)) Then
                cSQL = cSQL + ", Aksesuar=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='AKSESUAR DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='AKSESUAR DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.Aksesuar),0) From Devir2009 b Where b.stokno = x.stokno )"
            End If
            If CBool(CInt(nCase = 2) Or CInt(nCase = 3)) Then
                cSQL = cSQL + ", Mamul=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='MAMUL DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='MAMUL DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.mamul),0) From Devir2009 b Where b.stokno = x.stokno )  "
                cSQL = cSQL + ", Kırık=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='KIRIK DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='KIRIK DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.kırık),0) From Devir2009 b Where b.stokno = x.stokno )  "
                cSQL = cSQL + ", Defo=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='II.KALITE DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='II.KALITE DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.Defo),0) From Devir2009 b Where b.stokno = x.stokno )  "
                cSQL = cSQL + ", TUrun=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='T.URUN DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='T.URUN DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.Turun),0) From Devir2009 b Where b.stokno = x.stokno )  "
                cSQL = cSQL + ", TKırık=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='T.KIRIK DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='T.KIRIK DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.Tkırık),0) From Devir2009 b Where b.stokno = x.stokno )  "
                cSQL = cSQL + ", TDefo=(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Giris' " + _
                              "and b.depo='T.II.KALITE DEPO' " + _
                              "and b.stokno = x.stokno ) - " + _
                              "(Select coalesce(sum(b.netmiktar1),0) " + _
                              "From Stokfis a, Stokfislines b " + _
                              "Where a.stokfisno=b.stokfisno " + _
                              "and a.fistarihi>='" + BslTarihi + "' " + _
                              "and a.fistarihi<='" + BtsTarihi + "' " + _
                              "and a.stokfistipi='Cikis' " + _
                              "and b.depo='T.II.KALITE DEPO' " + _
                              "and b.stokno = x.stokno ) + " + _
                              "(Select coalesce(sum(b.Tdefo),0) From Devir2009 b Where b.stokno = x.stokno )  "
            End If
            '      cSQL = cSQL + "into " + cTableName + " from stok x  Where (x.kapandi = '' or x.kapandi='H' or x.kapandi is null) "
            cSQL = cSQL + "into " + cTableName + " from stok x  "

            If cStokNo = "" Then
                Select Case nCase
                    Case 0 : cSQL = cSQL + " Where x.anastokgrubu in('KUMAS','ASTAR')"
                    Case 1 : cSQL = cSQL + " Where x.anastokgrubu='AKSESUAR'"
                    Case 2 : cSQL = cSQL + " Where x.anastokgrubu='MAMUL'"
                End Select
            Else
                cSQL = cSQL + " Where x.entegrekodu = '" + cStokNo.Trim + "' "
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

            CloseConn(ConnYage)
            StokDurumuTarih = cTableName

        Catch Err As Exception
            StokDurumuTarih = "Hata"
            ErrDisp("Error DonemselStokDurumu " + Err.Message)
        End Try
    End Function
End Module

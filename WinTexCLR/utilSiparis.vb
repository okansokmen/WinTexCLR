Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module utilSiparis
    Private Structure oMRB
        Dim cModelNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cSTF As String
        Dim nAdet As Double
        Dim dSevkTar As Date
    End Structure

    Public Sub GetSipAnaKumas(ByVal cSiparisNo As String, ByVal cModelNo As String, ByVal cRenk As String, ByRef cAnaKumas As String, ByRef cAnaKumasRenk As String)

        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cSQL As String = ""
        Dim cReceteNo As String = ""

        cAnaKumas = ""
        cAnaKumasRenk = ""

        Try
            ConnYage = OpenConn()

            cSQL = "select receteno " + _
                    " from sipmodel " + _
                    " where siparisno = '" + cSiparisNo.Trim + "' " + _
                    " and modelno = '" + cModelNo.Trim + "' " + _
                    " and renk = '" + cRenk.Trim + "' "

            cReceteNo = SQLGetStringConnected(cSQL, ConnYage)

            If cReceteNo.Trim = "" Then
                cSQL = "select hammaddekodu, hammadderenk " + _
                        " from modelhammadde " + _
                        " where modelno = '" + cModelNo.Trim + "' " + _
                        " and (modelrenk = '" + cRenk.Trim + "' or modelrenk = 'HEPSI')" + _
                        " and anakumas = 'E' "
            Else
                cSQL = "select hammaddekodu, hammadderenk " + _
                        " from modelhammadde2 " + _
                        " where modelno = '" + cModelNo.Trim + "' " + _
                        " and (modelrenk = '" + cRenk.Trim + "' or modelrenk = 'HEPSI')" + _
                        " and receteno = '" + cReceteNo.Trim + "'" + _
                        " and anakumas = 'E' "
            End If

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                cAnaKumas = SQLReadString(oReader, "hammaddekodu")
                If SQLReadString(oReader, "hammadderenk") = "HEPSI" Then
                    cAnaKumasRenk = cRenk
                Else
                    cAnaKumasRenk = SQLReadString(oReader, "hammadderenk")
                End If
            End If
            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp("GetSipAnaKumas : " + ex.Message.Trim + vbCrLf + cSQL)
        End Try
    End Sub

    Public Sub GetSipAnaKumTed(ByVal cSiparisNo As String, ByVal cModelNo As String, ByVal cRenk As String, _
                               ByRef cFirma As String, ByRef dTermin As Date, ByRef nIhtiyac As Double, ByRef nKarsilanan As Double, ByRef dGiris As Date)

        Dim cSQL As String = ""
        Dim cAnaKumas As String = ""
        Dim cAnaKumasRenk As String = ""
        Dim cMTF As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        cFirma = ""
        dTermin = CDate("01.01.1950")
        nIhtiyac = 0
        nKarsilanan = 0
        dGiris = CDate("01.01.1950")

        Try
            GetSipAnaKumas(cSiparisNo, cModelNo, cRenk, cAnaKumas, cAnaKumasRenk)

            If cAnaKumas = "" Then Exit Sub

            ConnYage = OpenConn()

            cSQL = "select malzemetakipno, plfirma, bitistarihi, " + _
                    " ihtiyac = sum(coalesce(ihtiyac,0)), " + _
                    " karsilanan = sum(coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0)) " + _
                    " from mtkfislines " + _
                    " where stokno = '" + cAnaKumas + "' " + _
                    " and renk = '" + cAnaKumasRenk + "' " + _
                    " and malzemetakipno in (select malzemetakipno " + _
                                            " from sipmodel " + _
                                            " where siparisno = '" + cSiparisNo + "' " + _
                                            " and modelno = '" + cModelNo + "' " + _
                                            " and renk = '" + cRenk + "') group by malzemetakipno, plfirma, bitistarihi "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                cMTF = SQLReadString(oReader, "malzemetakipno")
                cFirma = SQLReadString(oReader, "plfirma")
                dTermin = SQLReadDate(oReader, "BitisTarihi")
                nIhtiyac = SQLReadDouble(oReader, "ihtiyac")
                nKarsilanan = SQLReadDouble(oReader, "Karsilanan")
            End If
            oReader.Close()
            oReader = Nothing

            cSQL = "select a.firma, b.termintarihi " + _
                    " from isemri a, isemrilines b " + _
                    " where a.isemrino = b.isemrino " + _
                    " and b.stokno = '" + cAnaKumas + "' " + _
                    " and b.renk = '" + cAnaKumasRenk + "' " + _
                    " and b.malzemetakipno in (select malzemetakipno " + _
                                            " from sipmodel " + _
                                            " where siparisno = '" + cSiparisNo + "' " + _
                                            " and modelno = '" + cModelNo + "' " + _
                                            " and renk = '" + cRenk + "')" + _
                    " order by b.termintarihi desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                cFirma = SQLReadString(oReader, "Firma")
                dTermin = SQLReadDate(oReader, "termintarihi")
            End If
            oReader.Close()
            oReader = Nothing

            cSQL = "select a.fistarihi " + _
                    " from stokfis a, stokfislines b " + _
                    " where a.stokfisno = b.stokfisno " + _
                    " and b.stokno = '" + cAnaKumas + "' " + _
                    " and b.renk = '" + cAnaKumasRenk + "' " + _
                    " and b.malzemetakipkodu in (select malzemetakipno " + _
                                            " from sipmodel " + _
                                            " where siparisno = '" + cSiparisNo + "' " + _
                                            " and modelno = '" + cModelNo + "' " + _
                                            " and renk = '" + cRenk + "')" + _
                    " order by a.fistarihi "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                dGiris = SQLReadDate(oReader, "fistarihi")
            End If
            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp("GetSipAnaKumTed : " + ex.Message.Trim + vbCrLf + cSQL)
        End Try
    End Sub

    Public Function GetUretimDurumu(ByVal cSiparisNo As String, Optional ByVal cModelNo As String = "", Optional ByVal cRenk As String = "") As String

        Dim cSQL As String = ""

        GetUretimDurumu = "Üretime Girmemiş"

        Try
            cSQL = "select a.girisdept " + _
                    " from uretharfis a, uretharrba b " + _
                    " where a.uretfisno = b.uretfisno " + _
                    " and b.uretimtakipno in (select uretimtakipno from sipmodel where siparisno = '" + cSiparisNo + "') " + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cRenk.Trim = "", "", " and renk = '" + cRenk.Trim + "' ").ToString + _
                    " order by a.fistarihi desc "

            GetUretimDurumu = SQLGetString(cSQL)

        Catch ex As Exception
            ErrDisp("GetUretimDurumu : " + ex.Message.Trim + vbCrLf + cSQL)
        End Try
    End Function

    Public Sub GetDeptFason(ByVal cSiparisNo As String, ByRef cFirma As String, ByRef nFiyat As Double, ByVal cDepartman As String)

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        cFirma = ""
        nFiyat = 0

        Try
            ConnYage = OpenConn()
            ' önce üretim hareketlerine bak
            cSQL = "select a.girisfirm_atl, b.fiyati, b.fiyatdoviz " + _
                    " from uretharfis a, uretharfislines b " + _
                    " where a.uretfisno = b.uretfisno " + _
                    " and a.girisdept like '%" + cDepartman.Trim + "%' " + _
                    " and b.uretimtakipno in (select uretimtakipno from sipmodel where siparisno = '" + cSiparisNo.Trim + "') " + _
                    " and a.girisfirm_atl is not null " + _
                    " and a.girisfirm_atl <> '' " + _
                    " order by a.fistarihi desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                cFirma = SQLReadString(oReader, "girisfirm_atl")
                nFiyat = SQLReadDouble(oReader, "fiyati")
            End If
            oReader.Close()
            ' sonra üretim işemirlerine bak
            cSQL = "select a.firma, b.fiyati, b.doviz " + _
                    " from uretimisemri a, uretimisdetayi b " + _
                    " where a.isemrino = b.isemrino " + _
                    " and a.departman like '%" + cDepartman.Trim + "%' " + _
                    " and a.uretimtakipno in (select uretimtakipno from sipmodel where siparisno = '" + cSiparisNo.Trim + "') " + _
                    " and a.firma is not null " + _
                    " and a.firma <> '' " + _
                    " order by a.tarih desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                If cFirma.Trim = "" Then
                    cFirma = SQLReadString(oReader, "firma")
                End If
                If nFiyat = 0 Then
                    nFiyat = SQLReadDouble(oReader, "fiyati")
                End If
            End If
            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp("GetDeptFason : " + ex.Message.Trim + vbCrLf + cSQL)
        End Try
    End Sub

    Public Function GetSipAksDurum(ByVal cSiparisNo As String, ByVal cModelNo As String, ByVal cRenk As String) As String

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        GetSipAksDurum = "Planlanmamış"

        Try
            ConnYage = OpenConn()

            cSQL = "select b.stoktipi, " +
                    " ihtiyac = sum(coalesce(a.ihtiyac,0)), " +
                    " isemriverilen = sum(coalesce(a.isemriverilen,0)), " +
                    " isemriicingelen = sum(coalesce(a.isemriicingelen,0)), " +
                    " isemriharicigelen = sum(coalesce(a.isemriharicigelen,0)), " +
                    " uretimicincikis = sum(coalesce(a.uretimicincikis,0)), " +
                    " uretimdeniade = sum(coalesce(a.uretimdeniade,0)) " +
                    " from mtkfislines a, stok b " +
                    " where a.stokno = b.stokno " +
                    " and a.departman not like '%KESIM%' " +
                    " and a.malzemetakipno in (select malzemetakipno " +
                                            " from sipmodel " +
                                            " where siparisno = '" + cSiparisNo.Trim + "' " +
                                            " and modelno = '" + cModelNo.Trim + "' " +
                                            " and renk = '" + cRenk.Trim + "')" +
                    " and not (b.anastokgrubu like '%KUMAS%' or " +
                             " b.anastokgrubu like '%KUMAŞ%' or " +
                             " (b.stoktipi like '%TELA%' and b.kesimecik = 'E') or " +
                             " b.stoktipi like '%ASTAR%' or " +
                             " b.stoktipi like '%CEPLIK%') " +
                    " group by b.stoktipi " +
                    " ORDER BY b.stoktipi "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                GetSipAksDurum = "Tamamlandı"
                If SQLReadDouble(oReader, "ihtiyac") > SQLReadDouble(oReader, "isemriicingelen") + SQLReadDouble(oReader, "isemriharicigelen") Then
                    GetSipAksDurum = "Eksik"
                    Exit Do
                End If
            Loop
            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp("GetSipAksDurum : " + ex.Message.Trim + vbCrLf + cSQL)
        End Try
    End Function

    Public Function GetSipAdet(Optional cSiparisNo As String = "", Optional cModelNo As String = "", Optional cRenk As String = "", Optional cBeden As String = "",
                               Optional cUTF As String = "", Optional cMTF As String = "", Optional cSTF As String = "", Optional cFirma As String = "", Optional dIlkSevkTar As Date = #1/1/1950#) As Double
        Dim cSQL As String = ""

        GetSipAdet = 0

        Try
            cSQL = "select sum(coalesce(adet,0)) " +
                " from sipmodel " +
                " where adet is not null " +
                " and adet <> 0 " +
                IIf(cSiparisNo = "", "", " and siparisno = '" + Trim(cSiparisNo) + "' ").ToString +
                IIf(cModelNo = "", "", " and modelno = '" + Trim(cModelNo) + "' ").ToString +
                IIf(cRenk = "" Or cRenk = "HEPSI", "", " and renk = '" + Trim(cRenk) + "' ").ToString +
                IIf(cBeden = "" Or cBeden = "HEPSI", "", " and beden = '" + Trim(cBeden) + "' ").ToString +
                IIf(cUTF = "", "", " and malzemetakipno = '" + Trim(cUTF) + "' ").ToString +
                IIf(cMTF = "", "", " and uretimtakipno = '" + Trim(cMTF) + "' ").ToString +
                IIf(cSTF = "", "", " and sevkiyattakipno = '" + Trim(cSTF) + "' ").ToString +
                IIf(cFirma = "", "", " and firma = '" + Trim(cFirma) + "' ").ToString +
                IIf(dIlkSevkTar = #1/1/1950#, "", " and ilksevktar = '" + CStr(dIlkSevkTar) + "' ").ToString

            GetSipAdet = SQLGetDouble(cSQL)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipAdet", cSQL)
        End Try
    End Function

    Public Function GetSipUretimCikis(Optional cSiparisNo As String = "", Optional cDept As String = "", Optional cModelNo As String = "", Optional cRenk As String = "", Optional cBeden As String = "",
                                      Optional lExactDepartmentName As Boolean = False, Optional cFirma As String = "", Optional cUTF As String = "", Optional cExcludeUretFisNo As String = "") As Double
        Dim cSQL As String = ""

        GetSipUretimCikis = 0

        Try
            If cDept.Trim = "" Then Exit Function

            cSQL = "select sum(coalesce(a.adet,0)) " +
                " from uretharrba a , uretharfis b, uretharfislines c  " +
                " where a.uretfisno = b.uretfisno " +
                " and b.uretfisno = c.uretfisno " +
                " and a.ulineno = c.ulineno " +
                IIf(cFirma.Trim = "", "", " and b.cikisfirm_atl = '" + cFirma.Trim + "' ").ToString +
                IIf(cModelNo.Trim = "", "", " and a.modelno = '" + cModelNo.Trim + "' ").ToString +
                IIf(cRenk.Trim = "" Or cRenk.Trim = "HEPSI", "", " and a.renk = '" + cRenk.Trim + "' ").ToString +
                IIf(cBeden.Trim = "" Or cBeden.Trim = "HEPSI", "", " and a.beden = '" + cBeden.Trim + "' ").ToString +
                IIf(cUTF.Trim = "", "", " and a.uretimtakipno = '" + cUTF.Trim + "' ").ToString +
                IIf(cExcludeUretFisNo.Trim = "", "", " and b.uretfisno <> '" + cExcludeUretFisNo.Trim + "' ").ToString

            If cSiparisNo.Trim <> "" Then
                cSQL = cSQL +
                " and a.uretimtakipno in (select uretimtakipno " +
                                        " from sipmodel " +
                                        " where siparisno = '" + cSiparisNo.Trim + "') "
            End If

            If lExactDepartmentName Then
                If cDept = "KESİM" Or cDept = "KESIM" Then
                    cSQL = cSQL + " and b.cikisdept in ('KESİM','KESIM') "
                ElseIf cDept = "DİKİM" Or cDept = "DIKIM" Then
                    cSQL = cSQL + " and b.cikisdept in ('DİKİM','DIKIM') "
                Else
                    cSQL = cSQL + " and b.cikisdept = '" + cDept.Trim + "' "
                End If
            Else
                cSQL = cSQL + " and (b.cikisdept like '%" + cDept.Trim + "%' or b.cikisdept = 'KOMPLE') "
            End If

            GetSipUretimCikis = SQLGetDouble(cSQL)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipUretimCikis", cSQL)
        End Try
    End Function

    Public Function GetSipUretIsemri(ConnYage As SqlConnection, Optional cSiparisNo As String = "", Optional cDept As String = "", Optional cModelNo As String = "", Optional cRenk As String = "", Optional cBeden As String = "",
                                    Optional lExactDepartmentName As Boolean = False, Optional cFirma As String = "", Optional cUTF As String = "") As Double
        Dim cSQL As String = ""

        GetSipUretIsemri = 0

        Try
            If cDept.Trim = "" Then Exit Function

            cSQL = "select sum(coalesce(c.adet,0)) " +
                    " from uretimisemri a, uretimisdetayi b, uretimisrba c " +
                    " where a.isemrino = b.isemrino " +
                    " and b.isemrino = c.isemrino " +
                    " and b.ulineno = c.ulineno " +
                    IIf(cFirma.Trim = "", "", " and a.firma = '" + cFirma.Trim + "' ").ToString +
                    IIf(cModelNo.Trim = "", "", " and c.modelno = '" + cModelNo.Trim + "' ").ToString +
                    IIf(cRenk.Trim = "", "", " and c.renk = '" + cRenk.Trim + "' ").ToString +
                    IIf(cBeden.Trim = "", "", " and c.beden = '" + cBeden.Trim + "' ").ToString +
                    IIf(cUTF.Trim = "", "", " and c.uretimtakipno = '" + cUTF.Trim + "' ").ToString

            If cSiparisNo.Trim <> "" Then
                cSQL = cSQL +
                    " and c.uretimtakipno in (select uretimtakipno " +
                                        " from sipmodel " +
                                        " where siparisno = '" + cSiparisNo.Trim + "') "
            End If

            If lExactDepartmentName Then
                If cDept = "KESİM" Or cDept = "KESIM" Then
                    cSQL = cSQL + " and a.departman in ('KESİM','KESIM') "
                ElseIf cDept = "DİKİM" Or cDept = "DIKIM" Then
                    cSQL = cSQL + " and a.departman in ('DİKİM','DIKIM') "
                Else
                    cSQL = cSQL + " and a.departman = '" + cDept.Trim + "' "
                End If
            Else
                cSQL = cSQL + " and a.departman like '%" + cDept.Trim + "%' "
            End If

            GetSipUretIsemri = SQLGetDoubleConnected(cSQL, ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipUretIsemri", cSQL)
        End Try
    End Function

    Public Function GetSevkAdet(Optional cModelNo As String = "", Optional cUTF As String = "", Optional cRenk As String = "", Optional cBeden As String = "",
                                Optional dTarih As Date = #1/1/1950#, Optional cSiparisNo As String = "") As Double
        Dim cSQL As String = ""
        Dim cFilter As String = ""

        GetSevkAdet = 0

        Try
            cFilter = "select sevkiyattakipno " +
                        " from sipmodel " +
                        " where sevkiyattakipno is not null " +
                        " and sevkiyattakipno <> '' " +
                        IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString +
                        IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString +
                        IIf(cSiparisNo.Trim = "", "", " and siparisno = '" + cSiparisNo.Trim + "' ").ToString +
                        IIf(cRenk.Trim = "", "", " and renk = '" + cRenk.Trim + "' ").ToString +
                        IIf(cBeden.Trim = "", "", " and beden = '" + cBeden.Trim + "' ").ToString

            cSQL = "select toplam = sum((b.koliend - b.kolibeg + 1) * c.adet) " +
                " from sevkform a, sevkformlines b, sevkformlinesrba c " +
                " where a.sevkformno = b.sevkformno " +
                " and b.sevkformno = c.sevkformno " +
                " and b.ulineno = c.ulineno " +
                " and a.ok = 'E' " +
                IIf(cSiparisNo.Trim = "", "", " and b.siparisno = '" + cSiparisNo.Trim + "' ").ToString +
                IIf(cModelNo.Trim = "", "", " and b.modelno = '" + cModelNo.Trim + "' ").ToString +
                IIf(cRenk.Trim = "", "", " and c.renk = '" + cRenk.Trim + "' ").ToString +
                IIf(cBeden.Trim = "", "", " and c.beden = '" + cBeden.Trim + "' ").ToString +
                IIf(dTarih = #1/1/1950#, "", " and a.sevktar <= '" + CStr(dTarih) + "' ").ToString +
                " and (a.sevkiyattakipno in (" + cFilter.Trim + " ) or b.sevkiyattakipno in (" + cFilter.Trim + " )) "

            GetSevkAdet = SQLGetDouble(cSQL)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSevkAdet", cSQL)
        End Try
    End Function

    Public Function GetSipSevkTarih2(cSiparisNo As String, Optional cSevkTarihFilter As String = "") As Date
        ' ilk gerçekleşmiş sevkiyat tarihini alır
        Dim cSQL As String = ""

        GetSipSevkTarih2 = #1/1/1950#

        Try
            cSQL = "select top 1 a.sevktar " +
                    " from sevkform a, sevkformlines b " +
                    " where a.sevkformno = b.sevkformno " +
                    " and b.siparisno = '" + Trim(cSiparisNo) + "' " +
                    IIf(cSevkTarihFilter = "", "", cSevkTarihFilter).ToString +
                    " order by a.sevktar  "

            GetSipSevkTarih2 = SQLGetDate(cSQL)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipSevkTarih2", cSQL)
        End Try
    End Function

    Public Function GetSipTutarDvz(cSiparisNo As String, Optional cModelNo As String = "", Optional cRenk As String = "",
                                   Optional cBeden As String = "", Optional cSevkiyatTakipNo As String = "", Optional cDvz As String = "TL", Optional cSTFs As String = "") As Double
        Dim cSQL As String = ""
        Dim nFiyat As Double = 0
        Dim cDoviz As String = ""
        Dim nKur As Double = 0
        Dim dTarih As Date = #1/1/1950#
        Dim nDvzFiyat As Double = 0
        Dim nDvzTutar As Double = 0
        Dim nDvzKur As Double = 0
        Dim aMRB() As oMRB
        Dim nCnt As Integer = -1
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        GetSipTutarDvz = 0

        Try
            nDvzTutar = 0

            cSQL = "select siparistarihi " +
                " from siparis " +
                " where kullanicisipno = '" + cSiparisNo.Trim + "' "

            dTarih = SQLGetDate(cSQL)

            If dTarih = #1/1/1950# Then
                dTarih = Today
            End If

            cSQL = "select modelno, renk, beden, sevkiyattakipno, " +
                " adet = sum(coalesce(adet,0)) " +
                " from sipmodel " +
                " where siparisno = '" + cSiparisNo.Trim + "' " +
                IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString +
                IIf(cRenk.Trim = "", "", " and renk = '" + cRenk.Trim + "' ").ToString +
                IIf(cBeden.Trim = "", "", " and beden = '" + cBeden.Trim + "' ").ToString +
                IIf(cSevkiyatTakipNo.Trim = "", "", " and sevkiyattakipno = '" + cSevkiyatTakipNo.Trim + "' ").ToString +
                IIf(cSTFs.Trim = "", "", " and sevkiyattakipno in (" + cSTFs.Trim + ") ").ToString +
                " group by modelno, renk, beden, sevkiyattakipno "

            If Not CheckExists(cSQL) Then Exit Function

            ConnYage = OpenConn()

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aMRB(nCnt)

                aMRB(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aMRB(nCnt).cRenk = SQLReadString(oReader, "renk")
                aMRB(nCnt).cBeden = SQLReadString(oReader, "beden")
                aMRB(nCnt).cSTF = SQLReadString(oReader, "sevkiyattakipno")
                aMRB(nCnt).nAdet = SQLReadDouble(oReader, "adet")
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aMRB)
                nFiyat = 0
                cDoviz = ""

                cSQL = "select satisfiyati, satisdoviz " +
                    " from sipfiyat " +
                    " where siparisno = '" + cSiparisNo.Trim + "' " +
                    " and modelkodu = '" + aMRB(nCnt).cModelNo + "' " +
                    " and (renk = '" + aMRB(nCnt).cRenk + "' or renk = 'HEPSI') " +
                    " and (beden = '" + aMRB(nCnt).cBeden + "' or beden = 'HEPSI') " +
                    " and satisfiyati is not null " +
                    " and satisfiyati <> 0 " +
                    " and (sevkiyattakipno = '" + aMRB(nCnt).cSTF + "' or sevkiyattakipno = 'HEPSI') "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "satisfiyati")
                    cDoviz = SQLReadString(oReader, "satisdoviz")
                End If
                oReader.Close()

                If nFiyat = 0 Then
                    cSQL = "select satisfiyati, satisdoviz " +
                        " from sipfiyat " +
                        " where siparisno = '" + cSiparisNo.Trim + "' " +
                        " and modelkodu = '" + aMRB(nCnt).cModelNo + "' " +
                        " and (renk = '" + aMRB(nCnt).cRenk + "' or renk = 'HEPSI') " +
                        " and (beden = '" + aMRB(nCnt).cBeden + "' or beden = 'HEPSI') " +
                        " and satisfiyati is not null " +
                        " and satisfiyati <> 0 "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        nFiyat = SQLReadDouble(oReader, "satisfiyati")
                        cDoviz = SQLReadString(oReader, "satisdoviz")
                    End If
                    oReader.Close()
                End If

                If nFiyat <> 0 Then
                    If cDoviz = cDvz Then
                        nDvzTutar = nDvzTutar + (aMRB(nCnt).nAdet * nFiyat)
                    Else
                        nKur = GetKurConnected(ConnYage, cDoviz, dTarih)
                        nDvzKur = GetKurConnected(ConnYage, cDvz, dTarih)
                        If nDvzKur <> 0 Then
                            nDvzFiyat = nFiyat * nKur / nDvzKur
                            nDvzTutar = nDvzTutar + (aMRB(nCnt).nAdet * nDvzFiyat)
                        End If
                    End If

                End If
            Next

            GetSipTutarDvz = nDvzTutar

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipTutarDvz", cSQL)
        End Try
    End Function

    Public Function GetKurConnected(ConnYage As SqlConnection, Optional cDoviz As String = "", Optional dTarih As DateTime = #1/1/1950#, Optional cKurCinsi As String = "", Optional cFirma As String = "") As Double

        Dim cSQL As String = ""

        GetKurConnected = 1

        Try
            cDoviz = cDoviz.Trim
            cKurCinsi = cKurCinsi.Trim
            cFirma = cFirma.Trim

            If cDoviz = "" Or cDoviz = "TL" Or cDoviz = "YTL" Then
                Exit Function
            End If

            If dTarih = #1/1/1950# Then dTarih = Today

            If cKurCinsi.Trim = "" And cFirma.Trim <> "" Then
                cSQL = "select KurCinsi " +
                        " from firma " +
                        " where firma = '" + cFirma + "' "

                cKurCinsi = SQLGetStringConnected(cSQL, ConnYage)
            End If

            If cKurCinsi = "" Or cKurCinsi = "Alis Kuru" Then
                cKurCinsi = "Kur"
            End If
            If InStr(LCase(cKurCinsi), "satis") > 0 And cKurCinsi <> "Efektif Satis Kuru" Then
                cKurCinsi = "Satis Kuru"
            End If

            cSQL = "set dateformat dmy " +
                    " select top 1 kur " +
                    " from dovkur " +
                    " where doviz = '" + cDoviz.Trim + "' " +
                    " and kurcinsi = '" + cKurCinsi.Trim + "' " +
                    " and tarih <= '" + SQLWriteDate(dTarih) + "' " +
                    " and kur is not null " +
                    " and kur <> 0 " +
                    " order by tarih desc "

            GetKurConnected = SQLGetDoubleConnected(cSQL, ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message, "GetKurConnected", cSQL)
        End Try
    End Function

    Public Function GetSipSevkDVZTutar(cSiparisNo As String, Optional cModelNo As String = "", Optional cDvz As String = "TL", Optional cSTFs As String = "") As Double

        Dim cSQL As String = ""
        Dim cFilter As String = ""
        Dim nKur As Double = 0
        Dim nFiyat As Double = 0
        Dim cDoviz As String = ""
        Dim nDvzKur As Double = 0
        Dim nCnt As Integer = -1
        Dim aMRB() As oMRB
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        GetSipSevkDVZTutar = 0

        Try
            ConnYage = OpenConn()

            cSQL = "select a.sevktar, a.sevkiyattakipno, b.modelno, c.renk, c.beden, " +
                " adet = sum((b.koliend - b.kolibeg + 1) * c.adet) " +
                " from sevkform a, sevkformlines b, sevkformlinesrba c " +
                " where a.sevkformno = b.sevkformno " +
                " and b.sevkformno = c.sevkformno " +
                " and b.ulineno = c.ulineno " +
                " and a.ok = 'E' " +
                " and b.siparisno = '" + cSiparisNo.Trim + "' " +
                IIf(cModelNo.Trim = "", "", " and b.modelno = '" + cModelNo.Trim + "' ").ToString +
                IIf(cSTFs.Trim = "", "", " and a.sevkiyattakipno in (" + cSTFs.Trim + ") ").ToString +
                " group by a.sevktar, a.sevkiyattakipno, b.modelno, c.renk, c.beden "

            If Not CheckExistsConnected(cSQL, ConnYage) Then
                ConnYage.Close()
                Exit Function
            End If

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aMRB(nCnt)

                aMRB(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aMRB(nCnt).cRenk = SQLReadString(oReader, "renk")
                aMRB(nCnt).cBeden = SQLReadString(oReader, "beden")
                aMRB(nCnt).cSTF = SQLReadString(oReader, "sevkiyattakipno")
                aMRB(nCnt).nAdet = SQLReadDouble(oReader, "adet")
                aMRB(nCnt).dSevkTar = SQLReadDate(oReader, "sevktar")
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aMRB)
                nFiyat = 0
                cDoviz = ""

                cSQL = "select satisfiyati, satisdoviz " +
                    " from sipfiyat " +
                    " where siparisno = '" + cSiparisNo.Trim + "' " +
                    " and modelkodu = '" + aMRB(nCnt).cModelNo + "' " +
                    " and (renk = '" + aMRB(nCnt).cRenk + "' or renk = 'HEPSI') " +
                    " and (beden = '" + aMRB(nCnt).cBeden + "' or beden = 'HEPSI') " +
                    " and satisfiyati is not null " +
                    " and satisfiyati <> 0 " +
                    " and (sevkiyattakipno = '" + aMRB(nCnt).cSTF + "' or sevkiyattakipno = 'HEPSI') "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "satisfiyati")
                    cDoviz = SQLReadString(oReader, "satisdoviz")
                End If
                oReader.Close()

                If nFiyat = 0 Then
                    cSQL = "select satisfiyati, satisdoviz " +
                        " from sipfiyat " +
                        " where siparisno = '" + cSiparisNo.Trim + "' " +
                        " and modelkodu = '" + aMRB(nCnt).cModelNo + "' " +
                        " and (renk = '" + aMRB(nCnt).cRenk + "' or renk = 'HEPSI') " +
                        " and (beden = '" + aMRB(nCnt).cBeden + "' or beden = 'HEPSI') " +
                        " and satisfiyati is not null " +
                        " and satisfiyati <> 0 "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        nFiyat = SQLReadDouble(oReader, "satisfiyati")
                        cDoviz = SQLReadString(oReader, "satisdoviz")
                    End If
                    oReader.Close()
                End If

                If nFiyat <> 0 Then
                    If cDoviz = cDvz Then
                        GetSipSevkDVZTutar = GetSipSevkDVZTutar + (nFiyat * aMRB(nCnt).nAdet)
                    Else
                        nKur = GetKurConnected(ConnYage, cDoviz, aMRB(nCnt).dSevkTar)
                        nDvzKur = GetKurConnected(ConnYage, cDvz, aMRB(nCnt).dSevkTar)

                        If nDvzKur <> 0 Then
                            GetSipSevkDVZTutar = GetSipSevkDVZTutar + (nFiyat * nKur / nDvzKur * aMRB(nCnt).nAdet)
                        End If
                    End If
                End If
            Next

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipSevkDVZTutar", cSQL)
        End Try
    End Function

    Public Function GetSipSevkDVZTutarConnected(ConnYage As SqlConnection, cSiparisNo As String, Optional cModelNo As String = "", Optional cDvz As String = "TL", Optional cSTFs As String = "") As Double

        Dim cSQL As String = ""
        Dim cFilter As String = ""
        Dim nKur As Double = 0
        Dim nFiyat As Double = 0
        Dim cDoviz As String = ""
        Dim nDvzKur As Double = 0
        Dim nCnt As Integer = -1
        Dim aMRB() As oMRB
        Dim oReader As SqlDataReader

        GetSipSevkDVZTutarConnected = 0

        Try
            cSQL = "select a.sevktar, a.sevkiyattakipno, b.modelno, c.renk, c.beden, " +
                    " adet = sum((b.koliend - b.kolibeg + 1) * c.adet) " +
                    " from sevkform a, sevkformlines b, sevkformlinesrba c " +
                    " where a.sevkformno = b.sevkformno " +
                    " and b.sevkformno = c.sevkformno " +
                    " and b.ulineno = c.ulineno " +
                    " and a.ok = 'E' " +
                    " and b.siparisno = '" + cSiparisNo.Trim + "' " +
                    IIf(cModelNo.Trim = "", "", " and b.modelno = '" + cModelNo.Trim + "' ").ToString +
                    IIf(cSTFs.Trim = "", "", " and a.sevkiyattakipno in (" + cSTFs.Trim + ") ").ToString +
                    " group by a.sevktar, a.sevkiyattakipno, b.modelno, c.renk, c.beden "

            If Not CheckExistsConnected(cSQL, ConnYage) Then
                Exit Function
            End If

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aMRB(nCnt)

                aMRB(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aMRB(nCnt).cRenk = SQLReadString(oReader, "renk")
                aMRB(nCnt).cBeden = SQLReadString(oReader, "beden")
                aMRB(nCnt).cSTF = SQLReadString(oReader, "sevkiyattakipno")
                aMRB(nCnt).nAdet = SQLReadDouble(oReader, "adet")
                aMRB(nCnt).dSevkTar = SQLReadDate(oReader, "sevktar")
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aMRB)
                nFiyat = 0
                cDoviz = ""

                cSQL = "select satisfiyati, satisdoviz " +
                    " from sipfiyat " +
                    " where siparisno = '" + cSiparisNo.Trim + "' " +
                    " and modelkodu = '" + aMRB(nCnt).cModelNo + "' " +
                    " and (renk = '" + aMRB(nCnt).cRenk + "' or renk = 'HEPSI') " +
                    " and (beden = '" + aMRB(nCnt).cBeden + "' or beden = 'HEPSI') " +
                    " and satisfiyati is not null " +
                    " and satisfiyati <> 0 " +
                    " and (sevkiyattakipno = '" + aMRB(nCnt).cSTF + "' or sevkiyattakipno = 'HEPSI') "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "satisfiyati")
                    cDoviz = SQLReadString(oReader, "satisdoviz")
                End If
                oReader.Close()

                If nFiyat = 0 Then
                    cSQL = "select satisfiyati, satisdoviz " +
                        " from sipfiyat " +
                        " where siparisno = '" + cSiparisNo.Trim + "' " +
                        " and modelkodu = '" + aMRB(nCnt).cModelNo + "' " +
                        " and (renk = '" + aMRB(nCnt).cRenk + "' or renk = 'HEPSI') " +
                        " and (beden = '" + aMRB(nCnt).cBeden + "' or beden = 'HEPSI') " +
                        " and satisfiyati is not null " +
                        " and satisfiyati <> 0 "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        nFiyat = SQLReadDouble(oReader, "satisfiyati")
                        cDoviz = SQLReadString(oReader, "satisdoviz")
                    End If
                    oReader.Close()
                End If

                If nFiyat <> 0 Then
                    If cDoviz = cDvz Then
                        GetSipSevkDVZTutarConnected = GetSipSevkDVZTutarConnected + (nFiyat * aMRB(nCnt).nAdet)
                    Else
                        nKur = GetKurConnected(ConnYage, cDoviz, aMRB(nCnt).dSevkTar)
                        nDvzKur = GetKurConnected(ConnYage, cDvz, aMRB(nCnt).dSevkTar)

                        If nDvzKur <> 0 Then
                            GetSipSevkDVZTutarConnected = GetSipSevkDVZTutarConnected + (nFiyat * nKur / nDvzKur * aMRB(nCnt).nAdet)
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipSevkDVZTutarConnected", cSQL)
        End Try
    End Function

    Public Sub GetSipFiyat(ConnYage As SqlConnection, cSiparisNo As String, ByRef nSatFiyat As Double, Optional ByRef cSatDoviz As String = "", Optional ByVal cHedefDoviz As String = "", Optional ByVal cModelNo As String = "",
                        Optional ByRef dSiparisTarih As Date = #1/1/1950#, Optional ByRef nSatKur As Double = 0, Optional ByRef nHdfKur As Double = 0,
                        Optional ByRef nOrjinalSatFiyat As Double = 0, Optional ByRef nOrjinalSatDoviz As String = "TL", Optional ByRef cOnMaliyetModelNo As String = "",
                        Optional ByVal lGetSipFiyatFromDokumaOnMaliyet As Boolean = False)

        Dim cSQL As String = ""
        Dim oReader As SqlDataReader

        Try
            cSQL = "select satisfiyati, satisdoviz, onmaliyetmodelno " +
                    " from sipfiyat " +
                    " where siparisno = '" + cSiparisNo.Trim + "' " +
                    IIf(cModelNo.Trim = "", "", " and (modelkodu = '" + cModelNo.Trim + "' or modelkodu = 'HEPSI') ").ToString +
                    " and satisfiyati is not null " +
                    " and satisfiyati > 0 "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nOrjinalSatFiyat = SQLReadDouble(oReader, "satisfiyati")
                nOrjinalSatDoviz = SQLReadString(oReader, "satisdoviz")
                cOnMaliyetModelNo = SQLReadString(oReader, "onmaliyetmodelno")
            End If
            oReader.Close()

            If lGetSipFiyatFromDokumaOnMaliyet Then

                nOrjinalSatFiyat = 0
                nOrjinalSatDoviz = "EUR"

                cSQL = "select targetfiyat, doviz " +
                        " from maliyetheader " +
                        " where calismano = '" + cOnMaliyetModelNo.Trim + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nOrjinalSatFiyat = SQLReadDouble(oReader, "targetfiyat")
                    nOrjinalSatDoviz = SQLReadString(oReader, "doviz")
                End If
                oReader.Close()
            End If

            nSatFiyat = nOrjinalSatFiyat
            cSatDoviz = nOrjinalSatDoviz

            cSQL = "select siparistarihi " +
                    " from siparis " +
                    " where kullanicisipno = '" + cSiparisNo.Trim + "' "

            dSiparisTarih = SQLGetDateConnected(cSQL, ConnYage)

            If dSiparisTarih = #1/1/1950# Then
                dSiparisTarih = Today
            End If

            If cHedefDoviz = "" Then
                nSatKur = GetKurConnected(ConnYage, cSatDoviz, dSiparisTarih)
                nHdfKur = nSatKur
            Else
                If cSatDoviz = cHedefDoviz Then
                    nSatKur = GetKurConnected(ConnYage, cSatDoviz, dSiparisTarih)
                    nHdfKur = nSatKur
                Else
                    nSatKur = GetKurConnected(ConnYage, cSatDoviz, dSiparisTarih)
                    nHdfKur = GetKurConnected(ConnYage, cHedefDoviz, dSiparisTarih)
                End If

                If nHdfKur <> 0 Then
                    nSatFiyat = nSatFiyat * nSatKur / nHdfKur
                    cSatDoviz = cHedefDoviz
                End If
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "GetSipFiyat", cSQL)
        End Try
    End Sub

End Module

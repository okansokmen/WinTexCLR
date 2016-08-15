Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module utilSiparis

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

            cSQL = "select b.stoktipi, " + _
                    " ihtiyac = sum(coalesce(a.ihtiyac,0)), " + _
                    " isemriverilen = sum(coalesce(a.isemriverilen,0)), " + _
                    " isemriicingelen = sum(coalesce(a.isemriicingelen,0)), " + _
                    " isemriharicigelen = sum(coalesce(a.isemriharicigelen,0)), " + _
                    " uretimicincikis = sum(coalesce(a.uretimicincikis,0)), " + _
                    " uretimdeniade = sum(coalesce(a.uretimdeniade,0)) " + _
                    " from mtkfislines a, stok b " + _
                    " where a.stokno = b.stokno " + _
                    " and a.departman not like '%KESIM%' " + _
                    " and a.malzemetakipno in (select malzemetakipno " + _
                                            " from sipmodel " + _
                                            " where siparisno = '" + cSiparisNo.Trim + "' " + _
                                            " and modelno = '" + cModelNo.Trim + "' " + _
                                            " and renk = '" + cRenk.Trim + "')" + _
                    " and not (b.anastokgrubu like '%KUMAS%' or " + _
                             " b.anastokgrubu like '%KUMAŞ%' or " + _
                             " (b.stoktipi like '%TELA%' and b.kesimecik = 'E') or " + _
                             " b.stoktipi like '%ASTAR%' or " + _
                             " b.stoktipi like '%CEPLIK%') " + _
                    " group by b.stoktipi " + _
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
End Module

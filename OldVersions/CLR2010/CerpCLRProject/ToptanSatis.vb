Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module ToptanSatis

    Structure FirmaAsorti
        Dim cBedenSeti As String
        Dim cBeden As String
        Dim nAdet As Double
    End Structure

    Public Function MagazaSendData(ByVal cModelNoSelected As String, ByVal cRenkSelected As String, ByVal cAsortiNoSelected As String, _
                               ByVal cSiparisNo As String, ByVal cModelNo As String, ByVal cRenk As String, ByVal cBeden As String, _
                               ByVal nAdet As Double, ByVal cMagaza As String, ByVal cAsorti As String, ByVal nSatirNo As Double, _
                               ByVal nFiyat As Double, ByVal cSiparisTarihi As String, ByVal cIlkSevkTarihi As String, _
                               ByVal cMusteriNo As String, ByVal cUserName As String, ByVal cGenelNotlar As String) As String

        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cSQL As String

        Dim dSimdi As Date
        Dim cBedenSeti As String = ""
        Dim nAAdet As Double
        Dim nCnt As Integer
        Dim aBedenSeti() As String
        Dim lFirstBedenSeti As Boolean = True
        Dim aFirmaAsorti() As FirmaAsorti
        Dim lBedenSeti As Boolean
        Dim nBedenSetiAdedi As Integer
        Dim aBedenSeti2() As String
        Dim cBuffer As String
        Dim nCnt2 As Integer
        Dim lFirstFirmaAsorti As Boolean = True
        Dim nAsortiToplami As Double
        Dim nBSCount As Integer
        Dim nMaxBedenSetiAdedi As Integer = 50

        MagazaSendData = "OK"
        cSQL = ""

        Try
            ConnYage = OpenConn()

            ReDim aFirmaAsorti(0)
            ReDim aBedenSeti(0)

            dSimdi = GetNowFromServer(ConnYage)
            If nSatirNo <> 0 Then
                nCnt = CInt(nSatirNo)
            End If

            ' update tsipmodel

            If cAsorti = "TEKLEME" Then

                cBedenSeti = cBeden
                cAsortiNoSelected = cBedenSeti + "-1"

                ' bedenseti yoksa ekle

                cSQL = "select bedenseti " + _
                        " from bedenseti " + _
                        " where bedenseti = '" + cBedenSeti + "' "

                If Not CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "insert into bedenseti (bedenseti,b01) " + _
                            " values ('" + cBedenSeti + "', " + _
                            " '" + cBeden + "') "

                    Call ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                ' asorti yoksa ekle

                cSQL = "select asortino " + _
                        " from firmaasorti " + _
                        " where asortino = '" + cAsortiNoSelected + "' " + _
                        " and adet = 1 "

                If Not CheckExistsConnected(cSQL, ConnYage) Then

                    cSQL = "delete from firmaasorti " + _
                            " where asortino = '" + cAsortiNoSelected + "' "

                    Call ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "insert into firmaasorti (asortino, bedenseti, renk, beden, adet) " + _
                            " values ('" + cAsortiNoSelected + "', " + _
                            " '" + cBedenSeti + "', " + _
                            " 'HEPSI', " + _
                            " '" + cBeden + "', " + _
                            " 1 ) "

                    Call ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                ' modelasorti yoksa ekle

                cSQL = "select modelno " + _
                        " from modelasorti " + _
                        " where modelno = '" + cModelNoSelected + "' " + _
                        " and asortino = '" + cAsortiNoSelected + "' "

                If Not CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "insert into modelasorti (modelno,asortino,sayac) " + _
                            " values ('" + cModelNoSelected + "', " + _
                            " '" + cAsortiNoSelected + "', " + _
                            " 0 ) "

                    Call ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                ' tsipmodel update 
                ' teklemede 1 adet beden oluyor mutlaka 

                cSQL = "delete tsipmodel " + _
                        " where siparisno = '" + cSiparisNo + "' " + _
                        " and modelno = '" + cModelNo + "' " + _
                        " and renk = '" + cRenk + "' " + _
                        " and beden = '" + cBeden + "' " + _
                        " and bedenseti = '" + cBedenSeti + "' " + _
                        " and asortino = '" + cAsortiNoSelected + "' "

                Call ExecuteSQLCommandConnected(cSQL, ConnYage)

                cSQL = " insert into tsipmodel (siparisno, modelno, renk, beden, bedenseti, " + _
                        " adet, fiyat, doviz, asortino, asortiadet, aciklama, receteno, kur) " + _
                        " values ( " + _
                        " '" + cSiparisNo + "', " + _
                        " '" + cModelNo + "', " + _
                        " '" + cRenk + "', " + _
                        " '" + cBeden + "', " + _
                        " '" + cBedenSeti + "', " + _
                        SQLWriteDecimal(nAdet) + ", " + _
                        SQLWriteDecimal(nFiyat) + ", " + _
                        " 'USD', " + _
                        " '" + cAsortiNoSelected + "', " + _
                        SQLWriteDecimal(nAdet) + ",'','',0 ) "

                Call ExecuteSQLCommandConnected(cSQL, ConnYage)

            Else

                nAsortiToplami = 0

                cSQL = "select bedenseti = coalesce(bedenseti,''), " + _
                        " beden = coalesce(beden,''), " + _
                        " adet = coalesce(adet,0) " + _
                        " from firmaasorti " + _
                        " where asortino = '" + cAsorti + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read

                    cBedenSeti = SQLReadString(oReader, "bedenseti").Trim ' SQLReadIntegeroReader.GetString(oReader.GetOrdinal("bedenseti")).Trim
                    cBeden = SQLReadString(oReader, "beden").Trim 'oReader.GetString(oReader.GetOrdinal("beden")).Trim
                    nAAdet = SQLReadDouble(oReader, "adet") 'CDbl(oReader.GetValue(oReader.GetOrdinal("adet")))

                    If lFirstFirmaAsorti Then
                        aFirmaAsorti(0).cBedenSeti = cBedenSeti
                        aFirmaAsorti(0).cBeden = cBeden
                        aFirmaAsorti(0).nAdet = nAAdet
                        lFirstFirmaAsorti = False
                    Else
                        ReDim Preserve aFirmaAsorti(UBound(aFirmaAsorti) + 1)
                        aFirmaAsorti(UBound(aFirmaAsorti)).cBedenSeti = cBedenSeti
                        aFirmaAsorti(UBound(aFirmaAsorti)).cBeden = cBeden
                        aFirmaAsorti(UBound(aFirmaAsorti)).nAdet = nAAdet
                    End If

                    nAsortiToplami = nAsortiToplami + nAAdet
                Loop
                oReader.Close()
                oReader = Nothing
                ' oCommand = Nothing

                If lFirstFirmaAsorti Or nAsortiToplami = 0 Then
                    'MagazaSendData = "Hata : " + cAsorti + " asorti tanımı hatalı"
                Else
                    For nCnt = 0 To UBound(aFirmaAsorti)

                        cSQL = "delete tsipmodel " + _
                                " where siparisno = '" + cSiparisNo + "' " + _
                                " and modelno = '" + cModelNo + "' " + _
                                " and renk = '" + cRenk + "' " + _
                                " and beden = '" + aFirmaAsorti(nCnt).cBeden + "' " + _
                                " and bedenseti = '" + aFirmaAsorti(nCnt).cBedenSeti + "' " + _
                                " and asortino = '" + cAsorti + "' "

                        Call ExecuteSQLCommandConnected(cSQL, ConnYage)

                        cSQL = "insert into tsipmodel (siparisno, modelno, renk, beden, bedenseti, " + _
                                " adet, fiyat, doviz, asortino, asortiadet, aciklama, receteno, kur) " + _
                                " values ( " + _
                                " '" + cSiparisNo + "', " + _
                                " '" + cModelNo + "', " + _
                                " '" + cRenk + "', " + _
                                " '" + aFirmaAsorti(nCnt).cBeden + "', " + _
                                " '" + aFirmaAsorti(nCnt).cBedenSeti + "', " + _
                                SQLWriteDecimal(nAdet * aFirmaAsorti(nCnt).nAdet) + ", " + _
                                SQLWriteDecimal(nFiyat) + ", " + _
                                " 'USD', " + _
                                " '" + cAsorti + "', " + _
                                SQLWriteDecimal(nAdet) + ",'','',0 ) "

                        Call ExecuteSQLCommandConnected(cSQL, ConnYage)
                    Next
                End If
            End If

            ' update tsiparis 

            cSQL = "select siparisno from tsiparis where siparisno = '" + cSiparisNo + "'"

            If CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = " update tsiparis set " + _
                       " siparistarihi = '" + cSiparisTarihi + "', " + _
                       " musterino = '" + cMusteriNo + "', " + _
                       " ilksevktarihi = '" + cIlkSevkTarihi + "', " + _
                       " sorumlu = '" + cUserName + "', " + _
                       " acenta = '" + cMagaza + "', " + _
                       " modificationdate = '" + SQLWriteDate(dSimdi) + "', " + _
                       " username = '" + cUserName + "', " + _
                       " genelnotlar = '" + cGenelNotlar + "' " + _
                       " where siparisno = '" + cSiparisNo + "'"

                ExecuteSQLCommandConnected(cSQL, ConnYage, True)

                ' ilgili beden seti sipariş kartında yoksa açalım

                nBedenSetiAdedi = 0
                lBedenSeti = False

                ' sql server index optimization ok

                cSQL = "select top 1 "

                For nBSCount = 1 To nMaxBedenSetiAdedi
                    cSQL = cSQL + " bedenseti" + nBSCount.ToString + " = coalesce(bedenseti" + nBSCount.ToString + ",'') "
                    If nBSCount <> nMaxBedenSetiAdedi Then
                        cSQL = cSQL + ","
                    End If
                Next

                cSQL = cSQL + _
                        " from tsiparis " + _
                        " where siparisno = '" + cSiparisNo + "' "

                oReader = GetSQLReader(cSQL, ConnYage)


                If oReader.Read() Then
                    nCnt2 = 0
                    For nCnt = 1 To nMaxBedenSetiAdedi
                        cBuffer = SQLReadString(oReader, "bedenseti" + nCnt.ToString).Trim 'oReader.GetString(oReader.GetOrdinal("bedenseti" + nCnt.ToString)).Trim
                        If cBuffer <> "" Then
                            nBedenSetiAdedi = nBedenSetiAdedi + 1
                            If cBedenSeti = cBuffer Then
                                lBedenSeti = True
                            End If
                            If lFirstBedenSeti Then
                                aBedenSeti(0) = cBuffer
                                lFirstBedenSeti = False
                            Else
                                nCnt2 = nCnt2 + 1
                                ReDim Preserve aBedenSeti(nCnt2)
                                aBedenSeti(nCnt2) = cBuffer
                            End If
                        End If
                    Next
                    If Not lBedenSeti Then
                        If lFirstBedenSeti Then
                            aBedenSeti(0) = cBedenSeti
                            lFirstBedenSeti = False
                        Else
                            nCnt2 = nCnt2 + 1
                            ReDim Preserve aBedenSeti(nCnt2)
                            aBedenSeti(nCnt2) = cBedenSeti
                        End If
                    End If
                End If
                oReader.Close()
                oReader = Nothing
                '  oCommand = Nothing

                ReDim aBedenSeti2(0)
                nCnt2 = 0
                lFirstBedenSeti = True
                For nCnt = 0 To UBound(aBedenSeti)

                    ' sql server index optimization ok

                    cSQL = "select modelno " + _
                            " from tsipmodel " + _
                            " where siparisno = '" + cSiparisNo + "' " + _
                            " and bedenseti = '" + aBedenSeti(nCnt) + "' "

                    If CheckExistsConnected(cSQL, ConnYage) Then
                        If lFirstBedenSeti Then
                            aBedenSeti2(0) = aBedenSeti(nCnt)
                            lFirstBedenSeti = False
                        Else
                            nCnt2 = nCnt2 + 1
                            ReDim Preserve aBedenSeti2(nCnt2)
                            aBedenSeti2(nCnt2) = aBedenSeti(nCnt)
                        End If
                    End If
                Next

                Debug.WriteLine("nBedenSetiAdedi : " + nBedenSetiAdedi.ToString)

                cSQL = "update tsiparis set "

                For nCnt = 0 To UBound(aBedenSeti2)
                    cSQL = cSQL + " bedenseti" + CStr(nCnt + 1) + " = '" + aBedenSeti2(nCnt) + "', "
                Next

                For nCnt = UBound(aBedenSeti2) + 1 To nMaxBedenSetiAdedi - 1
                    cSQL = cSQL + " bedenseti" + CStr(nCnt + 1) + " = '', "
                Next
                cSQL = cSQL + _
                        " kilitle = 'H' " + _
                        " where siparisno = '" + cSiparisNo + "'"

                Call ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else

                cSQL = " insert into tsiparis (siparisno, siparistarihi, musterino, dosyakapandi, ilksevktarihi, " + _
                         " sorumlu, acenta, creationdate, modificationdate, username, bedenseti1, genelnotlar) " + _
                         " values " + _
                         " ('" + cSiparisNo + "', " + _
                         " '" + cSiparisTarihi + "', " + _
                         " '" + cMusteriNo + "', " + _
                         " 'H', " + _
                         " '" + cIlkSevkTarihi + "', " + _
                         " '" + cUserName + "', " + _
                         " '" + cMagaza + "', " + _
                         " '" + SQLWriteDate(dSimdi) + "', " + _
                         " '" + SQLWriteDate(dSimdi) + "', " + _
                         " '" + cUserName + "', " + _
                         " '" + cBedenSeti + "', " + _
                         " '" + cGenelNotlar + "' )"

                Call ExecuteSQLCommandConnected(cSQL, ConnYage, True)
            End If

            Call CloseConn(ConnYage)
            MagazaSendData = "OK"

        Catch Err As Exception
            MagazaSendData = "Hata"
            ErrDisp("MagazaSendData : " + Err.Message.Trim + vbCrLf + cSQL)
        End Try
    End Function


End Module

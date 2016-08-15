Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module MTF

    Public Const G_isemriicinGelenGiris = " (stokhareketkodu = '04 Mlz Uretimden Giris' " + _
                                        " or stokhareketkodu = '06 Tamirden Giris' " + _
                                        " or stokhareketkodu = '02 Tedarikten Giris'  " + _
                                        " or stokhareketkodu = '05 Diger Giris' ) "

    Public Const G_isemriicinGelenCikis = " (stokhareketkodu = '02 Tedarikten iade' " + _
                                        " or stokhareketkodu = '06 Tamire Cikis' " + _
                                        " or stokhareketkodu = '04 Mlz Uretime iade' " + _
                                        " or stokhareketkodu = '05 Diger Cikis' ) "

    Public Const G_isemriHariciGelenGiris = " (stokhareketkodu = '05 Diger Giris' " + _
                                        " or stokhareketkodu = '02 Tedarikten Giris' " + _
                                        " or stokhareketkodu = '04 Mlz Uretimden Giris' " + _
                                        " or stokhareketkodu = '06 Tamirden Giris' " + _
                                        " or stokhareketkodu = '55 Kontrol Oncesi Giris' " + _
                                        " or stokhareketkodu = '77 Top Bolme Giris' " + _
                                        " or stokhareketkodu = '77 Aksesuar Bolme Giris' " + _
                                        " or stokhareketkodu = '08 SAYIM GIRIS' " + _
                                        " or stokhareketkodu = '90 Trans/Rezv Giris') "

    Public Const G_isemriHariciGelenCikis = " (stokhareketkodu = '05 Diger Cikis' " + _
                                        " or stokhareketkodu = '02 Tedarikten iade' " + _
                                        " or stokhareketkodu = '04 Mlz Uretime iade' " + _
                                        " or stokhareketkodu = '06 Tamire Cikis' " + _
                                        " or stokhareketkodu = '55 Kontrol Oncesi Cikis' " + _
                                        " or stokhareketkodu = '77 Top Bolme Cikis' " + _
                                        " or stokhareketkodu = '77 Aksesuar Bolme Cikis' " + _
                                        " or stokhareketkodu = '08 SAYIM CIKIS' " + _
                                        " or stokhareketkodu = '90 Trans/Rezv Cikis') "

    Public Const G_uretimicincikis = " stokhareketkodu = '01 Uretime Cikis' "

    Public Const G_uretimdeniade = " stokhareketkodu = '01 Uretimden iade' "

    Private Structure MRB
        Dim cModelNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cReceteNo As String
        Dim nAdet As Double
    End Structure

    Private Structure MTF
        Dim cStokNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cUDept As String
        Dim cMDept As String
        Dim nMiktar As Double
        Dim cBirim As String
        Dim nFire As Double
        Dim cHesaplama As String
        Dim nSiraNo As Double
    End Structure

    Public Sub MTFFastGenerateAll()

        Dim cSQL As String = ""
        Dim aMTF() As String = Nothing
        Dim nCnt As Integer = 0
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""

        Try
            lAltModelDetay = (GetSysPar("altmodeltakibi") = "1")

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            cSQL = "select distinct a.malzemetakipno " + _
                    " from " + cSipModelTableName + " a, siparis b  " + _
                    " where a.siparisno = b.kullanicisipno " + _
                    " and a.malzemetakipno is not null " + _
                    " and a.malzemetakipno <> '' " + _
                    " and (b.dosyakapandi = 'H' or b.dosyakapandi = '' or b.dosyakapandi is null) " + _
                    " order by a.malzemetakipno "

            If CheckExists(cSQL) Then
                aMTF = SQLtoStringArray(cSQL)
                For nCnt = 0 To UBound(aMTF)
                    MTKFastGenerate(aMTF(nCnt))
                Next
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "MTFFastGenerateAll", cSQL)
        End Try
    End Sub

    Public Function MTKFastGenerate(ByVal cMTF As String) As Integer

        Dim nCnt1 As Integer = 0
        Dim aMRBA() As MRB = Nothing
        Dim aMRBRA() As MRB = Nothing
        Dim aMTF() As MTF = Nothing
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim cSQL As String = ""
        Dim nCnt As Integer = 0
        Dim cMusteri As String = ""
        Dim lKesileneGore As Boolean = False
        Dim lKesIsemrineGore As Boolean = False
        Dim nKesilenAdet As Double = 0
        Dim nSiparisAdet As Double = 0
        Dim nReceteAdet As Double = 0
        Dim nUretimToleransi As Double = 0
        Dim nFire As Double = 0
        Dim nMalzemeFireFactor As Double = 1
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""

        MTKFastGenerate = 0

        Try
            If cMTF.Trim = "" Then Exit Function

            ConnYage = OpenConn()

            JustForLog("MTF bakimi basladi : " + cMTF.Trim)

            lKesileneGore = (GetSysParConnected("mtfkesilenegore", ConnYage) = "1")
            lKesIsemrineGore = (GetSysParConnected("mtfkesisemrinegore", ConnYage) = "1")
            lAltModelDetay = (GetSysParConnected("altmodeltakibi", ConnYage) = "1")

            Debug.WriteLine("lKesileneGore " + GetSysParConnected("mtfkesilenegore", ConnYage))
            Debug.WriteLine("lKesIsemrineGore " + GetSysParConnected("mtfkesisemrinegore", ConnYage))
            Debug.WriteLine("lAltModelDetay " + GetSysParConnected("altmodeltakibi", ConnYage))

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            cSQL = "select distinct musterino " + _
                    " from siparis " + _
                    " where kullanicisipno in (select siparisno " + _
                                            " from  " + cSipModelTableName + _
                                            " where malzemetakipno = '" + cMTF.Trim + "') " + _
                    " and musterino is not null " + _
                    " and musterino <> '' "

            cMusteri = SQLGetStringConnected(cSQL, ConnYage)

            cSQL = "select malzemetakipno " + _
                    " from mtkfis " + _
                     " where malzemetakipno = '" + cMTF.Trim + "' "

            If CheckExistsConnected(cSQL, ConnYage) Then
                If cMusteri <> "" Then

                    cSQL = "update mtkfis " + _
                            " set musteri = '" + SQLWriteString(cMusteri, 30) + "' " + _
                            " where malzemetakipno = '" + cMTF.Trim + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Else
                cSQL = "insert into mtkfis " + _
                        " (malzemetakipno, dosyakapandi, musteri) " + _
                        " values ('" + cMTF.Trim + "', " + _
                        " 'H', " + _
                        " '" + SQLWriteString(cMusteri, 30) + "') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            ' kilitlenmemiş satırların ihtiyaçlarını sıfırla
            cSQL = "update mtkfislines " + _
                    " set ihtiyac = 0 " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and (kilitle is null or kilitle = 'H' or kilitle = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set hesaplananihtiyac = 0 " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' model adetlerini siparişten hesapla

            cSQL = "select modelno, renk, beden, receteno, adet = sum(coalesce(adet,0)) " + _
                    " from " + cSipModelTableName + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and adet is not null " + _
                    " and adet <> 0 " + _
                    " group by modelno, renk, beden, receteno " + _
                    " order by modelno, renk, beden, receteno "

            nCnt = 0

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ' Final RBA
                ReDim Preserve aMRBRA(nCnt)
                aMRBRA(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aMRBRA(nCnt).cRenk = SQLReadString(oReader, "renk")
                aMRBRA(nCnt).cBeden = SQLReadString(oReader, "beden")
                aMRBRA(nCnt).cReceteNo = SQLReadString(oReader, "receteno")
                aMRBRA(nCnt).nAdet = SQLReadDouble(oReader, "adet")
                ' sipariş RBA
                ReDim Preserve aMRBA(nCnt)
                aMRBA(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aMRBA(nCnt).cRenk = SQLReadString(oReader, "renk")
                aMRBA(nCnt).cBeden = SQLReadString(oReader, "beden")
                aMRBA(nCnt).cReceteNo = SQLReadString(oReader, "receteno")
                aMRBA(nCnt).nAdet = SQLReadDouble(oReader, "adet")

                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            If nCnt = 0 Then
                ' sipariş adetleri girilmemiş
                ConnYage.Close()
                Exit Function
            End If

            If lKesileneGore Then
                ' kesim tamalandıysa kesilene göre adetleri al
                cSQL = "select departman, okbilgisi " + _
                        " from uretpllines " + _
                        " where uretimtakipno in (select uretimtakipno " + _
                                                " from " + cSipModelTableName + _
                                                " where malzemetakipno = '" + cMTF.Trim + "') " + _
                        " and departman like '%KESIM%' "

                If CheckExistsConnected(cSQL, ConnYage) Then

                    cSQL = "select departman, okbilgisi " + _
                            " from uretpllines " + _
                            " where uretimtakipno in (select uretimtakipno " + _
                                                    " from " + cSipModelTableName + _
                                                    " where malzemetakipno = '" + cMTF.Trim + "') " + _
                            " and departman like '%KESIM%' " + _
                            " and (okbilgisi is null or okbilgisi = '' or okbilgisi = 'H') "

                    If Not CheckExistsConnected(cSQL, ConnYage) Then
                        ' bütün kesimler kapanmıştır
                        ' model, renk ve bedene göre kesilen adet tablosu
                        nUretimToleransi = 0

                        cSQL = "select max(uretimtoleransi) " + _
                                " from modeluretim " + _
                                " where modelno in (select modelno " + _
                                                    " from " + cSipModelTableName + _
                                                    " where malzemetakipno = '" + cMTF.Trim + "') "

                        nUretimToleransi = SQLGetDoubleConnected(cSQL, ConnYage)

                        For nCnt = 0 To UBound(aMRBA)

                            nReceteAdet = aMRBA(nCnt).nAdet

                            cSQL = "select adet = sum(coalesce(x.adet,0)) " + _
                                    " from uretharrba x, uretharfis y " + _
                                    " where x.uretfisno = y.uretfisno " + _
                                    " and x.uretimtakipno in (select uretimtakipno " + _
                                                            " from " + cSipModelTableName + _
                                                            " where malzemetakipno = '" + cMTF.Trim + "') " + _
                                    " and y.cikisdept like '%KESIM%' " + _
                                    " and x.modelno = '" + aMRBA(nCnt).cModelNo + "' " + _
                                    " and x.renk = '" + aMRBA(nCnt).cRenk + "' " + _
                                    " and x.beden = '" + aMRBA(nCnt).cBeden + "' "

                            nKesilenAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                            cSQL = "select adet = sum(coalesce(adet,0)) " + _
                                    " from " + cSipModelTableName + _
                                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                                    " and adet is not null " + _
                                    " and adet <> 0 " + _
                                    " and modelno = '" + aMRBA(nCnt).cModelNo + "' " + _
                                    " and renk = '" + aMRBA(nCnt).cRenk + "' " + _
                                    " and beden = '" + aMRBA(nCnt).cBeden + "' "

                            nSiparisAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                            ReDim Preserve aMRBRA(nCnt)
                            aMRBRA(nCnt).cModelNo = aMRBA(nCnt).cModelNo
                            aMRBRA(nCnt).cRenk = aMRBA(nCnt).cRenk
                            aMRBRA(nCnt).cBeden = aMRBA(nCnt).cBeden
                            aMRBRA(nCnt).cReceteNo = aMRBA(nCnt).cReceteNo
                            aMRBRA(nCnt).nAdet = 0

                            If nKesilenAdet > 0 Then
                                ' sadece kesilen adet varsa ihtiyaç hesaplanır
                                If nReceteAdet > 0 Then
                                    aMRBRA(nCnt).nAdet = nKesilenAdet * nSiparisAdet / nReceteAdet
                                End If
                            End If

                            If aMRBRA(nCnt).nAdet = 0 Then
                                aMRBRA(nCnt).nAdet = nSiparisAdet
                            End If
                        Next
                    End If
                ElseIf lKesIsemrineGore Then
                    ' kesim işemrilerinin HEPSI onaylandıysa
                    ' en az 1 adet onaylı kesim işemri varsa
                    cSQL = "select isemrino " + _
                            " from uretimisemri " + _
                            " where uretimtakipno in (select uretimtakipno " + _
                                                    " from " + cSipModelTableName + _
                                                    " where malzemetakipno = '" + cMTF.Trim + "') " + _
                            " and departman like '%KESIM%' " + _
                            " and onay = 'E' "

                    If CheckExistsConnected(cSQL, ConnYage) Then
                        ' onaysız bir kesim işemri yoksa - bütün kesim işemrileri onaylıysa
                        cSQL = "select isemrino " + _
                                " from uretimisemri " + _
                                " where uretimtakipno in (select uretimtakipno " + _
                                                        " from " + cSipModelTableName + _
                                                        " where malzemetakipno = '" + cMTF.Trim + "') " + _
                                " and departman like '%KESIM%' " + _
                                " and (onay is null or onay = '' or onay = 'H') "

                        If Not CheckExistsConnected(cSQL, ConnYage) Then

                            cSQL = "select uretimtoleransi = max(uretimtoleransi) " + _
                                    " from modeluretim " + _
                                    " where modelno in (select modelno " + _
                                                        " from " + cSipModelTableName + _
                                                        " where malzemetakipno = '" + cMTF.Trim + "') "

                            nUretimToleransi = SQLGetDoubleConnected(cSQL, ConnYage)

                            For nCnt = 0 To UBound(aMRBA)

                                nReceteAdet = aMRBA(nCnt).nAdet

                                cSQL = "select adet = sum(coalesce(x.adet,0)) " + _
                                         " from uretimisrba x, uretimisemri y " + _
                                         " where x.isemrino = y.isemrino " + _
                                         " and x.uretimtakipno = y.uretimtakipno " + _
                                         " and x.uretimtakipno in (select uretimtakipno " + _
                                                                 " from " + cSipModelTableName + _
                                                                 " where malzemetakipno = '" + cMTF.Trim + "') " + _
                                         " and y.departman  like '%KESIM%' " + _
                                         " and x.modelno = '" + aMRBA(nCnt).cModelNo + "' " + _
                                         " and x.renk = '" + aMRBA(nCnt).cRenk + "' " + _
                                         " and x.beden = '" + aMRBA(nCnt).cBeden + "' "

                                nKesilenAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                                cSQL = "select adet = sum(coalesce(adet,0)) " + _
                                        " from " + cSipModelTableName + _
                                        " where malzemetakipno = '" + cMTF.Trim + "' " + _
                                        " and adet is not null " + _
                                        " and adet <> 0 " + _
                                        " and modelno = '" + aMRBA(nCnt).cModelNo + "' " + _
                                        " and renk = '" + aMRBA(nCnt).cRenk + "' " + _
                                        " and beden = '" + aMRBA(nCnt).cBeden + "' "

                                nSiparisAdet = SQLGetDoubleConnected(cSQL, ConnYage)

                                ReDim Preserve aMRBRA(nCnt)
                                aMRBRA(nCnt).cModelNo = aMRBA(nCnt).cModelNo
                                aMRBRA(nCnt).cRenk = aMRBA(nCnt).cRenk
                                aMRBRA(nCnt).cBeden = aMRBA(nCnt).cBeden
                                aMRBRA(nCnt).cReceteNo = aMRBA(nCnt).cReceteNo
                                aMRBRA(nCnt).nAdet = 0

                                If nKesilenAdet > 0 Then
                                    ' sadece kesilen adet varsa ihtiyaç hesaplanır
                                    If nReceteAdet > 0 Then
                                        aMRBRA(nCnt).nAdet = nKesilenAdet * nSiparisAdet / nReceteAdet
                                    End If
                                End If

                                If aMRBRA(nCnt).nAdet = 0 Then
                                    aMRBRA(nCnt).nAdet = nSiparisAdet
                                End If
                            Next
                        End If
                    End If
                End If
            End If

            For nCnt = 0 To UBound(aMRBRA)
                If aMRBRA(nCnt).nAdet > 0 Then
                    nCnt1 = 0

                    If aMRBRA(nCnt).cReceteNo = "" Then
                        cSQL = "select a.hammaddekodu, a.hammadderenk, a.hammaddebeden, b.maltakipesasi, a.temindept, " + _
                                " a.uretimdepartmani, a.kullanimmiktari, a.fire, a.hesaplama, a.miktarbirimi, " + _
                                " a.malzemetakipno " + _
                                " from modelhammadde a, stok b " + _
                                " where a.hammaddekodu = b.stokno " + _
                                " and a.modelno = '" + aMRBRA(nCnt).cModelNo + "' " + _
                                " and (a.modelrenk = 'HEPSI' or a.modelrenk = '" + aMRBRA(nCnt).cRenk + "') " + _
                                " and (a.modelbeden = 'HEPSI' or a.modelbeden = '" + aMRBRA(nCnt).cBeden + "') "
                    Else
                        cSQL = "select a.hammaddekodu, a.hammadderenk, a.hammaddebeden, b.maltakipesasi, a.temindept, " + _
                                " a.uretimdepartmani, a.kullanimmiktari, a.fire, a.hesaplama, a.miktarbirimi, " + _
                                " a.malzemetakipno " + _
                                " from modelhammadde2 a, stok b  " + _
                                " where a.hammaddekodu = b.stokno " + _
                                " and a.modelno = '" + aMRBRA(nCnt).cModelNo + "' " + _
                                " and (a.modelrenk = 'HEPSI' or a.modelrenk = '" + aMRBRA(nCnt).cRenk + "') " + _
                                " and (a.modelbeden = 'HEPSI' or a.modelbeden = '" + aMRBRA(nCnt).cBeden + "') " + _
                                " and a.receteno = '" + aMRBRA(nCnt).cReceteNo + "' "
                    End If

                    If CheckExistsConnected(cSQL, ConnYage) Then

                        oReader = GetSQLReader(cSQL, ConnYage)

                        Do While oReader.Read

                            ReDim Preserve aMTF(nCnt1)

                            aMTF(nCnt1).cStokNo = SQLReadString(oReader, "hammaddekodu")
                            aMTF(nCnt1).cRenk = IIf(SQLReadString(oReader, "hammadderenk") = "HEPSI", aMRBRA(nCnt).cRenk, SQLReadString(oReader, "hammadderenk")).ToString
                            aMTF(nCnt1).cBeden = IIf(SQLReadString(oReader, "hammaddebeden") = "HEPSI", aMRBRA(nCnt).cBeden, SQLReadString(oReader, "hammaddebeden")).ToString

                            Select Case SQLReadString(oReader, "maltakipesasi")
                                Case "1"
                                    aMTF(nCnt1).cRenk = "HEPSI"
                                    aMTF(nCnt1).cBeden = "HEPSI"
                                Case "2"
                                    aMTF(nCnt1).cBeden = "HEPSI"
                                Case "3"
                                    aMTF(nCnt1).cRenk = "HEPSI"
                            End Select

                            aMTF(nCnt1).cMDept = SQLReadString(oReader, "temindept")
                            aMTF(nCnt1).cUDept = SQLReadString(oReader, "uretimdepartmani")
                            aMTF(nCnt1).nMiktar = SQLReadDouble(oReader, "kullanimmiktari")
                            aMTF(nCnt1).nFire = SQLReadDouble(oReader, "fire")
                            aMTF(nCnt1).cHesaplama = SQLReadString(oReader, "hesaplama")
                            aMTF(nCnt1).cBirim = SQLReadString(oReader, "miktarbirimi")

                            nFire = 1
                            If lKesileneGore Or lKesIsemrineGore Then
                                If aMTF(nCnt1).nFire >= nUretimToleransi Then
                                    nFire = aMTF(nCnt1).nFire - nUretimToleransi
                                Else
                                    nFire = 0
                                End If
                            Else
                                nFire = aMTF(nCnt1).nFire
                            End If

                            Select Case SQLReadString(oReader, "hesaplama")
                                Case "1" : nMalzemeFireFactor = (1.0# + (nFire / 100.0#))       ' Yukardan asagıya
                                Case "2" : nMalzemeFireFactor = 1 / (1.0# - (nFire / 100.0#)) ' ' asagidan yukari hesaplansin
                                Case Else : nMalzemeFireFactor = (1.0# + (nFire / 100.0#))    ' Yukardan asagıya
                            End Select

                            nFire = nMalzemeFireFactor
                            ' fire, fire çarpanına dönüştüğünde 0 olamaz, en az 1 olabilir
                            If nFire = 0 Then
                                nFire = 1
                            End If

                            If SQLReadString(oReader, "malzemetakipno") = "" Then
                                aMTF(nCnt1).nMiktar = aMTF(nCnt1).nMiktar * nFire * aMRBRA(nCnt).nAdet
                            Else
                                ' eğer satıra MTF yazılmışsa toplam miktardır adet ile çarpılmaz
                                aMTF(nCnt1).nMiktar = aMTF(nCnt1).nMiktar * nFire
                            End If

                            nCnt1 = nCnt1 + 1
                        Loop
                        oReader.Close()
                        oReader = Nothing

                        For nCnt1 = 0 To UBound(aMTF)
                            UpdateMTKFisLines(ConnYage, cMTF, aMTF(nCnt1).cStokNo, aMTF(nCnt1).cRenk, aMTF(nCnt1).cBeden, aMTF(nCnt1).nMiktar, aMTF(nCnt1).cBirim, aMTF(nCnt1).cUDept, aMTF(nCnt1).cMDept)
                        Next
                    End If
                End If
            Next

            CloseConn(ConnYage)
            ' Post Process
            MTKPostProcess(cMTF.Trim, cSipModelTableName)
            ' işemri kontrol
            G_IsemriDeptKontrol(cMTF.Trim)
            ' Calc
            MTKLinesTopl(cMTF.Trim)
            ' The END
            MTKFastGenerate = 1
            JustForLog("MTF bakimi bitti : " + cMTF.Trim)

        Catch Err As Exception
            ErrDisp(Err.Message, "MTKFastGenerate", cSQL)
        End Try
    End Function

    Private Sub MTKPostProcess(cMTF As String, cSipModelTableName As String)

        Dim nYuvarla As Double = 0
        Dim cSQL As String = ""
        Dim cBirim As String = ""
        Dim cTamSayiYap As String = ""
        Dim cSipList As String = ""
        Dim nCnt As Integer = 0
        Dim oReader As SqlDataReader
        Dim aMTF() As MTF = Nothing
        Dim ConnYage As SqlConnection

        Try
            If cMTF.Trim = "" Then Exit Sub

            ConnYage = OpenConn()

            ' hesaplanan temel ihtiyaç hesaplananihtiyac hanesinde saklanır
            cSQL = "update mtkfislines " + _
                    " set hesaplananihtiyac = coalesce(ihtiyac,0) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set musteriihtiyac = (select sum(coalesce(b.miktar,0)) " + _
                                        " from mtkeklefis a, mtkeklefislines b " + _
                                        " where a.mtkeklefisno = b.mtkeklefisno " + _
                                        " and a.malzemetakipno = mtkfislines.malzemetakipno " + _
                                        " and b.stokno = mtkfislines.stokno " + _
                                        " and b.renk = mtkfislines.renk " + _
                                        " and b.beden = mtkfislines.beden " + _
                                        " and b.departman = mtkfislines.departman) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set ihtiyac = coalesce(musteriihtiyac,0) + coalesce(ihtiyatiihtiyac,0) + coalesce(hesaplananihtiyac,0) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            nCnt = 0

            cSQL = "select sirano, stokno " + _
                    " from mtkfislines " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aMTF(nCnt)
                aMTF(nCnt).nSiraNo = SQLReadDouble(oReader, "sirano")
                aMTF(nCnt).cStokNo = SQLReadString(oReader, "stokno")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aMTF)
                nYuvarla = 0
                cBirim = ""

                cSQL = "select yuvarla, birim1 " + _
                        " from stok " + _
                        " where stokno = '" + aMTF(nCnt).cStokNo + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nYuvarla = SQLReadDouble(oReader, "yuvarla")
                    cBirim = SQLReadString(oReader, "birim1")
                End If
                oReader.Close()
                oReader = Nothing

                cTamSayiYap = "H"
                If cBirim <> "" Then
                    cSQL = "select yuvarla " + _
                            " from birim " + _
                            " where birim = '" + cBirim + "' "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        cTamSayiYap = SQLReadString(oReader, "yuvarla")
                    End If
                    oReader.Close()
                End If

                If cTamSayiYap = "E" Then
                    cSQL = "update mtkfislines " + _
                            " set ihtiyac = ceiling(ihtiyac) " + _
                            " where sirano = " + SQLWriteDecimal(aMTF(nCnt).nSiraNo)

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    cSQL = "update mtkfislines " + _
                            " set ihtiyac = round(ihtiyac," + SQLWriteDecimal(nYuvarla) + ") " + _
                            " where sirano = " + SQLWriteDecimal(aMTF(nCnt).nSiraNo)

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            cSQL = "select notlar " + _
                    " from mtkfis " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            cSipList = SQLGetStringConnected(cSQL, ConnYage)

            cSQL = "select distinct siparisno " + _
                    " from " + cSipModelTableName + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and adet is not null " + _
                    " and adet <> 0 "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                If cSipList = "" Then
                    cSipList = SQLReadString(oReader, "siparisno")
                Else
                    If InStr(cSipList, SQLReadString(oReader, "siparisno")) = 0 Then
                        cSipList = cSipList + "," + SQLReadString(oReader, "siparisno")
                    End If
                End If
            Loop
            oReader.Close()
            oReader = Nothing

            cSQL = "update mtkfis " + _
                    " set notlar = '" + cSipList.Trim + "' " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)
            ' Malzeme zaman ve bütçe ön planlaması için
            ' Stok kartlarından temin süresi öndeğerleri alınır

            cSQL = "update mtkfislines " + _
                    " set teminsuresi = (select top 1 gelisgun " + _
                                        " from stok " + _
                                        " where stokno = mtkfislines.stokno) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and (teminsuresi is null or teminsuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set teminsuresi = (select top 1 a.gelisgun " + _
                                        " from stoktipi a, stok b " + _
                                        " where a.kod = b.stoktipi " + _
                                        " and b.stokno = mtkfislines.stokno) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and (teminsuresi is null or teminsuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set avanssuresi = (select top 1 avanssuresi " + _
                                        " from stok " + _
                                        " where stokno = mtkfislines.stokno) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and (teminsuresi is null or teminsuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set avanssuresi = (select top 1 a.avanssuresi " + _
                                        " from stoktipi a, stok b " + _
                                        " where a.kod = b.stoktipi " + _
                                        " and b.stokno = mtkfislines.stokno) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and (avanssuresi is null or avanssuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set oktipi = (select top 1 a.oktipi " + _
                                    " from stoktipi a, stok b " + _
                                    " where a.kod = b.stoktipi " + _
                                    " and b.stokno = mtkfislines.stokno) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and (oktipi is null or oktipi = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' Stokno, renk, beden aynı çıkış üretim departmanı farklıysa
            ' gelen malzemeyi bölüştürmek için
            ' ihtiyaç miktarıyla doğru orantılı bir katsayı çarpanı kullanılır
            cSQL = "update MtkFisLines " + _
                     " Set KatSayi = 1 " + _
                     " where malzemetakipno = '" + cMTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update MtkFisLines " + _
                    " set katsayi = coalesce(ihtiyac,0) / (select sum(coalesce(a.ihtiyac,0)) " + _
                                                            " from mtkfislines a " + _
                                                            " Where a.malzemetakipno = mtkfislines.malzemetakipno " + _
                                                            " and a.stokno = mtkfislines.stokno " + _
                                                            " and a.renk = mtkfislines.renk " + _
                                                            " and a.beden = mtkfislines.beden " + _
                                                            " and a.temindept = mtkfislines.temindept) " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and ihtiyac is not null " + _
                    " and ihtiyac <> 0 "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update MtkFisLines " + _
                    " Set katsayi = 1 " + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and (katsayi = 0 or katsayi is null) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from mtkfislines " + _
                   " where malzemetakipno = '" + cMTF.Trim + "' " + _
                   " and (ihtiyac = 0 or ihtiyac is null) " + _
                   " and (musteriihtiyac = 0 or musteriihtiyac is null) " + _
                   " and (ihtiyatiihtiyac = 0 or ihtiyatiihtiyac is null) " + _
                   " and (hesaplananihtiyac = 0 or hesaplananihtiyac is null) " + _
                   " and (isemriicingiden = 0 or isemriicingiden is null) " + _
                   " and (isemriharicigiden = 0 or isemriharicigiden is null) " + _
                   " and (uretimicincikis = 0 or uretimicincikis is null) " + _
                   " and (uretimdeniade = 0 or uretimdeniade is null) " + _
                   " and (isemriverilen = 0 or isemriverilen is null) " + _
                   " and (isemriicingelen = 0 or isemriicingelen is null) " + _
                   " and (isemriharicigelen = 0 or isemriharicigelen is null) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "MTKPostProcess", cSQL)
        End Try
    End Sub

    Private Sub UpdateMTKFisLines(ConnYage As SqlConnection, cMTF As String, cStokno As String, cHRenk As String, cHBeden As String, nMiktar As Double, _
                                  Optional cBirim As String = "", Optional cUDept As String = "", Optional cMDept As String = "", _
                                  Optional cTable As String = "", Optional lMalTakipEsasi As Boolean = True)

        Dim cSQL As String = ""
        Dim oReader As SqlDataReader
        Dim nMiktar2 As Double = 0
        Dim cTakipEsasi As String = ""
        Dim aMTF() As MTF = Nothing
        Dim nCnt1 As Integer = 0
        Dim nFire As Double = 0
        Dim nMalzemeFireFactor As Double = 1

        Try
            If cTable = "" Then
                cTable = "mtkfislines"
            End If

            cSQL = "select maltakipesasi, paratakipesasi, temindepartmani, uretimdepartmani, birim1 " + _
                    " from stok " + _
                    " where stokno = '" + cStokno.Trim + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                If lMalTakipEsasi Then
                    cTakipEsasi = SQLReadString(oReader, "maltakipesasi")
                Else
                    cTakipEsasi = SQLReadString(oReader, "paratakipesasi")
                End If

                Select Case cTakipEsasi
                    Case "1"
                        cHRenk = "HEPSI"
                        cHBeden = "HEPSI"
                    Case "2"
                        cHBeden = "HEPSI"
                    Case "3"
                        cHRenk = "HEPSI"
                End Select
                If cMDept.Trim = "" Then
                    cMDept = SQLReadString(oReader, "temindepartmani")
                End If
                If cUDept.Trim = "" Then
                    cUDept = SQLReadString(oReader, "uretimdepartmani")
                End If
                If cBirim.Trim = "" Then
                    cBirim = SQLReadString(oReader, "birim1")
                End If
            End If
            oReader.Close()
            oReader = Nothing

            cSQL = "select malzemetakipno " + _
                    " from " + cTable.Trim + _
                    " where malzemetakipno = '" + cMTF.Trim + "' " + _
                    " and stokno = '" + cStokno.Trim + "' " + _
                    " and renk = '" + cHRenk.Trim + "' " + _
                    " and beden = '" + cHBeden.Trim + "' " + _
                    " and temindept = '" + cMDept.Trim + "' " + _
                    " and departman = '" + cUDept.Trim + "' "

            If CheckExistsConnected(cSQL, ConnYage) Then
                cSQL = "update " + cTable.Trim + _
                        " set ihtiyac = coalesce(ihtiyac,0) + " + SQLWriteDecimal(nMiktar) + _
                        " where malzemetakipno = '" + cMTF.Trim + "' " + _
                        " and stokno = '" + cStokno.Trim + "' " + _
                        " and renk = '" + cHRenk.Trim + "' " + _
                        " and beden = '" + cHBeden.Trim + "' " + _
                        " and temindept = '" + cMDept.Trim + "' " + _
                        " and departman = '" + cUDept.Trim + "' " + _
                        " and (kilitle is null or kilitle = 'H' or kilitle = '') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "insert into " + cTable.Trim + " (malzemetakipno, stokno, renk, beden, ihtiyac, birim, departman, temindept, " + _
                        " isemriverilen, isemriicingelen, isemriharicigelen, isemriicingiden, isemriharicigiden, uretimicincikis, uretimdeniade, " + _
                        " hedefmlzbirimfiyati, hedefiscilikbirimfiyati, uretimecikisfiyati, musteriihtiyac, ihtiyatiihtiyac, hesaplananihtiyac) " + _
                        " values ('" + cMTF.Trim + "', " + _
                        " '" + cStokno.Trim + "', " + _
                        " '" + cHRenk.Trim + "', " + _
                        " '" + cHBeden.Trim + "', " + _
                        SQLWriteDecimal(nMiktar) + ", " + _
                        " '" + cBirim.Trim + "', " + _
                        " '" + cUDept.Trim + "', " + _
                        " '" + cMDept.Trim + "', " + _
                        " 0,0,0,0,0,0,0,0,0,0,0,0,0 )"

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            ' Recursive olarak hammadde ağacını hesapla
            cSQL = "select a.hammaddekodu, a.hamrenk, a.hambeden, a.fire, a.miktar, a.hesaplama, b.maltakipesasi, b.birim1 " + _
                    " from strecete a, stok b " + _
                    " where a.hammaddekodu = b.stokno " + _
                    " and a.anahammadde = '" + cStokno.Trim + "' " + _
                    " and a.hammaddekodu <> '" + cStokno.Trim + "' " + _
                    " and (a.anarenk = '" + cHRenk.Trim + "' or a.anarenk = 'HEPSI') " + _
                    " and (a.anabeden = '" + cHBeden.Trim + "' or a.anabeden = 'HEPSI') "

            If CheckExistsConnected(cSQL, ConnYage) Then

                oReader = GetSQLReader(cSQL, ConnYage)

                Do While oReader.Read

                    ReDim Preserve aMTF(nCnt1)

                    aMTF(nCnt1).cStokNo = SQLReadString(oReader, "hammaddekodu")
                    aMTF(nCnt1).cRenk = IIf(SQLReadString(oReader, "hammadderenk") = "HEPSI", cHRenk.Trim, SQLReadString(oReader, "hammadderenk")).ToString
                    aMTF(nCnt1).cBeden = IIf(SQLReadString(oReader, "hammaddebeden") = "HEPSI", cHBeden.Trim, SQLReadString(oReader, "hammaddebeden")).ToString

                    Select Case SQLReadString(oReader, "maltakipesasi")
                        Case "1"
                            aMTF(nCnt1).cRenk = "HEPSI"
                            aMTF(nCnt1).cBeden = "HEPSI"
                        Case "2"
                            aMTF(nCnt1).cBeden = "HEPSI"
                        Case "3"
                            aMTF(nCnt1).cRenk = "HEPSI"
                    End Select

                    aMTF(nCnt1).nMiktar = SQLReadDouble(oReader, "miktar")
                    aMTF(nCnt1).nFire = SQLReadDouble(oReader, "fire")
                    aMTF(nCnt1).cHesaplama = SQLReadString(oReader, "hesaplama")
                    aMTF(nCnt1).cBirim = SQLReadString(oReader, "birim1")

                    nFire = aMTF(nCnt1).nFire

                    nMalzemeFireFactor = 1
                    Select Case SQLReadString(oReader, "hesaplama")
                        Case "1" : nMalzemeFireFactor = (1.0# + (nFire / 100.0#))       ' Yukardan asagıya
                        Case "2" : nMalzemeFireFactor = 1 / (1.0# - (nFire / 100.0#)) ' ' asagidan yukari hesaplansin
                        Case Else : nMalzemeFireFactor = (1.0# + (nFire / 100.0#))    ' Yukardan asagıya
                    End Select

                    nFire = nMalzemeFireFactor
                    ' fire, fire çarpanına dönüştüğünde 0 olamaz, en az 1 olabilir
                    If nFire = 0 Then
                        nFire = 1
                    End If

                    aMTF(nCnt1).nMiktar = aMTF(nCnt1).nMiktar * nFire * nMiktar

                    nCnt1 = nCnt1 + 1
                Loop
                oReader.Close()
                oReader = Nothing

                For nCnt1 = 0 To UBound(aMTF)
                    ' recurse and recalc
                    UpdateMTKFisLines(ConnYage, cMTF, _
                                     aMTF(nCnt1).cStokNo, _
                                     aMTF(nCnt1).cRenk, _
                                     aMTF(nCnt1).cBeden, _
                                     aMTF(nCnt1).nMiktar, , , , cTable, lMalTakipEsasi)
                Next
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "UpdateMTKFisLines", cSQL)
        End Try
    End Sub

    Private Sub MTKLinesTopl(Optional cMTFNo As String = "")
        ' cMTFNo boş ise bütün MTF lerin durumu yeniden hesaplanır
        Dim cSQL As String = ""
        Dim cSql1 As String = ""
        Dim cSql2 As String = ""
        Dim cSql3 As String = ""
        Dim cSql4 As String = ""
        Dim cTempView As String = ""
        Dim G_MTKUretCikNoDept As Boolean = False
        Dim lSatisDusmesin As Boolean = False
        Dim ConnYage As SqlConnection

        Try
            ConnYage = OpenConn()

            G_MTKUretCikNoDept = (GetSysParConnected("mtkuretciknodept", ConnYage) = "1")
            lSatisDusmesin = (GetSysParConnected("mtfsatisdusmesin", ConnYage) = "1")

            cSQL = "update mtkfislines " + _
                    " set " + _
                    " isemriverilen = 0, " + _
                    " isemriicingelen = 0, " + _
                    " isemriharicigelen = 0, " + _
                    " isemriicingiden = 0, " + _
                    " isemriharicigiden = 0, " + _
                    " uretimicincikis = 0, " + _
                    " uretimdeniade = 0 " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSql1 = " select stokno = coalesce(b.stokno,'') , " + _
                            " renk = coalesce(b.renk ,'') , " + _
                            " beden = coalesce(b.beden,'') , " + _
                            " malzemetakipkodu = coalesce(b.malzemetakipkodu ,'') , " + _
                            " isemrino = coalesce(b.isemrino,'') , " + _
                            " giris = sum(coalesce(b.netmiktar1,0)) , " + _
                            " cikis = 0 , " + _
                            " stokhareketkodu = coalesce(b.stokhareketkodu,'') , " + _
                            " departman = coalesce(a.departman,'') " + _
                " from stokfis a , stokfislines b " + _
                " Where a.stokfisno = b.stokfisno " + _
                            IIf(cMTFNo.Trim = "", " and b.malzemetakipkodu is not null and b.malzemetakipkodu <> '' ", " and b.malzemetakipkodu = '" + cMTFNo.Trim + "' ").ToString + _
                            " and (a.iptal is null or a.iptal = '' or a.iptal = 'H') " + _
                            " and a.stokfistipi in ('Giris','02 Satis Iade','03 Defolu iade') " + _
                            IIf(lSatisDusmesin, " and not b.stokhareketkodu in ('07 Satis Iade','07 Satis') ", "").ToString + _
                " group by b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.isemrino, b.stokhareketkodu, a.departman "

            cSql2 = " select stokno = coalesce(b.stokno,'') , " + _
                            " renk = coalesce(b.renk ,'') , " + _
                            " beden = coalesce(b.beden,'') , " + _
                            " malzemetakipkodu = coalesce(b.malzemetakipkodu ,'') , " + _
                            " isemrino = coalesce(b.isemrino,'') , " + _
                            " giris = 0 , " + _
                            " cikis = sum(coalesce(b.netmiktar1,0)) , " + _
                            " stokhareketkodu = coalesce(b.stokhareketkodu,'') , " + _
                            " departman = coalesce(a.departman,'') " + _
                " from stokfis a , stokfislines b " + _
                " Where a.stokfisno = b.stokfisno " + _
                            IIf(cMTFNo.Trim = "", " and b.malzemetakipkodu is not null and b.malzemetakipkodu <> '' ", " and b.malzemetakipkodu = '" + cMTFNo.Trim + "' ").ToString + _
                            " and (a.iptal is null or a.iptal = '' or a.iptal = 'H') " + _
                            " and a.stokfistipi in ('Cikis','01 Satis') " + _
                            IIf(lSatisDusmesin, " and not b.stokhareketkodu in ('07 Satis Iade','07 Satis') ", "").ToString + _
                " group by b.stokno, b.renk, b.beden, b.malzemetakipkodu, b.isemrino, b.stokhareketkodu, a.departman "

            cSql3 = " select stokno = coalesce(stokno,'') , " + _
                            " renk = coalesce(renk ,'') , " + _
                            " beden = coalesce(beden ,'') , " + _
                            " malzemetakipkodu = coalesce(hedefmalzemetakipno,'') , " + _
                            " isemrino = '' , " + _
                            " giris = sum(coalesce(netmiktar1,0)), " + _
                            " cikis = 0 , " + _
                            " stokhareketkodu = '90 Trans/Rezv Giris' , " + _
                            " departman = '' " + _
                " From StokTransfer " + _
                IIf(cMTFNo.Trim = "", " where hedefmalzemetakipno is not null and hedefmalzemetakipno <> '' ", " where hedefmalzemetakipno = '" + cMTFNo.Trim + "' ").ToString + _
                " group by stokno, renk, beden, hedefmalzemetakipno "

            cSql4 = " select stokno = coalesce(stokno,'') , " + _
                            " renk = coalesce(renk ,'') , " + _
                            " beden = coalesce(beden,'') , " + _
                            " malzemetakipkodu = coalesce(kaynakmalzemetakipno,'') , " + _
                            " isemrino = '' , " + _
                            " giris = 0 , " + _
                            " cikis = sum(coalesce(netmiktar1,0)) , " + _
                            " stokhareketkodu = '90 Trans/Rezv Cikis' , " + _
                            " departman = '' " + _
                " From StokTransfer " + _
                IIf(cMTFNo.Trim = "", " where kaynakmalzemetakipno is not null and kaynakmalzemetakipno <> '' ", " where kaynakmalzemetakipno = '" + cMTFNo.Trim + "' ").ToString + _
                " group by stokno, renk, beden, kaynakmalzemetakipno "

            cSQL = cSql1 + " Union All " + _
                   cSql2 + " Union All " + _
                   cSql3 + " Union All " + _
                   cSql4

            cTempView = CreateTempView(ConnYage, cSQL)

            ' hareket kodlarina gore update

            ' 01 Uretime Cikis
            ' üretim departmanı belli olduğu için MTF de ilgili üretim departmanlı satır için çıkış olur
            cSQL = "update mtkfislines " + _
                    " set uretimicincikis  = (select coalesce(sum(coalesce(cikis,0)),0) " + _
                                                " from " + cTempView + " b " + _
                                                " Where mtkfislines.stokno = b.stokno  " + _
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu" + _
                                                " and mtkfislines.renk = b.renk" + _
                                                " and mtkfislines.beden = b.beden " + _
                                                IIf(G_MTKUretCikNoDept, "", " and mtkfislines.departman = b.departman ").ToString + _
                                                " and " + G_uretimicincikis + " ) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' 01 uretimden iade
            ' üretim departmanı belli olduğu için MTF de ilgili üretim departmanlı satıra iade olur
            cSQL = "update mtkfislines " + _
                    " set uretimdeniade  = (select coalesce(sum(coalesce(giris,0)),0) " + _
                                                " from " + cTempView + " b " + _
                                                " Where mtkfislines.stokno = b.stokno  " + _
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu" + _
                                                " and mtkfislines.renk = b.renk" + _
                                                " and mtkfislines.beden = b.beden " + _
                                                IIf(G_MTKUretCikNoDept, "", " and mtkfislines.departman = b.departman ").ToString + _
                                                " and " + G_uretimdeniade + " ) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' İşEmriİçinGelen alanı tedarikten yapılmış NET girişi gösterir
            ' 02 Tedarikten Giris
            ' 04 Mlz Uretimden Giris
            ' 05 Diger Giris            -> aslinda isemri no girdikten sonra diger giris olmamali, 02 veya 04 yapilmali
            ' 06 Tamirden Giris
            cSQL = "update mtkfislines " + _
                    " set isemriicingelen  = (select coalesce(sum(coalesce(giris,0)),0) " + _
                                        " from " + cTempView + " b " + _
                                        " Where mtkfislines.stokno = b.stokno  " + _
                                        " and mtkfislines.malzemetakipno = b.malzemetakipkodu" + _
                                        " and mtkfislines.renk = b.renk" + _
                                        " and mtkfislines.beden= b.beden" + _
                                        " and isemrino is not null " + _
                                        " and isemrino <> '' " + _
                                        " and " + G_isemriicinGelenGiris + " ) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' 02 Tedarikten iade
            ' 04 Mlz Uretime iade
            ' 05 Diger Cikis
            ' 06 Tamirden Giris
            cSQL = "update mtkfislines " + _
                    " set isemriicingelen  = coalesce(isemriicingelen,0) - (select coalesce(sum(coalesce(cikis,0)),0) " + _
                                                                            " from " + cTempView + " b " + _
                                                                            " Where mtkfislines.stokno = b.stokno  " + _
                                                                            " and mtkfislines.malzemetakipno = b.malzemetakipkodu" + _
                                                                            " and mtkfislines.renk = b.renk" + _
                                                                            " and mtkfislines.beden= b.beden" + _
                                                                            " and isemrino is not null " + _
                                                                            " and isemrino <> '' " + _
                                                                            " and " + G_isemriicinGelenCikis + " ) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' İşEmriHariciGelen alanı Karşılanan denkleminde kullanılan dengeleyici miktar olarak kullanılıyor
            ' Karşılanan = İşEmriİçinGelen + İşEmriHariciGelen
            ' 02 Tedarikten Giris       -> aslinda isemrino dolu olmali
            ' 04 Mlz Uretimden Giris    -> aslinda isemrino dolu olmali
            ' 05 Diger Giris            -> isemri no olmadan yapilan temel hareket
            ' 08 SAYIM GIRIS            -> sayimfarki giris fisi uzerinde isemrino bulunmuyor
            ' 55 Kontrol Oncesi Giris   -> kontrol tezgahina kumasi takmadan parca almak ile ilgili
            ' 77 Top Bolme Giris        -> otomatik uretilen top bolme hareketi bolunmus parcalarinin giris kodu, isemrino ya bagli olamaz
            ' 77 Aksesuar Bolme Giris   -> otomatik uretilen aksesuar bolme hareketi bolunmus aksesuarin giris kodu, isemrino ya bagli olamaz
            ' 90 Trans/Rezv Giris       -> transfer fisi uzerinde isemrino bulunmuyor
            'cSQL = "update mtkfislines " + _
            '        " set isemriharicigelen  = (select coalesce(sum(coalesce(giris,0)),0) " + _
            '                                " from " + cTempView + " b " + _
            '                                " Where mtkfislines.stokno = b.stokno  " + _
            '                                " and mtkfislines.renk = b.renk" + _
            '                                " and mtkfislines.beden= b.beden" + _
            '                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu" + _
            '                                " and (isemrino is null or isemrino = '') " + _
            '                                " and " + G_isemriHariciGelenGiris + " ) " + _
            '        IIf(cMTFNo = "", "", " where malzemetakipno = '" + cMTFNo + "' ")

            ' MTF ye yapılmış bütün rezervasyon girişleri
            cSQL = "update mtkfislines " + _
                    " set isemriharicigelen  = (select coalesce(sum(coalesce(giris,0)),0) " + _
                                            " from " + cTempView + " b " + _
                                            " Where mtkfislines.stokno = b.stokno  " + _
                                            " and mtkfislines.renk = b.renk " + _
                                            " and mtkfislines.beden= b.beden " + _
                                            " and mtkfislines.malzemetakipno = b.malzemetakipkodu) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' 02 Tedarikten iade       -> aslinda isemrino dolu olmali
            ' 04 Mlz Uretime iade      -> aslinda isemrino dolu olmali
            ' 05 Diger Cikis
            ' 08 SAYIM CIKIS
            ' 55 Kontrol Oncesi Cikis
            ' 77 Top Bolme Cikis
            ' 77 Aksesuar Bolme Cikis
            ' 90 Trans/Rezv Cikis
            'cSQL = "update mtkfislines " + _
            '        " set isemriharicigelen  = coalesce(isemriharicigelen,0) - (select coalesce(sum(coalesce(cikis,0)),0) " + _
            '                                                                    " from " + cTempView + " b " + _
            '                                                                    " Where mtkfislines.stokno = b.stokno  " + _
            '                                                                    " and mtkfislines.renk = b.renk" + _
            '                                                                    " and mtkfislines.beden= b.beden" + _
            '                                                                    " and mtkfislines.malzemetakipno = b.malzemetakipkodu" + _
            '                                                                    " and (isemrino is null or isemrino = '') " + _
            '                                                                    " and " + G_isemriHariciGelenCikis + " ) " + _
            '        IIf(cMTFNo = "", "", " where malzemetakipno = '" + cMTFNo + "' ")

            ' MTF den yapılmış bütün rezervasyon çıkışları
            cSQL = "update mtkfislines " + _
                    " set isemriharicigelen  = coalesce(isemriharicigelen,0) - (select coalesce(sum(coalesce(cikis,0)),0) " + _
                                                                                " from " + cTempView + " b " + _
                                                                                " Where mtkfislines.stokno = b.stokno  " + _
                                                                                " and mtkfislines.renk = b.renk " + _
                                                                                " and mtkfislines.beden= b.beden " + _
                                                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' Sonuç denklemde, Karşılanan = İşEmriİçinGelen + İşEmriHariciGelen
            ' Yani, Karşılanan = Üretime Net Çıkan + Elimizdeki Net Rezerve Miktar
            ' İşEmriHariciGelen = (ÜretimeÇıkan NET miktar)+ (MTF ye yapılmış NET rezervasyonlar (elimizdeki rezerve malzeme)) - (İşEmriİçinGelen NET miktar)
            cSQL = "update mtkfislines " + _
                    " set isemriharicigelen  = (coalesce(uretimicincikis,0) - coalesce(uretimdeniade,0)) " + _
                                                " + coalesce(isemriharicigelen,0) " + _
                                                " - coalesce(isemriicingelen,0) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set isemriicingiden  = (select coalesce(sum(coalesce(cikis,0)),0) " + _
                                                " from " + cTempView + " b " + _
                                                " Where mtkfislines.stokno = b.stokno " + _
                                                " and mtkfislines.renk = b.renk " + _
                                                " and mtkfislines.beden= b.beden " + _
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu " + _
                                                " and isemrino is not null " + _
                                                " and isemrino <> '' ) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines " + _
                    " set isemriharicigiden  = (select coalesce(sum(coalesce(cikis,0)),0) " + _
                                                " from " + cTempView + " b " + _
                                                " Where mtkfislines.stokno = b.stokno " + _
                                                " and mtkfislines.renk = b.renk " + _
                                                " and mtkfislines.beden= b.beden " + _
                                                " and mtkfislines.malzemetakipno = b.malzemetakipkodu " + _
                                                " and (isemrino is null or isemrino = '') ) " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)
            ' üretim departmanına göre işemri verilen adet
            ' her işemri satırında üretim departmanı olması gerekiyor
            cSQL = "update mtkfislines " + _
                    " set isemriverilen = (select coalesce(sum(coalesce(miktar1,0)),0) " + _
                                            " from isemrilines b " + _
                                            " Where mtkfislines.stokno = b.stokno  " + _
                                            " and mtkfislines.malzemetakipno = b.malzemetakipno " + _
                                            " and mtkfislines.renk = b.renk " + _
                                            " and mtkfislines.beden = b.beden " + _
                                            " and coalesce(mtkfislines.departman,'') = coalesce(b.departman,'') " + _
                                            " and b.isemrino is not null " + _
                                            " and b.isemrino <> '') " + _
                    IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' katsayılar doğru olarak hesaplanmış olmalıdır
            If G_MTKUretCikNoDept Then
                cSQL = "update mtkfislines " + _
                        " set isemriicingelen = coalesce(isemriicingelen,0) * coalesce(katsayi,0), " + _
                        " isemriharicigelen = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " + _
                        " isemriicingiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " + _
                        " isemriharicigiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " + _
                        " uretimicincikis = coalesce(uretimicincikis,0) * coalesce(katsayi,0), " + _
                        " uretimdeniade = coalesce(uretimdeniade,0) * coalesce(katsayi,0) " + _
                        IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString
            Else
                cSQL = "update mtkfislines " + _
                        " set isemriicingelen = coalesce(isemriicingelen,0) * coalesce(katsayi,0), " + _
                        " isemriharicigelen = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " + _
                        " isemriicingiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0), " + _
                        " isemriharicigiden = coalesce(isemriharicigelen,0) * coalesce(katsayi,0) " + _
                        IIf(cMTFNo.Trim = "", "", " where malzemetakipno = '" + cMTFNo.Trim + "' ").ToString
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' eğer yeni duruma göre ihtiyaç > karşılanan ise ilgili satırları aç
            cSQL = "update mtkfislines " + _
                    " set kapandi = 'H' " + _
                    " where coalesce(ihtiyac,0) > coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0) " + _
                    " and coalesce(ihtiyac,0) > coalesce(uretimicincikis,0) - coalesce(uretimdeniade,0) " + _
                    " and kapandi in ('E','e') " + _
                    IIf(cMTFNo.Trim = "", "", " and malzemetakipno = '" + cMTFNo.Trim + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            DropView(cTempView, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "MTKLinesTopl", cSQL)
        End Try
    End Sub

    Public Sub G_IsemriDeptKontrol(Optional cMTF As String = "")
        ' işemrinde üretime çıkış departmanı yoksa
        ' ilgili MTF deki en büyük ihtiyaç miktarına sahip üretim departmanını işemrindeki boş üretim departmanına atıyoruz
        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection

        Try
            ConnYage = OpenConn()
            ' departmanı BOŞ atılmış satırları tamamlar
            cSQL = "update isemrilines " + _
                    " set departman = (select top 1 departman " + _
                                        " from mtkfislines " + _
                                        " where malzemetakipno = isemrilines.malzemetakipno " + _
                                        " and stokno = isemrilines.stokno " + _
                                        " and renk = isemrilines.renk " + _
                                        " and beden = isemrilines.beden " + _
                                        " order by ihtiyac desc) " + _
                    " where (departman is null or departman = '') "

            If cMTF.Trim = "" Then
                cSQL = cSQL + _
                    " and malzemetakipno is not null " + _
                    " and malzemetakipno <> '' "
            Else
                cSQL = cSQL + _
                    " and malzemetakipno = '" + cMTF.Trim + "' "
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' departmanı yanlış girilmiş satırları toparlar
            cSQL = "update isemrilines " + _
                    " set departman = (select top 1 departman " + _
                                        " from mtkfislines " + _
                                        " where malzemetakipno = isemrilines.malzemetakipno " + _
                                        " and stokno = isemrilines.stokno " + _
                                        " and renk = isemrilines.renk " + _
                                        " and beden = isemrilines.beden " + _
                                        " order by ihtiyac desc) " + _
                    " where departman is not null " + _
                    " and departman <> '' " + _
                    " and not exists (select malzemetakipno " + _
                                    " from mtkfislines " + _
                                    " where StokNo = isemrilines.StokNo " + _
                                    " and renk = isemrilines.renk " + _
                                    " and beden = isemrilines.beden " + _
                                    " and coalesce(departman,'') = coalesce(isemrilines.departman,'')) "
            If cMTF.Trim = "" Then
                cSQL = cSQL + _
                    " and malzemetakipno is not null " + _
                    " and malzemetakipno <> '' "
            Else
                cSQL = cSQL + _
                    " and malzemetakipno = '" + cMTF.Trim + "' "
            End If

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "G_IsemriDeptKontrol", cSQL)
        End Try
    End Sub
End Module

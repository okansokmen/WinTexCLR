Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module UTF

    Private Structure oRBAA
        Dim cRenk As String
        Dim cBeden As String
        Dim nAdet As Double
        Dim nULineNo As Double
    End Structure

    Private Structure oIsEmri
        Dim cUTF As String
        Dim cModelNo As String
        Dim cDepartman As String
        Dim cParca As String
        Dim cBedenSeti As String
        Dim cCikisTakipEsasi As String
        Dim dBaslamaTarihi As Date
        Dim dBitisTarihi As Date
        Dim nFiyat As Double
        Dim cDoviz As String
        Dim nSira As Double
        Dim cPLFirma As String
        Dim cIEFirma1 As String
        Dim cIEEleman1 As String
        Dim cIEFirma2 As String
        Dim cIEEleman2 As String
    End Structure

    Private Structure oRota
        Dim cDepartman As String
        Dim cParca As String
        Dim nTolerans As Double
        Dim cGidenDepartman As String
        Dim nSira As Double
        Dim cGirisTakipEsasi As String
        Dim cCikisTakipEsasi As String
    End Structure

    Private Structure oUretPlLines
        Dim cUTF As String
        Dim cModelNo As String
        Dim cBedenSeti As String
        Dim cDepartman As String
        Dim nUretimToleransi As Double
        Dim cGirisTakipEsasi As String
        Dim cCikisTakipEsasi As String
        Dim cParca As String
        Dim nSira As Double
        Dim cGirisDepartmani As String
        Dim cCikisDepartmani As String
        Dim cGirisParcasi As String
        Dim cFirma As String
        Dim cYikamaKodu As String
        Dim nIscilikFiyat As Double
        Dim cIscilikDoviz As String
        Dim nUPLSiraNo As Double
    End Structure

    Private Structure oUretPlRBA
        Dim cRenk As String
        Dim cBeden As String
        Dim nAdet As Double
    End Structure

    Public Sub UTFFastGenerateAll()

        Dim cSQL As String = ""
        Dim aUTF() As String = Nothing
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

            cSQL = "select distinct a.uretimtakipno " + _
                    " from " + cSipModelTableName + " a, siparis b  " + _
                    " where a.siparisno = b.kullanicisipno " + _
                    " and a.uretimtakipno is not null " + _
                    " and a.uretimtakipno <> '' " + _
                    " and (b.dosyakapandi = 'H' or b.dosyakapandi = '' or b.dosyakapandi is null) " + _
                    " order by a.uretimtakipno "

            If CheckExists(cSQL) Then
                aUTF = SQLtoStringArray(cSQL)
                For nCnt = 0 To UBound(aUTF)
                    UTFGenerate(aUTF(nCnt))
                Next
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "UTFFastGenerateAll", cSQL)
        End Try
    End Sub

    Public Function UTFGenerate(cUTF As String) As Integer

        Dim cSQL As String = ""
        Dim aModel() As String = Nothing
        Dim aUretPlLines() As oUretPlLines = Nothing
        Dim aUretPlRBA() As oUretPlRBA = Nothing
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim nAdet As Double = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim lAltModelDetay As Boolean = False
        Dim cSipModelTableName As String = ""
        Dim nToleransCarpani As Double = 0

        UTFGenerate = 0

        Try
            lAltModelDetay = (GetSysPar("altmodeltakibi") = "1")

            If lAltModelDetay Then
                cSipModelTableName = "sipsubmodel"
            Else
                cSipModelTableName = "sipmodel"
            End If

            cSQL = "select distinct modelno " + _
                    " from " + cSipModelTableName + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and modelno is not null " + _
                    " and modelno <> '' "

            If CheckExists(cSQL) Then
                aModel = SQLtoStringArray(cSQL)
                For nCnt = 0 To UBound(aModel)
                    CheckModelRota(aModel(nCnt))
                Next
            End If

            ConnYage = OpenConn()

            cSQL = "select w.* " + _
                    " from (select distinct a.modelno, a.bedenseti, " + _
                            " b.departman, b.uretimtoleransi, b.giristakipesasi, b.cikistakipesasi, b.parca, " + _
                            " b.sira, b.girisdepartmani, b.cikisdepartmani, b.girisparcasi, b.firma, " + _
                            " b.yikamakodu, b.iscilikfiyat, b.iscilikdoviz, "
            cSQL = cSQL + _
                            " bssirano = (select top 1 rcount " + _
                                        " from sipmodel " + _
                                        " where uretimtakipno = a.uretimtakipno  " + _
                                        " and modelno =  a.modelno " + _
                                        " and departman = b.departman " + _
                                        " and parca = b.parca " + _
                                        " and bedenseti = a.bedenseti), "
            cSQL = cSQL + _
                            " uplsirano = (select top 1 sirano " + _
                                        " from uretpllines " + _
                                        " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                        " and modelno =  a.modelno " + _
                                        " and departman = b.departman " + _
                                        " and parca = b.parca " + _
                                        " and bedenseti = a.bedenseti) "
            cSQL = cSQL + _
                            " from " + cSipModelTableName + " a, modeluretim b " + _
                            " where a.modelno = b.modelno " + _
                            " and a.uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and a.modelno is not null " + _
                            " and a.modelno <> '' " + _
                            " and a.bedenseti is not null " + _
                            " and a.bedenseti <> '' " + _
                            " and b.departman is not null " + _
                            " and b.departman <> '') w " + _
                    " order by w.modelno, w.bssirano, w.sira "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aUretPlLines(nCnt)
                aUretPlLines(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aUretPlLines(nCnt).cBedenSeti = SQLReadString(oReader, "bedenseti")
                aUretPlLines(nCnt).cDepartman = SQLReadString(oReader, "departman")
                aUretPlLines(nCnt).nUretimToleransi = SQLReadDouble(oReader, "uretimtoleransi")
                aUretPlLines(nCnt).cGirisTakipEsasi = SQLReadString(oReader, "giristakipesasi")
                aUretPlLines(nCnt).cCikisTakipEsasi = SQLReadString(oReader, "cikistakipesasi")
                aUretPlLines(nCnt).cParca = SQLReadString(oReader, "parca")
                aUretPlLines(nCnt).nSira = SQLReadDouble(oReader, "sira")
                aUretPlLines(nCnt).cGirisDepartmani = SQLReadString(oReader, "girisdepartmani")
                aUretPlLines(nCnt).cCikisDepartmani = SQLReadString(oReader, "cikisdepartmani")
                aUretPlLines(nCnt).cGirisParcasi = SQLReadString(oReader, "girisparcasi")
                aUretPlLines(nCnt).cFirma = SQLReadString(oReader, "firma")
                aUretPlLines(nCnt).cYikamaKodu = SQLReadString(oReader, "yikamakodu")
                aUretPlLines(nCnt).nIscilikFiyat = SQLReadDouble(oReader, "iscilikfiyat")
                aUretPlLines(nCnt).cIscilikDoviz = SQLReadString(oReader, "iscilikdoviz")
                aUretPlLines(nCnt).nUPLSiraNo = SQLReadDouble(oReader, "uplsirano")

                If aUretPlLines(nCnt).cFirma = "" Then aUretPlLines(nCnt).cFirma = "DAHILI"
                If aUretPlLines(nCnt).cParca = "" Then aUretPlLines(nCnt).cParca = "KOMPLE"
                If aUretPlLines(nCnt).cGirisTakipEsasi = "" Then aUretPlLines(nCnt).cGirisTakipEsasi = "4"
                If aUretPlLines(nCnt).cCikisTakipEsasi = "" Then aUretPlLines(nCnt).cCikisTakipEsasi = "4"

                nCnt = nCnt + 1
            Loop
            oReader.Close()

            If nCnt = 0 Then
                ' kayit bulunamadi
                ConnYage.Close()
                UTFGenerate = 1
                Exit Function
            End If

            cSQL = "delete from uretplmaliyet " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from uretplrba " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "select uretimtakipno " + _
                    " from uretplfis " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "insert uretplfis (uretimtakipno , dosyakapandi , notlar) " + _
                     " values ('" + cUTF.Trim + "', " + _
                     " 'H', " + _
                     " 'CLR-otomatik UTF' ) "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            For nCnt = 0 To UBound(aUretPlLines)
                If aUretPlLines(nCnt).nUPLSiraNo <> 0 Then
                    cSQL = "update uretpllines " + _
                            " set sira = " + SQLWriteDecimal(aUretPlLines(nCnt).nSira) + ", " + _
                            " girisdepartmani = '" + aUretPlLines(nCnt).cGirisDepartmani + "', " + _
                            " cikisdepartmani = '" + aUretPlLines(nCnt).cCikisDepartmani + "', " + _
                            " girisparcasi = '" + aUretPlLines(nCnt).cGirisParcasi + "', " + _
                            " giristakipesasi = '" + aUretPlLines(nCnt).cGirisTakipEsasi + "', "

                    cSQL = cSQL + _
                            " cikistakipesasi = '" + aUretPlLines(nCnt).cCikisTakipEsasi + "', " + _
                            " islemkodu = '" + aUretPlLines(nCnt).cYikamaKodu + "', " + _
                            " plfirma = '" + aUretPlLines(nCnt).cFirma + "', " + _
                            " fiyati = " + SQLWriteDecimal(aUretPlLines(nCnt).nIscilikFiyat) + ", " + _
                            " doviz = '" + aUretPlLines(nCnt).cIscilikDoviz + "' "

                    cSQL = cSQL + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " + _
                            " and departman = '" + aUretPlLines(nCnt).cDepartman + "' " + _
                            " and parca = '" + aUretPlLines(nCnt).cParca + "' " + _
                            " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    cSQL = "insert uretpllines " + _
                            " (modelno, departman, uretimtakipno, uretimtoleransi, giristakipesasi, " + _
                            " cikistakipesasi, parca, sira, cikisdepartmani, girisdepartmani, " + _
                            " girisparcasi, toplamadet, isemriverilen, gelen, giden, " + _
                            " parcagelen, parcagiden, bedenseti, gelenparcacount, islemkodu, " + _
                            " plfirma, fiyati, doviz) "

                    cSQL = cSQL + _
                            " values  ('" + aUretPlLines(nCnt).cModelNo + "', " + _
                            " '" + aUretPlLines(nCnt).cDepartman + "', " + _
                            " '" + cUTF.Trim + "', " + _
                            SQLWriteDecimal(aUretPlLines(nCnt).nUretimToleransi) + ", " + _
                            " '" + aUretPlLines(nCnt).cGirisTakipEsasi + "',"

                    cSQL = cSQL + _
                            " '" + aUretPlLines(nCnt).cCikisTakipEsasi + "', " + _
                            " '" + aUretPlLines(nCnt).cParca + "', " + _
                            SQLWriteDecimal(aUretPlLines(nCnt).nSira) + ", " + _
                            " '" + aUretPlLines(nCnt).cCikisDepartmani + "', " + _
                            " '" + aUretPlLines(nCnt).cGirisDepartmani + "', "

                    cSQL = cSQL + _
                            " '" + aUretPlLines(nCnt).cGirisParcasi + "',0,0,0,0, "

                    cSQL = cSQL + _
                            " 0, " + _
                            " 0, " + _
                            " '" + aUretPlLines(nCnt).cBedenSeti + "', " + _
                            " 1, " + _
                            " '" + aUretPlLines(nCnt).cYikamaKodu + "', "

                    cSQL = cSQL + _
                            " '" + aUretPlLines(nCnt).cFirma + "', " + _
                            SQLWriteDecimal(aUretPlLines(nCnt).nIscilikFiyat) + ", " + _
                            " '" + aUretPlLines(nCnt).cIscilikDoviz + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                nToleransCarpani = 1 + (aUretPlLines(nCnt).nUretimToleransi / 100)

                Select Case aUretPlLines(nCnt).cCikisTakipEsasi
                    Case "1"
                        cSQL = "select sum(coalesce(adet,0)) " + _
                                " from " + cSipModelTableName + _
                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " + _
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' "

                        nAdet = SQLGetDoubleConnected(cSQL, ConnYage) * nToleransCarpani

                        cSQL = "insert uretplrba (uretimtakipno, departman, modelno, bedenseti, parca, renk, beden, adet) " + _
                                " values ('" + cUTF.Trim + "', " + _
                                " '" + aUretPlLines(nCnt).cDepartman + "', " + _
                                " '" + aUretPlLines(nCnt).cModelNo + "', " + _
                                " '" + aUretPlLines(nCnt).cBedenSeti + "', " + _
                                " '" + aUretPlLines(nCnt).cParca + "', " + _
                                " 'HEPSI', " + _
                                " 'HEPSI', " + _
                                SQLWriteDecimal(nAdet) + " ) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                    Case "2"
                        cSQL = "insert uretplrba (uretimtakipno, modelno, bedenseti, departman, parca, renk, beden, adet) " +
                                " select uretimtakipno, modelno, bedenseti, " +
                                " departman = '" + aUretPlLines(nCnt).cDepartman + "',  " +
                                " parca = '" + aUretPlLines(nCnt).cParca + "',  " +
                                " renk, " +
                                " beden = 'HEPSI',  " +
                                " adet = sum(coalesce(adet,0)) * " + SQLWriteDecimal(nToleransCarpani) +
                                " from " + cSipModelTableName +
                                " where uretimtakipno = '" + cUTF.Trim + "' " +
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " +
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' " +
                                " group by uretimtakipno, modelno, bedenseti, renk " +
                                " order by renk "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                    Case "3"
                        cSQL = "insert uretplrba (uretimtakipno, modelno, bedenseti, departman, parca, renk, beden, adet) " +
                                " select uretimtakipno, modelno, bedenseti, " +
                                " departman = '" + aUretPlLines(nCnt).cDepartman + "',  " +
                                " parca = '" + aUretPlLines(nCnt).cParca + "',  " +
                                " renk = 'HEPSI', " +
                                " beden " +
                                " adet = sum(coalesce(adet,0)) * " + SQLWriteDecimal(nToleransCarpani) +
                                " from  " + cSipModelTableName +
                                " where uretimtakipno = '" + cUTF.Trim + "' " +
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " +
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' " +
                                " group by uretimtakipno, modelno, bedenseti, beden " +
                                " order by beden "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                    Case Else
                        cSQL = "insert uretplrba (uretimtakipno, modelno, bedenseti, departman, parca, renk, beden, adet) " +
                                " select uretimtakipno, modelno, bedenseti, " +
                                " departman = '" + aUretPlLines(nCnt).cDepartman + "',  " +
                                " parca = '" + aUretPlLines(nCnt).cParca + "',  " +
                                " renk, " +
                                " beden,  " +
                                " adet = sum(coalesce(adet,0)) * " + SQLWriteDecimal(nToleransCarpani) +
                                " from " + cSipModelTableName +
                                " where uretimtakipno = '" + cUTF.Trim + "' " +
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " +
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' " +
                                " group by uretimtakipno, modelno, bedenseti, renk, beden " +
                                " order by renk, beden "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                End Select
            Next

            ' üreim sakat fişlerinide olaya dahil et
            cSQL = "update uretplrba " + _
                    " set adet = coalesce(adet,0) + coalesce((select sum(coalesce(b.adet,0)) " + _
                                                            " from usfis a, usfislines b " + _
                                                            " where a.usfisno = b.usfisno " + _
                                                            " and a.yenidenuretilsin = 'E' " + _
                                                            " and b.uretimtakipno = uretplrba.uretimtakipno " + _
                                                            " and b.modelno = uretplrba.modelno " + _
                                                            " and b.renk = uretplrba.renk " + _
                                                            " and b.beden = uretplrba.beden),0) " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and departman not in ('','SEVKIYAT') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            RecalcUretPlLines(ConnYage, cUTF.Trim)

            ConnYage.Close()

            UTFGenerate = 1
        Catch ex As Exception
            ErrDisp(ex.Message, "UTFGenerate", cSQL)
        End Try
    End Function

    Private Sub CheckModelRota(Optional cModelNo As String = "")
        ' AnaModelTipi ne göre varsayılan şablonu kullanılıyor
        Dim cSQL As String = ""
        Dim cFormNo As String = ""
        Dim aRota() As oRota = Nothing
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nCnt As Integer = 0

        Try
            If cModelNo.Trim = "" Then Exit Sub

            cSQL = "select modelno " + _
                    " from modeluretim " + _
                    " where modelno = '" + cModelNo.Trim + "' "

            If CheckExists(cSQL) Then
                Exit Sub
            End If

            cSQL = "select a.formno " + _
                    " from frmuretim a, ymodel b " + _
                    " where b.modelno = '" + cModelNo.Trim + "' " + _
                    " and (a.anamodeltipi = b.anamodeltipi or a.anamodeltipi = 'HEPSI') " + _
                    " and a.varsayilan = 'E' "

            cFormNo = SQLGetString(cSQL)

            If cFormNo.Trim = "" Then Exit Sub

            ConnYage = OpenConn()

            cSQL = "select a.departman, a.parca, a.tolerans, a.GidenDepartman, a.sira, " + _
                    " b.giristakipesasi, b.cikistakipesasi, deptsira = b.sira " + _
                    " from frmuretim a, departman b " + _
                    " where a.departman = b.departman " + _
                    " and a.formno = '" + cFormNo.Trim + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aRota(nCnt)
                aRota(nCnt).cDepartman = SQLReadString(oReader, "departman")
                aRota(nCnt).cParca = SQLReadString(oReader, "parca")
                aRota(nCnt).nTolerans = SQLReadDouble(oReader, "tolerans")
                aRota(nCnt).cGidenDepartman = SQLReadString(oReader, "GidenDepartman")
                aRota(nCnt).nSira = SQLReadDouble(oReader, "sira")
                aRota(nCnt).cGirisTakipEsasi = SQLReadString(oReader, "giristakipesasi")
                aRota(nCnt).cCikisTakipEsasi = SQLReadString(oReader, "cikistakipesasi")

                If aRota(nCnt).cParca.Trim = "" Then aRota(nCnt).cParca = "KOMPLE"
                If aRota(nCnt).nSira = 0 Then aRota(nCnt).nSira = SQLReadDouble(oReader, "deptsira")
                If aRota(nCnt).cGirisTakipEsasi = "" Then aRota(nCnt).cGirisTakipEsasi = "4"
                If aRota(nCnt).cCikisTakipEsasi = "" Then aRota(nCnt).cCikisTakipEsasi = "4"

                nCnt = nCnt + 1
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aRota)
                cSQL = "insert modeluretim " + _
                        " (modelno, departman, uretimtoleransi, giristakipesasi, cikistakipesasi, " + _
                        " parca, sira, girisdepartmani) "

                cSQL = cSQL + _
                        " values ('" + cModelNo.Trim + "', " + _
                        " '" + aRota(nCnt).cDepartman + "', " + _
                        SQLWriteDecimal(aRota(nCnt).nTolerans) + ", " + _
                        " '" + aRota(nCnt).cGirisTakipEsasi + "', " + _
                        " '" + aRota(nCnt).cCikisTakipEsasi + "', "

                cSQL = cSQL + _
                        " '" + aRota(nCnt).cParca + "', " + _
                        SQLWriteDecimal(aRota(nCnt).nSira) + ", " + _
                        " '" + aRota(nCnt).cGidenDepartman + "') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "CheckModelRota : " + cModelNo, cSQL)
        End Try

    End Sub

    Public Function UretimisEmriUret(ByVal cUTF As String, ByVal cAction As String, ByVal cDepartman As String, _
                                    ByVal cDefaultFirma As String, ByVal cDefaultPersonel As String, ByVal cPadNo As String, _
                                    ByVal cPartiNo As String, ByVal cUserName As String, ByVal cDepts As String, _
                                    ByVal cKesimSistemiNo As String) As Integer
        Dim cSQL As String = ""
        Dim cTable As String = ""
        Dim aBedenSeti() As String = Nothing
        Dim aFirma() As String = Nothing
        Dim aIsemri() As oIsEmri = Nothing
        Dim aRBAA() As oRBAA = Nothing
        Dim cBedenSeti As String = ""
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = -1
        Dim nCnt2 As Integer = -1
        Dim nCnt3 As Integer = -1
        Dim nFound As Integer = -1
        Dim lOK As Boolean = True
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim lSipmodel As Boolean = True
        Dim nFirma As Integer = 0
        Dim lNewIsEmri As Boolean = False
        Dim cIsEmriNo As String = ""
        Dim nIsemriAdedi As Double = 0
        Dim cFirma As String = ""
        Dim cPersonel As String = ""
        Dim nUlineNo As Double = 0
        Dim nOncekiAdet As Double = 0

        UretimisEmriUret = 0

        Try
            ConnYage = OpenConn()

            ErrDispConnected(ConnYage, "Basla : UretimisEmriUret", cUTF.Trim)

            If cDefaultFirma.Trim = "" Then
                cDefaultFirma = GetSysParConnected("dahili", ConnYage)
            End If

            cSQL = "select bedenseti1, bedenseti2, bedenseti3, bedenseti4, bedenseti5, " + _
                    " bedenseti6, bedenseti7, bedenseti8, bedenseti9, bedenseti10 " + _
                    " from siparis " + _
                    " where kullanicisipno in (select siparisno " + _
                                            " from sipmodel " + _
                                            " where uretimtakipno = '" + cUTF.Trim + "') "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                For nCnt = 1 To 10
                    cBedenSeti = "bedenseti" + CStr(nCnt)
                    If SQLReadString(oReader, cBedenSeti) <> "" Then

                        If nCnt1 = -1 Then
                            nCnt1 = nCnt1 + 1
                            ReDim Preserve aBedenSeti(nCnt1)
                            aBedenSeti(nCnt1) = SQLReadString(oReader, cBedenSeti)
                        Else
                            nFound = -1
                            For nCnt2 = 0 To UBound(aBedenSeti)
                                If aBedenSeti(nCnt2) = SQLReadString(oReader, cBedenSeti) Then
                                    nFound = nCnt2
                                    Exit For
                                End If
                            Next

                            If nFound = -1 Then
                                nCnt1 = nCnt1 + 1
                                ReDim Preserve aBedenSeti(nCnt1)
                                aBedenSeti(nCnt1) = SQLReadString(oReader, cBedenSeti)
                            End If
                        End If
                    End If
                Next
            Loop
            oReader.Close()

            If nCnt1 = -1 Then
                ' beden seti yok
                UretimisEmriUret = 2
                ' önce uretpllines yeniden hesaplanır
                RecalcUretPlLines(ConnYage, cUTF, , cDepartman, , , cDepts)
                ' üretim işemirleri hesaplanırken önceden hesaplanmış uretpllines kılavuz olarak alınır
                RecalcUretimIsemriDetails(ConnYage, cUTF, , cDepartman, , , , cDepts)
                ' topla çık
                ErrDispConnected(ConnYage, "BITIS beden seti yok : UretimisEmriUret", cUTF.Trim)
                ConnYage.Close()
                Exit Function
            End If

            cSQL = "(bedenseti char(30) null, sira decimal(10,0) null)"
            cTable = CreateTempTable(ConnYage, cSQL)

            For nCnt = 0 To UBound(aBedenSeti)
                lOK = True

                If cPadNo.Trim <> "" Then
                    cSQL = "select bedenseti " + _
                            " from pastalasortidagilimidetayi " + _
                            " where padno = '" + cPadNo.Trim + "' " + _
                            " and partino = '" + cPartiNo.Trim + "' " + _
                            " and bedenseti = '" + aBedenSeti(nCnt) + "' " + _
                            " and adet is not null " + _
                            " and adet <> 0 "

                    lOK = CheckExistsConnected(cSQL, ConnYage)
                End If

                If lOK Then
                    cSQL = "insert into " + cTable + " (bedenseti,sira) " +
                           " values ('" + aBedenSeti(nCnt) + "', " +
                           SQLWriteDecimal(nCnt) + ") "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Next

            Select Case cAction
                Case "add"
                    lSipmodel = True

                    cSQL = "select distinct firma " + _
                            " from sipmodel " + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and firma is not null " + _
                            " and firma <> '' " + _
                            " and adet is not null " + _
                            " and adet <> 0 " + _
                            " order by firma "

                    If CheckExistsConnected(cSQL, ConnYage) Then
                        aFirma = SQLtoStringArrayConnected(cSQL, ConnYage)
                    Else
                        ReDim aFirma(0)
                        aFirma(0) = ""
                        lSipmodel = False
                    End If
                Case Else
                    ReDim aFirma(0)
                    aFirma(0) = ""
                    lSipmodel = False
            End Select

            ErrDispConnected(ConnYage, "Basla uretpllines okunuyor", cUTF.Trim)

            nCnt = -1

            cSQL = "Select w.* " + _
                    " from (Select distinct sirano, uretimtakipno, modelno, departman, parca, " + _
                            " bedenseti, cikistakipesasi, baslamatarihi, bitistarihi, fiyati, " + _
                            " doviz, plfirma, "

            cSQL = cSQL + _
                            " iefirma1 = (select top 1 a.firma  " + _
                                        " from uretimisemri a,  uretimisdetayi b " + _
                                        " where a.isemrino = b.isemrino " + _
                                        " and a.uretimtakipno = b.uretimtakipno " + _
                                        " and a.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and a.departman = uretpllines.departman " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and a.firma <> '' " + _
                                        " and a.firma is not null " + _
                                        " order by a.tarih desc), "
            cSQL = cSQL + _
                            " ieeleman1 = (select top 1 a.eleman  " + _
                                        " from uretimisemri a,  uretimisdetayi b " + _
                                        " where a.isemrino = b.isemrino " + _
                                        " and a.uretimtakipno = b.uretimtakipno " + _
                                        " and a.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and a.departman = uretpllines.departman " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and a.eleman <> '' " + _
                                        " and a.eleman is not null " + _
                                        " order by a.tarih desc), "
            cSQL = cSQL + _
                            " iefirma2 = (select top 1 a.firma  " + _
                                        " from uretimisemri a,  uretimisdetayi b " + _
                                        " where a.isemrino = b.isemrino " + _
                                        " and a.uretimtakipno = b.uretimtakipno " + _
                                        " and a.departman = uretpllines.departman " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and a.firma <> '' " + _
                                        " and a.firma is not null " + _
                                        " order by a.tarih desc), "
            cSQL = cSQL + _
                            " ieeleman2 = (select top 1 a.eleman  " + _
                                        " from uretimisemri a,  uretimisdetayi b " + _
                                        " where a.isemrino = b.isemrino " + _
                                        " and a.uretimtakipno = b.uretimtakipno " + _
                                        " and a.departman = uretpllines.departman " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and a.eleman <> '' " + _
                                        " and a.eleman is not null " + _
                                        " order by a.tarih desc), "
            cSQL = cSQL + _
                           " isemriverilen2 = (select sum (coalesce(adet,0)) " + _
                                        " from uretimisrba " + _
                                        " where uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and modelno = uretpllines.modelno " + _
                                        " and departman = uretpllines.departman " + _
                                        " and bedenseti = uretpllines.bedenseti " + _
                                        " and parca = uretpllines.parca), "
            cSQL = cSQL + _
                           " toplamadet2 = (select sum (coalesce(adet,0)) " + _
                                        " from uretplrba " + _
                                        " where uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and modelno = uretpllines.modelno " + _
                                        " and departman = uretpllines.departman " + _
                                        " and bedenseti = uretpllines.bedenseti " + _
                                        " and parca = uretpllines.parca), "
            cSQL = cSQL + _
                            " sira = (select min(sira) from " + cTable + " where bedenseti = uretpllines.bedenseti) "

            cSQL = cSQL + _
                            " from uretpllines " + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            IIf(cDepartman = "", "", " and departman = '" + cDepartman + "' ").ToString + _
                            IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString + " ) w "

            cSQL = cSQL + _
                    " order by w.modelno, w.sira, w.sirano, w.departman "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                lOK = True

                If cPadNo.Trim = "" Then
                    lOK = SQLReadDouble(oReader, "toplamadet2") > SQLReadDouble(oReader, "isemriverilen2")
                End If

                If lOK Then
                    nCnt = nCnt + 1
                    ReDim Preserve aIsemri(nCnt)

                    aIsemri(nCnt).cUTF = SQLReadString(oReader, "uretimtakipno")
                    aIsemri(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                    aIsemri(nCnt).cDepartman = SQLReadString(oReader, "departman")
                    aIsemri(nCnt).cParca = SQLReadString(oReader, "parca")
                    aIsemri(nCnt).cBedenSeti = SQLReadString(oReader, "bedenseti")
                    aIsemri(nCnt).cCikisTakipEsasi = SQLReadString(oReader, "cikistakipesasi")
                    aIsemri(nCnt).cDoviz = SQLReadString(oReader, "doviz")
                    aIsemri(nCnt).cPLFirma = SQLReadString(oReader, "plfirma")
                    aIsemri(nCnt).cIEFirma1 = SQLReadString(oReader, "iefirma1")
                    aIsemri(nCnt).cIEEleman1 = SQLReadString(oReader, "ieeleman1")
                    aIsemri(nCnt).cIEFirma2 = SQLReadString(oReader, "iefirma2")
                    aIsemri(nCnt).cIEEleman2 = SQLReadString(oReader, "ieeleman2")

                    aIsemri(nCnt).dBaslamaTarihi = SQLReadDate(oReader, "baslamatarihi")
                    aIsemri(nCnt).dBitisTarihi = SQLReadDate(oReader, "bitistarihi")

                    aIsemri(nCnt).nFiyat = SQLReadDouble(oReader, "fiyati")
                    aIsemri(nCnt).nSira = SQLReadDouble(oReader, "sira")
                End If
            Loop
            oReader.Close()

            If nCnt = -1 Then
                ' işemri verilecek adet yok
                UretimisEmriUret = 3
                ' önce uretpllines yeniden hesaplanır
                RecalcUretPlLines(ConnYage, cUTF, , cDepartman, , , cDepts)
                ' üretim işemirleri hesaplanırken önceden hesaplanmış uretpllines kılavuz olarak alınır
                RecalcUretimIsemriDetails(ConnYage, cUTF, , cDepartman, , , , cDepts)
                ErrDispConnected(ConnYage, "BITIS adet yok : UretimisEmriUret", cUTF.Trim)
                DropTable(cTable, ConnYage)
                ConnYage.Close()
                Exit Function
            End If

            ErrDispConnected(ConnYage, "Basla ana dongu", cUTF.Trim)

            If cAction = "add" Or cAction = "revise" Then
                cSQL = "select uretimtakipno " + _
                        " from uretimisemrifis " + _
                        " where uretimtakipno = '" + cUTF.Trim + "' "

                If CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "update uretimisemrifis " +
                           " set notlar = 'otomatik uretim isemri-clr', " +
                           " modificationdate = getdate(), " +
                           " username = '" + cUserName.Trim + "' " +
                           " where uretimtakipno = '" + cUTF.Trim + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    cSQL = "set dateformat dmy " + _
                           " insert into uretimisemrifis (uretimtakipno, notlar, modificationdate, username) " + _
                           " values ('" + cUTF.Trim + "', " + _
                           " 'otomatik uretim isemri-clr', " + _
                           " getdate(), " + _
                           " '" + cUserName.Trim + "') "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            End If

            For nFirma = 0 To UBound(aFirma)
                For nCnt = 0 To UBound(aIsemri)
                    If cAction = "add" Or cAction = "revise" Then

                        If cDefaultPersonel.Trim = "" Then
                            If aIsemri(nCnt).cIEEleman1 <> "" Then
                                cPersonel = aIsemri(nCnt).cIEEleman1
                            ElseIf aIsemri(nCnt).cIEEleman2 <> "" Then
                                cPersonel = aIsemri(nCnt).cIEEleman2
                            End If
                        Else
                            cPersonel = cDefaultPersonel.Trim
                        End If

                        If aFirma(nFirma) <> "" Then
                            cFirma = aFirma(nFirma)
                        Else
                            If cDefaultFirma.Trim = "" Then
                                If aIsemri(nCnt).cIEFirma1 <> "" Then
                                    cFirma = aIsemri(nCnt).cIEFirma1
                                ElseIf aIsemri(nCnt).cIEFirma2 <> "" Then
                                    cFirma = aIsemri(nCnt).cIEFirma2
                                End If
                            Else
                                cFirma = cDefaultFirma.Trim
                            End If
                        End If

                        lNewIsEmri = False

                        If cAction = "revise" Then
                            ' revise uretimisemri
                            cSQL = "select a.isemrino " + _
                                   " from uretimisemri a, uretimisdetayi  b " + _
                                   " where a.uretimtakipno = b.uretimtakipno " + _
                                   " and a.isemrino = b.isemrino " + _
                                   " and a.uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " + _
                                   " and a.departman = '" + aIsemri(nCnt).cDepartman + "' " + _
                                   " and b.bedenseti = '" + aIsemri(nCnt).cBedenSeti + "' "

                            If cFirma.Trim <> "" Then
                                cSQL = cSQL + _
                                    " and a.firma = '" + cFirma.Trim + "' "
                            End If

                            cSQL = cSQL + _
                                   " order by a.isemrino desc "

                            cIsEmriNo = SQLGetStringConnected(cSQL, ConnYage)

                            lNewIsEmri = (cIsEmriNo.Trim = "")
                        Else
                            lNewIsEmri = True
                        End If

                        If lNewIsEmri Then
                            ' add uretimisemri
                            cIsEmriNo = GetLastIsemriNo(ConnYage, aIsemri(nCnt).cUTF, aIsemri(nCnt).cDepartman, aIsemri(nCnt).cBedenSeti)

                            cSQL = "select isemrino " + _
                                    " from uretimisemri " + _
                                    " where uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " + _
                                    " and departman = '" + aIsemri(nCnt).cDepartman + "' " + _
                                    " and isemrino = '" + cIsEmriNo.Trim + "' "

                            If Not CheckExistsConnected(cSQL, ConnYage) Then
                                cSQL = "set dateformat dmy " +
                                        " insert into uretimisemri " +
                                        " (uretimtakipno, isemrino, tarih, departman, firma, " +
                                        " eleman, ok, padno, partino, cikistakipesasi, " +
                                        " modificationdate, username) "

                                cSQL = cSQL +
                                       " values ('" + SQLWriteString(aIsemri(nCnt).cUTF, 30) + "', " +
                                       " '" + SQLWriteString(cIsEmriNo, 30) + "', " +
                                       " '" + SQLWriteDate(Today) + "', " +
                                       " '" + SQLWriteString(aIsemri(nCnt).cDepartman, 30) + "', " +
                                       " '" + SQLWriteString(cFirma, 30) + "', "

                                cSQL = cSQL +
                                       " '" + SQLWriteString(cPersonel, 30) + "', " +
                                       " 'H', " +
                                       " '" + SQLWriteString(cPadNo, 30) + "', " +
                                       " '" + SQLWriteString(cPartiNo, 30) + "', " +
                                       " '" + SQLWriteString(aIsemri(nCnt).cCikisTakipEsasi, 1) + "', "

                                cSQL = cSQL +
                                       " getdate(), " +
                                       " '" + SQLWriteString(cUserName, 30) + "') "

                                ExecuteSQLCommandConnected(cSQL, ConnYage, True)
                            End If
                        End If

                        cSQL = "select UlineNo " + _
                                " from uretimisdetayi " + _
                                " where uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " + _
                                " and departman = '" + aIsemri(nCnt).cDepartman + "' " + _
                                " and isemrino = '" + cIsEmriNo.Trim + "' " + _
                                " and modelno = '" + aIsemri(nCnt).cModelNo + "' " + _
                                " and bedenseti = '" + aIsemri(nCnt).cBedenSeti + "' " + _
                                " and parca = '" + aIsemri(nCnt).cParca + "' "

                        nUlineNo = SQLGetDoubleConnected(cSQL, ConnYage)

                        If nUlineNo = 0 Then
                            nUlineNo = getuHARlineno(ConnYage)

                            cSQL = "insert into uretimisdetayi " + _
                                   " (uretimtakipno, isemrino, modelno, bedenseti, parca, " + _
                                   " toplamadet, departman, cikistakipesasi, partino, ulineno) "

                            cSQL = cSQL +
                                   " values ('" + SQLWriteString(aIsemri(nCnt).cUTF, 30) + "', " +
                                   " '" + SQLWriteString(cIsEmriNo, 30) + "', " +
                                   " '" + SQLWriteString(aIsemri(nCnt).cModelNo, 30) + "', " +
                                   " '" + SQLWriteString(aIsemri(nCnt).cBedenSeti, 30) + "', " +
                                   " '" + SQLWriteString(aIsemri(nCnt).cParca, 30) + "', "

                            cSQL = cSQL +
                                   " 0, " +
                                   " '" + SQLWriteString(aIsemri(nCnt).cDepartman, 30) + "', " +
                                   " '" + SQLWriteString(aIsemri(nCnt).cCikisTakipEsasi, 1) + "', " +
                                   " '" + SQLWriteString(cKesimSistemiNo, 30) + "', " +
                                   SQLWriteDecimal(nUlineNo) + ") "

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End If

                        nCnt3 = -1

                        If cPadNo.Trim = "" Then
                            If lSipmodel Then
                                ' firma sipmodel den geliyorsa adetleri de sipmodel den al
                                cSQL = "select uretimtakipno, modelno, renk, beden, firma, bedenseti, " + _
                                        " adet = round((sum(coalesce(adet,0)) * (1+(select top 1 coalesce(uretimtoleransi,0) " + _
                                                                                    " from modeluretim " + _
                                                                                    " where modelno = sipmodel.modelno " + _
                                                                                    " and departman = '" + aIsemri(nCnt).cDepartman + "') / 100)),0), "
                                cSQL = cSQL + _
                                        " oncekiadet = (select sum(coalesce(adet,0)) " + _
                                                        " from uretimisemri a, uretimisdetayi b, uretimisrba c " + _
                                                        " where a.isemrino = b.isemrino " + _
                                                        " and b.isemrino = c.isemrino " + _
                                                        " and b.ulineno = c.ulineno " + _
                                                        " and c.uretimtakipno = sipmodel.uretimtakipno " + _
                                                        " and c.departman = '" + aIsemri(nCnt).cDepartman + "' " + _
                                                        " and c.modelno = sipmodel.modelno " + _
                                                        " and c.bedenseti = sipmodel.bedenseti " + _
                                                        " and c.parca = '" + aIsemri(nCnt).cParca + "' " + _
                                                        " and c.renk = sipmodel.renk " + _
                                                        " and c.beden = sipmodel.beden " + _
                                                        " and a.firma = sipmodel.firma ), "
                                cSQL = cSQL + _
                                        " ulineno = (select ulineno " + _
                                                        " from uretimisrba " + _
                                                        " where isemrino = '" + cIsEmriNo.Trim + "' " + _
                                                        " and uretimtakipno = sipmodel.uretimtakipno " + _
                                                        " and departman = '" + aIsemri(nCnt).cDepartman + "' " + _
                                                        " and modelno = sipmodel.modelno " + _
                                                        " and bedenseti = sipmodel.bedenseti " + _
                                                        " and parca = '" + aIsemri(nCnt).cParca + "' " + _
                                                        " and renk = sipmodel.renk " + _
                                                        " and beden = sipmodel.beden ) "
                                cSQL = cSQL + _
                                        " from sipmodel " + _
                                        " where uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " + _
                                        " and modelno = '" + aIsemri(nCnt).cModelNo + "' " + _
                                        " and bedenseti = '" + aIsemri(nCnt).cBedenSeti + "' " + _
                                        " and firma = '" + aFirma(nFirma) + "' "

                                cSQL = cSQL + _
                                        " group by uretimtakipno, modelno, renk, beden, firma, bedenseti "
                            Else
                                ' Sipmodel den gelmiyorsa büyük ihtimal departmanda tek firma var demektir
                                cSQL = "select uretimtakipno, departman, bedenseti, parca, modelno, renk, beden, " + _
                                        " adet = sum(coalesce(adet,0)), "

                                cSQL = cSQL + _
                                        " oncekiadet = (select sum(coalesce(c.adet,0)) " + _
                                                        " from uretimisemri a, uretimisdetayi b, uretimisrba c " + _
                                                        " where a.isemrino = b.isemrino " + _
                                                        " and b.isemrino = c.isemrino " + _
                                                        " and b.ulineno = c.ulineno " + _
                                                        " and c.uretimtakipno = uretplrba.uretimtakipno " + _
                                                        " and c.departman = uretplrba.departman " + _
                                                        " and c.modelno = uretplrba.modelno " + _
                                                        " and c.bedenseti = uretplrba.bedenseti " + _
                                                        " and c.parca = uretplrba.parca " + _
                                                        " and c.renk = uretplrba.renk " + _
                                                        " and c.beden = uretplrba.beden ), "
                                cSQL = cSQL + _
                                        " ulineno = (select ulineno " + _
                                                        " from uretimisrba " + _
                                                        " where isemrino = '" + cIsEmriNo.Trim + "' " + _
                                                        " and uretimtakipno = uretplrba.uretimtakipno " + _
                                                        " and departman = uretplrba.departman " + _
                                                        " and modelno = uretplrba.modelno " + _
                                                        " and bedenseti = uretplrba.bedenseti " + _
                                                        " and parca = uretplrba.parca " + _
                                                        " and renk = uretplrba.renk " + _
                                                        " and beden = uretplrba.beden ) "
                                cSQL = cSQL + _
                                        " from uretplrba " + _
                                        " where uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " + _
                                        " and departman = '" + aIsemri(nCnt).cDepartman + "' " + _
                                        " and modelno = '" + aIsemri(nCnt).cModelNo + "' " + _
                                        " and bedenseti = '" + aIsemri(nCnt).cBedenSeti + "' " + _
                                        " and parca = '" + aIsemri(nCnt).cParca + "' "

                                cSQL = cSQL + _
                                        " group by uretimtakipno, departman, bedenseti, parca, modelno, renk, beden "
                            End If
                        Else
                            ' pastal asorti dağılımından geliyorsa adetleri PAD çalışmasından al
                            cSQL = "select y.renk, y.beden, " + _
                                    " adet = sum(coalesce(y.adet,0) * coalesce(y.katsayi,0)), " + _
                                    " oncekiadet = 0, "
                            cSQL = cSQL + _
                                    " ulineno = (select ulineno " + _
                                                " from uretimisrba " + _
                                                " where isemrino = '" + cIsEmriNo.Trim + "' " + _
                                                " and uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " + _
                                                " and departman = '" + aIsemri(nCnt).cDepartman + "' " + _
                                                " and modelno = '" + aIsemri(nCnt).cModelNo + "' " + _
                                                " and bedenseti = '" + aIsemri(nCnt).cBedenSeti + "' " + _
                                                " and parca = '" + aIsemri(nCnt).cParca + "' " + _
                                                " and renk = y.renk " + _
                                                " and beden = y.beden ) "
                            cSQL = cSQL + _
                                    " from pastalasortidagilimi x, pastalasortidagilimidetayi y " + _
                                    " where x.padno = y.padno" + _
                                    " and x.padno = '" + cPadNo.Trim + "' " + _
                                    " and y.partino = '" + cPartiNo.Trim + "' " + _
                                    " and x.modelno = '" + aIsemri(nCnt).cModelNo + "' " + _
                                    " and y.bedenseti = '" + aIsemri(nCnt).cBedenSeti + "' "

                            cSQL = cSQL + _
                                    " group by y.renk, y.beden "
                        End If

                        oReader = GetSQLReader(cSQL, ConnYage)

                        Do While oReader.Read
                            If cPadNo.Trim = "" Then
                                nOncekiAdet = SQLReadDouble(oReader, "oncekiadet")
                            Else
                                nOncekiAdet = 0
                            End If
                            If SQLReadDouble(oReader, "adet") > nOncekiAdet Then
                                nCnt3 = nCnt3 + 1
                                ReDim Preserve aRBAA(nCnt3)
                                aRBAA(nCnt3).cRenk = SQLReadString(oReader, "renk")
                                aRBAA(nCnt3).cBeden = SQLReadString(oReader, "beden")
                                aRBAA(nCnt3).nAdet = SQLReadDouble(oReader, "adet") - nOncekiAdet
                                aRBAA(nCnt3).nULineNo = SQLReadDouble(oReader, "ulineno")
                            End If
                        Loop
                        oReader.Close()

                        If nCnt3 > -1 Then
                            For nCnt3 = 0 To UBound(aRBAA)
                                If aRBAA(nCnt3).nULineNo <> 0 Then
                                    ' uretimisrba tablosunda daha önce kayıt açılmış ise
                                    cSQL = "update uretimisrba " +
                                           " set adet = coalesce(adet,0) + " + SQLWriteDecimal(aRBAA(nCnt3).nAdet) +
                                           " where isemrino = '" + cIsEmriNo.Trim + "' " +
                                           " and uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " +
                                           " and departman = '" + aIsemri(nCnt).cDepartman + "' " +
                                           " and modelno = '" + aIsemri(nCnt).cModelNo + "' " +
                                           " and bedenseti = '" + aIsemri(nCnt).cBedenSeti + "' " +
                                           " and parca = '" + aIsemri(nCnt).cParca + "' " +
                                           " and renk = '" + aRBAA(nCnt3).cRenk + "' " +
                                           " and beden = '" + aRBAA(nCnt3).cBeden + "' "

                                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                                Else
                                    cSQL = "insert into uretimisrba " + _
                                           " (uretimtakipno, isemrino, modelno, bedenseti, parca, " + _
                                           " departman, renk, beden, adet, ulineno) "

                                    cSQL = cSQL + _
                                           " values ('" + aIsemri(nCnt).cUTF + "', " + _
                                           " '" + cIsEmriNo.Trim + "', " + _
                                           " '" + aIsemri(nCnt).cModelNo + "', " + _
                                           " '" + aIsemri(nCnt).cBedenSeti + "', " + _
                                           " '" + aIsemri(nCnt).cParca + "', "

                                    cSQL = cSQL +
                                           " '" + aIsemri(nCnt).cDepartman + "', " +
                                           " '" + aRBAA(nCnt3).cRenk + "', " +
                                           " '" + aRBAA(nCnt3).cBeden + "', " +
                                           SQLWriteDecimal(aRBAA(nCnt3).nAdet) + ", " +
                                           SQLWriteDecimal(nUlineNo) + ") "

                                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                                End If
                            Next
                        End If

                        ' SIFIR adetli işemri üretildiyse işemrini sil
                        cSQL = "select sum(coalesce(adet,0)) " + _
                                " from uretimisrba " + _
                                " where uretimtakipno = '" + aIsemri(nCnt).cUTF + "' " + _
                                " and isemrino = '" + cIsEmriNo.Trim + "' "

                        nIsemriAdedi = SQLGetDoubleConnected(cSQL, ConnYage)

                        If nIsemriAdedi = 0 Then
                            DeleteSingleUretimIsemri(ConnYage, aIsemri(nCnt).cUTF, cIsEmriNo, , , , , cDepts, False)
                        End If

                    ElseIf cAction = "delete" Then
                        DeleteSingleUretimIsemri(ConnYage, aIsemri(nCnt).cUTF, , aIsemri(nCnt).cDepartman, aIsemri(nCnt).cModelNo, aIsemri(nCnt).cBedenSeti, aIsemri(nCnt).cParca, cDepts, False)
                    End If
                Next
            Next
            ' önce uretpllines yeniden hesaplanır
            ErrDispConnected(ConnYage, "Basla : RecalcUretPlLines", cUTF.Trim)
            RecalcUretPlLines(ConnYage, cUTF, , cDepartman, , , cDepts)
            ' üretim işemirleri hesaplanırken önceden hesaplanmış uretpllines kılavuz olarak alınır
            ErrDispConnected(ConnYage, "Basla : RecalcUretimIsemriDetails", cUTF.Trim)
            RecalcUretimIsemriDetails(ConnYage, cUTF, , cDepartman, , , , cDepts)

            ErrDispConnected(ConnYage, "BITIS : UretimisEmriUret", cUTF.Trim)

            DropTable(cTable, ConnYage)
            ConnYage.Close()

            UretimisEmriUret = 1

        Catch ex As Exception
            ErrDisp(ex.Message, "UretimisEmriUret", cSQL)
        End Try
    End Function

    Private Function GetLastIsemriNo(ConnYage As SqlConnection, cUTF As String, Optional cDept As String = "", Optional cBedenSeti As String = "") As String

        Dim cSQL As String = ""
        Dim nIsEmriNo As Double = 0
        Dim cIsEmriNo As String = ""
        Dim oReader As SqlDataReader

        GetLastIsemriNo = ""

        Try
            If cDept.Trim = "" Or cBedenSeti.Trim = "" Then
                cSQL = "select isemrino " + _
                       " from uretimisemri " + _
                       " where uretimtakipno = '" + cUTF.Trim + "' " + _
                       " order by isemrino desc "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    cIsEmriNo = SQLReadString(oReader, "isemrino")
                    cIsEmriNo = Mid(cIsEmriNo, Len(cIsEmriNo) - 1, 2)
                    nIsEmriNo = CDbl(cIsEmriNo)
                End If
                oReader.Close()

                Do While True
                    nIsEmriNo = nIsEmriNo + 1
                    cIsEmriNo = Trim(Mid(cUTF, 1, 27)) + "_" + Microsoft.VisualBasic.Format(nIsEmriNo, "00")

                    cSQL = "select isemrino " + _
                           " from uretimisemri " + _
                           " where isemrino = '" + cIsEmriNo.Trim + "' "

                    If Not CheckExistsConnected(cSQL, ConnYage) Then
                        Exit Do
                    End If
                Loop
            Else
                cSQL = "select isemrino " + _
                       " from uretimisdetayi " + _
                       " where uretimtakipno = '" + cUTF.Trim + "' " + _
                       " and departman = '" + cDept.Trim + "' " + _
                       " and bedenseti = '" + cBedenSeti.Trim + "' " + _
                       " order by isemrino desc "

                If CheckExistsConnected(cSQL, ConnYage) Then

                    cIsEmriNo = SQLGetStringConnected(cSQL, ConnYage)

                    If InStr(cIsEmriNo, "@") > 0 Then
                        ' daha önce ek işemri üretilmiş
                        nIsEmriNo = Val(Mid(cIsEmriNo, Len(cIsEmriNo) - 1, 2))
                    Else
                        nIsEmriNo = 0
                    End If

                    Do While True
                        nIsEmriNo = nIsEmriNo + 1
                        If InStr(cIsEmriNo, "@") > 0 Then
                            cIsEmriNo = Mid(cIsEmriNo, 1, System.Math.Min(Len(cIsEmriNo), 30) - 3) + "@" + Microsoft.VisualBasic.Format(nIsEmriNo, "00")
                        Else
                            cIsEmriNo = Mid(cIsEmriNo, 1, System.Math.Min(Len(cIsEmriNo), 27)) + "@" + Microsoft.VisualBasic.Format(nIsEmriNo, "00")
                        End If

                        cSQL = "select isemrino " + _
                               " from uretimisemri " + _
                               " where isemrino = '" + cIsEmriNo.Trim + "' "

                        If Not CheckExistsConnected(cSQL, ConnYage) Then
                            Exit Do
                        End If
                    Loop
                Else
                    Do While True
                        nIsEmriNo = nIsEmriNo + 1
                        cIsEmriNo = Trim(Mid(cUTF, 1, 27)) + "_" + Microsoft.VisualBasic.Format(nIsEmriNo, "00")

                        cSQL = "select isemrino " + _
                               " from uretimisemri " + _
                               " where isemrino = '" + cIsEmriNo.Trim + "' "

                        If Not CheckExistsConnected(cSQL, ConnYage) Then
                            Exit Do
                        End If
                    Loop
                End If
            End If

            GetLastIsemriNo = cIsEmriNo.Trim

        Catch ex As Exception
            ErrDisp(ex.Message, "GetLastIsemriNo", cSQL)
        End Try
    End Function

    Private Function getuHARlineno(ConnYage As SqlConnection) As Double

        Dim cSQL As String = ""
        Dim nFisNo As Double = 0

        getuHARlineno = 0

        Try
            cSQL = "select uharlineno from sysinfo "
            nFisNo = SQLGetDoubleConnected(cSQL, ConnYage)
            nFisNo = nFisNo + 1

            cSQL = "update sysinfo set uharlineno = " + SQLWriteDecimal(nFisNo)
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            getuHARlineno = nFisNo

        Catch ex As Exception
            ErrDisp(ex.Message, "getuHARlineno", cSQL)
        End Try
    End Function

    Private Sub RecalcUretPlLines(ConnYage As SqlConnection, cUTF As String, Optional cModelNo As String = "", Optional cDepartman As String = "", _
                               Optional cBedenSeti As String = "", Optional cParca As String = "", Optional cDepts As String = "")

        Dim cSQL As String = ""
        Dim aUretPlLines() As oUretPlLines = Nothing
        Dim oReader As SqlDataReader
        Dim nCnt As Integer = -1

        Try
            ' temizle

            cSQL = "select distinct uretimtakipno, modelno, departman " + _
                    " from uretpllines " + _
                    " where not exists (select a.modelno " + _
                                        " from sipmodel a, modeluretim b " + _
                                        " where a.modelno = b.modelno " + _
                                        " and a.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and b.departman = uretpllines.departman) "
            cSQL = cSQL + _
                    " and not exists (select a.uretfisno " + _
                                        " from uretharfis a, uretharfislines b " + _
                                        " where a.uretfisno = b.uretfisno " + _
                                        " and (a.cikisdept = uretpllines.departman or a.girisdept = uretpllines.departman) " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and b.modelno = uretpllines.modelno) "
            cSQL = cSQL + _
                    " and not exists (select a.fofisno " + _
                                        " from fasonorgu a, fasonorgulines b " + _
                                        " where a.fofisno = b.fofisno " + _
                                        " and a.departman = uretpllines.departman " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and b.modelno = uretpllines.modelno) "
            cSQL = cSQL + _
                    " and uretimtakipno = '" + cUTF.Trim + "' " + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                nCnt = nCnt + 1
                ReDim Preserve aUretPlLines(nCnt)
                aUretPlLines(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aUretPlLines(nCnt).cDepartman = SQLReadString(oReader, "departman")
            Loop
            oReader.Close()

            If nCnt > -1 Then
                For nCnt = 0 To UBound(aUretPlLines)
                    cSQL = "delete from uretimisdetayi " + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and departman = '" + aUretPlLines(nCnt).cDepartman + "' " + _
                            " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete from uretimisrba " + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and departman = '" + aUretPlLines(nCnt).cDepartman + "' " + _
                            " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete from uretimisrbapos " + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and departman = '" + aUretPlLines(nCnt).cDepartman + "' " + _
                            " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete from uretimisemri " + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and departman = '" + aUretPlLines(nCnt).cDepartman + "' " + _
                            " and not exists (select isemrino " + _
                                            " from uretimisdetayi " + _
                                            " where isemrino = uretimisemri.isemrino " + _
                                            " and uretimtakipno = uretimisemri.uretimtakipno) "
                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete from uretimisemrifis " + _
                            " where uretimtakipno = '" + cUTF.Trim + "'  " + _
                            " and not exists (select uretimtakipno " + _
                                            " from uretimisemri " + _
                                            " where uretimtakipno = uretimisemrifis.uretimtakipno) "
                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    cSQL = "delete uretpllines " + _
                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                            " and departman = '" + aUretPlLines(nCnt).cDepartman + "' " + _
                            " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Next
            End If

            cSQL = "update uretpllines set " + _
                    " fiyati = (select top 1 coalesce(iscilikfiyat,0) " + _
                                " from modeluretim " + _
                                " where modelno = uretpllines.modelno " + _
                                " and departman = uretpllines.departman) , " + _
                    " doviz = (select top 1 coalesce(iscilikdoviz,'TL') " + _
                                " from modeluretim " + _
                                " where modelno = uretpllines.modelno " + _
                                " and departman = uretpllines.departman) "
            cSQL = cSQL + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (fiyati is null or fiyati = 0) " + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretpllines " + _
                    " set avanssuresi = (select top 1 avanssuresi " + _
                                        " from departman " + _
                                        " where departman = uretpllines.departman " + _
                                        " and avanssuresi is not null " + _
                                        " and avanssuresi > 0 ) "
            cSQL = cSQL + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (avanssuresi is null or avanssuresi = 0)  " + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretpllines " + _
                    " set oktipi = (select top 1 oktipi " + _
                                        " from departman " + _
                                        " where departman = uretpllines.departman " + _
                                        " and oktipi is not null " + _
                                        " and oktipi <> '') "
            cSQL = cSQL + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (oktipi is null or oktipi = '') " + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretpllines " + _
                   " set isemriverilen = (select sum (coalesce(adet,0)) " + _
                                        " from uretimisrba " + _
                                        " where uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and modelno = uretpllines.modelno " + _
                                        " and departman = uretpllines.departman " + _
                                        " and bedenseti = uretpllines.bedenseti " + _
                                        " and parca = uretpllines.parca), "
            cSQL = cSQL + _
                   " toplamadet = (select sum (coalesce(adet,0)) " + _
                                        " from uretplrba " + _
                                        " where uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and modelno = uretpllines.modelno " + _
                                        " and departman = uretpllines.departman " + _
                                        " and bedenseti = uretpllines.bedenseti " + _
                                        " and parca = uretpllines.parca), "
            cSQL = cSQL + _
                    " gelen = (select sum (coalesce(toplamadet,0)) " + _
                                        " from uretharfislines b, uretharfis a " + _
                                        " where a.uretfisno = b.uretfisno " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and a.girisdept = uretpllines.departman " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and b.bedenseti = uretpllines.bedenseti), "
            cSQL = cSQL + _
                    " giden = (select sum (coalesce(toplamadet,0)) " + _
                                        " from uretharfislines b, uretharfis a " + _
                                        " where a.uretfisno = b.uretfisno " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and a.cikisdept = uretpllines.departman " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and b.bedenseti = uretpllines.bedenseti), "
            cSQL = cSQL + _
                    " parcagelen = (select sum (coalesce(toplamadet,0)) " + _
                                    " from uretharfislines b, uretharfis a " + _
                                    " where a.uretfisno = b.uretfisno " + _
                                    " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                    " and  a.girisdept = uretpllines.departman " + _
                                    " and b.modelno = uretpllines.modelno " + _
                                    " and b.bedenseti = uretpllines.bedenseti " + _
                                    " and b.parca <> 'KOMPLE'), "
            cSQL = cSQL +
                    " parcagiden = (select sum (coalesce(toplamadet,0)) " +
                                    " from uretharfislines b, uretharfis a " +
                                    " where a.uretfisno = b.uretfisno " +
                                    " and b.uretimtakipno = uretpllines.uretimtakipno " +
                                    " and a.cikisdept = uretpllines.departman " +
                                    " and b.modelno = uretpllines.modelno " +
                                    " and b.bedenseti = uretpllines.bedenseti " +
                                    " and b.parca <> 'KOMPLE') "

            cSQL = cSQL + _
                   " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from uretpllines " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and toplamadet = 0 " + _
                    " and isemriverilen = 0 " + _
                    " and gelen = 0 " + _
                    " and parcagelen = 0 " + _
                    " and parcagiden = 0 " + _
                    " and giden = 0 " + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' hedef işçilik fiyatlarını dokuma ön maliyetten al

            cSQL = "update uretpllines " +
                    " set fiyati = (select top 1 fiyat  " +
                            " from maliyeturetim " +
                            " where islemler = uretpllines.departman  " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.modelno = uretpllines.modelno " +
                                                " and a.uretimtakipno = uretpllines.uretimtakipno) " +
                            " order by fiyat), "
            cSQL = cSQL +
                    " doviz = (select top 1 doviz " +
                            " from maliyeturetim " +
                            " where islemler = uretpllines.departman  " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.modelno = uretpllines.modelno " +
                                                " and a.uretimtakipno = uretpllines.uretimtakipno) " +
                            " order by fiyat), "
            cSQL = cSQL +
                    " plfirma = (select top 1 firma " +
                            " from maliyeturetim " +
                            " where islemler = uretpllines.departman  " +
                            " and fiyat is not null " +
                            " and fiyat <> 0 " +
                            " and calismano in (select b.maliyetcalismano " +
                                                " from sipmodel a, ymodel b " +
                                                " where a.modelno = b.modelno " +
                                                " and a.modelno = uretpllines.modelno " +
                                                " and a.uretimtakipno = uretpllines.uretimtakipno) " +
                            " order by fiyat) "
            cSQL = cSQL +
                   " where uretimtakipno = '" + cUTF.Trim + "' " +
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString +
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString +
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString +
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString +
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString +
                    " and (fiyati is null or fiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' hedef işçilik fiyatlarını model kartından al

            cSQL = "update uretpllines " +
                    " set fiyati = (select top 1 iscilikfiyat  " +
                            " from modeluretim " +
                            " where departman = uretpllines.departman  " +
                            " and iscilikfiyat is not null " +
                            " and iscilikfiyat <> 0 " +
                            " and modelno = uretpllines.modelno " +
                            " order by iscilikfiyat), "
            cSQL = cSQL +
                    " doviz = (select top 1 iscilikdoviz  " +
                            " from modeluretim " +
                            " where departman = uretpllines.departman  " +
                            " and iscilikfiyat is not null " +
                            " and iscilikfiyat <> 0 " +
                            " and modelno = uretpllines.modelno " +
                            " order by iscilikfiyat), "
            cSQL = cSQL +
                    " plfirma = (select top 1 firma  " +
                            " from modeluretim " +
                            " where departman = uretpllines.departman  " +
                            " and iscilikfiyat is not null " +
                            " and iscilikfiyat <> 0 " +
                            " and modelno = uretpllines.modelno " +
                            " order by iscilikfiyat) "
            cSQL = cSQL +
                   " where uretimtakipno = '" + cUTF.Trim + "' " +
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString +
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString +
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString +
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString +
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString +
                    " and (fiyati is null or fiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message, "RecalcUretPlLines", cSQL)
        End Try
    End Sub

    Private Function DeleteSingleUretimIsemri(ConnYage As SqlConnection, Optional cUTF As String = "", Optional cIsEmriNo As String = "", Optional cDepartman As String = "", _
                                                Optional cModelNo As String = "", Optional cBedenSeti As String = "", Optional cParca As String = "", Optional cDepts As String = "", _
                                                Optional lRecalc As Boolean = True) As Boolean

        Dim cFilter As String = ""
        Dim cFilter2 As String = ""
        Dim cFilter3 As String = ""
        Dim cSQL As String = ""

        DeleteSingleUretimIsemri = False

        Try
            If cUTF.Trim = "" Then Exit Function

            cFilter = " isemrino is not null " + _
                        IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString + _
                        IIf(cIsEmriNo.Trim = "", "", " and isemrino = '" + cIsEmriNo.Trim + "' ").ToString + _
                        IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString

            cFilter2 = " isemrino is not null " + _
                        IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString + _
                        IIf(cIsEmriNo.Trim = "", "", " and isemrino = '" + cIsEmriNo.Trim + "' ").ToString + _
                        IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                        IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString() + _
                        IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString() + _
                        IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString()

            cFilter3 = " isemrino is not null " + _
                        IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString + _
                        IIf(cIsEmriNo.Trim = "", "", " and isemrino = '" + cIsEmriNo.Trim + "' ").ToString + _
                        IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                        IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString()

            cSQL = "delete from uretimisemri      where " + cFilter
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from uretimisdetayi    where " + cFilter2
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from uretimisrba       where " + cFilter2
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from uretimisrbapos    where " + cFilter2
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete from uretimisadethedef where " + cFilter3
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            If lRecalc Then
                RecalcUretimIsemriDetails(ConnYage, cUTF, cIsEmriNo, cDepartman, cModelNo, cBedenSeti, cParca, cDepts)
                RecalcUretPlLines(ConnYage, cUTF, cModelNo, cDepartman, cBedenSeti, cParca, cDepts)
            End If

            DeleteSingleUretimIsemri = True

        Catch ex As Exception
            ErrDisp(ex.Message, "DeleteSingleUretimIsemri", cSQL)
        End Try
    End Function

    Private Sub RecalcUretimIsemriDetails(ConnYage As SqlConnection, Optional cUTF As String = "", Optional cIsEmriNo As String = "", Optional cDepartman As String = "", _
                                                Optional cModelNo As String = "", Optional cBedenSeti As String = "", Optional cParca As String = "", Optional cDepts As String = "")
        Dim cSQL As String = ""

        Try
            cSQL = "update uretimisdetayi " + _
                   " set toplamadet = (select sum (coalesce(adet,0)) " + _
                                       " from uretimisrba " + _
                                       " where isemrino = uretimisdetayi.isemrino " + _
                                       " and uretimtakipno = uretimisdetayi.uretimtakipno " + _
                                       " and modelno = uretimisdetayi.modelno " + _
                                       " and departman = uretimisdetayi.departman " + _
                                       " and bedenseti = uretimisdetayi.bedenseti " + _
                                       " and parca = uretimisdetayi.parca " + _
                                       " and ulineno = uretimisdetayi.ulineno) " + _
                    " where  isemrino is not null " + _
                    IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString + _
                    IIf(cIsEmriNo.Trim = "", "", " and isemrino = '" + cIsEmriNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretimisdetayi " + _
                    " set sozlesmeadedi = coalesce(toplamadet,0) " + _
                    " where  isemrino is not null " + _
                    IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString + _
                    IIf(cIsEmriNo.Trim = "", "", " and isemrino = '" + cIsEmriNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString() + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString() + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString + _
                    " and (sozlesmeadedi is null or sozlesmeadedi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretimisdetayi " + _
                    " set baslama_tar = (select top 1 baslamatarihi " + _
                                    " from uretpllines " + _
                                    " where uretimtakipno = uretimisdetayi.uretimtakipno " + _
                                    " and modelno = uretimisdetayi.modelno " + _
                                    " and departman = uretimisdetayi.departman " + _
                                    " and bedenseti = uretimisdetayi.bedenseti " + _
                                    " and parca = uretimisdetayi.parca " + _
                                    " and baslamatarihi is not null " + _
                                    " and baslamatarihi <> '01.01.1950'), "

            cSQL = cSQL + _
                    " bitis_tar =  (select top 1 bitistarihi " + _
                                    " from uretpllines " + _
                                    " where uretimtakipno = uretimisdetayi.uretimtakipno " + _
                                    " and modelno = uretimisdetayi.modelno " + _
                                    " and departman = uretimisdetayi.departman " + _
                                    " and bedenseti = uretimisdetayi.bedenseti " + _
                                    " and parca = uretimisdetayi.parca " + _
                                    " and bitistarihi is not null " + _
                                    " and bitistarihi <> '01.01.1950') "

            cSQL = cSQL + _
                    " where  isemrino is not null " + _
                    IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString + _
                    IIf(cIsEmriNo.Trim = "", "", " and isemrino = '" + cIsEmriNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString() + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString() + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString + _
                    " and (bitis_tar is null or bitis_tar = '01.01.1950') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretimisdetayi " + _
                    " set fiyati = (select top 1 fiyati " + _
                                    " from uretpllines " + _
                                    " where uretimtakipno = uretimisdetayi.uretimtakipno " + _
                                    " and modelno = uretimisdetayi.modelno " + _
                                    " and departman = uretimisdetayi.departman " + _
                                    " and bedenseti = uretimisdetayi.bedenseti " + _
                                    " and parca = uretimisdetayi.parca " + _
                                    " and fiyati is not null " + _
                                    " and fiyati <> 0), "
            cSQL = cSQL + _
                    " doviz = (select top 1 doviz " + _
                                    " from uretpllines " + _
                                    " where uretimtakipno = uretimisdetayi.uretimtakipno " + _
                                    " and modelno = uretimisdetayi.modelno " + _
                                    " and departman = uretimisdetayi.departman " + _
                                    " and bedenseti = uretimisdetayi.bedenseti " + _
                                    " and parca = uretimisdetayi.parca " + _
                                    " and fiyati is not null " + _
                                    " and fiyati <> 0) "
            cSQL = cSQL + _
                    " where  isemrino is not null " + _
                    IIf(cUTF.Trim = "", "", " and uretimtakipno = '" + cUTF.Trim + "' ").ToString + _
                    IIf(cIsEmriNo.Trim = "", "", " and isemrino = '" + cIsEmriNo.Trim + "' ").ToString + _
                    IIf(cDepartman.Trim = "", "", " and departman = '" + cDepartman.Trim + "' ").ToString + _
                    IIf(cModelNo.Trim = "", "", " and modelno = '" + cModelNo.Trim + "' ").ToString() + _
                    IIf(cBedenSeti.Trim = "", "", " and bedenseti = '" + cBedenSeti.Trim + "' ").ToString() + _
                    IIf(cParca.Trim = "", "", " and parca = '" + cParca.Trim + "' ").ToString + _
                    IIf(cDepts = "", "", " and departman in (" + cDepts + ") ").ToString + _
                    " and (fiyati is null or fiyati = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message, "RecalcUretimIsemriDetails", cSQL)
        End Try
    End Sub

End Module

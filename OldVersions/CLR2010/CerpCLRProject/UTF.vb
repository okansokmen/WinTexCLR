Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module UTF
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

            cSQL = "select distinct a.modelno, a.bedenseti, " + _
                    " b.departman, b.uretimtoleransi, b.giristakipesasi, b.cikistakipesasi, b.parca, b.sira, b.girisdepartmani, b.cikisdepartmani, b.girisparcasi " + _
                    " from " + cSipModelTableName + " a, modeluretim b " + _
                    " where a.modelno = b.modelno " + _
                    " and a.uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and a.modelno is not null " + _
                    " and a.modelno <> '' " + _
                    " and a.bedenseti is not null " + _
                    " and a.bedenseti <> '' " + _
                    " and a.departman is not null " + _
                    " and a.departman <> '' " + _
                    " order by a.modelno, a.bedenseti, b.sira "

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

                If aUretPlLines(nCnt).cParca = "" Then aUretPlLines(nCnt).cParca = "KOMPLE"
                If aUretPlLines(nCnt).cGirisTakipEsasi = "" Then aUretPlLines(nCnt).cGirisTakipEsasi = "4"
                If aUretPlLines(nCnt).cCikisTakipEsasi = "" Then aUretPlLines(nCnt).cCikisTakipEsasi = "4"

                nCnt = nCnt + 1
            Loop
            oReader.Close()

            For nCnt = 0 To UBound(aUretPlLines)
                cSQL = "select uretimtakipno " + _
                        " from uretpllines " + _
                        " where uretimtakipno = '" + cUTF.Trim + "' " + _
                        " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " + _
                        " and departman = '" + aUretPlLines(nCnt).cDepartman + "' " + _
                        " and parca = '" + aUretPlLines(nCnt).cParca + "' " + _
                        " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' "

                If CheckExistsConnected(cSQL, ConnYage) Then
                    cSQL = "update uretpllines " + _
                            " set sira = " + SQLWriteDecimal(aUretPlLines(nCnt).nSira) + ", " + _
                            " girisdepartmani = '" + aUretPlLines(nCnt).cGirisDepartmani + "', " + _
                            " cikisdepartmani = '" + aUretPlLines(nCnt).cCikisDepartmani + "', " + _
                            " girisparcasi = '" + aUretPlLines(nCnt).cGirisParcasi + "', " + _
                            " giristakipesasi = '" + aUretPlLines(nCnt).cGirisTakipEsasi + "', " + _
                            " cikistakipesasi = '" + aUretPlLines(nCnt).cCikisTakipEsasi + "' " + _
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
                            " parcagelen, parcagiden, bedenseti, gelenparcacount) "

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
                            " 0,0,'" + aUretPlLines(nCnt).cBedenSeti + "',1)"

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If

                Select Case aUretPlLines(nCnt).cCikisTakipEsasi
                    Case "1"
                        cSQL = "select sum(coalesce(adet,0)) " + _
                                " from " + cSipModelTableName + _
                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " + _
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' "

                        nAdet = SQLGetDoubleConnected(cSQL, ConnYage)

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
                        cSQL = "insert uretplrba (uretimtakipno, modelno, bedenseti, departman, parca, renk, beden, adet) " + _
                                " select uretimtakipno, modelno, bedenseti, " + _
                                " departman = '" + aUretPlLines(nCnt).cDepartman + "',  " + _
                                " parca = '" + aUretPlLines(nCnt).cParca + "',  " + _
                                " renk, " + _
                                " beden = 'HEPSI',  " + _
                                " adet = sum(coalesce(adet,0)) " + _
                                " from " + cSipModelTableName + _
                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " + _
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' " + _
                                " group by renk " + _
                                " order by renk "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                    Case "3"
                        cSQL = "insert uretplrba (uretimtakipno, modelno, bedenseti, departman, parca, renk, beden, adet) " + _
                                " select uretimtakipno, modelno, bedenseti, " + _
                                " departman = '" + aUretPlLines(nCnt).cDepartman + "',  " + _
                                " parca = '" + aUretPlLines(nCnt).cParca + "',  " + _
                                " renk = 'HEPSI', " + _
                                " beden " + _
                                " adet = sum(coalesce(adet,0)) " + _
                                " from  " + cSipModelTableName + _
                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " + _
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' " + _
                                " group by beden " + _
                                " order by beden "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)

                    Case Else
                        cSQL = "insert uretplrba (uretimtakipno, modelno, bedenseti, departman, parca, renk, beden, adet) " + _
                                " select uretimtakipno, modelno, bedenseti, " + _
                                " departman = '" + aUretPlLines(nCnt).cDepartman + "',  " + _
                                " parca = '" + aUretPlLines(nCnt).cParca + "',  " + _
                                " renk, " + _
                                " beden,  " + _
                                " adet = sum(coalesce(adet,0)) " + _
                                " from " + cSipModelTableName + _
                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                " and modelno = '" + aUretPlLines(nCnt).cModelNo + "' " + _
                                " and bedenseti = '" + aUretPlLines(nCnt).cBedenSeti + "' " + _
                                " group by renk, beden " + _
                                " order by renk, beden "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                End Select
            Next

            cSQL = "update uretpllines " + _
                    " set toplamadet = (select sum(coalesce(adet,0)) " + _
                                    " from uretplrba " + _
                                    " where uretimtakipno = uretpllines.uretimtakipno " + _
                                    " and departman = uretpllines.departman " + _
                                    " and modelno = uretpllines.modelno " + _
                                    " and bedenseti = uretpllines.bedenseti " + _
                                    " and parca = uretpllines.parca), "
            cSQL = cSQL + _
                    " isemriverilen = (select sum (coalesce(toplamadet,0)) " + _
                                    " from uretimisdetayi " + _
                                    " where uretimtakipno = uretpllines.uretimtakipno " + _
                                    " and departman = uretpllines.departman " + _
                                    " and modelno = uretpllines.modelno " + _
                                    " and bedenseti = uretpllines.bedenseti " + _
                                    " and parca = uretpllines.parca), "
            cSQL = cSQL + _
                    " gelen = (select sum (coalesce(b.toplamadet,0)) " + _
                                    " from uretharfislines b, uretharfis a " + _
                                    " where a.uretfisno = b.uretfisno  " + _
                                    " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                    " and a.girisdept = uretpllines.departman " + _
                                    " and b.modelno = uretpllines.modelno " + _
                                    " and b.bedenseti = uretpllines.bedenseti), "
            cSQL = cSQL + _
                    " giden = (select sum (coalesce(b.toplamadet,0)) " + _
                                    " from uretharfislines b, uretharfis a " + _
                                    " where a.uretfisno = b.uretfisno  " + _
                                    " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                    " and a.cikisdept = uretpllines.departman " + _
                                    " and b.modelno = uretpllines.modelno " + _
                                    " and b.bedenseti = uretpllines.bedenseti), "
            cSQL = cSQL + _
                    " parcagelen = (select sum (coalesce(b.toplamadet,0)) " + _
                                    " from uretharfislines b, uretharfis a " + _
                                    " where a.uretfisno = b.uretfisno  " + _
                                    " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                    " and a.girisdept = uretpllines.departman " + _
                                    " and b.modelno = uretpllines.modelno " + _
                                    " and b.bedenseti = uretpllines.bedenseti " + _
                                    " and b.parca <> 'KOMPLE'), "
            cSQL = cSQL + _
                    " parcagiden = (select sum (coalesce(b.toplamadet,0)) " + _
                                    " from uretharfislines b, uretharfis a " + _
                                    " where a.uretfisno = b.uretfisno  " + _
                                    " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                    " and a.cikisdept = uretpllines.departman " + _
                                    " and b.modelno = uretpllines.modelno " + _
                                    " and b.bedenseti = uretpllines.bedenseti " + _
                                    " and b.parca <> 'KOMPLE') "
            cSQL = cSQL + _
                    " where uretimtakipno = '" + cUTF.Trim + "' "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            ' fiyat / doviz
            cSQL = "update uretpllines " + _
                    " set fiyati = (select top 1 coalesce(iscilikfiyat,0) from modeluretim WHERE modelno = uretpllines.modelno and departman = uretpllines.departman) , " + _
                    " doviz = (select top 1 coalesce(iscilikdoviz,'') from modeluretim WHERE modelno = uretpllines.modelno and departman = uretpllines.departman) " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (fiyati is null or fiyati = 0)"

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretpllines " + _
                    " set teminsuresi = (select top 1 teminsuresi " + _
                                        " from departman " + _
                                        " where departman = uretpllines.departman " + _
                                        " and teminsuresi is not null " + _
                                        " and teminsuresi > 0 ) " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (teminsuresi is null or teminsuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretpllines " + _
                    " set avanssuresi = (select top 1 avanssuresi " + _
                                        " from departman " + _
                                        " where departman = uretpllines.departman " + _
                                        " and avanssuresi is not null " + _
                                        " and avanssuresi > 0 ) " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (avanssuresi is null or avanssuresi = 0) "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretpllines " + _
                    " set oktipi = (select top 1 oktipi " + _
                                        " from departman " + _
                                        " where departman = uretpllines.departman " + _
                                        " and oktipi is not null " + _
                                        " and oktipi <> '' ) " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (oktipi is null or oktipi = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update uretpllines " + _
                    " set plfirma = 'DAHILI' " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and (plfirma is null or plfirma = '') "

            ExecuteSQLCommandConnected(cSQL, ConnYage)

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

End Module

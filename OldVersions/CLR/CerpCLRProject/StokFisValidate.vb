Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module StokFisValidate

    Private Structure StokFisLines
        Dim cStokNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cDepo As String
        Dim cMTF As String
        Dim cPartiNo As String
        Dim cTopNo As String
        Dim cStokHareketKodu As String
        Dim nNetMiktar1 As Double
        Dim nNetMiktar2 As Double
        Dim nNetMiktar3 As Double
        Dim nBrutMiktar1 As Double
        Dim nBrutMiktar2 As Double
        Dim nBrutMiktar3 As Double
        Dim nFiyat As Double
        Dim cDoviz As String
        Dim nIscilik As Double
        Dim cIDoviz As String
        Dim nReferansBirimFiyat As Double
        Dim nReferansIscilik As Double
        Dim cUTF As String
        Dim cModelNo As String
        Dim cParca As String
        Dim cUretDept As String
        Dim cUretIsemriNo As String
        Dim cIsemriNo As String
        Dim nAgirlik As Double
    End Structure

    Public Function MultiStokFisValidate(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String) As SqlInt32

        Dim cSQL As String
        Dim aStokFisNo() As String
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim ConnYage As SqlConnection
        Dim lFound As Boolean = False

        MultiStokFisValidate = 0

        Try
            ConnYage = OpenConn()

            nCnt = 0
            ReDim aStokFisNo(0)

            cSQL = "Select distinct a.stokfisno " + _
                    " from stokfis a, stokfislines b " + _
                    " where a.stokfisno = b.stokfisno " + _
                    IIf(cStokNo = "", "", " and b.stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and b.renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and b.beden = '" + cBeden + "' ").ToString

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aStokFisNo(nCnt)
                aStokFisNo(nCnt) = SQLReadString(oReader, "stokfisno")
                lFound = True
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

            MultiStokFisValidate = 1

            If lFound Then
                For nCnt = 0 To UBound(aStokFisNo)
                    If aStokFisNo(nCnt).Trim <> "" Then
                        If SingleStokFisValidate("validate", aStokFisNo(nCnt), cStokNo, cRenk, cBeden) = 0 Then
                            MultiStokFisValidate = 0
                            Exit For
                        End If
                    End If
                Next
            End If

        Catch
            MultiStokFisValidate = 0
            ErrDisp("Error MultiStokFisValidate " + Err.Description.Trim)
        End Try
    End Function

    Public Function SingleStokFisValidate(ByVal cAction As String, ByVal cStokFisNo As String, ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String) As SqlInt32

        Dim ConnYage As SqlConnection

        SingleStokFisValidate = 0

        Try
            ConnYage = OpenConn()
            SingleStokFisValidate = SingleStokFisValidateConnected(ConnYage, cAction, cStokFisNo, cStokNo, cRenk, cBeden)
            CloseConn(ConnYage)
        Catch
            SingleStokFisValidate = 0
            ErrDisp("Error SingleStokFisValidate " + Err.Description.Trim)
        End Try
    End Function

    Public Function SingleStokFisValidateConnected(ByVal ConnYage As SqlConnection, ByVal cAction As String, ByVal cStokFisNo As String, ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String) As SqlInt32

        Dim oReader As SqlDataReader
        Dim cSQL As String = ""

        Dim nTLFiyat As Double
        Dim nTLIscilik As Double
        Dim nTLTutar As Double
        Dim nFiyatKur As Double
        Dim nIscilikKur As Double
        Dim nDvzFiyat As Double
        Dim cStokFisTipi As String = ""
        Dim cDepartman As String = ""
        Dim cFirma As String = ""
        Dim dTarih As Date
        Dim cFilter As String
        Dim oStokRB As New stokrb
        Dim oSysFlags As New General.SysFlags
        Dim lIsemri As Boolean = False
        Dim nIsEmriCnt As Integer
        Dim aFis() As StokFisLines
        Dim nCnt As Integer
        Dim nCnt2 As Integer
        Dim lFoundIsemri As Boolean
        Dim lFound As Boolean = False
        Dim nMlytTutar As Double
        Dim nKur As Double
        Dim aIsEmri() As String
        Dim cStatus As String = "Critical"
        Dim cMTF As String = ""

        SingleStokFisValidateConnected = 0

        Try
            JustForLog("StokFis-" + cAction.Trim + "-Begin " + cStokFisNo)

            If cStokFisNo.Trim = "" Then
                SingleStokFisValidateConnected = 1
                JustForLog("StokFis-" + cAction.Trim + "-End-EmptyStokFisno")
                Exit Function
            End If

            ReDim aIsEmri(0)
            nIsEmriCnt = 0

            cAction = LCase(cAction.Trim)

            ReDim aFis(0)

            ReadSysFlags(oSysFlags, ConnYage)

            dTarih = Today

            cSQL = "Select StokFisTipi, Departman, Firma, FisTarihi " + _
                    " From StokFis " + _
                    " where StokFisNo = '" + cStokFisNo.Trim + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read() Then
                cStokFisTipi = LCase(SQLReadString(oReader, "stokfistipi"))
                cDepartman = SQLReadString(oReader, "departman")
                cFirma = SQLReadString(oReader, "firma")
                dTarih = SQLReadDate(oReader, "fistarihi")
                lFound = True
            End If
            oReader.Close()
            oReader = Nothing

            If Not lFound Then
                SingleStokFisValidateConnected = 1
                JustForLog("StokFis-" + cAction.Trim + "-End-NoHeader " + cStokFisNo)
                Exit Function
            End If

            cStatus = "Begin Read Master Record"

            lFound = False
            nCnt = 0
            cSQL = "select stokno,renk,beden,depo,malzemetakipkodu,partino,topno,stokhareketkodu,NetMiktar1,NetMiktar2,NetMiktar3, " + _
                    " BrutMiktar1,BrutMiktar2,BrutMiktar3,birimfiyat,dovizcinsi,iscilikfiyat,iscilikdoviz,referansbirimfiyat,referansiscilik, " + _
                    " uretimtakipno,modelno,parca,departman,urtisemrino,isemrino,agirlik " + _
                    " from StokFisLines " + _
                    " where stokfisno = '" + cStokFisNo + "' " + _
                    " and stokno is not null " + _
                    " and stokno <> '' " + _
                    IIf(cStokNo = "", "", " and stokno = '" + cStokNo.Trim + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk.Trim + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden.Trim + "' ").ToString

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                ReDim Preserve aFis(nCnt)

                aFis(nCnt).cStokNo = SQLReadString(oReader, "stokno")
                aFis(nCnt).cRenk = SQLReadString(oReader, "renk")
                aFis(nCnt).cBeden = SQLReadString(oReader, "beden")
                aFis(nCnt).cDepo = SQLReadString(oReader, "depo")
                aFis(nCnt).cMTF = SQLReadString(oReader, "malzemetakipkodu")
                aFis(nCnt).cPartiNo = SQLReadString(oReader, "partino")
                aFis(nCnt).cTopNo = SQLReadString(oReader, "topno")
                aFis(nCnt).cStokHareketKodu = SQLReadString(oReader, "stokhareketkodu")
                aFis(nCnt).nNetMiktar1 = SQLReadDouble(oReader, "NetMiktar1")
                aFis(nCnt).nNetMiktar2 = SQLReadDouble(oReader, "NetMiktar2")
                aFis(nCnt).nNetMiktar3 = SQLReadDouble(oReader, "NetMiktar3")
                aFis(nCnt).nBrutMiktar1 = SQLReadDouble(oReader, "BrutMiktar1")
                aFis(nCnt).nBrutMiktar2 = SQLReadDouble(oReader, "BrutMiktar2")
                aFis(nCnt).nBrutMiktar3 = SQLReadDouble(oReader, "BrutMiktar3")
                aFis(nCnt).nFiyat = SQLReadDouble(oReader, "birimfiyat")
                aFis(nCnt).cDoviz = SQLReadString(oReader, "dovizcinsi")
                aFis(nCnt).nIscilik = SQLReadDouble(oReader, "iscilikfiyat")
                aFis(nCnt).cIDoviz = SQLReadString(oReader, "iscilikdoviz")
                aFis(nCnt).nReferansBirimFiyat = SQLReadDouble(oReader, "referansbirimfiyat")
                aFis(nCnt).nReferansIscilik = SQLReadDouble(oReader, "referansiscilik")
                aFis(nCnt).cUTF = SQLReadString(oReader, "uretimtakipno")
                aFis(nCnt).cModelNo = SQLReadString(oReader, "modelno")
                aFis(nCnt).cParca = SQLReadString(oReader, "parca")
                aFis(nCnt).cUretDept = SQLReadString(oReader, "departman")
                aFis(nCnt).cUretIsemriNo = SQLReadString(oReader, "urtisemrino")
                aFis(nCnt).cIsemriNo = SQLReadString(oReader, "isemrino")
                aFis(nCnt).nAgirlik = SQLReadDouble(oReader, "agirlik")

                lFound = True
                If aFis(nCnt).cIsemriNo <> "" Then
                    lIsemri = True
                    lFoundIsemri = False
                    For nCnt2 = 0 To UBound(aIsEmri)
                        If aFis(nCnt).cIsemriNo = aIsEmri(nCnt2) Then
                            lFoundIsemri = True
                            Exit For
                        End If
                    Next
                    If Not lFoundIsemri Then
                        ReDim Preserve aIsEmri(nIsEmriCnt)
                        aIsEmri(nIsEmriCnt) = aFis(nCnt).cIsemriNo
                        nIsEmriCnt = nIsEmriCnt + 1
                    End If
                End If
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            cStatus = "End Read Master Record"

            If Not lFound Then
                SingleStokFisValidateConnected = 1
                JustForLog("StokFis-" + cAction.Trim + "-End-NoDetail " + cStokFisNo)
                Exit Function
            End If

            For nCnt = 0 To UBound(aFis)

                cStatus = "Satir No : " + nCnt.ToString

                nTLFiyat = 0
                nTLIscilik = 0
                nFiyatKur = GetKur(aFis(nCnt).cDoviz, dTarih, ConnYage)
                nIscilikKur = GetKur(aFis(nCnt).cIDoviz, dTarih, ConnYage)
                nDvzFiyat = 0

                If aFis(nCnt).nFiyat <> 0 Then
                    nTLFiyat = aFis(nCnt).nFiyat * nFiyatKur
                End If

                If aFis(nCnt).nNetMiktar1 <> 0 And aFis(nCnt).nIscilik <> 0 Then
                    aFis(nCnt).nIscilik = aFis(nCnt).nIscilik / aFis(nCnt).nNetMiktar1
                    nTLIscilik = aFis(nCnt).nIscilik * nIscilikKur
                End If

                nTLTutar = (nTLFiyat + nTLIscilik) * aFis(nCnt).nNetMiktar1

                If nFiyatKur <> 0 Then
                    nDvzFiyat = (nTLFiyat + nTLIscilik) / nFiyatKur
                End If

                nMlytTutar = nTLTutar
                nKur = GetKur(oSysFlags.G_OrtMlytDoviz, dTarih, ConnYage)

                If nKur > 0 Then
                    nMlytTutar = nMlytTutar / nKur
                End If

                CheckRenkBedenValidate(aFis(nCnt).cStokNo, aFis(nCnt).cRenk, aFis(nCnt).cBeden, ConnYage)

                oStokRB.cTableName = "stokrb"
                oStokRB.cStokno = aFis(nCnt).cStokNo
                oStokRB.cRenk = aFis(nCnt).cRenk
                oStokRB.cBeden = aFis(nCnt).cBeden
                oStokRB.cPartiNo = aFis(nCnt).cPartiNo
                oStokRB.cDepo = aFis(nCnt).cDepo
                oStokRB.cMtk = aFis(nCnt).cMTF
                oStokRB.cTopNo = aFis(nCnt).cTopNo

                cMTF = aFis(nCnt).cMTF

                If LCase(aFis(nCnt).cStokHareketKodu) = "02 satis iade" Or LCase(aFis(nCnt).cStokHareketKodu) = "03 defolu iade" Then cStokFisTipi = "giris"
                If LCase(aFis(nCnt).cStokHareketKodu) = "01 satis" Then cStokFisTipi = "cikis"

                ' stok
                If UpdateStokRB(ConnYage, cStokFisTipi, cAction, "stok", aFis(nCnt).cStokNo, aFis(nCnt).nNetMiktar1, aFis(nCnt).nNetMiktar2, aFis(nCnt).nNetMiktar3) = 0 Then
                    JustForLog("StokFis-" + cAction.Trim + "-End-Exit Function Stok " + cStokFisNo)
                    SingleStokFisValidateConnected = 0
                    Exit Function
                End If
                cStatus = "Stok table update OK"

                ' stokrb 
                cFilter = CheckInsertStokRB(oStokRB, oSysFlags, ConnYage)
                If UpdateStokRB(ConnYage, cStokFisTipi, cAction, "stokrb", aFis(nCnt).cStokNo, aFis(nCnt).nNetMiktar1, aFis(nCnt).nNetMiktar2, aFis(nCnt).nNetMiktar3, _
                             cFilter, aFis(nCnt).nAgirlik, dTarih, aFis(nCnt).nFiyat, aFis(nCnt).cDoviz, cDepartman, cFirma) = 0 Then
                    JustForLog("StokFis-" + cAction.Trim + "-End-Exit Function StokRB " + cStokFisNo)
                    SingleStokFisValidateConnected = 0
                    Exit Function
                End If
                cStatus = "StokRB table update OK"

                ' barkodlu kumaş 
                If G_BarkodluKumas(aFis(nCnt).cStokNo, oSysFlags, ConnYage) Then
                    oStokRB.cTableName = "stoktoprb"
                    If oSysFlags.G_WFNoMTK Then
                        oStokRB.cMtk = ""
                    End If
                    cFilter = CheckInsertStokRB(oStokRB, oSysFlags, ConnYage)
                    If UpdateStokRB(ConnYage, cStokFisTipi, cAction, "stoktoprb", aFis(nCnt).cStokNo, aFis(nCnt).nNetMiktar1, aFis(nCnt).nNetMiktar2, aFis(nCnt).nNetMiktar3, _
                                 cFilter, aFis(nCnt).nAgirlik, dTarih, aFis(nCnt).nFiyat, aFis(nCnt).cDoviz, cDepartman, cFirma) = 0 Then
                        JustForLog("StokFis-" + cAction.Trim + "-End-Exit Function stoktoprb " + cStokFisNo)
                        SingleStokFisValidateConnected = 0
                        Exit Function
                    End If
                    oStokRB.cMtk = cMTF
                    cStatus = "StokTopRB table update OK"
                End If

                ' barkodlu aksesuar 
                If G_BarkodluAksesuar(aFis(nCnt).cStokNo, oSysFlags, ConnYage) Then
                    oStokRB.cTableName = "stokaksesuarrb"
                    If oSysFlags.G_WANoMTK Then
                        oStokRB.cMtk = ""
                    End If
                    cFilter = CheckInsertStokRB(oStokRB, oSysFlags, ConnYage)
                    If UpdateStokRB(ConnYage, cStokFisTipi, cAction, "stokaksesuarrb", aFis(nCnt).cStokNo, aFis(nCnt).nNetMiktar1, aFis(nCnt).nNetMiktar2, aFis(nCnt).nNetMiktar3, _
                                 cFilter, aFis(nCnt).nAgirlik, dTarih, aFis(nCnt).nFiyat, aFis(nCnt).cDoviz, cDepartman, cFirma) = 0 Then
                        JustForLog("StokFis-" + cAction.Trim + "-End-Exit Function stokaksesuarrb " + cStokFisNo)
                        SingleStokFisValidateConnected = 0
                        Exit Function
                    End If
                    oStokRB.cMtk = cMTF
                    cStatus = "StokAksesuarRB table update OK"
                End If

                ' MTF 
                If UpdateMTF(ConnYage, cAction, cStokFisTipi, cDepartman, aFis(nCnt).cStokHareketKodu, aFis(nCnt).cIsemriNo, _
                          aFis(nCnt).cMTF, aFis(nCnt).cStokNo, aFis(nCnt).cRenk, aFis(nCnt).cBeden, aFis(nCnt).nNetMiktar1) = 0 Then
                    JustForLog("StokFis-" + cAction.Trim + "-End-Exit Function MTF " + cStokFisNo)
                    SingleStokFisValidateConnected = 0
                    Exit Function
                End If
                cStatus = "MTF update OK"

                ' işemri
                If UpdateIsemri(ConnYage, cAction, cStokFisTipi, aFis(nCnt).cStokHareketKodu, aFis(nCnt).cIsemriNo, _
                               aFis(nCnt).cMTF, aFis(nCnt).cStokNo, aFis(nCnt).cRenk, aFis(nCnt).cBeden, aFis(nCnt).nNetMiktar1) = 0 Then
                    JustForLog("StokFis-" + cAction.Trim + "-End-Exit Function isEmri " + cStokFisNo)
                    SingleStokFisValidateConnected = 0
                    Exit Function
                End If
                cStatus = "IsEmri update OK"

                ' stokrb2
                If Not oSysFlags.G_NoUpdateStokRB2 Then

                    cFilter = " where stokno = '" + aFis(nCnt).cStokNo + "' " + _
                            " and renk = '" + aFis(nCnt).cRenk + "' " + _
                            " and beden = '" + aFis(nCnt).cBeden + "' "

                    If CheckExistsConnected("select stokno from stokrb2 " + cFilter, ConnYage) Then

                        If LCase(cStokFisTipi) = "giris" Then
                            cSQL = "update stokrb2 " + _
                                    " set alisfiyati1 = " + SQLWriteDecimal(aFis(nCnt).nFiyat) + ", " + _
                                    " alisdovizi1 = '" + aFis(nCnt).cDoviz + "' " + _
                                    cFilter
                        Else
                            cSQL = "update stokrb2 " + _
                                    " set satisfiyati1 = " + SQLWriteDecimal(aFis(nCnt).nFiyat) + ", " + _
                                    " satisdovizi1 = '" + aFis(nCnt).cDoviz + "' " + _
                                    cFilter
                        End If
                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    Else
                        If LCase(cStokFisTipi) = "giris" Then
                            cSQL = "insert into stokrb2 " + _
                                    " (stokno,renk,beden,satisfiyati1,satisdovizi1,satisfiyati2,satisdovizi2,satisfiyati3,satisdovizi3,satisfiyati4, " + _
                                    " satisdovizi4,alisfiyati1,alisdovizi1,alisfiyati2,alisdovizi2,alisfiyati3,alisdovizi3,alisfiyati4,alisdovizi4) " + _
                                    " values (" + _
                                    " '" + aFis(nCnt).cStokNo + "', " + _
                                    " '" + aFis(nCnt).cRenk + "', " + _
                                    " '" + aFis(nCnt).cBeden + "', " + _
                                    " 0,'',0,'',0,'',0,'', " + _
                                     SQLWriteDecimal(aFis(nCnt).nFiyat) + ", " + _
                                    " '" + aFis(nCnt).cDoviz + "', " + _
                                    " 0,'',0,'',0,'')"
                        Else
                            cSQL = "insert into stokrb2 " + _
                                    " (stokno,renk,beden,satisfiyati1,satisdovizi1,satisfiyati2,satisdovizi2,satisfiyati3,satisdovizi3,satisfiyati4, " + _
                                    " satisdovizi4,alisfiyati1,alisdovizi1,alisfiyati2,alisdovizi2,alisfiyati3,alisdovizi3,alisfiyati4,alisdovizi4) " + _
                                    " values (" + _
                                    " '" + aFis(nCnt).cStokNo + "', " + _
                                    " '" + aFis(nCnt).cRenk + "', " + _
                                    " '" + aFis(nCnt).cBeden + "', " + _
                                     SQLWriteDecimal(aFis(nCnt).nFiyat) + ", " + _
                                    " '" + aFis(nCnt).cDoviz + "', " + _
                                    " 0,'',0,'',0,'', " + _
                                    " 0,'',0,'',0,'',0,'')"
                        End If
                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                End If
                cStatus = "StokRB2 table update OK"

                ' ortalama maliyet

                If Not oSysFlags.G_OrtalamaCalcIptal And _
                    LCase(cStokFisTipi) = "giris" And _
                    (aFis(nCnt).cStokHareketKodu = "02 Tedarikten Giris" _
                     Or aFis(nCnt).cStokHareketKodu = "05 Diger Giris" _
                     Or aFis(nCnt).cStokHareketKodu = "06 Tamirden Giris" _
                     Or aFis(nCnt).cStokHareketKodu = "04 Mlz Uretimden Giris" _
                     Or aFis(nCnt).cStokHareketKodu = "90 Trans/Rezv Giris") And _
                     aFis(nCnt).nNetMiktar1 > 0 And _
                     aFis(nCnt).nFiyat > 0 Then

                    cSQL = " update stokrb " + _
                           " set alismiktari1 =  coalesce(alismiktari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + " , " + _
                           " alistutari1 = coalesce(alistutari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nMlytTutar) + " " + _
                           " where stokno = '" + aFis(nCnt).cStokNo + "' " + _
                           " and renk = '" + aFis(nCnt).cRenk + "' " + _
                           " and beden = '" + aFis(nCnt).cBeden + "' " + _
                           " and depo = '" + aFis(nCnt).cDepo + "' " + _
                           " and malzemetakipkodu = '" + aFis(nCnt).cMTF + "' " + _
                           " and partino = '" + aFis(nCnt).cPartiNo + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)

                    oStokRB.cTableName = "stokrbstatic"
                    oStokRB.cStokno = aFis(nCnt).cStokNo
                    oStokRB.cRenk = aFis(nCnt).cRenk
                    oStokRB.cBeden = aFis(nCnt).cBeden
                    oStokRB.cPartiNo = aFis(nCnt).cPartiNo
                    oStokRB.cDepo = aFis(nCnt).cDepo
                    oStokRB.cMtk = aFis(nCnt).cMTF
                    oStokRB.cAy = Microsoft.VisualBasic.Format(Month(dTarih), "00")
                    oStokRB.cYil = Microsoft.VisualBasic.Format(Year(dTarih), "00")

                    cFilter = CheckInsertStokRB(oStokRB, oSysFlags, ConnYage)

                    cSQL = " update stokrbstatic " + _
                           " set alismiktari1 =  coalesce(alismiktari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + " , " + _
                           " alistutari1 = coalesce(alistutari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nMlytTutar) + " " + _
                           " where " + cFilter

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
                cStatus = "Ortalama Maliyet update OK"

                ' üretime çıkış maliyetleri

                If Not oSysFlags.G_NoUpdateStokRB2 And _
                    (aFis(nCnt).cStokHareketKodu = "01 Uretime Cikis" Or _
                     aFis(nCnt).cStokHareketKodu = "01 Uretimden iade") Then

                    cFilter = " where stokno = '" + aFis(nCnt).cStokNo + "' " + _
                              " and uretimtakipno = '" + aFis(nCnt).cUTF + "' " + _
                              " and modelno = '" + aFis(nCnt).cModelNo + "' " + _
                              " and parca = '" + aFis(nCnt).cParca + "' " + _
                              " and departman = '" + aFis(nCnt).cUretDept + "' " + _
                              " and uretimisemrino = '" + aFis(nCnt).cUretIsemriNo + "' " + _
                              " and malzemetakipkodu = '" + aFis(nCnt).cMTF + "' " + _
                              " and renk = '" + aFis(nCnt).cRenk + "' " + _
                              " and beden = '" + aFis(nCnt).cBeden + "' " + _
                              " and partino = '" + aFis(nCnt).cPartiNo + "' " + _
                              " and depo = ' " + aFis(nCnt).cDepo + "' "

                    If Not CheckExistsConnected("select stokno from uretstokrb " + cFilter, ConnYage) Then

                        cSQL = " insert into uretstokrb " + _
                                " (StokNo , uretimtakipno, ModelNo, parca, departman, uretimisemrino, malzemetakipkodu, renk, beden, partino, " + _
                                " depo , SonGirisTarihi, songirisfiyati, songirisdovizi, songirisdovizfiyati, SonGirisDept, SonGirisFirmasi, donemgiris1, donemcikis1, donemgiris2, " + _
                                " donemcikis2 , donemgiris3, donemcikis3) "

                        cSQL = cSQL + " values (" + _
                                " '" + aFis(nCnt).cStokNo + "', " + _
                                " '" + aFis(nCnt).cUTF + "', " + _
                                " '" + aFis(nCnt).cModelNo + "', " + _
                                " '" + aFis(nCnt).cParca + "', " + _
                                " '" + aFis(nCnt).cUretDept + "', " + _
                                " '" + aFis(nCnt).cUretIsemriNo + "', " + _
                                " '" + aFis(nCnt).cMTF + "', " + _
                                " '" + aFis(nCnt).cRenk + "', " + _
                                " '" + aFis(nCnt).cBeden + "', " + _
                                " '" + aFis(nCnt).cPartiNo + "', "

                        cSQL = cSQL + _
                                " '" + aFis(nCnt).cDepo + "', " + _
                                " '" + SQLWriteDate(dTarih) + "', " + _
                                SQLWriteDecimal(nDvzFiyat) + ", " + _
                                " '" + aFis(nCnt).cDoviz + "', " + _
                                SQLWriteDecimal(aFis(nCnt).nFiyat) + ", " + _
                                " '" + cDepartman + "', " + _
                                " '" + cFirma + "', " + _
                                " 0,0,0,0,0,0) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage, True)
                    End If

                    If LCase(cStokFisTipi) = "giris" Then
                        If LCase(cAction) = "validate" Then

                            cSQL = " update uretstokrb set " + _
                                    " SonGirisTarihi = '" + SQLWriteDate(dTarih) + "', " + _
                                    " songirisfiyati = " + SQLWriteDecimal(nDvzFiyat) + ", " + _
                                    " songirisdovizi = '" + aFis(nCnt).cDoviz + "', " + _
                                    " SonGirisDept = '" + cDepartman + "', " + _
                                    " SonGirisFirmasi = '" + cFirma + "' " + _
                                    cFilter + _
                                    " and SonGirisTarihi < '" + SQLWriteDate(dTarih) + "'"

                            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

                            cSQL = "update uretstokrb set " + _
                                    " donemgiris1 = coalesce(donemgiris1,0) + " + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + ", " + _
                                    " donemgiris2 = coalesce(donemgiris2,0) + " + SQLWriteDecimal(aFis(nCnt).nNetMiktar2) + ", " + _
                                    " donemgiris3 = coalesce(donemgiris3,0) + " + SQLWriteDecimal(aFis(nCnt).nNetMiktar3) + _
                                    cFilter

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        Else
                            cSQL = "update uretstokrb set " + _
                                    " donemgiris1 = coalesce(donemgiris1,0) - " + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + ", " + _
                                    " donemgiris2 = coalesce(donemgiris2,0) - " + SQLWriteDecimal(aFis(nCnt).nNetMiktar2) + ", " + _
                                    " donemgiris3 = coalesce(donemgiris3,0) - " + SQLWriteDecimal(aFis(nCnt).nNetMiktar3) + _
                                    cFilter

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End If
                    Else
                        ' üretim maliyetleri malzeme çıkışında olur

                        If LCase(cAction) = "validate" Then

                            cSQL = "update uretstokrb set " + _
                                    " donemcikis1 = coalesce(donemcikis1,0) + " + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + ", " + _
                                    " donemcikis2 = coalesce(donemcikis2,0) + " + SQLWriteDecimal(aFis(nCnt).nNetMiktar2) + ", " + _
                                    " donemcikis3 = coalesce(donemcikis3,0) + " + SQLWriteDecimal(aFis(nCnt).nNetMiktar3) + ", " + _
                                    " girismiktari1 = coalesce(girismiktari1,0) + " + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + ", " + _
                                    " giristutari1 = coalesce(giristutari1,0) + " + SQLWriteDecimal(nMlytTutar) + _
                                    cFilter

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        Else
                            cSQL = "update uretstokrb set " + _
                                    " donemcikis1 = coalesce(donemcikis1,0) - " + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + ", " + _
                                    " donemcikis2 = coalesce(donemcikis2,0) - " + SQLWriteDecimal(aFis(nCnt).nNetMiktar2) + ", " + _
                                    " donemcikis3 = coalesce(donemcikis3,0) - " + SQLWriteDecimal(aFis(nCnt).nNetMiktar3) + ", " + _
                                    " girismiktari1 = coalesce(girismiktari1,0) - " + SQLWriteDecimal(aFis(nCnt).nNetMiktar1) + ", " + _
                                    " giristutari1 = coalesce(giristutari1,0) - " + SQLWriteDecimal(nMlytTutar) + _
                                    cFilter

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                        End If
                    End If
                End If
                cStatus = "UretStokRB table update OK"

            Next
    
            cStatus = "Main Update Completed"

            ' otomatik işemri kapatma
            If oSysFlags.G_OtoMlzIsemriKapat And lIsemri And cAction = "validate" And LCase(cStokFisTipi) = "giris" Then
                OtoisEmriKapat(ConnYage, oSysFlags, cStokFisNo, aIsEmri)
            End If
            cStatus = "OtoisEmriKapat Completed"

            SingleStokFisValidateConnected = 1
            JustForLog("StokFis-" + cAction.Trim + "-End Successful " + cStokFisNo)
        Catch
            SingleStokFisValidateConnected = 0
            ErrDisp("Error SingleStokFisValidateConnected " + Err.Description.Trim + vbCrLf + _
                    "SQL : " + cSQL + vbCrLf + _
                    "Action : " + cAction.Trim + vbCrLf + _
                    "StokFisNo : " + cStokFisNo + vbCrLf + _
                    "Status : " + cStatus)
        End Try
    End Function

    Private Sub OtoisEmriKapat(ByVal ConnYage As SqlConnection, ByVal oSysFlags As SysFlags, ByVal cStokFisNo As String, ByVal aIsEmri() As String)

        Dim lKapat As Boolean
        Dim cSQL As String = ""
        Dim nTolerans As Double
        Dim nGeneralTolerans As Double
        Dim nIstenen As Double
        Dim nGelen As Double
        Dim nCnt As Integer
        Dim oReader As SqlDataReader

        ' stok fisi validateten sonra calisiyor
        ' ilgili is emirlerini kontrol ediyor ve kapanmasi gerekenleri kapatiyor
        Try

            nGeneralTolerans = 0
            If oSysFlags.G_YuzdeGirisKontrol Then
                nGeneralTolerans = oSysFlags.G_YuzdeGirisKontrolYuzde
            End If

            For nCnt = 0 To UBound(aIsEmri)

                lKapat = True

                cSQL = "select a.maxkabulorani, b.miktar1, b.TedarikGelen, b.UretimGelen " + _
                        " from isemri a, isemrilines b " + _
                        " where a.isemrino = b.isemrino " + _
                        " and a.isemrino = '" + aIsEmri(nCnt) + "' " + _
                        " and (a.isemriok is null or a.isemriok = 'H') "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.HasRows Then

                    Do While oReader.Read

                        nTolerans = nGeneralTolerans
                        If SQLReadDouble(oReader, "maxkabulorani") <> 0 Then
                            nTolerans = SQLReadDouble(oReader, "maxkabulorani")
                        End If

                        nIstenen = SQLReadDouble(oReader, "miktar1")
                        nGelen = SQLReadDouble(oReader, "TedarikGelen") + SQLReadDouble(oReader, "UretimGelen")
                        lKapat = (nIstenen - (nIstenen * nTolerans / 100) <= nGelen)
                        If Not lKapat Then Exit Do
                    Loop
                    oReader.Close()
                    oReader = Nothing

                    If lKapat Then
                        cSQL = "update isemri " + _
                                 " set isEmriOk = 'E', " + _
                                 " oktarihi = getdate() " + _
                                 " where isEmriNo = '" + aIsEmri(nCnt) + " '"

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                Else
                    oReader.Close()
                    oReader = Nothing
                End If
            Next

        Catch ex As Exception
            ErrDisp("OtoisEmriKapat " + ex.Message + vbCrLf + "SQL : " + cSQL)
        End Try
    End Sub

End Module

Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module TransferValidate

    Public Function TopluTransferValidate(ByVal cAction As String, ByVal cTopluFisNo As String) As SqlInt32

        Dim cSQL As String
        Dim ConnYage As SqlConnection
        Dim aTransferFisNo() As String
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim lFound As Boolean = False

        TopluTransferValidate = 0

        Try
            nCnt = 0
            ReDim aTransferFisNo(0)

            ConnYage = OpenConn()

            cSQL = "select distinct transferfisno " + _
                    " from stoktransfer " + _
                    " where toplufisno = '" + Trim(cTopluFisNo) + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aTransferFisNo(nCnt)
                aTransferFisNo(nCnt) = SQLReadString(oReader, "transferfisno")
                lFound = True
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            CloseConn(ConnYage)

            TopluTransferValidate = 1

            If lFound Then
                For nCnt = 0 To UBound(aTransferFisNo)
                    If aTransferFisNo(nCnt).Trim <> "" Then
                        If TransferFisValidate(cAction, aTransferFisNo(nCnt)) = 0 Then
                            TopluTransferValidate = 0
                            Exit For
                        End If
                    End If
                Next
            End If

        Catch
            TopluTransferValidate = 0
            ErrDisp("Error TopluTransferValidate " + Err.Description.Trim)
        End Try
    End Function

    Public Function MultiTransferFisValidate(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String) As SqlInt32

        Dim cSQL As String
        Dim aTransferFisNo() As String
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim ConnYage As SqlConnection
        Dim lFound As Boolean = False

        MultiTransferFisValidate = 0

        Try
            ConnYage = OpenConn()

            nCnt = 0
            ReDim aTransferFisNo(0)

            cSQL = "Select transferfisno " + _
                    " from stoktransfer " + _
                    IIf(cStokNo = "", "", " where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aTransferFisNo(nCnt)
                aTransferFisNo(nCnt) = SQLReadString(oReader, "transferfisno")
                nCnt = nCnt + 1
                lFound = True
            Loop
            oReader.Close()
            oReader = Nothing

            MultiTransferFisValidate = 1
            If lFound Then
                For nCnt = 0 To UBound(aTransferFisNo)
                    If TransferValidateConnected(ConnYage, "validate", aTransferFisNo(nCnt)) = 0 Then
                        MultiTransferFisValidate = 0
                        Exit For
                    End If
                Next
            End If
            CloseConn(ConnYage)
        Catch
            MultiTransferFisValidate = 0
            ErrDisp("Error MultiTransferFisValidate " + Err.Description.Trim)
        End Try
    End Function

    Public Function TransferFisValidate(ByVal cAction As String, ByVal cTransferFisNo As String) As SqlInt32

        Dim ConnYage As SqlConnection

        TransferFisValidate = 0

        Try
            ConnYage = OpenConn()
            TransferFisValidate = TransferValidateConnected(ConnYage, cAction, cTransferFisNo)
            CloseConn(ConnYage)
        Catch
            TransferFisValidate = 0
            ErrDisp("Error TransferFisValidate " + Err.Description.Trim)
        End Try
    End Function

    Private Function TransferValidateConnected(ByVal ConnYage As SqlConnection, ByVal cAction As String, ByVal cTransferFisNo As String) As SqlInt32

        Dim oStokRBKaynak As New stokrb
        Dim oStokRBHedef As New stokrb
        Dim oSysFlags As New General.SysFlags

        Dim oReader As SqlDataReader
        Dim cSQL As String
        Dim dTarih As Date
        Dim nMiktar1 As Double
        Dim nMiktar2 As Double
        Dim nMiktar3 As Double
        Dim nFiyat As Double = 0
        Dim cDoviz As String = "TL"
        Dim cFilterKaynak As String
        Dim cFilterHedef As String
        Dim nMlytTutar As Double
        Dim nKur As Double
        Dim cMRP As String = ""
        Dim cMTF As String = ""

        TransferValidateConnected = 0

        Try
            JustForLog("TransferValidate-Begin " + cTransferFisNo)

            If cTransferFisNo.Trim = "" Then
                JustForLog("TransferValidate-End-EmptyTransferFisNo ")
                TransferValidateConnected = 1
                Exit Function
            End If

            ReadSysFlags(oSysFlags, ConnYage)

            cSQL = "Select * From stoktransfer where transferfisno = '" + Trim(cTransferFisNo) + "' "

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read() Then

                dTarih = SQLReadDate(oReader, "tarih")
                nMiktar1 = SQLReadDouble(oReader, "NetMiktar1")
                nMiktar2 = SQLReadDouble(oReader, "NetMiktar2")
                nMiktar3 = SQLReadDouble(oReader, "NetMiktar3")
                nFiyat = SQLReadDouble(oReader, "BirimFiyat")
                cDoviz = SQLReadString(oReader, "DovizCinsi")
                cMRP = SQLReadString(oReader, "mprojeno")

                oStokRBKaynak.cTableName = "stokrb"
                oStokRBKaynak.cStokno = SQLReadString(oReader, "stokno")
                oStokRBKaynak.cRenk = SQLReadString(oReader, "renk")
                oStokRBKaynak.cBeden = SQLReadString(oReader, "beden")
                oStokRBKaynak.cPartiNo = SQLReadString(oReader, "KaynakPartino")
                oStokRBKaynak.cDepo = SQLReadString(oReader, "KaynakDepo")
                oStokRBKaynak.cMtk = SQLReadString(oReader, "kaynakmalzemetakipno")
                oStokRBKaynak.cTopNo = SQLReadString(oReader, "topno")

                oStokRBHedef.cTableName = "stokrb"
                oStokRBHedef.cStokno = SQLReadString(oReader, "stokno")
                oStokRBHedef.cRenk = SQLReadString(oReader, "renk")
                oStokRBHedef.cBeden = SQLReadString(oReader, "beden")
                oStokRBHedef.cPartiNo = SQLReadString(oReader, "HedefPartiNo")
                oStokRBHedef.cDepo = SQLReadString(oReader, "HedefDepo")
                oStokRBHedef.cMtk = SQLReadString(oReader, "HedefMalzemeTakipNo")
                oStokRBHedef.cTopNo = SQLReadString(oReader, "topno")
            End If
            oReader.Close()
            oReader = Nothing

            ' Kaynaktan Çık
            CheckRenkBedenValidate(oStokRBKaynak.cStokno, oStokRBKaynak.cRenk, oStokRBKaynak.cBeden, ConnYage)

            If UpdateStokRB(ConnYage, "cikis", cAction, "stok", oStokRBKaynak.cStokno, nMiktar1, nMiktar2, nMiktar3) = 0 Then
                TransferValidateConnected = 0
                Exit Function
            End If

            cFilterKaynak = CheckInsertStokRB(oStokRBKaynak, oSysFlags, ConnYage)
            If UpdateStokRB(ConnYage, "cikis", cAction, oStokRBKaynak.cTableName, oStokRBKaynak.cStokno, nMiktar1, nMiktar2, nMiktar3, cFilterKaynak) = 0 Then
                TransferValidateConnected = 0
                Exit Function
            End If

            ' Hedefe Gir
            CheckRenkBedenValidate(oStokRBHedef.cStokno, oStokRBHedef.cRenk, oStokRBHedef.cBeden, ConnYage)

            If UpdateStokRB(ConnYage, "giris", cAction, "stok", oStokRBHedef.cStokno, nMiktar1, nMiktar2, nMiktar3) = 0 Then
                TransferValidateConnected = 0
                Exit Function
            End If

            cFilterHedef = CheckInsertStokRB(oStokRBHedef, oSysFlags, ConnYage)
            If UpdateStokRB(ConnYage, "giris", cAction, oStokRBHedef.cTableName, oStokRBHedef.cStokno, nMiktar1, nMiktar2, nMiktar3, cFilterHedef) = 0 Then
                TransferValidateConnected = 0
                Exit Function
            End If

            ' barkodlu kumaş
            If oSysFlags.G_WinFabric Then
                If G_BarkodluKumas(oStokRBKaynak.cStokno, oSysFlags, ConnYage) Then
                    ' bk çık
                    cMTF = oStokRBKaynak.cMtk
                    oStokRBKaynak.cTableName = "stoktoprb"
                    If oSysFlags.G_WFNoMTK Then
                        oStokRBKaynak.cMtk = ""
                    End If
                    cFilterKaynak = CheckInsertStokRB(oStokRBKaynak, oSysFlags, ConnYage)
                    If UpdateStokRB(ConnYage, "cikis", cAction, oStokRBKaynak.cTableName, oStokRBKaynak.cStokno, nMiktar1, nMiktar2, nMiktar3, cFilterKaynak) = 0 Then
                        TransferValidateConnected = 0
                        Exit Function
                    End If
                    oStokRBKaynak.cMtk = cMTF
                    ' bk gir
                    cMTF = oStokRBHedef.cMtk
                    oStokRBHedef.cTableName = "stoktoprb"
                    If oSysFlags.G_WFNoMTK Then
                        oStokRBHedef.cMtk = ""
                    End If
                    cFilterHedef = CheckInsertStokRB(oStokRBHedef, oSysFlags, ConnYage)
                    If UpdateStokRB(ConnYage, "giris", cAction, oStokRBHedef.cTableName, oStokRBHedef.cStokno, nMiktar1, nMiktar2, nMiktar3, cFilterHedef) = 0 Then
                        TransferValidateConnected = 0
                        Exit Function
                    End If
                    oStokRBHedef.cMtk = cMTF
                End If
            End If


            ' barkodlu aksesuar
            If oSysFlags.G_WinAccessory Then
                If G_BarkodluAksesuar(oStokRBKaynak.cStokno, oSysFlags, ConnYage) Then
                    ' bk çık
                    cMTF = oStokRBKaynak.cMtk
                    oStokRBKaynak.cTableName = "stokaksesuarrb"
                    If oSysFlags.G_WANoMTK Then
                        oStokRBKaynak.cMtk = ""
                    End If
                    cFilterKaynak = CheckInsertStokRB(oStokRBKaynak, oSysFlags, ConnYage)
                    If UpdateStokRB(ConnYage, "cikis", cAction, oStokRBKaynak.cTableName, oStokRBKaynak.cStokno, nMiktar1, nMiktar2, nMiktar3, cFilterKaynak) = 0 Then
                        TransferValidateConnected = 0
                        Exit Function
                    End If
                    oStokRBKaynak.cMtk = cMTF
                    ' bk gir
                    cMTF = oStokRBHedef.cMtk
                    oStokRBHedef.cTableName = "stokaksesuarrb"
                    If oSysFlags.G_WANoMTK Then
                        oStokRBHedef.cMtk = ""
                    End If
                    cFilterHedef = CheckInsertStokRB(oStokRBHedef, oSysFlags, ConnYage)
                    If UpdateStokRB(ConnYage, "giris", cAction, oStokRBHedef.cTableName, oStokRBHedef.cStokno, nMiktar1, nMiktar2, nMiktar3, cFilterHedef) = 0 Then
                        TransferValidateConnected = 0
                        Exit Function
                    End If
                    oStokRBHedef.cMtk = cMTF
                End If
            End If

            ' mtk update
            If oStokRBKaynak.cMtk <> oStokRBHedef.cMtk Then
                If UpdateMTF(ConnYage, cAction, "cikis", "", "transfer", "", oStokRBKaynak.cMtk, oStokRBKaynak.cStokno, oStokRBKaynak.cRenk, oStokRBKaynak.cBeden, nMiktar1) = 0 Then
                    TransferValidateConnected = 0
                    Exit Function
                End If

                If UpdateMTF(ConnYage, cAction, "giris", "", "transfer", "", oStokRBHedef.cMtk, oStokRBHedef.cStokno, oStokRBHedef.cRenk, oStokRBHedef.cBeden, nMiktar1) = 0 Then
                    TransferValidateConnected = 0
                    Exit Function
                End If
            End If

            ' MalzemeRezervasyonProjesi (MRP) update
            If cMRP <> "" Then
                cSQL = "update mprojelines " + _
                            " set mtkmiktar = coalesce (mtkmiktar,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nMiktar1) + _
                            " where mprojeno = '" + cMRP + "' " + _
                            " and stokno = '" + oStokRBHedef.cStokno + "' " + _
                            " and renk = '" + oStokRBHedef.cRenk + "' " + _
                            " and beden = '" + oStokRBHedef.cBeden + "' "
                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

            ' ortalama maliyet
            If Not oSysFlags.G_OrtalamaCalcIptal And _
                 nMiktar1 > 0 And _
                 nFiyat > 0 Then

                nKur = GetKur(cDoviz, dTarih, ConnYage)
                nMlytTutar = nMiktar1 * nFiyat * nKur

                nKur = GetKur(oSysFlags.G_OrtMlytDoviz, dTarih, ConnYage)
                If nKur > 0 Then
                    nMlytTutar = nMlytTutar / nKur
                End If

                cSQL = "update stokrb " + _
                        " set alismiktari1 =  coalesce(alismiktari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nMiktar1) + " , " + _
                        " alistutari1 = coalesce(alistutari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nMlytTutar) + " " + _
                        " where stokno = '" + oStokRBHedef.cStokno + "' " + _
                        " and renk = '" + oStokRBHedef.cRenk + "' " + _
                        " and beden = '" + oStokRBHedef.cBeden + "' " + _
                        " and depo = '" + oStokRBHedef.cDepo + "' " + _
                        " and malzemetakipkodu = '" + oStokRBHedef.cMtk + "' " + _
                        " and partino = '" + oStokRBHedef.cPartiNo + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                oStokRBHedef.cTableName = "stokrbstatic"
                oStokRBHedef.cAy = Microsoft.VisualBasic.Format(Month(dTarih), "00")
                oStokRBHedef.cYil = Microsoft.VisualBasic.Format(Year(dTarih), "00")

                cFilterHedef = CheckInsertStokRB(oStokRBHedef, oSysFlags, ConnYage)

                cSQL = "update stokrbstatic " + _
                        " set alismiktari1 =  coalesce(alismiktari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nMiktar1) + " , " + _
                        " alistutari1 = coalesce(alistutari1,0) " + IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nMlytTutar) + " " + _
                        " where " + cFilterHedef

                ExecuteSQLCommandConnected(cSQL, ConnYage)

            End If
            TransferValidateConnected = 1
            JustForLog("TransferValidate-End " + cTransferFisNo)

        Catch
            TransferValidateConnected = 0
            ErrDisp("Error TransferValidateConnected " + Err.Description.Trim)
        End Try
    End Function

End Module

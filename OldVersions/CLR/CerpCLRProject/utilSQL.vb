Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

' clr entegrasyonunda bağlanan procesin connectionı kullanılıyor
' sadece 1 bağlantı olabiliyor
' o bağlantıda 1 datareader olabiliyor
'cConnStr = "context connection=true"

'edit my project 
' permission(level) -> external()

'run sql queries
' sp_configure 'show advanced options', 1
' reconfigure
' sp_configure 'clr enabled','1'
' reconfigure
' ALTER AUTHORIZATION ON DATABASE::wintex TO sa
' ALTER DATABASE wintex SET TRUSTWORTHY ON
' USE wintex
' EXEC sp_changedbowner 'sa'

'DROP PROCEDURE BarkodEskiStokDurumu
'DROP PROCEDURE buildmasterplan
'DROP PROCEDURE FastMTFRecalc 
'DROP PROCEDURE fastmtfbuild
'DROP PROCEDURE fastmtfbuildall
'DROP PROCEDURE fastutfbuild
'DROP PROCEDURE fastutfbuildall
'DROP PROCEDURE faststfbuild
'DROP PROCEDURE faststfbuildall
'DROP PROCEDURE gettoplamsiparisview
'DROP PROCEDURE hizlistokbakimi
'DROP PROCEDURE kumascikisbarkodlu 
'DROP PROCEDURE mtfhesaplax 
'DROP PROCEDURE planlamaforwardall 
'DROP PROCEDURE planlamaotomatikkapatall 
'DROP PROCEDURE siparistakip1 
'DROP PROCEDURE tarihstokdurumu 
'DROP PROCEDURE tekstoktoplam 
'DROP PROCEDURE validatemultistokfis 
'DROP PROCEDURE validatemultitransferfis 
'DROP PROCEDURE validatestokfis 
'DROP PROCEDURE validatetoplutransferfis
'DROP PROCEDURE validatetransferfis 
'DROP PROCEDURE DMTFIsemriKilavuzu 
'DROP FUNCTION RaporSiparisDurumu1
'DROP ASSEMBLY WinTex

'CREATE ASSEMBLY WinTex from 'D:\okan\CLR\CerpCLRProject\obj\release\WinTex.dll' WITH PERMISSION_SET = EXTERNAL_ACCESS

'CREATE PROCEDURE dbo.PlanlamaOtomatikKapatAll AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].PlanlamaOtomatikKapatAll
'CREATE PROCEDURE dbo.PlanlamaForwardAll  @cSiparisNo NVARCHAR (4000) AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].PlanlamaForwardAll
'CREATE PROCEDURE dbo.HizliStokBakimi AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].HizliStokBakimi
'CREATE PROCEDURE dbo.FastSTFBuildAll AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].FastSTFBuildAll
'CREATE PROCEDURE dbo.FastMTFBuildAll AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].FastMTFBuildAll
'CREATE PROCEDURE dbo.FastUTFBuildAll AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].FastUTFBuildAll
'CREATE PROCEDURE dbo.FastSTFBuild @cSTF NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].FastSTFBuild
'CREATE PROCEDURE dbo.FastMTFBuild @cMTF NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].FastMTFBuild
'CREATE PROCEDURE dbo.FastUTFBuild @cUTF NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].FastUTFBuild
'CREATE PROCEDURE dbo.SiparisTakip1 AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].SiparisTakip1
'CREATE PROCEDURE dbo.GetToplamSiparisView @cFilter NVARCHAR (4000), @cTableName NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].GetToplamSiparisView
'CREATE PROCEDURE dbo.MTFHesaplax @cFilter1 NVARCHAR (4000), @cFilter2 NVARCHAR (4000), @cTSip NVARCHAR (4000),  @cTableName NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].MTFHesaplax
'CREATE PROCEDURE dbo.BuildMasterPlan @cFilter NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].BuildMasterPlan
'CREATE PROCEDURE dbo.DMTFIsemriKilavuzu @cDetayIhtiyacTable NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME WinTex.[WinTex.StoredProcedures].DMTFIsemriKilavuzu
'CREATE PROCEDURE [dbo].[KumasCikisBarkodlu] @cStokFisNo NVARCHAR (4000), @cStokHareketKodu NVARCHAR (4000), @cDepartman NVARCHAR (4000), @cFirma NVARCHAR (4000), @cNotlar NVARCHAR (4000) AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[KumasCikisBarkodlu]
'CREATE PROCEDURE [dbo].[TarihStokDurumu] @BslTarihi NVARCHAR (4000), @BtsTarihi NVARCHAR (4000), @StokNo NVARCHAR (4000), @nCase INT, @cTableName NVARCHAR (4000) OUTPUT AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[TarihStokDurumu]
'CREATE PROCEDURE [dbo].[TekStokToplam] @cStokNo NVARCHAR (4000), @cRenk NVARCHAR (4000), @cBeden NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[TekStokToplam]
'CREATE PROCEDURE [dbo].[ValidateMultiStokFis] @cStokNo NVARCHAR (4000), @cRenk NVARCHAR (4000), @cBeden NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[ValidateMultiStokFis]
'CREATE PROCEDURE [dbo].[ValidateMultiTransferFis] @cStokNo NVARCHAR (4000), @cRenk NVARCHAR (4000), @cBeden NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[ValidateMultiTransferFis]
'CREATE PROCEDURE [dbo].[ValidateStokFis] @cAction NVARCHAR (4000), @cStokFisNo NVARCHAR (4000), @cStokNo NVARCHAR (4000), @cRenk NVARCHAR (4000), @cBeden NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[ValidateStokFis]
'CREATE PROCEDURE [dbo].[ValidateTopluTransferFis] @cAction NVARCHAR (4000), @cTTFisNo NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[ValidateTopluTransferFis]
'CREATE PROCEDURE [dbo].[ValidateTransferFis] @cAction NVARCHAR (4000), @cTransferFisNo NVARCHAR (4000), @nSonuc INT OUTPUT AS EXTERNAL NAME [WinTex].[WinTex.StoredProcedures].[ValidateTransferFis]

'CREATE FUNCTION [dbo].[RaporSiparisDurumu1]
'(@nOpenClose INT, @nMonth INT)
'RETURNS 
'     TABLE (
'        [Siparis_Planlanan_Svk_Hft]    DECIMAL (10)   NULL,
'        [Siparis_Son_Svk_Hft]          DECIMAL (10)   NULL,
'        [Siparis_Gelis_Tarihi]         DATETIME       NULL,
'        [Siparis_Planlanan_Svk_Tarihi] DATETIME       NULL,
'        [Son_Sevk_Tarihi]              DATETIME       NULL,
'        [Sevkiyat_Siparis_Gun_Frk]     DECIMAL (10)   NULL,
'        [Musteri]                      NVARCHAR (30)  NULL,
'        [Musteri_Temsilcisi]           NVARCHAR (30)  NULL,
'        [Siparis_No]                   NVARCHAR (30)  NULL,
'        [Model_No]                     NVARCHAR (30)  NULL,
'        [Renk]                         NVARCHAR (30)  NULL,
'        [Siparis_Durumu]               NVARCHAR (30)  NULL,
'        [Kumas_Tedarikcisi]            NVARCHAR (30)  NULL,
'        [Kumas_Termini]                DATETIME       NULL,
'        [Kumas_Gelis_Tarihi]           DATETIME       NULL,
'        [Kumas_Tamamlanma_Yuzdesi]     DECIMAL (10)   NULL,
'        [Aksesuar_Durumu]              NVARCHAR (30)  NULL,
'        [Kesim_OK_Tarihi]              DATETIME       NULL,
'        [Kesim_Tarihi]                 DATETIME       NULL,
'        [Siparis_Adedi]                DECIMAL (10)   NULL,
'        [Kesim_Adedi]                  DECIMAL (10)   NULL,
'        [Dikim_Giren]                  DECIMAL (10)   NULL,
'        [Dikim_Adedi]                  DECIMAL (10)   NULL,
'        [Sevk_Adedi]                   DECIMAL (10)   NULL,
'        [Kesim_Sevkiyat_Farki]         DECIMAL (10)   NULL,
'        [Dikim_Sevkiyat_Farki]         DECIMAL (10)   NULL,
'        [Sevkiyat_Siparis_Farki]       DECIMAL (10)   NULL,
'        [Kesim_Yuzdesi]                DECIMAL (10)   NULL,
'        [Dikim_Yuzdesi]                DECIMAL (10)   NULL,
'        [Sevkiyat_Yuzdesi]             DECIMAL (10)   NULL,
'        [Sevkiyat_Yuzdesi2]            DECIMAL (10)   NULL,
'        [Kesim_Atolyesi]               NVARCHAR (30)  NULL,
'        [Dikim_Atolyesi]               NVARCHAR (30)  NULL,
'        [Gerceklesen_Sevkiyat]         DATETIME       NULL,
'        [Ikinci_Kalite]                DECIMAL (10)   NULL,
'        [Atolye_Iade]                  DECIMAL (10)   NULL,
'        [Uretim_Kaybi]                 DECIMAL (10)   NULL,
'        [SiparisTakipNotlari]          NVARCHAR (250) NULL)
'AS EXTERNAL NAME [WinTex].[WinTex.UserDefinedFunctions].[RaporSiparisDurumu1]


Module utilSQL

    Public Function SQLGetRowCount(ByVal cTableName As String, ByVal ConnYage As SqlConnection) As Double

        Dim cSQL As String = ""

        cSQL = "select count(*) from " + cTableName.Trim

        SQLGetRowCount = SQLGetDoubleConnected(cSQL, ConnYage)
    End Function

    Public Function SQLGetQueryCount(ByVal cSQL As String, ByVal ConnYage As SqlConnection) As Double

        Dim cSQL2 As String = ""

        cSQL2 = "select count(*) from (" + cSQL.Trim + ") as x"

        SQLGetQueryCount = SQLGetDoubleConnected(cSQL2, ConnYage)
    End Function

    Public Function GetSQLReader(ByVal cSQL As String, ByVal ConnYage As SqlConnection) As SqlDataReader

        Dim nCnt As Integer = 0
        Dim nTimeOut As Integer = 0
        Dim lOK As Boolean = False

        GetSQLReader = Nothing

        Try
            For nCnt = 1 To 100
                nTimeOut = 50 + (10 * nCnt)
                GetSQLReader = SQLReaderHandled(cSQL, ConnYage, nTimeOut)
                If GetSQLReader Is Nothing Then
                    lOK = False
                Else
                    lOK = True
                    Exit For
                End If
            Next

            If Not lOK Then
                ErrDisp("TimeOut : " + nTimeOut.ToString, "GetSQLReader TimeOUT Problem", cSQL)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "GetSQLReader", cSQL)
        End Try
    End Function

    Private Function SQLReaderHandled(ByVal cSQL As String, ByVal ConnYage As SqlConnection, ByVal nTimeOut As Integer) As SqlDataReader

        Dim oCommand As SqlCommand

        SQLReaderHandled = Nothing

        Try
            oCommand = New SqlCommand
            oCommand.CommandText = cSQL
            oCommand.Connection = ConnYage
            oCommand.CommandTimeout = nTimeOut
            SQLReaderHandled = oCommand.ExecuteReader

        Catch ex As Exception
            'ErrDisp("SQLReaderHandled : " + ex.Message + vbCrLf + _
            '        "SQL : " + cSQL + vbCrLf + _
            '        "TimeOut : " + nTimeOut.ToString)
            ' do nothing
        End Try
    End Function

    Public Function OpenConn() As SqlConnection

        OpenConn = Nothing

        Try
            OpenConn = New SqlConnection("context connection=true")
            OpenConn.Open()

        Catch Err As Exception
            ErrDisp("OpenConn : " + Err.Message)
        End Try
    End Function

    Public Sub CloseConn(ByVal oMyConnection As SqlConnection)

        Try
            oMyConnection.Close()
            oMyConnection = Nothing

        Catch Err As Exception
            ErrDisp("CloseConn : " + Err.Message)
        End Try
    End Sub

    Public Function CheckNullString(ByVal oValue As Object) As String
        CheckNullString = ""
        Try
            If IsDBNull(oValue) Then
                CheckNullString = ""
            ElseIf IsNothing(oValue) Then
                CheckNullString = ""
            Else
                CheckNullString = oValue.ToString.Trim()
            End If
        Catch ex As Exception
            ErrDisp("CheckNullString : " + ex.Message)
        End Try
    End Function

    Public Function fPercent(ByVal Val1 As Double, ByVal Val2 As Double, Optional ByVal nCase As Integer = 0) As Double
        fPercent = 0

        If Val1 <> 0 Then
            If nCase = 0 Then
                fPercent = Math.Abs((Val1 - Val2) / Val1) * 100
            ElseIf nCase = 1 Then
                fPercent = Math.Abs((Val1 - Val2) / Val1) * 100
                fPercent = 100 - fPercent
            End If
        End If
    End Function

    Public Function fSeek(ByVal cSQL As String) As Object

        Dim oReader As SqlDataReader
        Dim nRow As Integer = 0
        Dim nCol As Integer = 0
        Dim ConnYage As SqlConnection
        Dim oRecord(0, 0) As String
        Dim nFieldCount As Integer = 0

        FSeek = oRecord

        Try
            ConnYage = OpenConn()

            oReader = GetSQLReader(cSQL, ConnYage)

            If oReader.Read Then
                nFieldCount = oReader.FieldCount
                ReDim oRecord(nFieldCount, 0)

                nRow = 0
                Do While oReader.Read
                    ReDim Preserve oRecord(nFieldCount, nRow)
                    For nCol = 0 To nFieldCount - 1
                        If IsDBNull(oReader.GetValue(nCol)) Then
                            oRecord(nCol, nRow) = ""
                        Else
                            oRecord(nCol, nRow) = oReader.GetValue(nCol).ToString.Trim
                        End If
                    Next
                    nRow = nRow + 1
                Loop
                oReader.Close()
                oReader = Nothing
            End If
            CloseConn(ConnYage)

            FSeek = oRecord
        Catch Err As Exception
            ErrDisp("Error FSeek" + Err.Message)
        End Try
    End Function

    Public Function ExecuteSQLCommand(ByVal cSQL As String, Optional ByVal DateFormat As Boolean = False) As Boolean

        Dim ConnYage As SqlConnection

        ExecuteSQLCommand = False

        Try
            If cSQL.Trim = "" Then Exit Function

            ConnYage = OpenConn()
            ExecuteSQLCommand = ExecuteSQLCommandConnected(cSQL, ConnYage, DateFormat)
            CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp("ExecuteSQLCommand : " + Err.Message + vbCrLf + cSQL)
        End Try
    End Function

    Public Function ExecuteSQLCommandConnected(ByVal cSQL As String, ByVal ConnYage As SqlConnection, Optional ByVal DateFormat As Boolean = False) As Boolean

        Dim nCnt As Integer = 0
        Dim nTimeOut As Integer = 0

        ExecuteSQLCommandConnected = False

        Try
            If cSQL.Trim = "" Then Exit Function

            For nCnt = 1 To 10
                nTimeOut = 50 + (10 * nCnt)
                ExecuteSQLCommandConnected = ExecuteSQLCommandHandled(cSQL, ConnYage, DateFormat, nTimeOut)
                If ExecuteSQLCommandConnected Then
                    Exit For
                End If
            Next
            If Not ExecuteSQLCommandConnected Then
                ErrDisp("ExecuteSQLCommandHandled Problem : TimeOUT " + vbCrLf + _
                        "SQL : " + cSQL + vbCrLf + _
                        "TimeOut : " + nTimeOut.ToString)
            End If
        Catch ex As Exception
            ExecuteSQLCommandConnected = False
            ErrDisp("ExecuteSQLCommandConnected Error : " + ex.Message + vbCrLf + _
                    "SQL : " + cSQL)
        End Try
    End Function

    Private Function ExecuteSQLCommandHandled(ByVal cSQL As String, ByVal ConnYage As SqlConnection, Optional ByVal DateFormat As Boolean = False, Optional ByVal nTimeOut As Integer = 60) As Boolean

        Dim oCommand As SqlCommand
        Dim returnValue As Integer

        ExecuteSQLCommandHandled = False

        If cSQL.Trim = "" Then Exit Function

        Try
            If DateFormat Then cSQL = "Set dateformat 'dmy'  " + cSQL

            oCommand = New SqlCommand
            oCommand.CommandText = cSQL
            oCommand.Connection = ConnYage
            oCommand.CommandTimeout = nTimeOut
            returnValue = oCommand.ExecuteNonQuery()
            oCommand = Nothing
            ExecuteSQLCommandHandled = True

        Catch Err As Exception
            'ErrDisp("ExecuteSQLCommandHandled : " + Err.Message + vbCrLf + _
            '        "SQL : " + cSQL + vbCrLf + _
            '        "TimeOut : " + nTimeOut.ToString)
            ' do nothing
        End Try
    End Function

    Public Function CheckExists(ByVal cSQL As String) As Boolean

        Dim ConnYage As SqlConnection

        CheckExists = False

        Try
            ConnYage = OpenConn()
            CheckExists = CheckExistsConnected(cSQL, ConnYage)
            CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp("CheckExists : " + Err.Message + vbCrLf + cSQL)
        End Try
    End Function

    Public Function CheckExistsConnected(ByVal cSQL As String, ByVal ConnYage As SqlConnection) As Boolean

        Dim oReader As SqlDataReader

        CheckExistsConnected = False

        If cSQL.Trim = "" Then Exit Function

        Try
            oReader = GetSQLReader(cSQL, ConnYage)
            If oReader.Read() Then
                CheckExistsConnected = True
            End If
            oReader.Close()
            oReader = Nothing

        Catch Err As Exception
            ErrDisp("CheckExistsConnected : " + Err.Message + vbCrLf + "SQL : " + cSQL)
        End Try
    End Function

    Public Function GetNowFromServer(ByVal ConnYage As SqlConnection) As Date

        Dim oReader As SqlDataReader
        Dim cSQL As String = ""

        GetNowFromServer = CDate("01.01.50 00:00:00")

        Try
            cSQL = "select getdate() "
            oReader = GetSQLReader(cSQL, ConnYage)
            If oReader.Read() Then
                GetNowFromServer = oReader.GetDateTime(0)
            End If
            oReader.Close()
            oReader = Nothing

        Catch Err As Exception
            ErrDisp("GetNowFromServer : " + Err.Message)
        End Try
    End Function

    Public Function GetKur(ByVal cDoviz As String, ByVal dTarih As Date, ByVal ConnYage As SqlConnection) As Double

        Dim cSQL As String

        Try
            GetKur = 0

            If cDoviz = "" Then
                GetKur = 0
            ElseIf cDoviz = "TL" Or cDoviz = "YTL" Then
                GetKur = 1
            Else
                cSQL = "set dateformat dmy " + _
                         " select kur = coalesce(kur,0)  " + _
                         " from dovkur " + _
                         " where doviz = '" + cDoviz + "' " + _
                         " and kurcinsi = 'Kur' " + _
                         " and tarih = '" + SQLWriteDate(dTarih) + "' "

                GetKur = SQLGetDoubleConnected(cSQL, ConnYage)
            End If

        Catch Err As Exception
            ErrDisp("GetKur : " + cDoviz + "/" + Err.Message)
        End Try
    End Function

    Public Function G_CBool(ByVal cParam As String) As Boolean
        G_CBool = False
        If cParam = "1" Or cParam = "2" Or cParam = "E" Or cParam = "e" Then
            G_CBool = True
        End If
    End Function

    Public Sub ReturnSingleRow(ByVal cMessage As String)
        Dim Record As New SqlDataRecord(New SqlMetaData("stringcol", SqlDbType.NVarChar, 4000))
        Record.SetSqlString(0, Mid(cMessage, 1, 4000).Trim)
        SqlContext.Pipe.Send(Record)
    End Sub

    Public Sub ReturnSQLDouble(ByVal nSonuc As Double)
        Dim Record As New SqlDataRecord(New SqlMetaData("doublecol", SqlDbType.Decimal))
        Record.SetSqlDouble(0, nSonuc)
        SqlContext.Pipe.Send(Record)
    End Sub

    Public Function fMin(ByVal ParamArray Values() As Double) As Double
        Dim nCnt As Integer
        fMin = 0
        If IsArray(Values) Then
            For nCnt = 0 To UBound(Values)
                If fMin > Values(nCnt) Then
                    fMin = Values(nCnt)
                End If
            Next
        End If
    End Function

    Public Function fMax(ByVal ParamArray Values() As Double) As Double
        Dim nCnt As Integer
        fMax = 0
        If IsArray(Values) Then
            For nCnt = 0 To UBound(Values)
                If fMax < Values(nCnt) Then
                    fMax = Values(nCnt)
                End If
            Next
        End If
    End Function

    Public Function fAvg(ByVal ParamArray Values() As Double) As Double
        Dim nCnt As Integer
        Dim nTotal As Double
        fAvg = 0
        If IsArray(Values) Then
            For nCnt = 0 To UBound(Values)
                nTotal = nTotal + Values(nCnt)
            Next
            fAvg = nTotal / nCnt
        End If
    End Function

    Private Function IsEmpty(ByVal lThisItem As Object) As Boolean
        IsEmpty = True
    End Function

    Public Sub ErrDisp(Optional ByVal cExplanation As String = "", Optional ByVal cKey As String = "", Optional ByVal cSQL As String = "")
        Try
            File.AppendAllText("c:\WinTexClrErr.txt", "V.1.0-Err : " + Now.ToString + ";" + cKey.Trim + ";" + cExplanation.Trim + ";" + cSQL + vbCrLf)
        Catch ex As Exception
            ' on error do nothing
        End Try
    End Sub

    Public Sub JustForLog(Optional ByVal cMessage As String = "")
        Try
            File.AppendAllText("c:\WinTexClrLog.txt", "V.1.0-Log : " + Now.ToString + ";" + cMessage.Trim + vbCrLf)
        Catch ex As Exception
            ' on error do nothing
        End Try
    End Sub

    Public Sub ErrDispConnected(ByVal ConnYage As SqlConnection, Optional ByVal cExplanation As String = "", Optional ByVal cKey As String = "")

        Dim cSQL As String = ""

        Try
            cSQL = "insert into clrerrlog (tarih, recordkey, errorexp) " + _
                    " values (" + _
                    " '" + SQLWriteDate(Now) + "', " + _
                    " '" + cKey.Trim + "', " + _
                    " '" + cExplanation.Trim + "') "

            ExecuteSQLCommandConnected(cSQL, ConnYage, True)

        Catch ex As Exception
            Debug.WriteLine("ErrDisp " + ex.Message + " " + cKey.Trim + " " + cExplanation.Trim)
        End Try
    End Sub

 
    Public Function StrStrip(cText As String) As String
        StrStrip = ""
        StrStrip = Replace(cText, Chr(13), " ")
    End Function

    Public Function StrStrip2(cText As String) As String
        StrStrip2 = ""
        StrStrip2 = Replace(cText, Chr(13), " ")
        StrStrip2 = Replace(StrStrip2, Chr(10), " ")
    End Function

    Public Function StrStripLettersNumbers(cText As String, Optional lReplaceBadCharactersWithBlank As Boolean = True, Optional lDeleteSpace As Boolean = False, Optional lMaxLen As Integer = 0) As String

        Dim nCnt As Integer
        Dim nMaxLen As Integer
        Dim cBuffer As String

        nMaxLen = Len(cText)
        StrStripLettersNumbers = ""

        For nCnt = 1 To nMaxLen
            cBuffer = Mid(cText, nCnt, 1)

            If (Asc(cBuffer) > 47 And Asc(cBuffer) < 58) Or _
                (Asc(cBuffer) > 64 And Asc(cBuffer) < 91) Or _
                (Asc(cBuffer) > 96 And Asc(cBuffer) < 123) Then
                If lDeleteSpace Then
                    If cBuffer <> " " Then
                        StrStripLettersNumbers = StrStripLettersNumbers + cBuffer
                    End If
                Else
                    StrStripLettersNumbers = StrStripLettersNumbers + cBuffer
                End If
            Else
                If lReplaceBadCharactersWithBlank Then StrStripLettersNumbers = StrStripLettersNumbers + " "
            End If
        Next
        If lMaxLen > 0 Then
            StrStripLettersNumbers = Mid(StrStripLettersNumbers, 1, lMaxLen)
        End If
        StrStripLettersNumbers = Trim(StrStripLettersNumbers)
    End Function

    Public Function StrStripNumbers(cText As String, Optional lReplaceBadCharactersWithBlank As Boolean = False) As String

        Dim nCnt As Integer
        Dim nMaxLen As Integer
        Dim cBuffer As String

        nMaxLen = Len(cText)
        StrStripNumbers = ""

        For nCnt = 1 To nMaxLen
            cBuffer = Mid(cText, nCnt, 1)
            If (Asc(cBuffer) > 47 And Asc(cBuffer) < 58) Then
                StrStripNumbers = StrStripNumbers + cBuffer
            Else
                If lReplaceBadCharactersWithBlank Then StrStripNumbers = StrStripNumbers + " "
            End If
        Next
        StrStripNumbers = Trim(StrStripNumbers)
    End Function

 
End Module
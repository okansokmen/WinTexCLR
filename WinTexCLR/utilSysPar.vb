Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module utilSysPar

    Public Function GetSysPar(ByVal cParameterName As String) As String

        Dim ConnYage As SqlClient.SqlConnection

        ConnYage = OpenConn()
        GetSysPar = GetSysParConnected(cParameterName, ConnYage)
        CloseConn(ConnYage)
    End Function

    Public Sub SetSysPar(ByVal cParameterName As String, ByVal cParameterValue As String)

        Dim ConnYage As SqlClient.SqlConnection

        ConnYage = OpenConn()
        SetSysParConnected(cParameterName, cParameterValue, ConnYage)
        CloseConn(ConnYage)
    End Sub

    ' connected

    Public Function GetSysParConnected(ByVal cParameterName As String, ByVal ConnYage As SqlConnection) As String

        Dim cSQL As String = ""

        GetSysParConnected = ""

        Try
            cSQL = "select parametervalue " + _
                        " from syspar " + _
                        " where parametername = '" + cParameterName.Trim + "' "

            GetSysParConnected = SQLGetStringConnected(cSQL, ConnYage)

            If GetSysParConnected = "" Then
                GetSysParConnected = SysDefault(cParameterName)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "GetSysParConnected", cSQL)
        End Try
    End Function

    Public Sub SetSysParConnected(ByVal cParameterName As String, ByVal cParameterValue As String, ByVal ConnYage As SqlConnection)

        Dim cSQL As String = ""

        Try

            cSQL = "delete syspar where parametername = '" + cParameterName.Trim + "' "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "insert into syspar (parametername,parametervalue) " + _
                                " values ('" + cParameterName.Trim + "', " + _
                                        " '" + cParameterValue.Trim + "') "
            ExecuteSQLCommandConnected(cSQL, ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "SetSysParConnected", cSQL)
        End Try
    End Sub

    Public Function GetFisNo(ByVal cKeyField As String, Optional ByVal cFormat As String = "", Optional ByVal cParameterType As String = "") As String

        Dim ConnYage As SqlClient.SqlConnection

        GetFisNo = ""

        Try
            ConnYage = OpenConn()
            GetFisNo = GetFisNoConnected(ConnYage, cKeyField, cFormat)
            Call CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "GetFisNo")
        End Try
    End Function

    Public Function GetFisNoConnected(ByVal ConnYage As SqlClient.SqlConnection, ByVal cKeyField As String, Optional ByVal cFormat As String = "") As String

        Dim cSQL As String = ""
        Dim nFisNo As Double = 0

        GetFisNoConnected = ""

        Try

            cSQL = "select parametervalue " + _
                    " from syspar " + _
                    " where parametername = '" + cKeyField + "' "

            If CheckExistsConnected(cSQL, ConnYage) Then
                nFisNo = CDbl(SQLGetStringConnected(cSQL, ConnYage))
                nFisNo = nFisNo + 1
                If cFormat = "" Then
                    GetFisNoConnected = SQLWriteDecimal(nFisNo, True)
                Else
                    GetFisNoConnected = Microsoft.VisualBasic.Format(nFisNo, cFormat)
                End If
                GetFisNoConnected = GetFisNoConnected.Trim

                cSQL = "update syspar " + _
                        " set parametervalue = " + nFisNo.ToString + _
                        " where parametername = '" + cKeyField + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Else
                cSQL = "insert into syspar (parametername, parametervalue) " + _
                        " values ('" + cKeyField.Trim + "','1') "

                ExecuteSQLCommandConnected(cSQL, ConnYage)

                nFisNo = 1
                If cFormat = "" Then
                    GetFisNoConnected = "1"
                Else
                    GetFisNoConnected = Microsoft.VisualBasic.Format(nFisNo, cFormat)
                End If
                GetFisNoConnected = GetFisNoConnected.Trim
            End If

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "GetFisNoConnected", cSQL)
        End Try
    End Function
End Module

Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module utilSQLTemp

    Public Function CreateTempView(ByVal ConnYage As SqlConnection, Optional cSQL As String = "") As String

        Dim cView As String = ""

        CreateTempView = ""

        Try
            Randomize()
            cView = "TmpV_" + CStr(Int(Rnd() * 10000)) ' + CStr(CInt(Int((6 * Rnd()) + 10000)))
            DropView(cView, ConnYage)

            If cSQL.Trim <> "" Then
                cSQL = "create view " + cView + " as " + cSQL
                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If
            CreateTempView = cView

        Catch ex As Exception
            ErrDisp(ex.Message, "CreateTempView", cSQL)
        End Try
    End Function

    Public Function CreateTempTable(ByVal ConnYage As SqlConnection, Optional cSQL As String = "", Optional cTableName As String = "") As String

        CreateTempTable = ""

        Try
            If cTableName.Trim = "" Then
                Randomize()
                cTableName = "TmpT_" + CStr(Int(Rnd() * 10000)) 'CStr(CInt(Int((6 * Rnd()) + 10000)))
            End If
            DropTable(cTableName, ConnYage)

            If cSQL.Trim <> "" Then
                cSQL = "create table " + cTableName + " " + cSQL
                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If
            CreateTempTable = cTableName

        Catch ex As Exception
            ErrDisp(ex.Message, "CreateTempTable", cSQL)
        End Try
    End Function

    Public Sub DropView(ByVal cViewName As String, ByVal ConnYage As SqlConnection)

        Dim cSQL As String = ""

        cSQL = "IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[" + cViewName + "]')) " + _
               "DROP VIEW [dbo].[" + cViewName + "]"

        ExecuteSQLCommandConnected(cSQL, ConnYage)
    End Sub

    Public Sub DropTable(ByVal cTableName As String, ByVal ConnYage As SqlConnection)

        Dim cSQL As String = ""

        cSQL = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + cTableName + "]') AND type in (N'U')) " + _
               "DROP TABLE [dbo].[" + cTableName + "]"

        ExecuteSQLCommandConnected(cSQL, ConnYage)
    End Sub
End Module

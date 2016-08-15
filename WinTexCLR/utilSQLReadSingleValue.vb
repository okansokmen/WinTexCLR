Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module utilSQLReadSingleValue

    Public Function SQLGetString(ByVal cSQL As String) As String

        Dim ConnYage As SqlClient.SqlConnection

        SQLGetString = ""

        If cSQL.Trim = "" Then Exit Function

        Try
            ConnYage = OpenConn()
            SQLGetString = SQLGetStringConnected(cSQL, ConnYage)
            Call CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetString", cSQL)
        End Try
    End Function

    Public Function SQLGetDouble(ByVal cSQL As String) As Double

        Dim ConnYage As SqlClient.SqlConnection

        SQLGetDouble = 0

        If cSQL.Trim = "" Then Exit Function

        Try
            ConnYage = OpenConn()
            SQLGetDouble = SQLGetDoubleConnected(cSQL, ConnYage)
            Call CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetDouble", cSQL)
        End Try
    End Function

    Public Function SQLGetDate(ByVal cSQL As String) As DateTime

        Dim ConnYage As SqlClient.SqlConnection

        SQLGetDate = CDate("01.01.1950")

        If cSQL.Trim = "" Then Exit Function

        Try
            ConnYage = OpenConn()
            SQLGetDate = SQLGetDateConnected(cSQL, ConnYage)
            Call CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetDate", cSQL)
        End Try
    End Function

    Public Function SQLGetInteger(ByVal cSQL As String) As Integer

        Dim ConnYage As SqlClient.SqlConnection

        SQLGetInteger = 0

        Try
            If cSQL.Trim = "" Then Exit Function

            ConnYage = OpenConn()
            SQLGetInteger = SQLGetIntegerConnected(cSQL, ConnYage)
            Call CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetInteger", cSQL)
        End Try
    End Function

    ' connected

    Public Function SQLGetStringConnected(ByVal cSQL As String, ByVal ConnYage As SqlConnection) As String

        Dim oReader As SqlDataReader

        SQLGetStringConnected = ""

        If cSQL.Trim = "" Then Exit Function

        Try
            oReader = GetSQLReader(cSQL, ConnYage)
            If oReader.Read() Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLGetStringConnected = ""
                Else
                    SQLGetStringConnected = oReader.GetString(0).Trim()
                End If
            End If
            oReader.Close()
            oReader = Nothing

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetStringConnected", cSQL)
        End Try
    End Function

    Public Function SQLGetDoubleConnected(ByVal cSQL As String, ByVal ConnYage As SqlConnection) As Double

        Dim oReader As SqlDataReader

        SQLGetDoubleConnected = 0

        If cSQL.Trim = "" Then Exit Function

        Try
            oReader = GetSQLReader(cSQL, ConnYage)
            If oReader.Read() Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLGetDoubleConnected = 0
                ElseIf IsNumeric(oReader.GetValue(0)) Then
                    SQLGetDoubleConnected = oReader.GetDecimal(0)
                Else
                    SQLGetDoubleConnected = 0
                End If
            End If
            oReader.Close()
            oReader = Nothing

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetDoubleConnected", cSQL)
        End Try
    End Function


    Public Function SQLGetDateConnected(ByVal cSQL As String, ByVal ConnYage As SqlConnection) As DateTime

        Dim oReader As SqlDataReader

        SQLGetDateConnected = CDate("01.01.1950")

        If cSQL.Trim = "" Then Exit Function

        Try
            oReader = GetSQLReader(cSQL, ConnYage)
            If oReader.Read() Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLGetDateConnected = CDate("01.01.1950")
                ElseIf IsDate(oReader.GetValue(0)) Then
                    SQLGetDateConnected = oReader.GetDateTime(0)
                Else
                    SQLGetDateConnected = CDate("01.01.1950")
                End If
            End If
            oReader.Close()
            oReader = Nothing

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetDateConnected", cSQL)
        End Try
    End Function

    Public Function SQLGetIntegerConnected(ByVal cSQL As String, ByVal ConnYage As SqlConnection) As Integer

        Dim oReader As SqlDataReader

        SQLGetIntegerConnected = 0

        Try
            If cSQL.Trim = "" Then Exit Function

            oReader = GetSQLReader(cSQL, ConnYage)
            If oReader.Read() Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLGetIntegerConnected = 0
                Else
                    SQLGetIntegerConnected = oReader.GetInt32(0)
                End If
            End If
            oReader.Close()
            oReader = Nothing

        Catch Err As Exception
            ErrDisp(Err.Message.Trim, "SQLGetDateConnected", cSQL)
        End Try
    End Function

    Public Function SQLBuildFilterString2(ByVal ConnYage As SqlConnection, ByVal cSQL As String, Optional ByVal lFilterMode As Boolean = True, Optional ByVal lUnique As Boolean = True) As String

        Dim oReader As SqlDataReader
        Dim cFilter As String = ""

        SQLBuildFilterString2 = ""

        Try
            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                If Not IsDBNull(oReader.GetValue(0)) Then
                    BuildFilterString(cFilter, oReader.GetValue(0).ToString.Trim, lFilterMode, lUnique)
                End If
            Loop
            oReader.Close()

            SQLBuildFilterString2 = cFilter

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "SQLBuildFilterString2", cSQL)
        End Try
    End Function

    Public Function ArrayToString(ByRef aStrArray() As String, Optional ByVal lFilterMode As Boolean = True) As String

        Dim nCnt As Integer = 0

        ArrayToString = ""

        Try
            For nCnt = 0 To UBound(aStrArray)
                BuildFilterString(ArrayToString, Trim(aStrArray(nCnt)), lFilterMode)
            Next

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "ArrayToString")
        End Try
    End Function

    Public Sub BuildFilterString(ByRef cFilter As String, cValue As String, Optional ByVal lFilterMode As Boolean = True, Optional ByVal lUnique As Boolean = True)

        Try
            cValue = cValue.Trim
            cFilter = cFilter.Trim

            If cValue = "" Then Exit Sub

            If lUnique Then
                If InStr(cFilter, cValue) = 0 Then
                    If cFilter = "" Then
                        If lFilterMode Then
                            cFilter = "'" + cValue + "'"
                        Else
                            cFilter = cValue
                        End If
                    Else
                        If lFilterMode Then
                            cFilter = cFilter + ",'" + cValue + "'"
                        Else
                            cFilter = cFilter + "," + cValue
                        End If
                    End If
                End If
            Else
                If cFilter = "" Then
                    If lFilterMode Then
                        cFilter = "'" + cValue + "'"
                    Else
                        cFilter = cValue
                    End If
                Else
                    If lFilterMode Then
                        cFilter = cFilter + ",'" + cValue + "'"
                    Else
                        cFilter = cFilter + "," + cValue
                    End If
                End If
            End If

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "BuildFilterString")
        End Try
    End Sub

    Public Sub StrAddToArray(ByRef aStrArray() As String, ByVal cValue As String)

        Dim nCnt As Integer = 0
        Dim nFound As Integer = -1

        Try
            cValue = cValue.Trim

            If cValue = "" Then Exit Sub

            If aStrArray(0) = "" Then
                aStrArray(0) = cValue
            Else
                For nCnt = 0 To UBound(aStrArray)
                    If aStrArray(nCnt) = cValue Then
                        nFound = nCnt
                        Exit For
                    End If
                Next
                If nFound = -1 Then
                    nCnt = UBound(aStrArray) + 1
                    ReDim Preserve aStrArray(nCnt)
                    aStrArray(nCnt) = cValue
                End If
            End If

        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "StrAddToArray")
        End Try
    End Sub

End Module

Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module utilSQLRW
    ' read
    Public Function SQLReadString(ByVal oReader As SqlDataReader, Optional ByVal cFieldName As String = "", Optional ByVal nWidth As Integer = 0) As String
        SQLReadString = ""
        Try
            If cFieldName = "" Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLReadString = ""
                ElseIf IsNothing(oReader.GetValue(0)) Then
                    SQLReadString = ""
                Else
                    SQLReadString = oReader.GetString(0).Trim()
                End If
            Else
                If IsDBNull(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadString = ""
                ElseIf IsNothing(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadString = ""
                Else
                    SQLReadString = oReader.GetString(oReader.GetOrdinal(cFieldName)).Trim()
                End If
            End If
            SQLReadString = Replace(SQLReadString, "'", "")
            If nWidth > 0 Then
                SQLReadString = Mid(SQLReadString, 1, nWidth).Trim
            End If
            If IsNothing(SQLReadString) Then
                SQLReadString = ""
            End If
            If IsDBNull(SQLReadString) Then
                SQLReadString = ""
            End If

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLReadString", cFieldName)
        End Try
    End Function

    Public Function SQLReadDouble(ByVal oReader As SqlDataReader, Optional ByVal cFieldName As String = "") As Double
        SQLReadDouble = 0
        Try
            If cFieldName = "" Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLReadDouble = 0
                ElseIf IsNothing(oReader.GetValue(0)) Then
                    SQLReadDouble = 0
                ElseIf IsNumeric(oReader.GetValue(0)) Then
                    SQLReadDouble = oReader.GetDecimal(0)
                End If
            Else
                If IsDBNull(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadDouble = 0
                ElseIf IsNothing(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadDouble = 0
                ElseIf IsNumeric(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadDouble = oReader.GetDecimal(oReader.GetOrdinal(cFieldName))
                End If
            End If

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLReadDouble", cFieldName)
        End Try
    End Function

    Public Function SQLReadInteger(ByVal oReader As SqlDataReader, Optional ByVal cFieldName As String = "") As Integer
        SQLReadInteger = 0
        Try
            If cFieldName = "" Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLReadInteger = 0
                ElseIf IsNothing(oReader.GetValue(0)) Then
                    SQLReadInteger = 0
                ElseIf IsNumeric(oReader.GetValue(0)) Then
                    SQLReadInteger = oReader.GetInt32(0)
                End If
            Else
                If IsDBNull(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadInteger = 0
                ElseIf IsNothing(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadInteger = 0
                ElseIf IsNumeric(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadInteger = oReader.GetInt32(oReader.GetOrdinal(cFieldName))
                End If
            End If

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLReadInteger", cFieldName)
        End Try
    End Function

    Public Function SQLReadDate(ByVal oReader As SqlDataReader, Optional ByVal cFieldName As String = "") As Date
        SQLReadDate = #1/1/1950#
        Try
            If cFieldName = "" Then
                If IsDBNull(oReader.GetValue(0)) Then
                    SQLReadDate = #1/1/1950#
                ElseIf IsNothing(oReader.GetValue(0)) Then
                    SQLReadDate = #1/1/1950#
                ElseIf IsDate(oReader.GetDateTime(0)) Then
                    SQLReadDate = oReader.GetDateTime(0)
                End If
            Else
                If IsDBNull(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadDate = #1/1/1950#
                ElseIf IsNothing(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadDate = #1/1/1950#
                ElseIf IsDate(oReader.GetValue(oReader.GetOrdinal(cFieldName))) Then
                    SQLReadDate = oReader.GetDateTime(oReader.GetOrdinal(cFieldName))
                End If
            End If

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLReadDate", cFieldName)
        End Try
    End Function

    ' write value 

    Public Function SQLWriteString(ByVal cValue As String, Optional ByVal nLength As Integer = 0) As String
        SQLWriteString = ""
        Try
            If cValue.Trim = "" Then Exit Function

            SQLWriteString = Replace(cValue, "'", " ", 1, -1).Trim

            If nLength <> 0 Then
                SQLWriteString = Mid(SQLWriteString, 1, nLength).Trim
            End If
            SQLWriteString = SQLWriteString.Trim

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLWriteString")
        End Try
    End Function

    Public Function SQLWriteDecimal(ByVal nValue As Object, Optional ByVal lFullClean As Boolean = False) As String
        SQLWriteDecimal = "0"
        Try
            If IsNumeric(nValue) Then
                nValue = CDbl(nValue)
            Else
                nValue = 0
            End If
            SQLWriteDecimal = Microsoft.VisualBasic.Format(nValue, "G")
            SQLWriteDecimal = Replace(SQLWriteDecimal, ",", ".")
            'If LCase(SQLWriteDecimal) = "nan" Or SQLWriteDecimal = "Infinity" Then SQLWriteDecimal = "0"
            'If lFullClean Then
            '    SQLWriteDecimal = Replace(SQLWriteDecimal, ",", "")
            '    SQLWriteDecimal = Replace(SQLWriteDecimal, ".", "")
            'Else
            '    SQLWriteDecimal = Replace(SQLWriteDecimal, ",", "")
            'End If

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLWriteDecimal", SQLWriteDecimal)
        End Try
    End Function

    'Public Function SQLWriteStringDecimal(ByVal cValue As String, Optional ByVal lFullClean As Boolean = False) As String
    '    SQLWriteStringDecimal = "0"
    '    Try
    '        If cValue = "" Then cValue = "0"
    '        If Not IsNumeric(cValue) Then cValue = "0"
    '        SQLWriteStringDecimal = cValue
    '        If lFullClean Then
    '            SQLWriteStringDecimal = Replace(SQLWriteStringDecimal, ",", "")
    '            SQLWriteStringDecimal = Replace(SQLWriteStringDecimal, ".", "")
    '        Else
    '            SQLWriteStringDecimal = Replace(SQLWriteStringDecimal, ",", "")
    '        End If

    '    Catch Err As Exception
    '        ErrDisp(Err.Message, "SQLWriteStringDecimal", SQLWriteStringDecimal)
    '    End Try
    'End Function

    Public Function SQLWriteInteger(ByVal nValue As Integer) As String
        SQLWriteInteger = "0"
        Try
            If IsNumeric(nValue) Then
                SQLWriteInteger = nValue.ToString
                SQLWriteInteger = Replace(SQLWriteInteger, ",", ".")
            End If

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLWriteInteger")
        End Try
    End Function

    Public Function SQLWriteDate(ByVal dValue As Date) As String
        SQLWriteDate = "01.01.1950"
        Try
            If IsDate(dValue) Then
                SQLWriteDate = Microsoft.VisualBasic.Format(dValue.Date, "dd.MM.yyyy")
            End If

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLWriteDate")
        End Try
    End Function

    ' array

    Public Function SQLtoStringArrayConnected(ByVal cSQL As String, ConnYage As SqlClient.SqlConnection, Optional ByVal cDefault As String = "", Optional cVarType As String = "string") As String()
        Dim dr As SqlDataReader
        Dim aResult() As String = Nothing
        Dim nCnt As Integer = 0

        ReDim aResult(0)
        SQLtoStringArrayConnected = aResult

        Try

            nCnt = 0
            If cDefault.Trim <> "" Then
                ReDim Preserve aResult(nCnt)
                aResult(nCnt) = cDefault.Trim
                nCnt = nCnt + 1
            End If

            dr = New SqlCommand(cSQL, ConnYage).ExecuteReader

            Do While dr.Read

                If cVarType = "string" Then
                    If SQLReadString(dr) <> "" Then
                        ReDim Preserve aResult(nCnt)
                        aResult(nCnt) = SQLReadString(dr)
                        nCnt = nCnt + 1
                    End If
                ElseIf cVarType = "integer" Then
                    If SQLReadInteger(dr) <> 0 Then
                        ReDim Preserve aResult(nCnt)
                        aResult(nCnt) = SQLReadInteger(dr).ToString
                        nCnt = nCnt + 1
                    End If
                End If

            Loop
            dr.Close()
            SQLtoStringArrayConnected = aResult

        Catch Err As Exception
            ErrDisp(Err.Message, "SQLtoStringArrayConnected", cSQL)
        End Try
    End Function

    Public Function SQLtoStringArray(ByVal cSQL As String, Optional ByVal cDefault As String = "", Optional cVarType As String = "string") As String()

        Dim ConnYage As SqlClient.SqlConnection

        SQLtoStringArray = Nothing

        Try
            ConnYage = OpenConn()
            SQLtoStringArray = SQLtoStringArrayConnected(cSQL, ConnYage, cDefault, cVarType)
            CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp(ex.Message, "SQLtoStringArray", cSQL)
        End Try
    End Function

End Module

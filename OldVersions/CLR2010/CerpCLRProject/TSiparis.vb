Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Module TSiparis

    Public Function OtoTSipKapat(ByVal cSiparisNo As String) As String

        Dim ConnYage As SqlConnection

        OtoTSipKapat = "OK"

        Try
            ConnYage = OpenConn()
            OtoTSipKapat = OtoTSipKapatConnected(ConnYage, cSiparisNo)
            CloseConn(ConnYage)
        Catch Err As Exception
            OtoTSipKapat = "Hata"
            ErrDisp("Error OtoTSipKapatConnected " + Err.Message)
        End Try
    End Function

    Public Function OtoTSipKapatConnected(ByVal ConnYage As SqlConnection, ByVal cSiparisNo As String) As String

        Dim cSQL As String = ""

        OtoTSipKapatConnected = "OK"

        Try

            cSQL = "select a.siparisno " + _
                    " from tsiparis a, tsipmodel b " + _
                    " where a.siparisno = '" + cSiparisNo.Trim + "' " + _
                    " and a.siparisno = b.siparisno " + _
                    " and (a.dosyakapandi = 'H' or a.dosyakapandi is null) " + _
                    " and coalesce(b.adet,0) > coalesce(b.gidenadet,0) "

            If Not CheckExistsConnected(cSQL, ConnYage) Then

                cSQL = "update tsiparis " + _
                        " set dosyakapandi = 'E', " + _
                        " kapanistarihi = getdate() " + _
                        " where siparisno = '" + cSiparisNo.Trim + "' "

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            End If

        Catch Err As Exception
            OtoTSipKapatConnected = "Hata"
            ErrDisp("Error OtoTSipKapatConnected " + Err.Message)
        End Try
    End Function
End Module

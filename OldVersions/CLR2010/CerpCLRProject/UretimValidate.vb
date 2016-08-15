Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Module UretimValidate

    Private Structure oSiparis
        Dim cModelNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim nRCount As Double
    End Structure

    Public Function SiparisUretimAdet(ByVal SiparisNo As String) As String

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nUretilen As Double
        Dim nUrtKirik As Double
        Dim nUrtSakat As Double
        Dim oSysFlags As SysFlags = Nothing
        Dim aSipModel() As oSiparis
        Dim lFirst As Boolean = True
        Dim nCnt As Integer

        ReDim aSipModel(0)

        ReadSysFlagsMain(oSysFlags)

        ConnYage = OpenConn()

        cSQL = "Select * From sipmodel where siparisno='" + SiparisNo + "'"

        oReader = GetSQLReader(cSQL, ConnYage)

        Do While oReader.Read
            If lFirst Then
                lFirst = False
                aSipModel(0).cModelNo = SQLReadString(oReader, "ModelNo")
                aSipModel(0).cRenk = SQLReadString(oReader, "renk")
                aSipModel(0).cBeden = SQLReadString(oReader, "beden")
                aSipModel(0).nRCount = SQLReadDouble(oReader, "rcount")
            Else
                ReDim Preserve aSipModel(UBound(aSipModel) + 1)
                aSipModel(UBound(aSipModel)).cModelNo = SQLReadString(oReader, "ModelNo")
                aSipModel(UBound(aSipModel)).cRenk = SQLReadString(oReader, "renk")
                aSipModel(UBound(aSipModel)).cBeden = SQLReadString(oReader, "beden")
                aSipModel(UBound(aSipModel)).nRCount = SQLReadDouble(oReader, "rcount")
            End If
        Loop
        oReader.Close()
        oReader = Nothing

        For nCnt = 0 To UBound(aSipModel)

            cSQL = " select Miktar=coalesce(Sum(b.adet),0)" + _
                    " from uretharfis a,uretharrba b " + _
                    " Where a.uretfisno = b.uretfisno " + _
                    " and a.girisdept='" + oSysFlags.G_DeptMamul + "'" + _
                    " and b.uretimtakipno='" + SiparisNo + "' " + _
                    " and b.modelno='" + aSipModel(nCnt).cModelNo + "' " + _
                    " and b.renk='" + aSipModel(nCnt).cRenk + "' " + _
                    " and b.beden ='" + aSipModel(nCnt).cBeden + "' " + _
                    " and (b.sakatkodu='' or b.sakatkodu is null) "

            nUretilen = ReadSingleDoubleValueConnected(cSQL, ConnYage)

            cSQL = " select Miktar=coalesce(Sum(b.adet),0)" + _
                    " from uretharfis a,uretharrba b " + _
                    " Where a.uretfisno = b.uretfisno " + _
                    " and a.girisdept='" + oSysFlags.G_DeptMamul + "'" + _
                    " and b.uretimtakipno='" + SiparisNo + "' " + _
                    " and b.modelno='" + aSipModel(nCnt).cModelNo + "' " + _
                    " and b.renk='" + aSipModel(nCnt).cRenk + "' " + _
                    " and b.beden ='" + aSipModel(nCnt).cBeden + "' " + _
                    " and b.sakatkodu='Kırık'"
            nUrtKirik = ReadSingleDoubleValueConnected(cSQL, ConnYage)

            cSQL = " select Miktar=coalesce(Sum(b.adet),0)" + _
                    " from uretharfis a,uretharrba b " + _
                    " Where a.uretfisno = b.uretfisno " + _
                    " and a.girisdept='" + oSysFlags.G_DeptMamul + "'" + _
                    " and b.uretimtakipno='" + SiparisNo + "' " + _
                    " and b.modelno='" + aSipModel(nCnt).cModelNo + "' " + _
                    " and b.renk='" + aSipModel(nCnt).cRenk + "' " + _
                    " and b.beden ='" + aSipModel(nCnt).cBeden + "' " + _
                    " and b.sakatkodu in (select sakatkodu from sakat where(sakatdepo='II.KALITE DEPO' or sakatdepo='T.II.KALITE DEPO')) "
            nUrtSakat = ReadSingleDoubleValueConnected(cSQL, ConnYage)


            cSQL = " Update sipmodel " + _
                   " set uretilen='" + SQLWriteDecimal(nUretilen) + "', " + _
                   " uretilenkirik='" + SQLWriteDecimal(nUrtKirik) + "',  " + _
                   " uretilensakat='" + SQLWriteDecimal(nUrtSakat) + "'  " + _
                   " where rcount='" + SQLWriteDecimal(aSipModel(nCnt).nRCount) + "'"

            ExecuteSQLCommandConnected(cSQL, ConnYage)
        Next


        CloseConn(ConnYage)

        SiparisUretimAdet = "Ok"
    End Function
End Module

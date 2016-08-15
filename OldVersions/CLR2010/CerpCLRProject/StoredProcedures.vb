Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Collections
Imports System.Diagnostics
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports System.Runtime.InteropServices

Partial Public Class StoredProcedures

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub PlanlamaOtomatikKapatAll()
        OtomatikKapatmalar()
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub PlanlamaForwardAll(ByVal cSiparisNo As String)
        ForwardAll(cSiparisNo)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub HizliStokBakimi()
        HizliStokRBBakimi()
        HizliStokTopRBBakimi()
        HizliStokAksesuarRBBakimi()
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub FastSTFBuildAll()
        STFFastGenerateAll()
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub FastSTFBuild(ByVal cSTF As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = STFGenerate(cSTF)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub FastMTFBuildAll()
        MTFFastGenerateAll()
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub FastMTFBuild(ByVal cMTF As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = MTKFastGenerate(cMTF)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub FastUTFBuildAll()
        UTFFastGenerateAll()
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub FastUTFBuild(ByVal cUTF As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = UTFGenerate(cUTF)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub GetToplamSiparisView(cFilter As String, cTableName As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = GetToplamSiparisView_0(cFilter, cTableName)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub MTFHesaplax(cFilter1 As String, cFilter2 As String, cTSip As String, cTableName As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = MTFHesaplax_0(cFilter1, cFilter2, cTSip, cTableName)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub BuildMasterPlan(cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = GetMasterPlanData(cFilter)
    End Sub

     ' EndOfNewCode

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub ValidateStokFis(ByVal cAction As String, ByVal cStokFisNo As String, ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = SingleStokFisValidate(cAction, cStokFisNo, cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub ValidateMultiStokFis(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = MultiStokFisValidate(cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub ValidateTopluTransferFis(ByVal cAction As String, ByVal cTTFisNo As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = TopluTransferValidate(cAction, cTTFisNo)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub ValidateTransferFis(ByVal cAction As String, ByVal cTransferFisNo As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = TransferFisValidate(cAction, cTransferFisNo)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub ValidateMultiTransferFis(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = MultiTransferFisValidate(cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub TekStokToplam(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = SingleStokToplam(cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub BarkodEskiStokDurumu(ByVal cTarih As String, ByVal nKumasAks As Integer, <Out()> ByRef cTableName As String)
        cTableName = BarkodStokDurumu(cTarih, nKumasAks)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub KumasCikisBarkodlu(ByVal cStokFisNo As String, ByVal cStokHareketKodu As String, ByVal cDepartman As String, ByVal cFirma As String, ByVal cNotlar As String)
        Dim cSonuc As String
        cSonuc = BarkodluKumasCikis(cStokFisNo.Trim, cStokHareketKodu.Trim, cDepartman.Trim, cFirma.Trim, cNotlar.Trim)
        ReturnSingleRow(cSonuc)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub TarihStokDurumu(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal StokNo As String, ByVal nCase As Integer,
                                      <Out()> ByRef cTableName As SqlString)
        cTableName = StokDurumuTarih(BslTarihi, BtsTarihi, StokNo, nCase)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Sub SiparisTakip1()
        'r_SiparisDurumu1()
    End Sub

    
End Class

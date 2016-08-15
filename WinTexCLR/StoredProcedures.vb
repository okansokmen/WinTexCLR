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
Imports Microsoft.VisualBasic
Imports System.Reflection

Partial Public Class StoredProcedures
    '<Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyet8CLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
    '    nSonuc = 0
    '    Try
    '        nSonuc = STISonMaliyet8(cFilter)
    '    Catch ex As Exception
    '        ErrDisp(ex.Message.Trim, "STISonMaliyet8CLR")
    '    End Try
    'End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyetIhracatCLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISonMaliyetIhracat(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "StokFisKurTamamlaCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub StokFisKurTamamlaCLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = StokFisKurTamamla(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "StokFisKurTamamlaCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyetDigerMasrafCLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISonMaliyetDigerMasraf(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "STISonMaliyetDigerMasrafCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyet7CreateCLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISonMaliyet7Create(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "STISonMaliyet7CreateCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyetOnMaliyetCLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISonMaliyetOnMaliyet(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "STISonMaliyetOnMaliyetCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyetMalzemeCLR(ByVal cFilter As String, ByVal cFilter2 As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISonMaliyetMalzeme(cFilter, cFilter2)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "STISonMaliyetMalzemeCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyetUretimCLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISonMaliyetUretim(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "STISonMaliyetUretimCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISMPOpenRecordsCLR(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISMPOpenRecords(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "STISMPOpenRecordsCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub FastMTFBuildMulti(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = MTFFastGenerateMulti(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "FastMTFBuildMulti")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub STISonMaliyetInitialCleanupCLR(cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STISonMaliyetInitialCleanup(cFilter)
        Catch ex As Exception
            nSonuc = 0
            ErrDisp(ex.Message.Trim, "STISonMaliyetInitialCleanupCLR")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub WinTexCLRVersion(<Out()> ByRef nSonuc As SqlInt32)
        Try
            nSonuc = General.CLRVersion
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "WinTexCLRVersion")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub BuildMasterPlan(cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = GetMasterPlanData(cFilter)
        Catch ex As Exception
            nSonuc = 0
            ErrDisp(ex.Message.Trim, "BuildMasterPlan")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub PlanlamaOtomatikKapatAll()
        Try
            OtomatikKapatmalar()
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "PlanlamaOtomatikKapatAll")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub PlanlamaForwardAll(ByVal cSiparisNo As String)
        Try
            ForwardAll(cSiparisNo)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "PlanlamaForwardAll")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub HizliStokBakimi()
        Try
            HizliStokRBBakimi()
            HizliStokTopRBBakimi()
            HizliStokAksesuarRBBakimi()
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "HizliStokBakimi")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub FastSTFBuildAll(ByVal cFilter As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STFFastGenerateAll(cFilter)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "FastSTFBuildAll")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub FastSTFBuild(ByVal cSTF As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = STFGenerate(cSTF)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "FastSTFBuild")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub FastMTFBuildAll()
        Try
            MTFFastGenerateAll()
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "FastMTFBuildAll")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub FastMTFBuild(ByVal cMTF As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = MTKFastGenerate(cMTF)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "FastMTFBuild")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub FastUTFBuildAll()
        Try
            UTFFastGenerateAll()
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "FastUTFBuildAll")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub FastUTFBuild(ByVal cUTF As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = UTFGenerate(cUTF)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "FastUTFBuild")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub GetToplamSiparisView(cFilter As String, cTableName As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = GetToplamSiparisView_0(cFilter, cTableName)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "GetToplamSiparisView")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub MTFHesaplax(cFilter1 As String, cFilter2 As String, cTSip As String, cTableName As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = MTFHesaplax_0(cFilter1, cFilter2, cTSip, cTableName)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "MTFHesaplax")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub DMTFIsemriKilavuzu(ByVal cDetayIhtiyacTable As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = DMTFIsemriKilavuzuHesapla(cDetayIhtiyacTable)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "DMTFIsemriKilavuzu")
        End Try
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub UretimisEmriUretCLR(ByVal cUTF As String, ByVal cAction As String, ByVal cDepartman As String,
                                                                                    ByVal cDefaultFirma As String, ByVal cDefaultPersonel As String, ByVal cPadNo As String,
                                                                                    ByVal cPartiNo As String, ByVal cUserName As String, ByVal cDepts As String, ByVal cKesimSistemiNo As String,
                                                                                    <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = 0
        Try
            nSonuc = UretimisEmriUret(cUTF, cAction, cDepartman, cDefaultFirma, cDefaultPersonel, cPadNo, cPartiNo, cUserName, cDepts, cKesimSistemiNo)
        Catch ex As Exception
            ErrDisp(ex.Message.Trim, "UretimisEmriUretCLR")
        End Try
    End Sub

    ' EndOfNewCode

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub ValidateStokFis(ByVal cAction As String, ByVal cStokFisNo As String, ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = SingleStokFisValidate(cAction, cStokFisNo, cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub ValidateMultiStokFis(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = MultiStokFisValidate(cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub ValidateTopluTransferFis(ByVal cAction As String, ByVal cTTFisNo As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = TopluTransferValidate(cAction, cTTFisNo)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub ValidateTransferFis(ByVal cAction As String, ByVal cTransferFisNo As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = TransferFisValidate(cAction, cTransferFisNo)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub ValidateMultiTransferFis(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = MultiTransferFisValidate(cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub TekStokToplam(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, <Out()> ByRef nSonuc As SqlInt32)
        nSonuc = SingleStokToplam(cStokNo, cRenk, cBeden)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub BarkodEskiStokDurumu(ByVal cTarih As String, ByVal nKumasAks As Integer, <Out()> ByRef cTableName As String)
        cTableName = BarkodStokDurumu(cTarih, nKumasAks)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub KumasCikisBarkodlu(ByVal cStokFisNo As String, ByVal cStokHareketKodu As String, ByVal cDepartman As String, ByVal cFirma As String, ByVal cNotlar As String)
        Dim cSonuc As String
        cSonuc = BarkodluKumasCikis(cStokFisNo.Trim, cStokHareketKodu.Trim, cDepartman.Trim, cFirma.Trim, cNotlar.Trim)
        ReturnSingleRow(cSonuc)
    End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub TarihStokDurumu(ByVal BslTarihi As String, ByVal BtsTarihi As String, ByVal StokNo As String, ByVal nCase As Integer,
                                      <Out()> ByRef cTableName As SqlString)
        cTableName = StokDurumuTarih(BslTarihi, BtsTarihi, StokNo, nCase)
    End Sub

    '<Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub SiparisTakip1()
    '    r_SiparisDurumu1()
    'End Sub

    <Microsoft.SqlServer.Server.SqlProcedure> Public Shared Sub GenelGiderDagit(<Out()> ByRef nSonuc As SqlInt32)
        nSonuc = GenelGiderDagitimi()
    End Sub


End Class

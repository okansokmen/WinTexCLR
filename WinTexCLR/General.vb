Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic

Module General
    Public Const CLRVersion As Integer = 84
    Public Const lDebug As Boolean = True

    Public Const G_NumberFormat As String = "###,###,###,###,###,##0"
    Public Const G_Number1Format As String = "###,###,###,###,###,##0.0"
    Public Const G_Number2Format As String = "###,###,###,###,###,##0.00"
    Public Const G_Number3Format As String = "###,###,###,###,###,##0.000"
    Public Const G_Number4Format As String = "###,###,###,###,###,##0.0000"
    Public Const G_Number5Format As String = "###,###,###,###,###,##0.00000"
    Public Const G_Number6Format As String = "###,###,###,###,###,##0.000000"

    Public Structure oIplik
        Dim cHammadde As String
        Dim nYzd As Double
        Dim nYzdRnd As Double
    End Structure

    Public Structure stokrb
        Dim cTableName As String
        Dim cStokno As String
        Dim cRenk As String
        Dim cBeden As String
        Dim cPartiNo As String
        Dim cDepo As String
        Dim cMtk As String
        Dim cTopNo As String
        Dim cAy As String
        Dim cYil As String
        Dim nMiktar As Double
    End Structure

    Private Structure MTKIsemri
        Dim cMTF As String
        Dim cDepartman As String
        Dim cStokNo As String
        Dim cRenk As String
        Dim cBeden As String
        Dim nMiktar As Double
    End Structure

    Private Structure isemri
        Dim nSiraNo As Double
        Dim nMiktar As Double
    End Structure

    Public Structure SysFlags
        Dim G_WinAccessory As Boolean
        Dim G_WinFabric As Boolean
        Dim G_WANoMTK As Boolean
        Dim G_WFNoMTK As Boolean
        Dim G_AksDisableTopSiraNo As Boolean
        Dim G_KumasDisableTopSiraNo As Boolean
        Dim G_NoUpdateStokRB2 As Boolean
        Dim G_OrtalamaCalcIptal As Boolean
        Dim G_OrtMlytDoviz As String
        Dim G_OtoMlzIsemriKapat As Boolean
        Dim G_YuzdeGirisKontrol As Boolean
        Dim G_YuzdeGirisKontrolYuzde As Double
        Dim G_SevkStokDeposu As String
        Dim G_TSevkStokDeposu As String
        Dim G_FireCalcUp As Boolean
        Dim G_OnaysizFiyatFilter As String
        Dim G_DeptMamul As String
        Dim G_Date As String
        Dim G_SyeTolerans As Double
        Dim G_MTFKesileneGore As Boolean
        Dim G_MTFKesisEmrineGore As Boolean
        Dim G_Dahili As String
    End Structure

    Public Sub ReadSysFlags(ByRef oSysFlags As SysFlags, ByVal ConnYage As SqlConnection)
        oSysFlags.G_WinFabric = G_CBool(GetSysParConnected("kumastoptakibi", ConnYage))
        oSysFlags.G_WinAccessory = G_CBool(GetSysParConnected("aksesuartakibi", ConnYage))
        oSysFlags.G_WFNoMTK = G_CBool(GetSysParConnected("kumascikiskontrol", ConnYage))
        oSysFlags.G_WANoMTK = G_CBool(GetSysParConnected("aksesuarcikiskontrol", ConnYage))
        oSysFlags.G_AksDisableTopSiraNo = G_CBool(GetSysParConnected("aksdisabletopsirano", ConnYage))
        oSysFlags.G_KumasDisableTopSiraNo = G_CBool(GetSysParConnected("kumasdisabletopsirano", ConnYage))
        oSysFlags.G_NoUpdateStokRB2 = G_CBool(GetSysParConnected("noupdatestokrb2", ConnYage))
        oSysFlags.G_OrtalamaCalcIptal = G_CBool(GetSysParConnected("ortalamacalciptal", ConnYage))
        oSysFlags.G_OrtMlytDoviz = GetSysParConnected("ortalamamaliyetdovizi", ConnYage)
        oSysFlags.G_OtoMlzIsemriKapat = G_CBool(GetSysParConnected("otomlzisemrikapat", ConnYage))
        oSysFlags.G_YuzdeGirisKontrol = G_CBool(GetSysParConnected("yuzdegiriskontrol", ConnYage))
        oSysFlags.G_YuzdeGirisKontrolYuzde = CDbl(GetSysParConnected("yuzdegiriskontrolyuzde", ConnYage))
        oSysFlags.G_SevkStokDeposu = GetSysParConnected("sevkstokdeposu", ConnYage)
        oSysFlags.G_TSevkStokDeposu = GetSysParConnected("tsevkstokdeposu", ConnYage)
        oSysFlags.G_Dahili = GetSysParConnected("dahili", ConnYage)
        oSysFlags.G_FireCalcUp = G_CBool(GetSysParConnected("firecalcup", ConnYage))
        oSysFlags.G_DeptMamul = GetSysParConnected("sevkstokdepartmani", ConnYage)
        oSysFlags.G_SyeTolerans = CDbl(GetSysParConnected("syetolerans", ConnYage))
        oSysFlags.G_Date = "01.01.1950"
        oSysFlags.G_OnaysizFiyatFilter = " "

        'If G_CBool(GetSysPar("onaysizfiyatkullanilmasin", "integer", ConnYage)) Then
        '    oSysFlags.G_OnaysizFiyatFilter = "onay = 'E' "
        'Else
        '    oSysFlags.G_OnaysizFiyatFilter = " "
        'End If
    End Sub

    Public Sub CheckRenkBedenValidate(ByVal cStokNo As String, ByRef cRenk As String, ByRef cBeden As String, ByVal ConnYage As SqlConnection)

        Dim cSQL As String = ""
        Dim cMTE As String = ""

        Try
            If cStokNo.Trim = "" Then Exit Sub
            cSQL = "select maltakipesasi from stok where stokno = '" + cStokNo.Trim + "' "
            cMTE = SQLGetStringConnected(cSQL, ConnYage)
            Select Case cMTE
                Case "1"
                    cRenk = "HEPSI"
                    cBeden = "HEPSI"
                Case "2"
                    cBeden = "HEPSI"
                Case "3"
                    cRenk = "HEPSI"
            End Select
        Catch ex As Exception
            ' do nothing
        End Try
    End Sub

    Public Function GeTopSirano(ByVal cTopNo As String, ByVal cStokNo As String, ByVal cMTF As String, ByVal cRenk As String, ByVal cBeden As String, ByVal cPartiNo As String, ByVal oSysFlags As SysFlags, ByVal ConnYage As SqlConnection) As Double

        Dim cSQL As String

        GeTopSirano = 0

        Try
            If oSysFlags.G_KumasDisableTopSiraNo Then Exit Function

            cSQL = "select topsirano " + _
                        " from stoktoprb " + _
                        " where stokno = '" + cStokNo.Trim + "'" + _
                        " and renk = '" + cRenk.Trim + "'" + _
                        " and beden = '" + cBeden.Trim + "'" + _
                        " and partino = '" + cPartiNo.Trim + "'" + _
                        " and malzemetakipkodu = '" + cMTF.Trim + "' " + _
                        " and topsirano is not null " + _
                        " and topsirano <> 0 " + _
                        " order by topsirano desc "

            GeTopSirano = SQLGetDoubleConnected(cSQL, ConnYage) + 1
        Catch
            ErrDisp("Error GeTopSirano " + Err.Description.Trim)
        End Try
    End Function

    Public Function GetAksesuarSirano(ByVal cTopNo As String, ByVal cStokNo As String, ByVal cMTF As String, ByVal cRenk As String, ByVal cBeden As String, ByVal cPartiNo As String, ByVal oSysFlags As SysFlags, ByVal ConnYage As SqlConnection) As Double

        Dim cSQL As String

        GetAksesuarSirano = 0

        Try
            If oSysFlags.G_AksDisableTopSiraNo Then Exit Function

            cSQL = "select topsirano " + _
                        " from stokaksesuarrb " + _
                        " where stokno = '" + cStokNo.Trim + "'" + _
                        " and renk = '" + cRenk.Trim + "'" + _
                        " and beden = '" + cBeden.Trim + "'" + _
                        " and partino = '" + cPartiNo.Trim + "'" + _
                        " and malzemetakipkodu = '" + cMTF.Trim + "' " + _
                        " and topsirano is not null " + _
                        " and topsirano <> 0 " + _
                        " order by topsirano desc "

            GetAksesuarSirano = SQLGetDoubleConnected(cSQL, ConnYage) + 1
        Catch
            ErrDisp("Error GetAksesuarSirano " + Err.Description.Trim)
        End Try
    End Function


    Public Function SysDefault(ByVal cPName As String, Optional ByVal cPType As String = "string") As String

        SysDefault = ""

        Try
            Select Case LCase(cPName)
                Case "syetolerans"
                    SysDefault = "0"

                Case "smtpport"
                    SysDefault = "25"

                Case "urunbirim"
                    SysDefault = "AD"

                    ' model
                Case "altmodeltakibi"
                    SysDefault = "0"

                    ' diger
                Case "fislerdemaxgunfarki"
                    SysDefault = "30"

                    ' magaza
                Case "satissatirozet"
                    SysDefault = "0"

                    ' gomlek isi (Ravelli)
                Case "gomlekactive"
                    SysDefault = "1"
                Case "dokumaonmlytgomlek"
                    SysDefault = "0"

                    ' ous default mailer
                Case "smtpserveraddress"
                    SysDefault = "195.87.6.21"
                Case "smtpfromaddress"
                    SysDefault = "yagemailer@yage.com.tr"
                Case "smtpusername"
                    SysDefault = "yagemailer"
                Case "smtppassword"
                    SysDefault = "yage"

                    ' argox barkod yazici
                Case "printerport"
                    SysDefault = "COM1"
                Case "magazayazici"         ' magaza
                    SysDefault = "COM1"
                Case "magazaporthizi"       ' magaza
                    SysDefault = "9600"
                Case "magazaportparite"     ' magaza
                    SysDefault = "n"
                Case "magazaportdatabit"    ' magaza
                    SysDefault = "8"
                Case "magazaportstoppit"    ' magaza
                    SysDefault = "1"
                Case "barcodeprinterwait"    ' magaza
                    SysDefault = "3000"

                    ' terazi sistemi Hisar teraziye gore ayarlidir
                Case "ototartihisar"
                    SysDefault = "0"
                Case "winpacktarticom"
                    SysDefault = "1"
                Case "teraziporthizi"
                    SysDefault = "57600"
                Case "teraziportparite"
                    SysDefault = "n"
                Case "teraziportdatabit"
                    SysDefault = "8"
                Case "teraziportstoppit"
                    SysDefault = "1"

                Case "dahili"
                    SysDefault = "DAHiLi"
                Case "uretimisemriyaratkilavuz"
                    SysDefault = "0"
                Case "aksesuarbarkoddeposu"
                    SysDefault = "AKSESUAR DEPOSU"
                Case "kumasbarkoddeposu"
                    SysDefault = "KUMAS DEPOSU"
                Case "mamuldepo"
                    SysDefault = "MAMUL DEPO"
                Case "mtfraporcounter"
                    SysDefault = "0"
                Case "stokpricetype"
                    SysDefault = "0"
                Case "firecalcup"
                    SysDefault = "0"
                Case "wordtableno"
                    SysDefault = "18"
                Case "wordfontname"
                    SysDefault = "Courier New"
                Case "wordfontsize"
                    SysDefault = "10"
                Case "gensqllevel"
                    SysDefault = "1"
                Case "exceltableno"
                    SysDefault = "18"
                Case "excelfontname"
                    SysDefault = "Courier New"
                Case "excelfontsize"
                    SysDefault = "10"
                Case "textprgname"
                    SysDefault = ""
                Case "textfontname"
                    SysDefault = "Courier New"
                Case "textfontsize"
                    SysDefault = "10"
                Case "dontshowsipclosed"
                    SysDefault = "0"
                Case "dontshowisemriclosed"
                    SysDefault = "0"
                Case "stokdurumdetay"
                    SysDefault = "0"
                Case "repnoteshorizontal"
                    SysDefault = "0"
                Case "nUretFisRBAGather"
                    SysDefault = "0"
                Case "nUretFisRBALineNo"
                    SysDefault = "1"
                Case "word97"
                    SysDefault = "1"
                Case "defaultbedenseti"
                    SysDefault = "SML"
                Case "defaultunitset"
                    SysDefault = "AD"
                Case "excel97"
                    SysDefault = "1"
                Case "modifymtk"
                    SysDefault = "0"
                Case "modifyutk"
                    SysDefault = "0"
                Case "isemriotoparti"
                    SysDefault = "0"
                Case "isemriotopartidept"
                    SysDefault = ""
                Case "nisemripartino"
                    SysDefault = "0"
                Case "wordtop"
                    SysDefault = "25"
                Case "wordbottom"
                    SysDefault = "42"
                Case "wordleft"
                    SysDefault = "30"
                Case "wordright"
                    SysDefault = "30"
                Case "urtisrenkayri"
                    SysDefault = "0"
                Case "sipkayifoyuret"
                    SysDefault = "1"
                Case "cantescapefromselect"
                    SysDefault = "0"
                Case "linkirsaliyeno"
                    SysDefault = "0"
                Case "reportrichtext"
                    SysDefault = "0"
                Case "checkuretfistoplam"
                    SysDefault = "0"
                Case "otomatikuretimstokcikisi"
                    SysDefault = "0"
                Case "mtkserbestmiktar"
                    SysDefault = "0"
                Case "mtkreservemiktar"
                    SysDefault = "0"
                Case "utkirssingleprint"
                    SysDefault = "0"
                Case "kumasdepo"
                    SysDefault = "KUMAS"
                Case "kolidepo"
                    SysDefault = "KOLi"
                Case "autostokhareketi"
                    SysDefault = "0"
                Case "uretimtusutoplamgoster"
                    SysDefault = "0"
                Case "stokfisfiyatfromisemri"
                    SysDefault = "0"
                Case "modelaramadaresim"
                    SysDefault = "0"
                Case "parcakodupenceresi"
                    SysDefault = "0"
                Case "autodosyano"
                    SysDefault = "0"
                Case "uretimfirenoshow"
                    SysDefault = "0"
                Case "checkurthrkfis"
                    SysDefault = "0"
                Case "alternatifrecete"
                    SysDefault = "0"
                Case "uretimistogaisle"
                    SysDefault = "0"
                Case "sevkiyatistogaisle"
                    SysDefault = "0"
                Case "imalatciftkurtarihi"
                    SysDefault = "0"
                Case "ithsiptarihchecking"
                    SysDefault = "0"
                Case "repdatetimechecking"
                    SysDefault = "0"
                Case "sevkstokdepartmani"
                    SysDefault = "SEVKIYAT"
                Case "sevkstokdeposu"
                    SysDefault = "MAMUL"
                Case "uretimstokdeposu"
                    SysDefault = "MAMUL"
                Case "stcikiseldekimiktar"
                    SysDefault = "0"
                Case "getkalemfiyatfobcif"
                    SysDefault = "0"
                Case "ortalamamaliyetdovizi"
                    SysDefault = "TL"
                Case "getistatistikfiyatfobcif"
                    SysDefault = "0"
                Case "siparisaramaekranirenkli"
                    SysDefault = "0"
                Case "logogolddizini"
                    SysDefault = ""
                Case "logogatewaydizini"
                    SysDefault = ""
                Case "logousername"
                    SysDefault = ""
                Case "logouserpassword"
                    SysDefault = ""
                Case "logofirmnumber"
                    SysDefault = "1"
                Case "logoconnectionstatus"
                    SysDefault = "0"
                Case "logotransferupdate"
                    SysDefault = "0"
                Case "logotransfervisual"
                    SysDefault = "0"
                Case "logotransferfiyat"
                    SysDefault = "1"
                Case "logomagazaaktar"
                    SysDefault = "0"
                Case "logofiyatindirimli"
                    SysDefault = "0"
                Case "logomuhkoduaktar"
                    SysDefault = "0"
                Case "logokdvmuhaktar"
                    SysDefault = "0"
                Case "logocikiskdvmuh"
                    SysDefault = "391.01.001"
                Case "logogiriskdvmuh"
                    SysDefault = "191.02.001"
                Case "logocikiskdvalimiade"
                    SysDefault = "391.02.001"
                Case "logogiriskdvalim"
                    SysDefault = "191.01.001"
                Case "logoureticikoduaktar"
                    SysDefault = "0"
                Case "logokdvdahil"
                    SysDefault = "1"
                Case "logotoplufisaktar"
                    SysDefault = "0"
                Case "logosatirbirlestir"
                    SysDefault = "0"

                Case "stokkoduzunlugu"
                    SysDefault = "10"
                Case "satistakurus"
                    SysDefault = "0"
                Case "satisfiyatno"
                    SysDefault = "1"
                Case "toptansatisfiyatno"
                    SysDefault = "1"
                Case "alisfiyatno"
                    SysDefault = "1"
                Case "magazakodu"
                    SysDefault = ""
                Case "magazamerkeztelno"
                    SysDefault = "2744072"
                Case "magazakullanici"
                    SysDefault = "WinStore"
                Case "magazasifre"
                    SysDefault = "yage"
                Case "cantchangetextreport"
                    SysDefault = "0"
                Case "maliyetfiyatno"
                    SysDefault = "0"
                Case "dycls"
                    SysDefault = "01/01/2000"
                Case "barcodestokgirisi"    '
                    SysDefault = "0"
                Case "mgzstarttime"    '
                    SysDefault = "09:00"
                Case "mgzendtime"    '
                    SysDefault = "18:30"
                Case "stoktestfromisemri"
                    SysDefault = "1"
                Case "stokkodkontrol"
                    SysDefault = "0"
                Case "stokfirmakontrol"
                    SysDefault = "0"
                Case "satisfatura"
                    SysDefault = "0"
                Case "satisfaturaaltlimit"
                    SysDefault = "0"
                Case "logoaktarimdegismesin"
                    SysDefault = "0"
                Case "stkirsackayri"
                    SysDefault = "0"
                Case "stkirsayrac"
                    SysDefault = "0"
                Case "cevrimfonksiyonu"
                    SysDefault = "0"
                Case "barkodaktarimseriolsun"
                    SysDefault = "1"
                Case "logomagazasatisaktar"
                    SysDefault = "1"
                Case "gunkapatkontrol"
                    SysDefault = "0"
                Case "urtgunkapatkontrol"
                    SysDefault = "0"
                Case "utkisemrifoyikilemsin"
                    SysDefault = "0"
                Case "utkstokharekettipi"
                    SysDefault = "05 Diger Giris"
                Case "svkstokharekettipi"
                    SysDefault = "05 Diger Cikis"
                Case "stokfisallowtakipharici"
                    SysDefault = "0"
                Case "siparisaramaekranimodelli"
                    SysDefault = "1"
                Case "sqlbackupdirectory"
                    SysDefault = "C:\Mssql\Backup"
                Case "pricedecimal"
                    SysDefault = "2"
                Case "toplurezervasyongorunsun"
                    SysDefault = "0"
                Case "stokharodemegorunsun"
                    SysDefault = "0"
                Case "uretimharodemegorunsun"
                    SysDefault = "0"
                Case "sistemtarihidegissin"
                    SysDefault = "1"
                Case "kapalimtkfoyugorunmesin"
                    SysDefault = "0"
                Case "kapaliutkfoyugorunmesin"
                    SysDefault = "0"
                Case "repfilterbegtar"
                    SysDefault = "01.01.1950"
                Case "repfilterlasttar"
                    SysDefault = "01.01.2099"
                Case "iscilikfiyatolsun"
                    SysDefault = "0"
                Case "mtkhrkozelkodolsun"
                    SysDefault = "0"
                Case "stkirsayracsay"
                    SysDefault = "80"
                Case "stokfisindirimgorunsun"
                    SysDefault = "0"
                Case "uretimhareketfirmakontrol"
                    SysDefault = "0"
                Case "toplutransferhedefdepo"
                    SysDefault = "0"
                Case "toplutransferhedefpartino"
                    SysDefault = "0"
                Case "toplutransferhedefmtkno"
                    SysDefault = "0"
                Case "controluretfistoplam"
                    SysDefault = "0"
                Case "sqltimeout"
                    SysDefault = "600"
                Case "barcodecounterwrtfirma"
                    SysDefault = "1"
                Case "ncrstokdosyasi"
                    SysDefault = "c:\DATABASE\Condtimp.dat"
                Case Else
                    Select Case LCase(cPType)
                        Case "", "string"
                            SysDefault = ""
                        Case "long", "integer"
                            SysDefault = "0"
                        Case "float", "double"
                            SysDefault = "0"
                        Case "boolean"
                            SysDefault = "0"
                        Case "date"
                            SysDefault = "01.01.1950"
                    End Select
            End Select
        Catch
            ErrDisp("Error SysDefault " + Err.Description.Trim)
        End Try
    End Function

    Public Function UpdateStokRB(ByVal ConnYage As SqlConnection, Optional ByVal cHareketTipi As String = "", Optional ByVal cTrnType As String = "", Optional ByVal cTableName As String = "", _
                            Optional ByVal cStokno As String = "", Optional ByVal nNetMiktar1 As Double = 0, Optional ByVal nNetMiktar2 As Double = 0, Optional ByVal nNetMiktar3 As Double = 0, _
                            Optional ByVal cFilter As String = "", Optional ByVal nAgirlik As Double = 0, _
                            Optional ByVal dSonGirisTarihi As Date = #1/1/1950#, Optional ByVal nFiyat As Double = 0, _
                            Optional ByVal cDoviz As String = "", Optional ByVal cSonGirisDept As String = "", _
                            Optional ByVal cSonGirisFirmasi As String = "") As SqlInt32

        Dim cSQL As String = ""
        Dim cPhase As String = "Start"

        UpdateStokRB = 0

        Try

            'If cHareketTipi.Trim = "" Or cTrnType.Trim = "" Or cTableName.Trim = "" Or cStokno.Trim = "" Then
            '    ErrDisp("Unbelievable error " + cHareketTipi.Trim + "/" + cTrnType.Trim + "/" + cTableName.Trim + "/" + cStokno.Trim)
            '    Exit Function
            'End If

            cHareketTipi = LCase(cHareketTipi).Trim
            cTrnType = LCase(cTrnType).Trim
            cTableName = LCase(cTableName).Trim
            cStokno = cStokno.Trim

            cPhase = "Start with parameters : " + cHareketTipi + " " + cTrnType + " " + cTableName + " " + cStokno + "*"

            If cHareketTipi = "giris" Then
                Select Case cTrnType
                    Case "validate"
                        Select Case cTableName
                            Case "stok"
                                cSQL = "update " + cTableName + _
                                        " set donemgiris1 = coalesce(donemgiris1,0) + " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemgiris2 = coalesce(donemgiris2,0) + " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemgiris3 = coalesce(donemgiris3,0) + " + SQLWriteDecimal(nNetMiktar3) + _
                                        " where stokno = '" + cStokno + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "stokrb"
                                cSQL = "update " + cTableName + _
                                        " set donemgiris1 = coalesce(donemgiris1,0) + " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemgiris2 = coalesce(donemgiris2,0) + " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemgiris3 = coalesce(donemgiris3,0) + " + SQLWriteDecimal(nNetMiktar3) + _
                                        " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                                If nFiyat > 0 And dSonGirisTarihi <> #1/1/1950# Then
                                    cSQL = " update " + cTableName + _
                                            " set SonGirisTarihi = '" + SQLWriteDate(dSonGirisTarihi) + "', " + _
                                                " songirisfiyati = " + SQLWriteDecimal(nFiyat) + ", " + _
                                                " songirisdovizi = '" + cDoviz.Trim + "', " + _
                                                " SongirisDovizFiyati = " + SQLWriteDecimal(nFiyat) + ", " + _
                                                " SonGirisDept = '" + cSonGirisDept.Trim + "', " + _
                                                " SonGirisFirmasi = '" + cSonGirisFirmasi.Trim + "' " + _
                                            " where SonGirisTarihi <= '" + SQLWriteDate(dSonGirisTarihi) + "' " + _
                                            " and " + cFilter

                                    ExecuteSQLCommandConnected(cSQL, ConnYage, True)

                                End If

                            Case "stoktoprb", "stokaksesuarrb"
                                cSQL = "update " + cTableName + " " + _
                                        " set donemgiris1 = coalesce(donemgiris1,0) + " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemgiris2 = coalesce(donemgiris2,0) + " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemgiris3 = coalesce(donemgiris3,0) + " + SQLWriteDecimal(nNetMiktar3) + ", " + _
                                            " Agirlik = coalesce(Agirlik,0) + " + SQLWriteDecimal(nAgirlik) + _
                                        " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                                If nFiyat > 0 And dSonGirisTarihi <> #1/1/1950# Then
                                    cSQL = " update " + cTableName + " " + _
                                            " set SonGirisTarihi = '" + SQLWriteDate(dSonGirisTarihi) + "', " + _
                                                " songirisfiyati = " + SQLWriteDecimal(nFiyat) + ", " + _
                                                " songirisdovizi = '" + cDoviz.Trim + "', " + _
                                                " SongirisDovizFiyati = " + SQLWriteDecimal(nFiyat) + ", " + _
                                                " SonGirisDept = '" + cSonGirisDept.Trim + "', " + _
                                                " SonGirisFirmasi = '" + cSonGirisFirmasi.Trim + "' " + _
                                            " where SonGirisTarihi <= '" + SQLWriteDate(dSonGirisTarihi) + "' " + _
                                            " and " + cFilter

                                    ExecuteSQLCommandConnected(cSQL, ConnYage, True)

                                End If
                        End Select

                    Case "revert"

                        Select Case cTableName
                            Case "stok"
                                cSQL = "update " + cTableName + _
                                       " set donemgiris1 = coalesce(donemgiris1,0) - " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                           " donemgiris2 = coalesce(donemgiris2,0) - " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                           " donemgiris3 = coalesce(donemgiris3,0) - " + SQLWriteDecimal(nNetMiktar3) + _
                                       " where stokno = '" + cStokno + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "stokrb"
                                cSQL = "update " + cTableName + _
                                       " set donemgiris1 = coalesce(donemgiris1,0) - " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                           " donemgiris2 = coalesce(donemgiris2,0) - " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                           " donemgiris3 = coalesce(donemgiris3,0) - " + SQLWriteDecimal(nNetMiktar3) + _
                                       " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "stoktoprb", "stokaksesuarrb"
                                cSQL = "update " + cTableName + " " + _
                                           " set donemgiris1 = coalesce(donemgiris1,0) - " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                                " donemgiris2 = coalesce(donemgiris2,0) - " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                                " donemgiris3 = coalesce(donemgiris3,0) - " + SQLWriteDecimal(nNetMiktar3) + ", " + _
                                                " Agirlik = coalesce(Agirlik,0) - " + SQLWriteDecimal(nAgirlik) + _
                                        " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                        End Select
                End Select

            Else

                ' Çıkış hareketi

                Select Case cTrnType
                    Case "validate"
                        Select Case cTableName
                            Case "stok"
                                cSQL = "update " + cTableName + _
                                        " set donemcikis1 = coalesce(donemcikis1,0) + " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemcikis2 = coalesce(donemcikis2,0) + " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemcikis3 = coalesce(donemcikis3,0) + " + SQLWriteDecimal(nNetMiktar3) + _
                                        " where stokno = '" + cStokno + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "stokrb"
                                cSQL = "update " + cTableName + _
                                        " set donemcikis1 = coalesce(donemcikis1,0) + " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemcikis2 = coalesce(donemcikis2,0) + " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemcikis3 = coalesce(donemcikis3,0) + " + SQLWriteDecimal(nNetMiktar3) + _
                                        " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "stoktoprb", "stokaksesuarrb"
                                cSQL = "update " + cTableName + _
                                        " set donemcikis1 = coalesce(donemcikis1,0) + " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemcikis2 = coalesce(donemcikis2,0) + " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemcikis3 = coalesce(donemcikis3,0) + " + SQLWriteDecimal(nNetMiktar3) + ", " + _
                                            " Agirlik = coalesce(Agirlik,0) - " + SQLWriteDecimal(nAgirlik) + _
                                        " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                        End Select

                    Case "revert"

                        Select Case cTableName
                            Case "stok"
                                cSQL = "update " + cTableName + _
                                        " set donemcikis1 = coalesce(donemcikis1,0) - " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemcikis2 = coalesce(donemcikis2,0) - " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemcikis3 = coalesce(donemcikis3,0) - " + SQLWriteDecimal(nNetMiktar3) + _
                                        " where stokno = '" + cStokno + "' "

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "stokrb"
                                cSQL = "update " + cTableName + _
                                        " set donemcikis1 = coalesce(donemcikis1,0) - " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemcikis2 = coalesce(donemcikis2,0) - " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemcikis3 = coalesce(donemcikis3,0) - " + SQLWriteDecimal(nNetMiktar3) + _
                                        " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                            Case "stoktoprb", "stokaksesuarrb"
                                cSQL = "update " + cTableName + _
                                        " set donemcikis1 = coalesce(donemcikis1,0) - " + SQLWriteDecimal(nNetMiktar1) + ", " + _
                                            " donemcikis2 = coalesce(donemcikis2,0) - " + SQLWriteDecimal(nNetMiktar2) + ", " + _
                                            " donemcikis3 = coalesce(donemcikis3,0) - " + SQLWriteDecimal(nNetMiktar3) + ", " + _
                                            " Agirlik = coalesce(Agirlik,0) - " + SQLWriteDecimal(nAgirlik) + _
                                        " where " + cFilter

                                ExecuteSQLCommandConnected(cSQL, ConnYage)

                        End Select
                End Select
            End If
            UpdateStokRB = 1
            cPhase = "End"
        Catch
            UpdateStokRB = 0
            ErrDisp("Error UpdateStokRB " + Err.Description.Trim + vbCrLf + _
                                            "Phase : " + cPhase + vbCrLf + _
                                            "SQL : " + cSQL)
        End Try
    End Function

    Public Function G_BarkodluKumas(ByVal cStokNo As String, ByVal oSysFlags As SysFlags, ByVal ConnYaGe As SqlConnection) As Boolean

        Dim cSQL As String

        G_BarkodluKumas = False

        Try

            If Not oSysFlags.G_WinFabric Then Exit Function

            cSQL = "select stokno " + _
                    " from stok " + _
                    " where stokno = '" + cStokNo + "' " + _
                    " and toptakibi = 'E' "

            G_BarkodluKumas = CheckExistsConnected(cSQL, ConnYaGe)
        Catch
            ErrDisp("Error G_BarkodluKumas " + Err.Description.Trim)
        End Try
    End Function

    Public Function G_BarkodluAksesuar(ByVal cStokNo As String, ByVal oSysFlags As SysFlags, ByVal ConnYaGe As SqlConnection) As Boolean

        Dim cSQL As String

        G_BarkodluAksesuar = False
        Try

            If Not oSysFlags.G_WinAccessory Then Exit Function

            cSQL = "select stokno " + _
                    " from stok " + _
                    " where stokno = '" + cStokNo + "' " + _
                    " and aksesuartakibi = 'E' "

            G_BarkodluAksesuar = CheckExistsConnected(cSQL, ConnYaGe)
        Catch
            ErrDisp("Error G_BarkodluAksesuar " + Err.Description.Trim)
        End Try
    End Function

    Public Function UpdateMTF(ByVal ConnYage As SqlConnection, ByVal cAction As String, ByVal cStokFisTipi As String, ByVal cDepartman As String, _
                          ByVal cStokHareketKodu As String, ByVal cIsemriNo As String, ByVal cMTF As String, ByVal cStokNo As String, ByVal cRenk As String, _
                          ByVal cBeden As String, ByVal nNetMiktar1 As Double) As SqlInt32

        Dim cFilter As String
        Dim cSQL As String

        UpdateMTF = 0

        Try

            If cMTF.Trim = "" Then
                UpdateMTF = 1
                Exit Function
            End If

            ' replace nulls to zeros
            ExecuteSQLCommandConnected("update mtkfislines set isemriverilen = 0     where malzemetakipno ='" + cMTF + "' and isemriverilen is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set isemriicingelen = 0   where malzemetakipno ='" + cMTF + "' and isemriicingelen is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set isemriharicigelen = 0 where malzemetakipno ='" + cMTF + "' and isemriharicigelen is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set isemriicingelen = 0   where malzemetakipno ='" + cMTF + "' and isemriicingelen is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set isemriicingiden = 0   where malzemetakipno ='" + cMTF + "' and isemriicingiden is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set isemriharicigelen = 0 where malzemetakipno ='" + cMTF + "' and isemriharicigelen is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set isemriharicigiden = 0 where malzemetakipno ='" + cMTF + "' and isemriharicigiden is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set uretimicincikis = 0   where malzemetakipno ='" + cMTF + "' and uretimicincikis is null ", ConnYage)
            ExecuteSQLCommandConnected("update mtkfislines set uretimdeniade = 0     where malzemetakipno ='" + cMTF + "' and uretimdeniade is null ", ConnYage)

            cFilter = " where stokno = '" + cStokNo + "' " + _
                    " and malzemetakipno = '" + cMTF + "' " + _
                    " and renk = '" + cRenk + "' " + _
                    " and beden = '" + cBeden + "' "

            If LCase(cStokFisTipi) = "giris" Then

                ' ******** giris hareketleri

                If cIsemriNo.Trim = "" Then
                    If cStokHareketKodu = "02 Tedarikten Giris" Or _
                        cStokHareketKodu = "04 Mlz Uretimden Giris" Or _
                        cStokHareketKodu = "05 Diger Giris" Or _
                        cStokHareketKodu = "06 Tamirden Giris" Or _
                        cStokHareketKodu = "55 Kontrol Oncesi Giris" Or _
                        cStokHareketKodu = "77 Top Bolme Giris" Or _
                        cStokHareketKodu = "77 Aksesuar Bolme Giris" Or _
                        cStokHareketKodu = "08 SAYIM GIRIS" Or _
                        cStokHareketKodu = "90 Trans/Rezv Giris" Or _
                        cStokHareketKodu = "transfer" Then

                        cSQL = "update mtkfislines " + _
                                " set isemriharicigelen  = coalesce(isemriharicigelen,0) " + _
                                IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                                cFilter

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                Else
                    If cStokHareketKodu = "02 Tedarikten Giris" Or _
                        cStokHareketKodu = "04 Mlz Uretimden Giris" Or _
                        cStokHareketKodu = "05 Diger Giris" Or _
                        cStokHareketKodu = "06 Tamirden Giris" Then

                        cSQL = "update mtkfislines " + _
                                " set isemriicingelen = coalesce(isemriicingelen,0) " + _
                                IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                                cFilter

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If
                End If
                If cDepartman <> "" And cStokHareketKodu = "01 Uretimden iade" Then
                    cSQL = "update mtkfislines " + _
                            " set uretimdeniade = coalesce(uretimdeniade,0)  " + _
                            IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                            cFilter + _
                            " and departman = '" + cDepartman + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            Else

                ' ******** cikis hareketleri

                If cIsemriNo.Trim = "" Then
                    If cStokHareketKodu = "02 Tedarikten iade" Or _
                        cStokHareketKodu = "04 Mlz Uretime iade" Or _
                        cStokHareketKodu = "05 Diger Cikis" Or _
                        cStokHareketKodu = "06 Tamire Cikis" Or _
                        cStokHareketKodu = "55 Kontrol Oncesi Cikis" Or _
                        cStokHareketKodu = "77 Top Bolme Cikis" Or _
                        cStokHareketKodu = "77 Aksesuar Bolme Cikis" Or _
                        cStokHareketKodu = "08 SAYIM CIKIS" Or _
                        cStokHareketKodu = "90 Trans/Rezv Cikis" Or _
                        cStokHareketKodu = "transfer" Then

                        cSQL = "update mtkfislines " + _
                                " set isemriharicigelen  = coalesce(isemriharicigelen,0)  " + _
                                IIf(cAction = "validate", " - ", " + ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                                cFilter

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                    cSQL = "update mtkfislines " + _
                            " set isemriharicigiden = coalesce(isemriharicigiden,0) " + _
                            IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                            cFilter

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                Else
                    If cStokHareketKodu = "02 Tedarikten iade" Or _
                        cStokHareketKodu = "04 Mlz Uretime iade" Or _
                        cStokHareketKodu = "05 Diger Cikis" Or _
                        cStokHareketKodu = "06 Tamire Cikis" Then

                        cSQL = "update mtkfislines " + _
                                " set isemriicingelen = coalesce(isemriicingelen,0) " + _
                                IIf(cAction = "validate", " - ", " + ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                                cFilter

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                    cSQL = "update mtkfislines " + _
                           " set isemriicingiden = coalesce(isemriicingiden,0)  " + _
                           IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                           cFilter

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
                If cDepartman <> "" And cStokHareketKodu = "01 Uretime Cikis" Then
                    cSQL = "update mtkfislines " + _
                            " set uretimicincikis = coalesce(uretimicincikis,0)  " + _
                            IIf(cAction = "validate", " + ", " - ").ToString + SQLWriteDecimal(nNetMiktar1) + _
                            cFilter + _
                            " and departman = '" + cDepartman + "' "

                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                End If
            End If
            UpdateMTF = 1
        Catch
            UpdateMTF = 0
            ErrDisp("Error UpdateMTF " + Err.Description.Trim)
        End Try
    End Function

    Public Function UpdateIsemri(ByVal ConnYage As SqlConnection, ByVal cAction As String, ByVal cStokFisTipi As String, ByVal cStokHareketKodu As String, _
                            ByVal cIsemriNo As String, ByVal cMTF As String, ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String, _
                            ByVal nNetMiktar1 As Double) As SqlInt32

        Dim oReader As SqlDataReader
        Dim cFilter As String
        Dim nKalan As Double
        Dim nKullanilan As Double
        Dim nSiraNo As Double
        Dim nCnt1 As Integer
        Dim cSQL As String

        Dim aIsemri() As isemri
        Dim nIhtiyac As Double
        Dim nGelen As Double
        Dim lFound As Boolean

        UpdateIsemri = 0

        Try
            If cIsemriNo = "" Then
                UpdateIsemri = 1
                Exit Function
            End If

            cFilter = " where isemrino = '" + cIsemriNo + "' " + _
                       " and malzemetakipno = '" + cMTF + "' " + _
                       " and stokno = '" + cStokNo + "' " + _
                       " and renk = '" + cRenk + "' " + _
                       " and beden = '" + cBeden + "' "

            If LCase(cAction) = "validate" Then
                ' act
                If LCase(cStokFisTipi) = "giris" Then
                    Select Case cStokHareketKodu
                        Case "02 Tedarikten Giris", "04 Mlz Uretimden Giris"

                            nKalan = nNetMiktar1
                            nKullanilan = 0
                            nSiraNo = 0
                            nCnt1 = 0

                            ReDim aIsemri(0)
                            lFound = False

                            cSQL = "select sirano, miktar1, UretimGelen, TedarikGelen " + _
                                   " from isemrilines " + _
                                   cFilter + _
                                   " order by termintarihi "

                            oReader = GetSQLReader(cSQL, ConnYage)

                            Do While oReader.Read

                                nSiraNo = SQLReadDouble(oReader, "sirano")
                                nIhtiyac = SQLReadDouble(oReader, "miktar1") - SQLReadDouble(oReader, "UretimGelen") - SQLReadDouble(oReader, "TedarikGelen")

                                If nKalan > 0 Then
                                    If nIhtiyac > 0 Then
                                        If nIhtiyac > nKalan Then
                                            nKullanilan = nKalan
                                        Else
                                            nKullanilan = nIhtiyac
                                        End If
                                        nKalan = nKalan - nKullanilan

                                        lFound = True
                                        ReDim Preserve aIsemri(nCnt1)
                                        aIsemri(nCnt1).nSiraNo = nSiraNo
                                        aIsemri(nCnt1).nMiktar = nKullanilan
                                        nCnt1 = nCnt1 + 1
                                    End If
                                End If
                            Loop
                            oReader.Close()
                            oReader = Nothing

                            If lFound Then
                                For nCnt1 = 0 To UBound(aIsemri)
                                    cSQL = "update isemrilines "

                                    Select Case cStokHareketKodu
                                        Case "02 Tedarikten Giris" : cSQL = cSQL + " set TedarikGelen = coalesce(TedarikGelen,0) + " + SQLWriteDecimal(aIsemri(nCnt1).nMiktar)
                                        Case "04 Mlz Uretimden Giris" : cSQL = cSQL + " set UretimGelen  = coalesce(UretimGelen,0) + " + SQLWriteDecimal(aIsemri(nCnt1).nMiktar)
                                    End Select

                                    cSQL = cSQL + cFilter + _
                                            " and sirano = " + SQLWriteDecimal(aIsemri(nCnt1).nSiraNo)

                                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                                Next
                            End If

                            If nKalan > 0 Then
                                cSQL = "update isemrilines "

                                Select Case cStokHareketKodu
                                    Case "02 Tedarikten Giris" : cSQL = cSQL + " set TedarikGelen = coalesce(TedarikGelen,0) + " + SQLWriteDecimal(nKalan)
                                    Case "04 Mlz Uretimden Giris" : cSQL = cSQL + " set UretimGelen  = coalesce(UretimGelen,0) + " + SQLWriteDecimal(nKalan)
                                End Select

                                cSQL = cSQL + cFilter + _
                                        " and sirano = " + SQLWriteDecimal(nSiraNo)

                                ExecuteSQLCommandConnected(cSQL, ConnYage)
                            End If

                        Case "06 Tamirden Giris"
                            cSQL = "update isemrilines " + _
                                    " set tamirgelen = coalesce(tamirgelen,0) + " + SQLWriteDecimal(nNetMiktar1) + _
                                    cFilter

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End Select
                Else
                    Select Case cStokHareketKodu
                        Case "02 Tedarikten iade", "04 Mlz Uretime iade"

                            nKalan = nNetMiktar1
                            nKullanilan = 0
                            nSiraNo = 0
                            nCnt1 = 0

                            ReDim aIsemri(0)
                            lFound = False

                            cSQL = "select sirano, miktar1, UretimGelen, TedarikGelen " + _
                                   " from isemrilines " + _
                                    cFilter + _
                                   " order by termintarihi desc "

                            oReader = GetSQLReader(cSQL, ConnYage)

                            Do While oReader.Read

                                nSiraNo = SQLReadDouble(oReader, "sirano")
                                Select Case cStokHareketKodu
                                    Case "02 Tedarikten iade" : nGelen = SQLReadDouble(oReader, "TedarikGelen")
                                    Case "04 Mlz Uretime iade" : nGelen = SQLReadDouble(oReader, "UretimGelen")
                                End Select
                                If nGelen > 0 And nKalan > 0 Then
                                    If nGelen > nKalan Then
                                        nKullanilan = nKalan
                                    Else
                                        nKullanilan = nGelen
                                    End If
                                    nKalan = nKalan - nKullanilan

                                    lFound = True
                                    ReDim Preserve aIsemri(nCnt1)
                                    aIsemri(nCnt1).nSiraNo = nSiraNo
                                    aIsemri(nCnt1).nMiktar = nKullanilan
                                    nCnt1 = nCnt1 + 1
                                End If

                            Loop
                            oReader.Close()
                            oReader = Nothing

                            If lFound Then
                                For nCnt1 = 0 To UBound(aIsemri)

                                    cSQL = "update isemrilines "

                                    Select Case cStokHareketKodu
                                        Case "02 Tedarikten iade" : cSQL = cSQL + " set TedarikGelen = coalesce(TedarikGelen,0) - " + SQLWriteDecimal(aIsemri(nCnt1).nMiktar)
                                        Case "04 Mlz Uretime iade" : cSQL = cSQL + " set UretimGelen  = coalesce(UretimGelen,0) - " + SQLWriteDecimal(aIsemri(nCnt1).nMiktar)
                                    End Select

                                    cSQL = cSQL + cFilter + _
                                            " and sirano = " + SQLWriteDecimal(aIsemri(nCnt1).nSiraNo)

                                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                                Next
                            End If
                    End Select
                End If
            Else
                ' revert
                If LCase(cStokFisTipi) = "giris" Then
                    Select Case cStokHareketKodu
                        Case "02 Tedarikten Giris", "04 Mlz Uretimden Giris"
                            nKalan = nNetMiktar1
                            nSiraNo = 0
                            nCnt1 = 0

                            ReDim aIsemri(0)
                            lFound = False

                            cSQL = "select sirano, miktar1, UretimGelen, TedarikGelen " + _
                                   " from isemrilines " + _
                                   cFilter + _
                                   " order by termintarihi desc "

                            oReader = GetSQLReader(cSQL, ConnYage)

                            Do While oReader.Read

                                nSiraNo = SQLReadDouble(oReader, "sirano")
                                Select Case cStokHareketKodu
                                    Case "02 Tedarikten Giris" : nGelen = SQLReadDouble(oReader, "TedarikGelen")
                                    Case "04 Mlz Uretimden Giris" : nGelen = SQLReadDouble(oReader, "UretimGelen")
                                End Select
                                If nGelen > 0 And nKalan > 0 Then
                                    If nGelen > nKalan Then
                                        nKullanilan = nKalan
                                    Else
                                        nKullanilan = nGelen
                                    End If
                                    nKalan = nKalan - nKullanilan

                                    lFound = True
                                    ReDim Preserve aIsemri(nCnt1)
                                    aIsemri(nCnt1).nSiraNo = nSiraNo
                                    aIsemri(nCnt1).nMiktar = nKullanilan
                                    nCnt1 = nCnt1 + 1
                                End If
                            Loop
                            oReader.Close()
                            oReader = Nothing

                            If lFound Then
                                For nCnt1 = 0 To UBound(aIsemri)
                                    cSQL = "update isemrilines "

                                    Select Case cStokHareketKodu
                                        Case "02 Tedarikten Giris" : cSQL = cSQL + " set TedarikGelen = coalesce(TedarikGelen,0) - " + SQLWriteDecimal(aIsemri(nCnt1).nMiktar)
                                        Case "04 Mlz Uretimden Giris" : cSQL = cSQL + " set UretimGelen  = coalesce(UretimGelen,0) - " + SQLWriteDecimal(aIsemri(nCnt1).nMiktar)
                                    End Select

                                    cSQL = cSQL + cFilter + _
                                            " and sirano = " + SQLWriteDecimal(aIsemri(nCnt1).nSiraNo)

                                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                                Next
                            End If

                        Case "06 Tamirden Giris"
                            cSQL = "update isemrilines " + _
                                    " set tamirgelen = coalesce(tamirgelen,0) - " + SQLWriteDecimal(nNetMiktar1) + _
                                   cFilter

                            ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End Select
                Else
                    Select Case cStokHareketKodu
                        Case "02 Tedarikten iade", "04 Mlz Uretime iade"

                            nKalan = nNetMiktar1
                            nSiraNo = 0
                            nCnt1 = 0

                            ReDim aIsemri(0)
                            lFound = False

                            cSQL = "select sirano, miktar1, UretimGelen, TedarikGelen " + _
                                   " from isemrilines " + _
                                   cFilter + _
                                   " order by termintarihi "

                            oReader = GetSQLReader(cSQL, ConnYage)

                            Do While oReader.Read

                                nSiraNo = SQLReadDouble(oReader, "sirano")
                                nIhtiyac = SQLReadDouble(oReader, "miktar1") - SQLReadDouble(oReader, "UretimGelen") - SQLReadDouble(oReader, "TedarikGelen")

                                If nKalan > 0 Then
                                    If nIhtiyac > 0 Then
                                        If nIhtiyac > nKalan Then
                                            nKullanilan = nKalan
                                        Else
                                            nKullanilan = nIhtiyac
                                        End If
                                        nKalan = nKalan - nKullanilan

                                        lFound = True
                                        ReDim Preserve aIsemri(nCnt1)
                                        aIsemri(nCnt1).nSiraNo = nSiraNo
                                        aIsemri(nCnt1).nMiktar = nKullanilan
                                        nCnt1 = nCnt1 + 1
                                    End If
                                End If

                            Loop
                            oReader.Close()
                            oReader = Nothing

                            If lFound Then
                                For nCnt1 = 0 To UBound(aIsemri)
                                    cSQL = "update isemrilines "

                                    Select Case cStokHareketKodu
                                        Case "02 Tedarikten iade" : cSQL = cSQL + " set TedarikGelen = coalesce(TedarikGelen,0) + " + SQLWriteDecimal(nKullanilan)
                                        Case "04 Mlz Uretime iade" : cSQL = cSQL + " set UretimGelen  = coalesce(UretimGelen,0) + " + SQLWriteDecimal(nKullanilan)
                                    End Select

                                    cSQL = cSQL + cFilter + _
                                            " and sirano = " + SQLWriteDecimal(aIsemri(nCnt1).nSiraNo)

                                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                                Next

                                If nKalan > 0 Then
                                    cSQL = "update isemrilines "

                                    Select Case cStokHareketKodu
                                        Case "02 Tedarikten iade" : cSQL = cSQL + " set TedarikGelen = coalesce(TedarikGelen,0) + " + SQLWriteDecimal(nKalan)
                                        Case "04 Mlz Uretime iade" : cSQL = cSQL + " set UretimGelen  = coalesce(UretimGelen,0) + " + SQLWriteDecimal(nKalan)
                                    End Select

                                    cSQL = cSQL + cFilter + _
                                            " and sirano = " + SQLWriteDecimal(aIsemri(UBound(aIsemri)).nSiraNo)

                                    ExecuteSQLCommandConnected(cSQL, ConnYage)
                                End If
                            End If
                    End Select
                End If
            End If
            UpdateIsemri = 1
        Catch
            UpdateIsemri = 0
            ErrDisp("Error UpdateIsemri " + Err.Description.Trim)
        End Try
    End Function

    Public Function SingleStokToplam(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String) As SqlInt32

        Dim ConnYage As SqlConnection
        Dim cSQL As String

        SingleStokToplam = 0

        Try
            ConnYage = OpenConn()

            cSQL = "update stok set " + _
                    " donemgiris1 = 0 , donemgiris2 = 0 , donemgiris3 = 0 , " + _
                    " donemcikis1 = 0 , donemcikis2 = 0 , donemcikis3 = 0 " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete StokRB " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete StokTopRB " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "delete StokAksesuarRB " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update isemrilines set " + _
                    " uretimgelen = 0 , tedarikgelen = 0 " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            cSQL = "update mtkfislines set " + _
                    " isemriverilen = 0 , isemriicingelen = 0 , isemriharicigelen = 0 , " + _
                    " isemriicingiden = 0 , isemriharicigiden = 0 , " + _
                    " uretimicincikis = 0 ,uretimdeniade = 0 " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)


            cSQL = "delete uretstokrb " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)


            cSQL = "delete StokRBStatic " + _
                    IIf(cStokNo = "", "", " Where stokno = '" + cStokNo + "' ").ToString + _
                    IIf(cRenk = "", "", " and renk = '" + cRenk + "' ").ToString + _
                    IIf(cBeden = "", "", " and beden = '" + cBeden + "' ").ToString

            ExecuteSQLCommandConnected(cSQL, ConnYage)

            CloseConn(ConnYage)

            If MultiStokFisValidate(cStokNo, cRenk, cBeden) = 0 Then
                SingleStokToplam = 0
                Exit Function
            End If

            If MultiTransferFisValidate(cStokNo, cRenk, cBeden) = 0 Then
                SingleStokToplam = 0
                Exit Function
            End If

            If actMTKIsemri(cStokNo, cRenk, cBeden) = 0 Then
                SingleStokToplam = 0
                Exit Function
            End If

            SingleStokToplam = 1
        Catch
            SingleStokToplam = 0
            ErrDisp("Error SingleStokToplam " + Err.Description.Trim)
        End Try
    End Function

    Private Function actMTKIsemri(ByVal cStokNo As String, ByVal cRenk As String, ByVal cBeden As String) As SqlInt32

        Dim cSQL As String
        Dim oReader As SqlDataReader
        Dim nCnt As Integer
        Dim aData() As MTKIsemri
        Dim ConnYage As SqlConnection

        actMTKIsemri = 0

        Try

            ConnYage = OpenConn()

            nCnt = 0
            ReDim aData(0)

            cSQL = "Select distinct a.isemrino, a.departman,  " + _
                    " b.Stokno, b.malzemetakipno, b.renk, b.beden, b.miktar1 " + _
                    " From isemri a , isemrilines b " + _
                    " where a.isemrino = b.isemrino" + _
            IIf(cStokNo = "", "", " and b.stokno = '" + cStokNo + "' ").ToString + _
            IIf(cRenk = "", "", " and b.renk = '" + cRenk + "' ").ToString + _
            IIf(cBeden = "", "", " and b.beden = '" + cBeden + "' ").ToString

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read
                ReDim Preserve aData(nCnt)
                aData(nCnt).cMTF = SQLReadString(oReader, "malzemetakipno")
                aData(nCnt).cDepartman = SQLReadString(oReader, "departman")
                aData(nCnt).cStokNo = SQLReadString(oReader, "Stokno")
                aData(nCnt).cRenk = SQLReadString(oReader, "renk")
                aData(nCnt).cBeden = SQLReadString(oReader, "beden")
                aData(nCnt).nMiktar = SQLReadDouble(oReader, "miktar1")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            oReader = Nothing

            For nCnt = 0 To UBound(aData)

                CheckRenkBedenValidate(aData(nCnt).cStokNo, aData(nCnt).cRenk, aData(nCnt).cBeden, ConnYage)

                cSQL = "update mtkfislines " + _
                        " set isemriverilen = coalesce(isemriverilen,0) + " + SQLWriteDecimal(aData(nCnt).nMiktar) + _
                        " where stokno = '" + aData(nCnt).cStokNo + "' " + _
                        " and malzemetakipno = '" + aData(nCnt).cMTF + "' " + _
                        " and temindept = '" + aData(nCnt).cDepartman + "' " + _
                        IIf(UCase(aData(nCnt).cRenk) = "HEPSI", "", " and renk = '" + aData(nCnt).cRenk + "' ").ToString + _
                        IIf(UCase(aData(nCnt).cBeden) = "HEPSI", "", " and beden = '" + aData(nCnt).cBeden + "' ").ToString

                ExecuteSQLCommandConnected(cSQL, ConnYage)
            Next
            CloseConn(ConnYage)
            actMTKIsemri = 1
        Catch
            actMTKIsemri = 0
            ErrDisp("Error actMTKIsemri " + Err.Description.Trim)
        End Try
    End Function

    Public Function GetStokFisNo() As String

        Dim ConnYage As SqlConnection
        Dim cSQL As String
        Dim nFisNo As Double

        GetStokFisNo = "0000000000"

        Try
            ConnYage = OpenConn()

            cSQL = "select stokfisno from sysinfo "
            nFisNo = SQLGetDoubleConnected(cSQL, ConnYage)
            nFisNo = nFisNo + 1

            cSQL = "update sysinfo set stokfisno = " + SQLWriteDecimal(nFisNo)
            ExecuteSQLCommandConnected(cSQL, ConnYage)

            GetStokFisNo = Microsoft.VisualBasic.Format(nFisNo, "0000000000")

            CloseConn(ConnYage)

        Catch Err As Exception
            ErrDisp("GetStokFisNo : " + Err.Message)
        End Try

    End Function

    Public Function GetStokFiyat(ByRef nFiyat As Double, ByRef cDoviz As String, ByVal StokNo As String, ByVal Renk As String, _
                                 ByVal beden As String, ByVal MtkNo As String, Optional ByVal isEmriNo As String = "", _
                                 Optional ByVal Firma As String = "", Optional ByVal HareketKodu As String = "") As Boolean

        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nStokFiyatNo As Integer = 0
        Dim cParaTakipEsasi As String = "1"
        Dim cFilter As String = ""
        Dim cSQL As String
        Dim oSysFlags As SysFlags = Nothing

        GetStokFiyat = True

        Try
            If nFiyat <> 0 Then Exit Function

            ConnYage = OpenConn()

            ReadSysFlags(oSysFlags, ConnYage)

            GetStokFiyat = False

            cSQL = "select paratakipesasi from stok where stokno = '" + StokNo + "' "
            cParaTakipEsasi = SQLGetStringConnected(cSQL, ConnYage)

            ' Eğer isemrino varsa isemrinki fiyatı getir

            If isEmriNo <> "" Then
                cSQL = "select fiyat, doviz " + _
                        " from isemrilines " + _
                        " where stokno = '" + StokNo + "' " + _
                        " and isemrino = '" + isEmriNo + "' " + _
                         IIf(Trim(Renk) = "", "", " and renk = '" + Renk + "' ").ToString + _
                         IIf(Trim(beden) = "", "", "  and beden = '" + beden + "' ").ToString + _
                         IIf(Trim(MtkNo) = "", "", "  and malzemetakipno = '" + MtkNo + "' ").ToString

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "fiyat")
                    cDoviz = SQLReadString(oReader, "doviz")
                End If
                oReader.Close()
                oReader = Nothing

                GetStokFiyat = True
                Exit Function
            End If

            ' eğer işemri yoksa hareket koduna göre fiyat istesinden ilgili stok koduna ait fiyatı getir

            If HareketKodu <> "" Then
                cSQL = "select StokFiyatNo from stokhareketkodu where kod = '" + HareketKodu + "' "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nStokFiyatNo = SQLReadInteger(oReader, "StokFiyatNo")
                End If
                oReader.Close()
                oReader = Nothing

                If nStokFiyatNo <> 0 Then

                    cSQL = "select top 1 fiyat1,doviz1,fiyat2,doviz2,fiyat3,doviz3,fiyat4,doviz4 " + _
                            " from stokfiyat " + _
                            " where stokno = '" + StokNo + "' " + _
                            IIf(cParaTakipEsasi = "2" Or cParaTakipEsasi = "4", " and renk = '" + Renk + "' ", "").ToString + _
                            IIf(cParaTakipEsasi = "3" Or cParaTakipEsasi = "4", " and beden = '" + beden + "' ", "").ToString + _
                            IIf(oSysFlags.G_OnaysizFiyatFilter = "", " ", " and " + oSysFlags.G_OnaysizFiyatFilter).ToString + _
                            " order by tarih desc "

                    oReader = GetSQLReader(cSQL, ConnYage)

                    If oReader.Read Then
                        Select Case nStokFiyatNo
                            Case 1 'Alış Fiyatı
                                nFiyat = SQLReadDouble(oReader, "fiyat1")
                                cDoviz = SQLReadString(oReader, "doviz1")
                                GetStokFiyat = True
                            Case 2 'Ürt.Çıkış Fİyatı
                                nFiyat = SQLReadDouble(oReader, "fiyat2")
                                cDoviz = SQLReadString(oReader, "doviz2")
                                GetStokFiyat = True
                            Case 3 'Satış Fiyatı
                                nFiyat = SQLReadDouble(oReader, "fiyat3")
                                cDoviz = SQLReadString(oReader, "doviz3")
                                GetStokFiyat = True
                            Case 4 'Diğer Fiyat
                                nFiyat = SQLReadDouble(oReader, "fiyat4")
                                cDoviz = SQLReadString(oReader, "doviz4")
                                GetStokFiyat = True
                        End Select
                    End If
                    oReader.Close()
                    oReader = Nothing
                End If
            End If

            'eğer işemri ve hareket kodu yoksa fiyat istesinden firmanın ilgili sok koduna ait fiyatı getir

            If HareketKodu = "" And isEmriNo = "" Then

                cSQL = "select top 1 fiyat1, doviz1 " + _
                        " from stokfiyat " + _
                        " where stokno = '" + StokNo + "' " + _
                        " and Firma = '" + Firma + "' " + _
                        IIf(cParaTakipEsasi = "2" Or cParaTakipEsasi = "4", " and renk = '" + Renk + "' ", "").ToString + _
                        IIf(cParaTakipEsasi = "3" Or cParaTakipEsasi = "4", " and beden = '" + beden + "' ", "").ToString + _
                        IIf(oSysFlags.G_OnaysizFiyatFilter = "", " ", " and " + oSysFlags.G_OnaysizFiyatFilter).ToString + _
                        " order by tarih desc "

                oReader = GetSQLReader(cSQL, ConnYage)

                If oReader.Read Then
                    nFiyat = SQLReadDouble(oReader, "fiyat1")
                    cDoviz = SQLReadString(oReader, "doviz1")
                End If
                oReader.Close()
                oReader = Nothing

                GetStokFiyat = True
            End If

            CloseConn(ConnYage)

        Catch Err As Exception
            GetStokFiyat = False
            ErrDisp("GetStokFiyat : " + Err.Message)
        End Try
    End Function

    Public Function CheckInsertStokRB(ByVal oStokRB As stokrb, ByVal oSysFlags As SysFlags, ByVal ConnYage As SqlConnection) As String

        Dim cFilter As String
        Dim cSQL As String
        Dim cTopSiraNo As String

        CheckInsertStokRB = ""

        Try
            cFilter = ""

            Select Case LCase(oStokRB.cTableName)

                Case "stokrb"

                    cFilter = " stokno = '" + Trim(oStokRB.cStokno) + "' " + _
                            " and malzemetakipkodu = '" + Trim(oStokRB.cMtk) + "' " + _
                            " and renk = '" + Trim(oStokRB.cRenk) + "' " + _
                            " and beden = '" + Trim(oStokRB.cBeden) + "' " + _
                            " and partino = '" + Trim(oStokRB.cPartiNo) + "' " + _
                            " and depo = '" + Trim(oStokRB.cDepo) + "' "

                    cSQL = "select stokno " + _
                            " from  " + oStokRB.cTableName + _
                            " where " + cFilter

                    If Not CheckExistsConnected(cSQL, ConnYage) Then

                        cSQL = "insert into " + oStokRB.cTableName + _
                                " (Stokno,malzemetakipkodu,renk,beden,partino,depo, " + _
                                " donemgiris1,donemcikis1,donemgiris2,donemcikis2,donemgiris3,donemcikis3, " + _
                                " devirgiris1,devircikis1,devirgiris2,devircikis2,devirgiris3,devircikis3, " + _
                                " alismiktari1,alistutari1 ) "

                        cSQL = cSQL + " values ('" + Trim(oStokRB.cStokno) + "', " + _
                                    " '" + Trim(oStokRB.cMtk) + "', " + _
                                    " '" + Trim(oStokRB.cRenk) + "', " + _
                                    " '" + Trim(oStokRB.cBeden) + "', " + _
                                    " '" + Trim(oStokRB.cPartiNo) + "', " + _
                                    " '" + Trim(oStokRB.cDepo) + "', "

                        cSQL = cSQL + " 0,0,0,0,0,0,0,0,0,0,0,0,0,0) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                Case "stoktoprb", "stokaksesuarrb"

                    cFilter = " stokno = '" + Trim(oStokRB.cStokno) + "' " + _
                            " and malzemetakipkodu = '" + Trim(oStokRB.cMtk) + "' " + _
                            " and renk = '" + Trim(oStokRB.cRenk) + "' " + _
                            " and beden = '" + Trim(oStokRB.cBeden) + "' " + _
                            " and partino = '" + Trim(oStokRB.cPartiNo) + "' " + _
                            " and depo = '" + Trim(oStokRB.cDepo) + "' " + _
                            " and topno = '" + Trim(oStokRB.cTopNo) + "' "

                    cSQL = "select stokno " + _
                            " from  " + oStokRB.cTableName + _
                            " where " + cFilter

                    If Not CheckExistsConnected(cSQL, ConnYage) Then

                        cTopSiraNo = "0"

                        Select Case oStokRB.cTableName
                            Case "stoktoprb"
                                cTopSiraNo = GeTopSirano(Trim(oStokRB.cTopNo), Trim(oStokRB.cStokno), Trim(oStokRB.cMtk), Trim(oStokRB.cRenk), _
                                                         Trim(oStokRB.cBeden), Trim(oStokRB.cPartiNo), oSysFlags, ConnYage).ToString
                            Case "stokaksesuarrb"
                                cTopSiraNo = GetAksesuarSirano(Trim(oStokRB.cTopNo), Trim(oStokRB.cStokno), Trim(oStokRB.cMtk), Trim(oStokRB.cRenk), _
                                                               Trim(oStokRB.cBeden), Trim(oStokRB.cPartiNo), oSysFlags, ConnYage).ToString
                        End Select

                        cSQL = "insert into " + oStokRB.cTableName + _
                                " (Stokno,malzemetakipkodu,renk,beden,partino,depo,topno,topsirano, " + _
                                " donemgiris1,donemcikis1,donemgiris2,donemcikis2,donemgiris3,donemcikis3, " + _
                                " devirgiris1,devircikis1,devirgiris2,devircikis2,devirgiris3,devircikis3, " + _
                                " alismiktari1,alistutari1,agirlik) "

                        cSQL = cSQL + " values ('" + Trim(oStokRB.cStokno) + "', " + _
                                    " '" + Trim(oStokRB.cMtk) + "', " + _
                                    " '" + Trim(oStokRB.cRenk) + "', " + _
                                    " '" + Trim(oStokRB.cBeden) + "', " + _
                                    " '" + Trim(oStokRB.cPartiNo) + "', " + _
                                    " '" + Trim(oStokRB.cDepo) + "', " + _
                                    " '" + Trim(oStokRB.cTopNo) + "', " + _
                                    Trim(cTopSiraNo) + ", "

                        cSQL = cSQL + " 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

                Case "stokrbstatic"

                    cFilter = " stokno = '" + Trim(oStokRB.cStokno) + "' " + _
                            " and malzemetakipkodu = '" + Trim(oStokRB.cMtk) + "' " + _
                            " and renk = '" + Trim(oStokRB.cRenk) + "' " + _
                            " and beden = '" + Trim(oStokRB.cBeden) + "' " + _
                            " and partino = '" + Trim(oStokRB.cPartiNo) + "' " + _
                            " and depo = '" + Trim(oStokRB.cDepo) + "' " + _
                            " and ay = '" + Trim(oStokRB.cAy) + "' " + _
                            " and yil = '" + Trim(oStokRB.cYil) + "' "

                    cSQL = "select stokno " + _
                            " from  " + oStokRB.cTableName + _
                            " where " + cFilter

                    If Not CheckExistsConnected(cSQL, ConnYage) Then

                        cSQL = "insert into " + oStokRB.cTableName + _
                                " (stokno,malzemetakipkodu,renk,beden,partino,depo, " + _
                                " ay,yil,alismiktari1,alistutari1 ) " + _
                                " values ('" + Trim(oStokRB.cStokno) + " ' ," + _
                                " '" + Trim(oStokRB.cMtk) + " ' ," + _
                                " '" + Trim(oStokRB.cRenk) + " ' ," + _
                                " '" + Trim(oStokRB.cBeden) + " ' ," + _
                                " '" + Trim(oStokRB.cPartiNo) + " ' ," + _
                                " '" + Trim(oStokRB.cDepo) + " ' ," + _
                                " '" + Trim(oStokRB.cAy) + " ' ," + _
                                " '" + Trim(oStokRB.cYil) + " ',0,0 ) "

                        ExecuteSQLCommandConnected(cSQL, ConnYage)
                    End If

            End Select

            CheckInsertStokRB = cFilter
        Catch
            ErrDisp("Error CheckInsertStokRB " + Err.Description.Trim)
        End Try
    End Function

    Public Function GetTasarimKarisim1(ByVal cTasarimNo As String) As String

        Dim cSQL As String = ""
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim aIplik(0) As oIplik
        Dim nCnt As Integer = -1
        Dim nMax As Double = 0
        Dim nMaxIndex As Integer = 0
        Dim nToplam As Double = 0

        GetTasarimKarisim1 = ""

        Try
            If cTasarimNo.Trim = "" Then Exit Function

            ConnYage = OpenConn()

            cSQL = "select w.iplikhammadde, yuzde = round(w.miktar / w.toplam * 100,6) " + _
                    " from (select  a.tasarimno, b.iplikhammadde, " + _
                            " miktar =  sum(a.miktar * b.karisimyuzdesi / 100), " + _
                            " toplam = (select sum(coalesce(miktar,0)) from tasarimip x where x.tasarimno = a.tasarimno and (karisimakatilmasin is null or karisimakatilmasin = '' or karisimakatilmasin = 'H')) " + _
                            " from tasarimip a, stokiplikhammade b " + _
                            " where a.stokno = b.stokno " + _
                            " and a.tasarimno = '" + cTasarimNo.Trim + "' " + _
                            " and b.iplikhammadde is not null " + _
                            " and b.iplikhammadde <> '' " + _
                            " and b.karisimyuzdesi is not null " + _
                            " and b.karisimyuzdesi <> 0 " + _
                            " and (a.karisimakatilmasin is null or a.karisimakatilmasin = '' or a.karisimakatilmasin = 'H') " + _
                            " group by a.tasarimno, b.iplikhammadde) w  " + _
                    " where w.toplam is not null " + _
                    " and w.toplam <> 0 " + _
                    " order by round(w.miktar / w.toplam * 100,0) desc "

            oReader = GetSQLReader(cSQL, ConnYage)

            Do While oReader.Read

                nCnt = nCnt + 1
                ReDim Preserve aIplik(nCnt)
                aIplik(nCnt).cHammadde = SQLReadString(oReader, "iplikhammadde")
                aIplik(nCnt).nYzd = SQLReadDouble(oReader, "yuzde")
                aIplik(nCnt).nYzdRnd = Math.Round(SQLReadDouble(oReader, "yuzde"), 0)

                nToplam = nToplam + Math.Round(SQLReadDouble(oReader, "yuzde"), 0)

                If Math.Round(SQLReadDouble(oReader, "yuzde"), 0) > nMax Then
                    nMax = Math.Round(SQLReadDouble(oReader, "yuzde"), 0)
                    nMaxIndex = nCnt
                End If

            Loop
            oReader.Close()
            oReader = Nothing

            If nToplam > 99 And nToplam < 100 Then
                aIplik(nMaxIndex).nYzdRnd = aIplik(nMaxIndex).nYzdRnd + (100 - nToplam)
            ElseIf nToplam > 100 And nToplam < 101 Then
                aIplik(nMaxIndex).nYzdRnd = aIplik(nMaxIndex).nYzdRnd - (nToplam - 100)
            End If

            For nCnt = 0 To UBound(aIplik)
                If GetTasarimKarisim1 = "" Then
                    GetTasarimKarisim1 = aIplik(nCnt).cHammadde + " %" + Microsoft.VisualBasic.Format(aIplik(nCnt).nYzdRnd, G_NumberFormat)
                Else
                    GetTasarimKarisim1 = GetTasarimKarisim1 + ";" + aIplik(nCnt).cHammadde + " %" + Microsoft.VisualBasic.Format(aIplik(nCnt).nYzdRnd, G_NumberFormat)
                End If
            Next

            CloseConn(ConnYage)

        Catch ex As Exception
            ErrDisp("Error GetTasarimKarisim1 " + Err.Description.Trim)
            GetTasarimKarisim1 = Err.Description.Trim
        End Try
    End Function

    Public Function Conv_Tr_Char(s As String) As String

        Dim cOut As String = ""

        Conv_Tr_Char = ""

        Try
            cOut = s
            cOut = Replace(cOut, "ğ", "g")
            cOut = Replace(cOut, "Ğ", "G")
            cOut = Replace(cOut, "ü", "u")
            cOut = Replace(cOut, "Ü", "U")
            cOut = Replace(cOut, "ı", "i")
            cOut = Replace(cOut, "İ", "I")
            cOut = Replace(cOut, "ş", "s")
            cOut = Replace(cOut, "Ş", "S")
            cOut = Replace(cOut, "ç", "c")
            cOut = Replace(cOut, "Ç", "C")
            cOut = Replace(cOut, "ö", "o")
            cOut = Replace(cOut, "Ö", "O")
            cOut = Replace(cOut, "%", "Y")

            Conv_Tr_Char = cOut

        Catch ex As Exception
            ErrDisp("Error Conv_Tr_Char " + Err.Description.Trim)
        End Try
    End Function
End Module

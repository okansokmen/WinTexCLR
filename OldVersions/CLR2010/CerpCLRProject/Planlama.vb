Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

' sipok
' plgonderitarihi   - planlanan gönderi tarihi
' oktar             - gerçekleşen gönderi tarihi
' pltarihi          - planlanan ok tarihi
' oktar2            - gerçekleşen ok tarihi
' ok                - planlaması kapandi
' pltarihkilitle    - planlanan tarihleri kilitle 
' teminsuresi       - plgonderitarihi ile pltarihi arasındaki işgünü adedi
' avanssuresi       -

' mtkfislines
' pltarihi          - planlanan malzeme işemri veriliş tarihi
' baslamatarihi     - planlanan başlama (işemri veriliş) tarihi
' bitistarihi       - planlanan bitiş (depoya giriş) tarihi 
' kapanistarihi     - gerçekleşen kapanma tarihi
' pltarihkilitle    - planlanan tarihleri kilitle 
' teminsuresi       - baslamatarihi ile bitistarihi arasındaki işgünü adedi
' avanssuresi       - malzeme depoya girdikten kaç işgünü sonra üretime çıkacak
' okbilgisi         - planlaması kapandi
' kapandi           - dinamik MTF satır kapanışı (sadece dinamik malzeme ihtiyaç hesabını bağlar)
' orjbaslamatarihi  - orjinal planlanan başlama (işemri veriliş) tarihi
' orjbitistarihi    - orjinal planlanan bitiş (depoya giriş) tarihi 

' uretpllines
' pltarihi          - planlanan üretim işemri veriliş tarihi
' baslamatarihi     - planlanan üretim başlama tarihi
' bitistarihi       - planlanan üretim bitiş tarihi 
' kapanistarihi     - gerçekleşen kapanma tarihi
' pltarihkilitle    - planlanan tarihleri kilitle 
' teminsuresi       - baslamatarihi ile bitistarihi arasındaki işgünü adedi
' avanssuresi       - iki üretim departmanı arasındaki geçişte bekleme süresi
' okbilgisi         - planlaması kapandi
' orjbaslamatarihi  - orjinal planlanan üretime başlama tarihi
' orjbitistarihi    - orjinal planlanan üretim bitiş tarihi 

' sevkplfislines
' ilkesevktar       - planlanan ilk sevkiyat tarihi
' sonsevktar        - planlanan son sevkiyat tarihi
' ok                - planlaması kapandi
' oktarihi          - kapanış tarihi
' pltarihkilitle    - planlanan tarihleri kilitle 

Module Planlama
    Private Structure oSS
        Dim nSiraNo As Double
        Dim nTeminSuresi As Double
        Dim dBitis As Date
    End Structure

    Private Structure oCP
        Dim cSiparisNo As String
        Dim cOkTipi As String
        Dim cRenk As String
        Dim cBeden As String
        Dim dPlTarihi As Date
        Dim dOkTar2 As Date
        Dim nAvansSuresi As Double
    End Structure

    Private Structure oMTF
        Dim cMTF As String
        Dim cDepartman As String
        Dim nSira As Double
    End Structure

    Private Structure oUTF
        Dim cUTF As String
        Dim cModelNo As String
        Dim cDepartman As String
        Dim dBaslamaTarihi As Date
        Dim dBitisTarihi As Date
        Dim nTeminSuresi As Double
        Dim cOkBilgisi As String
        Dim dKapanisTarihi As Date
        Dim cOkTipi As String
        Dim cPlTarihKilitle As String
    End Structure

    Private aTatil As Date()

    Public Sub ForwardAll(Optional cSiparisNo As String = "")

        JustForLog("ForwardAll start " + cSiparisNo)

        LoadTatiller()
        AllinitWithReality(cSiparisNo)
        OtomatikKapatmalar(cSiparisNo)
        PlCPdenUretim(cSiparisNo)
        PlMalzemedenUretim(cSiparisNo)

        JustForLog("ForwardAll finish " + cSiparisNo)
    End Sub

    Private Sub AllinitWithReality(Optional cSiparisno As String = "")
        JustForLog("AllinitWithReality start")
        CPinitWithReality(cSiparisno)
        MTFinitWithReality(cSiparisno)
        UTFinitWithReality(cSiparisno)
        JustForLog("AllinitWithReality finish")
    End Sub

    ' gerçekleşen tarihleri planlamaya yansıt

    Private Sub CPinitWithReality(Optional cSiparisno As String = "")

        Dim cSql As String = ""

        Try
            ' önce kayıp bitiş tarihleri bulunur
            cSql = "set dateformat dmy " + _
                    " update sipok " + _
                    " set oktar2 = (select top 1 oktarihi " + _
                                    " from sipokhistory " + _
                                    " where siparisno = sipok.siparisno " + _
                                    " and modelkodu = sipok.modelkodu " + _
                                    " and oktipi = sipok.oktipi " + _
                                    " and renk = sipok.renk " + _
                                    " and beden = sipok.beden " + _
                                    " and ok = 'E' " + _
                                    " and oktarihi is not null " + _
                                    " and oktarihi <> '01.01.1950' " + _
                                    " order by oktarihi) " + _
                    " where (oktar2 is null or oktar2 = '01.01.1950') " + _
                    " and ok = 'E' "
            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            ExecuteSQLCommand(cSql)

            ' önce gerçekleşen tarihler güncellenir
            ' gerçekleşen gönderi tarihi ilk giden numunenin gidiş tarihidir
            cSql = "set dateformat dmy " + _
                    " update sipok " + _
                    " set oktar = (select top 1 gonderitarihi " + _
                                " from sipokhistory " + _
                                " where siparisno = sipok.siparisno " + _
                                " and modelkodu = sipok.modelkodu " + _
                                " and oktipi = sipok.oktipi " + _
                                " and renk = sipok.renk " + _
                                " and beden = sipok.beden " + _
                                " and gonderitarihi is not null " + _
                                " and gonderitarihi <> '01.01.1950' " + _
                                " order by gonderitarihi) " + _
                    " where  (oktar is null or oktar = '01.01.1950') "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            ExecuteSQLCommand(cSql)

            ' gerçekleşen OK tarihi kontrolü
            cSql = "set dateformat dmy " + _
                    " update sipok " + _
                    " set oktar2 = (select top 1 oktarihi " + _
                                    " from sipokhistory " + _
                                    " where siparisno = sipok.siparisno " + _
                                    " and modelkodu = sipok.modelkodu " + _
                                    " and oktipi = sipok.oktipi " + _
                                    " and renk = sipok.renk " + _
                                    " and beden = sipok.beden " + _
                                    " and ok = 'E' " + _
                                    " and oktarihi is not null " + _
                                    " and oktarihi <> '01.01.1950' " + _
                                    " order by oktarihi) " + _
                    " where (oktar2 is null or oktar2 = '01.01.1950') "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            ExecuteSQLCommand(cSql)

            ' sonra planlanan tarihler güncellenir
            ' planlanan gönderi tarihi ilk giden numunenin gidiş tarihidir
            cSql = "set dateformat dmy " + _
                    " update sipok " + _
                    " set plgonderitarihi = (select top 1 gonderitarihi " + _
                                            " from sipokhistory " + _
                                            " where siparisno = sipok.siparisno " + _
                                            " and modelkodu = sipok.modelkodu " + _
                                            " and oktipi = sipok.oktipi " + _
                                            " and renk = sipok.renk " + _
                                            " and beden = sipok.beden " + _
                                            " and gonderitarihi is not null " + _
                                            " and gonderitarihi <> '01.01.1950' " + _
                                            " order by gonderitarihi) " + _
                    " where (plgonderitarihi is null or plgonderitarihi = '01.01.1950') " + _
                    " and (ok is null or ok = 'H' or ok = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            ExecuteSQLCommand(cSql)

            ' planlanan OK tarihi gerçekleşen OK tarihidir
            cSql = "set dateformat dmy " + _
                    " update sipok " + _
                    " set pltarihi = (select top 1 oktarihi " + _
                                    " from sipokhistory " + _
                                    " where siparisno = sipok.siparisno " + _
                                    " and modelkodu = sipok.modelkodu " + _
                                    " and oktipi = sipok.oktipi " + _
                                    " and renk = sipok.renk " + _
                                    " and beden = sipok.beden " + _
                                    " and ok = 'E' " + _
                                    " and oktarihi is not null " + _
                                    " and oktarihi <> '01.01.1950' " + _
                                    " order by oktarihi) " + _
                    " where (pltarihi is null or pltarihi = '01.01.1950') " + _
                    " and (ok is null or ok = 'H' or ok = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            ExecuteSQLCommand(cSql)

            ' OK tarihi varsa OK le
            cSql = "update sipok " + _
                    " set ok = 'E' " + _
                    " where oktar2 is not null " + _
                    " and oktar2 <> '01.01.1950' "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            ExecuteSQLCommand(cSql)

            ' OK lenmişse satır planlamaya kilitlenir
            cSql = "update sipok " + _
                    " set pltarihkilitle = 'E' " + _
                    " where ok = 'E' "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            ExecuteSQLCommand(cSql)

        Catch ex As Exception
            ErrDisp(ex.Message, "CPinitWithReality", cSql)
        End Try
    End Sub

    Private Sub MTFinitWithReality(Optional cSiparisno As String = "")

        Dim cSql As String = ""
        Dim cMTF As String = ""

        Try
            cMTF = GetOpenMTFFromSiparisNo(cSiparisno)
            If cMTF.Trim = "" Then Exit Sub

            ' önce kayıp bitiş tarihleri bulunur
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set bitistarihi = (select top 1 a.fistarihi " + _
                                       " from stokfis a, stokfislines b " + _
                                       " where a.stokfisno = b.stokfisno " + _
                                       " and b.malzemetakipkodu = mtkfislines.malzemetakipno " + _
                                       " and b.stokno = mtkfislines.stokno " + _
                                       " and b.renk = mtkfislines.renk " + _
                                       " and b.beden = mtkfislines.beden " + _
                                       " and b.stokhareketkodu in ('04 Mlz Uretimden Giris','06 Tamirden Giris','02 Tedarikten Giris','05 Diger Giris','90 Trans/Rezv Giris') " + _
                                       " and a.fistarihi is not null " + _
                                       " order by a.fistarihi desc ) " + _
                    " where malzemetakipno in (" + cMTF + ") " + _
                    " and (bitistarihi is null or bitistarihi = '01.01.1950') " + _
                    " and okbilgisi = 'E' " + _
                    " and (coalesce(ihtiyac,0) <= coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0)) "

            ExecuteSQLCommand(cSql)

            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set bitistarihi = (select top 1 tarih  " + _
                                      " from stoktransfer " + _
                                      " where hedefmalzemetakipno = mtkfislines.malzemetakipno " + _
                                      " and stokno = mtkfislines.stokno " + _
                                      " and renk = mtkfislines.renk " + _
                                      " and beden = mtkfislines.beden  " + _
                                      " and tarih is not null " + _
                                      " order by tarih desc ) " + _
                    " where malzemetakipno in (" + cMTF + ") " + _
                    " and (bitistarihi is null or bitistarihi = '01.01.1950') " + _
                    " and okbilgisi = 'E' " + _
                    " and (coalesce(ihtiyac,0) <= coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0)) "

            ExecuteSQLCommand(cSql)

            ' Malzemede başlangıç tarihi ilk ONAYLANMIŞ işemri veriliş tarihidir
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set baslamatarihi = (select top 1 a.tarih " + _
                                        " from isemri a, isemrilines b " + _
                                        " where a.isemrino = b.isemrino " + _
                                        " and b.malzemetakipno = mtkfislines.malzemetakipno " + _
                                        " and b.stokno = mtkfislines.stokno " + _
                                        " and b.renk = mtkfislines.renk " + _
                                        " and b.beden = mtkfislines.beden " + _
                                        " and a.departman = mtkfislines.departman " + _
                                        " and a.tarih is not null " + _
                                        " and a.onay = 'E' " + _
                                        " order by a.tarih ) " + _
                    " where malzemetakipno in (" + cMTF + ") " + _
                    " and (baslamatarihi is null or baslamatarihi = '01.01.1950') " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') "

            ExecuteSQLCommand(cSql)

            ' Malzeme bitiş tarihi ihtiyaç karşılandıysa en son giriş hareket tarihidir
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set bitistarihi = (select top 1 a.fistarihi " + _
                                        " from stokfis a, stokfislines b " + _
                                        " where a.stokfisno = b.stokfisno " + _
                                        " and b.malzemetakipkodu = mtkfislines.malzemetakipno " + _
                                        " and b.stokno = mtkfislines.stokno " + _
                                        " and b.renk = mtkfislines.renk " + _
                                        " and b.beden = mtkfislines.beden " + _
                                        " and b.stokhareketkodu in ('04 Mlz Uretimden Giris','06 Tamirden Giris','02 Tedarikten Giris','05 Diger Giris','90 Trans/Rezv Giris') " + _
                                        " and a.fistarihi is not null " + _
                                        " order by a.fistarihi desc ) " + _
                    " where malzemetakipno in (" + cMTF + ") " + _
                    " and (bitistarihi is null or bitistarihi = '01.01.1950') " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                    " and (coalesce(ihtiyac,0) <= coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0))  "

            ExecuteSQLCommand(cSql)

            ' Malzeme bitiş tarihi ihtiyaç karşılandıysa en son transfer hareket tarihidir
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set bitistarihi = (select top 1 tarih " + _
                                        " from stoktransfer " + _
                                        " where hedefmalzemetakipno = mtkfislines.malzemetakipno " + _
                                        " and stokno = mtkfislines.stokno " + _
                                        " and renk = mtkfislines.renk " + _
                                        " and beden = mtkfislines.beden " + _
                                        " and tarih is not null " + _
                                        " order by tarih desc ) " + _
                    " where malzemetakipno in (" + cMTF + ") " + _
                    " and (bitistarihi is null or bitistarihi = '01.01.1950') " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                    " and (coalesce(ihtiyac,0) <= coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0))  "

            ExecuteSQLCommand(cSql)

            ' ihtiyaç karşılanmışsa satır kapatılır
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set okbilgisi = 'E' " + _
                    " where malzemetakipno in (" + cMTF + ") " + _
                    " and baslamatarihi is not null " + _
                    " and bitistarihi is not null " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                    " and (coalesce(ihtiyac,0) <= coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0))  "

            ExecuteSQLCommand(cSql)

            ' ihtiyaç karşılanmışsa ve kapatılmışsa satır planlamaya kilitlenir
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set pltarihkilitle = 'E' " + _
                    " where malzemetakipno in (" + cMTF + ") " + _
                    " and baslamatarihi is not null " + _
                    " and bitistarihi is not null " + _
                    " and okbilgisi = 'E'  " + _
                    " and (coalesce(ihtiyac,0) <= coalesce(isemriicingelen,0) + coalesce(isemriharicigelen,0))  "

            ExecuteSQLCommand(cSql)

        Catch ex As Exception
            ErrDisp(ex.Message, "MTFinitWithReality", cSql)
        End Try
    End Sub

    Private Sub UTFinitWithReality(Optional cSiparisno As String = "")

        Dim cSql As String = ""
        Dim cUTF As String = ""

        Try
            cUTF = GetOpenUTFFromSiparisNo(cSiparisno)
            If cUTF.Trim = "" Then Exit Sub

            ' önce kayıp bitiş tarihleri bulunur
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set bitistarihi = (select top 1 a.fistarihi " + _
                                         " from uretharfis a, uretharfislines b " + _
                                         " where a.uretfisno = b.uretfisno " + _
                                         " and a.cikisdept = uretpllines.departman " + _
                                         " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                         " and b.modelno = uretpllines.modelno " + _
                                         " order by a.fistarihi desc )  " + _
                    " where uretimtakipno in (" + cUTF + ") " + _
                    " and (bitistarihi is null or bitistarihi = '01.01.1950') " + _
                    " and okbilgisi = 'E' " + _
                    " and coalesce(toplamadet,0) <= coalesce(giden,0) "

            ExecuteSQLCommand(cSql)

            ' UTFde ilk giriş tarihi, KESIM için ilk kumaş çıkış tarihidir
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set baslamatarihi = (select top 1 a.fistarihi " + _
                                        " from stokfis a, stokfislines b " + _
                                        " where a.stokfisno = b.stokfisno " + _
                                        " and a.departman = uretpllines.departman " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " and b.stokhareketkodu = '01 Uretime Cikis' " + _
                                        " and a.fistarihi is not null " + _
                                        " order by a.fistarihi ) " + _
                    " where uretimtakipno in (" + cUTF + ") " + _
                    " and (baslamatarihi is null or baslamatarihi = '01.01.1950') " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                    " and departman like '%KESIM%' "

            ExecuteSQLCommand(cSql)

            ' UTF de departman KESIM değilse ilk uretim giriş fişi tarihidir
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set baslamatarihi = (select top 1 a.fistarihi " + _
                                        " from uretharfis a, uretharfislines b " + _
                                        " where a.uretfisno = b.uretfisno " + _
                                        " and a.girisdept = uretpllines.departman  " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " order by a.fistarihi ) " + _
                    " where uretimtakipno in (" + cUTF + ") " + _
                    " and (baslamatarihi is null or baslamatarihi = '01.01.1950') " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                    " and departman not like '%KESIM%' "

            ExecuteSQLCommand(cSql)

            ' UTF bitiş tarihi ihtiyaç karşılandıysa son çıkış fiş tarihidir
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set bitistarihi = (select top 1 a.fistarihi " + _
                                        " from uretharfis a, uretharfislines b " + _
                                        " where a.uretfisno = b.uretfisno " + _
                                        " and a.cikisdept = uretpllines.departman  " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " order by a.fistarihi desc ) " + _
                    " where uretimtakipno in (" + cUTF + ") " + _
                    " and (bitistarihi is null or bitistarihi = '01.01.1950') " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                    " and coalesce(toplamadet,0) <= coalesce(giden,0) "

            ExecuteSQLCommand(cSql)

            ' UTF işemri tarihi
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set pltarihi = (select top 1 a.tarih " + _
                                        " from uretimisemri a, uretimisdetayi b " + _
                                        " where a.isemrino = b.isemrino " + _
                                        " and a.departman = uretpllines.departman " + _
                                        " and b.uretimtakipno = uretpllines.uretimtakipno " + _
                                        " and b.modelno = uretpllines.modelno " + _
                                        " order by a.tarih ) " + _
                    " where uretimtakipno in (" + cUTF + ") " + _
                    " and (pltarihi is null or pltarihi = '01.01.1950') " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') "

            ExecuteSQLCommand(cSql)

            ' ihtiyaç karşılanmışsa satır kapatılır 
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set okbilgisi = 'E' " + _
                    " where uretimtakipno in (" + cUTF + ") " + _
                    " and baslamatarihi is not null " + _
                    " and bitistarihi is not null " + _
                    " and (okbilgisi is null or okbilgisi = 'H' or okbilgisi = '') " + _
                    " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                    " and coalesce(toplamadet,0) <= coalesce(giden,0) "

            ExecuteSQLCommand(cSql)

            ' ihtiyaç karşılanmışsa ve kapatılmışsa planlamaya kilitlenir
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set pltarihkilitle = 'E' " + _
                    " where uretimtakipno in (" + cUTF + ") " + _
                    " and baslamatarihi is not null " + _
                    " and bitistarihi is not null " + _
                    " and okbilgisi = 'E' " + _
                    " and coalesce(toplamadet,0) <= coalesce(giden,0) "

            ExecuteSQLCommand(cSql)

        Catch ex As Exception
            ErrDisp(ex.Message, "UTFinitWithReality", cSql)
        End Try
    End Sub

    Private Function GetOpenMTFFromSiparisNo(Optional cSiparisno As String = "") As String

        Dim cSql As String = ""
        Dim ConnYage As SqlConnection

        GetOpenMTFFromSiparisNo = ""

        Try
            ConnYage = OpenConn()

            cSql = "select distinct a.malzemetakipno " + _
                    " from sipmodel a, mtkfis b " + _
                    " where a.malzemetakipno = b.malzemetakipno " + _
                    " and a.malzemetakipno is not null " + _
                    " and a.malzemetakipno <> '' " + _
                    " and (b.dosyakapandi is null or b.dosyakapandi = '' or b.dosyakapandi = 'H') "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and a.siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and a.siparisno = '" + cSiparisno.Trim + "' "
            End If

            GetOpenMTFFromSiparisNo = SQLBuildFilterString2(ConnYage, cSql)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "GetOpenMTFFromSiparisNo", cSql)
        End Try
    End Function

    Private Function GetOpenUTFFromSiparisNo(cSiparisno As String) As String

        Dim cSql As String = ""
        Dim ConnYage As SqlConnection

        GetOpenUTFFromSiparisNo = ""

        Try
            ConnYage = OpenConn()

            cSql = "select distinct a.uretimtakipno " + _
                    " from sipmodel a, uretplfis b " + _
                    " where a.uretimtakipno = b.uretimtakipno " + _
                    " and a.uretimtakipno is not null " + _
                    " and a.uretimtakipno <> '' " + _
                    " and (b.dosyakapandi is null or b.dosyakapandi = '' or b.dosyakapandi = 'H') "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and a.siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and a.siparisno = '" + cSiparisno.Trim + "' "
            End If

            GetOpenUTFFromSiparisNo = SQLBuildFilterString2(ConnYage, cSql)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "GetOpenUTFFromSiparisNo", cSql)
        End Try
    End Function

    Private Function GetOpenSTFFromSiparisNo(cSiparisno As String) As String

        Dim cSql As String = ""
        Dim ConnYage As SqlConnection

        GetOpenSTFFromSiparisNo = ""

        Try
            ConnYage = OpenConn()

            cSql = "select distinct a.sevkiyattakipno " + _
                    " from sipmodel a, sevkplfis b " + _
                    " where a.sevkiyattakipno = b.sevkiyattakipno " + _
                    " and a.sevkiyattakipno is not null " + _
                    " and a.sevkiyattakipno <> '' " + _
                    " and (b.ok is null or b.ok = '' or b.ok = 'H') "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and a.siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and a.siparisno = '" + cSiparisno.Trim + "' "
            End If

            GetOpenSTFFromSiparisNo = SQLBuildFilterString2(ConnYage, cSql)

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "GetOpenSTFFromSiparisNo", cSql)
        End Try
    End Function

    Public Sub PlCPdenUretim(Optional cSiparisno As String = "")

        Dim cSql As String = ""
        Dim dTarih As Date = #1/1/1950#
        Dim dBitis As Date = #1/1/1950#
        Dim cSiraNo As String = ""
        Dim nSure As Double = 0
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim aCP() As oCP = Nothing
        Dim aSS() As oSS = Nothing
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        Try
            JustForLog("PlCPdenUretim start")

            cSql = "select siparisno, oktipi, renk, beden, pltarihi, oktar2, avanssuresi " + _
                    " from sipok " + _
                    " where oktipi is not null " + _
                    " and oktipi <> '' "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and siparisno in (select kullanicisipno " + _
                                                " from siparis " + _
                                                " where (siparis.dosyakapandi is null or siparis.dosyakapandi = 'H' or siparis.dosyakapandi = '') " + _
                                                " and (siparis.planlamaok  = 'E') " + _
                                                " and (siparis.plkapanis is null or siparis.plkapanis = 'H' or siparis.plkapanis = '') ) "
            Else
                cSql = cSql + " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            cSql = cSql + " order by siparisno, sira "

            If Not CheckExists(cSql) Then Exit Sub

            ConnYage = OpenConn()
            oReader = GetSQLReader(cSql, ConnYage)

            Do While oReader.Read
                ReDim Preserve aCP(nCnt)
                aCP(nCnt).cSiparisNo = SQLReadString(oReader, "siparisno")
                aCP(nCnt).cOkTipi = SQLReadString(oReader, "oktipi")
                aCP(nCnt).cRenk = SQLReadString(oReader, "renk")
                aCP(nCnt).cBeden = SQLReadString(oReader, "beden")
                aCP(nCnt).dPlTarihi = SQLReadDate(oReader, "pltarihi")
                aCP(nCnt).dOkTar2 = SQLReadDate(oReader, "oktar2")
                aCP(nCnt).nAvansSuresi = SQLReadDouble(oReader, "avanssuresi")
                nCnt = nCnt + 1
            Loop
            oReader.Close()
            ConnYage.Close()

            For nCnt = 0 To UBound(aCP)

                dTarih = #1/1/1950#
                If aCP(nCnt).dOkTar2 > #1/1/1950# Then
                    dTarih = aCP(nCnt).dOkTar2
                ElseIf aCP(nCnt).dPlTarihi > #1/1/1950# Then
                    dTarih = aCP(nCnt).dPlTarihi
                End If
                dTarih = GetNBD(dTarih, aCP(nCnt).nAvansSuresi)

                If dTarih > #1/1/1950# Then

                    ' Kapatılmamış MTF satırları için kaydırma yapılacak
                    cSql = "select sirano, teminsuresi " + _
                            " from mtkfislines " + _
                            " where oktipi = '" + aCP(nCnt).cOkTipi + "' " + _
                            " and (renk = '" + aCP(nCnt).cRenk + "' or renk = 'HEPSI') " + _
                            " and (beden = '" + aCP(nCnt).cBeden + "' or beden = 'HEPSI') " + _
                            " and (kapandi is null or kapandi = 'H' or kapandi = '') " + _
                            " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') " + _
                            " and malzemetakipno in (select malzemetakipno " + _
                                                    " from sipmodel " + _
                                                    " where siparisno = '" + aCP(nCnt).cSiparisNo + "') "

                    If CheckExists(cSql) Then

                        nCnt1 = 0
                        ConnYage = OpenConn()
                        oReader = GetSQLReader(cSql, ConnYage)

                        Do While oReader.Read
                            ReDim Preserve aSS(nCnt1)
                            aSS(nCnt1).nSiraNo = SQLReadDouble(oReader, "sirano")
                            aSS(nCnt1).nTeminSuresi = SQLReadDouble(oReader, "teminsuresi")
                            aSS(nCnt1).dBitis = GetNBD(dTarih, SQLReadDouble(oReader, "teminsuresi"))

                            nCnt = nCnt + 1
                        Loop
                        oReader.Close()
                        ConnYage.Close()

                        For nCnt1 = 0 To UBound(aSS)
                            cSql = "set dateformat dmy " + _
                                    " update mtkfislines " + _
                                    " set baslamatarihi = '" + SQLWriteDate(dTarih) + "', " + _
                                    " bitistarihi = '" + SQLWriteDate(aSS(nCnt1).dBitis) + "' " + _
                                    " where sirano = " + SQLWriteDecimal(aSS(nCnt1).nSiraNo)

                            ExecuteSQLCommand(cSql)

                        Next
                    Else
                        ' Kapatılmamış UTF satırları için kaydırma yapılacak
                        cSql = "select sirano, teminsuresi " + _
                                " from uretpllines " + _
                                " where oktipi = '" + aCP(nCnt).cOkTipi + "' " + _
                                " and uretimtakipno in (select uretimtakipno " + _
                                                        " from sipmodel " + _
                                                        " where siparisno = '" + aCP(nCnt).cSiparisNo + "') " + _
                                " and (okbilgisi is null or okbilgisi = 'H') " + _
                                " and (pltarihkilitle is null or pltarihkilitle = 'H' or pltarihkilitle = '') "

                        If CheckExists(cSql) Then

                            nCnt1 = 0
                            ConnYage = OpenConn()
                            oReader = GetSQLReader(cSql, ConnYage)

                            Do While oReader.Read
                                ReDim Preserve aSS(nCnt1)
                                aSS(nCnt1).nSiraNo = SQLReadDouble(oReader, "sirano")
                                aSS(nCnt1).nTeminSuresi = SQLReadDouble(oReader, "teminsuresi")
                                aSS(nCnt1).dBitis = GetNBD(dTarih, SQLReadDouble(oReader, "teminsuresi"))

                                nCnt = nCnt + 1
                            Loop
                            oReader.Close()
                            ConnYage.Close()

                            For nCnt1 = 0 To UBound(aSS)
                                cSql = "set dateformat dmy " + _
                                        " update uretpllines " + _
                                        " set baslamatarihi = '" + SQLWriteDate(dTarih) + "', " + _
                                        " bitistarihi = '" + SQLWriteDate(aSS(nCnt1).dBitis) + "' " + _
                                        " where sirano = " + SQLWriteDecimal(aSS(nCnt1).nSiraNo)

                                ExecuteSQLCommand(cSql)
                            Next
                        End If
                    End If
                End If
            Next

            JustForLog("PlCPdenUretim finish")

        Catch ex As Exception
            ErrDisp(ex.Message, "PlCPdenUretim", cSql)
        End Try
    End Sub

    ' Gün Hesabı

    Public Sub LoadTatiller()

        Dim cSql As String = ""
        Dim nCnt As Integer = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        Try
            ReDim aTatil(0)
            aTatil(0) = #2/2/1950#
            nCnt = 0

            cSql = "select tatilgunu " + _
                    " from tatilgunleri " + _
                    " where tatilgunu is not null " + _
                    " and tatilgunu <> '01.01.1950' " + _
                    " order by tatilgunu "

            ConnYage = OpenConn()

            oReader = GetSQLReader(cSql, ConnYage)

            Do While oReader.Read
                ReDim Preserve aTatil(nCnt)
                aTatil(nCnt) = SQLReadDate(oReader, "tatilgunu")
                nCnt = nCnt + 1
            Loop
            oReader.Close()

            ConnYage.Close()

        Catch ex As Exception
            ErrDisp(ex.Message, "LoadTatiller", cSql)
        End Try
    End Sub

    Public Function GetTeminSuresi(dBasla As Date, dBitir As Date) As Double

        Dim dTarih As Date = #1/1/1950#
        Dim lTatil As Boolean = False
        Dim nCnt As Integer = 0

        GetTeminSuresi = 0

        Try
            If dBitir <= dBasla Then Exit Function

            dTarih = dBasla

            Do While True
                lTatil = False

                For nCnt = 0 To UBound(aTatil)
                    If dTarih = aTatil(nCnt) Then
                        lTatil = True
                        Exit For
                    ElseIf dTarih < aTatil(nCnt) Then
                        Exit For
                    End If
                Next

                If Not lTatil Then
                    GetTeminSuresi = GetTeminSuresi + 1
                End If
                dTarih = dTarih.AddDays(1)
                If dTarih = dBitir Then
                    Exit Do
                End If
            Loop

        Catch ex As Exception
            ErrDisp(ex.Message, "GetTeminSuresi")
        End Try
    End Function

    Public Function GetFBD(ByVal dTarih As Date, ByVal nDays As Double) As Date
        ' get first available business day, BACKWARD
        Dim nCnt As Double

        GetFBD = dTarih

        Try
            If dTarih = #1/1/1950# Then Exit Function

            If Not BusinessDay(dTarih) Then
                dTarih = GetPrevBusinessDay(dTarih)
            End If

            If nDays > 0 Then
                For nCnt = 1 To nDays
                    dTarih = dTarih.AddDays(-1) ' Subtract(System.TimeSpan.FromDays(1))
                    dTarih = GetPrevBusinessDay(dTarih)
                Next
            End If

            GetFBD = dTarih

        Catch ex As Exception
            ErrDisp(ex.Message, "GetFBD")
        End Try
    End Function

    Private Function GetPrevBusinessDay(ByVal dTarih As Date) As Date

        Dim nCnt As Integer = 0
        Dim lTatil As Boolean = False

        GetPrevBusinessDay = dTarih

        Try
            For nCnt = 0 To UBound(aTatil)
                If dTarih = aTatil(nCnt) Then
                    lTatil = True
                    Exit For
                ElseIf dTarih < aTatil(nCnt) Then
                    Exit For
                End If
            Next

            If lTatil Then
                dTarih = dTarih.AddDays(-1)
                dTarih = GetPrevBusinessDay(dTarih)
            End If
            GetPrevBusinessDay = dTarih

        Catch ex As Exception
            ErrDisp(ex.Message, "GetPrevBusinessDay")
        End Try
    End Function

    Public Function GetNBD(ByVal dTarih As Date, ByVal nDays As Double) As Date
        ' get next available business day, FORWARD
        Dim nCnt As Double = 0

        GetNBD = dTarih

        Try
            If dTarih = #1/1/1950# Then Exit Function

            If Not BusinessDay(dTarih) Then
                dTarih = GetNextBusinessDay(dTarih)
            End If

            If nDays > 0 Then
                For nCnt = 1 To nDays
                    dTarih = dTarih.AddDays(1)
                    dTarih = GetNextBusinessDay(dTarih)
                Next
            End If

            GetNBD = dTarih

        Catch ex As Exception
            ErrDisp(ex.Message, "GetNBD")
        End Try
    End Function

    Private Function BusinessDay(ByVal dTarih As Date) As Boolean

        Dim nCnt As Integer = 0
        Dim lTatil As Boolean = False

        BusinessDay = True

        Try
            For nCnt = 0 To UBound(aTatil)
                If dTarih = aTatil(nCnt) Then
                    lTatil = True
                    Exit For
                ElseIf dTarih < aTatil(nCnt) Then
                    Exit For
                End If
            Next

            BusinessDay = Not lTatil

        Catch ex As Exception
            ErrDisp(ex.Message, "BusinessDay")
        End Try
    End Function

    Private Function GetNextBusinessDay(ByVal dTarih As Date) As Date

        Dim nCnt As Integer = 0
        Dim lTatil As Boolean = False

        GetNextBusinessDay = dTarih

        Try
            For nCnt = 0 To UBound(aTatil)
                If dTarih = aTatil(nCnt) Then
                    lTatil = True
                    Exit For
                ElseIf dTarih < aTatil(nCnt) Then
                    Exit For
                End If
            Next

            If lTatil Then
                dTarih = dTarih.AddDays(1)
                dTarih = GetNextBusinessDay(dTarih)
            End If
            GetNextBusinessDay = dTarih

        Catch ex As Exception
            ErrDisp(ex.Message, "GetNextBusinessDay")
        End Try
    End Function

    Public Sub PlMalzemedenUretim(Optional cSiparisno As String = "")

        Dim cSql As String = ""
        Dim dTarih As Date = #1/1/1950#
        Dim nAvansSuresi As Double = 0
        Dim aMTF() As oMTF = Nothing
        Dim aUTF() As oUTF = Nothing
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        Try
            JustForLog("PlMalzemedenUretim start")

            ' MTF deki Bitiş Tarihlerinin değişmesini UTF lere yansıtır

            cSql = "select distinct a.malzemetakipno, a.departman, b.sira " + _
                    " from mtkfislines a, departman b " + _
                    " where a.departman = b.departman " + _
                    " and a.departman is not null " + _
                    " and a.departman <> '' " + _
                    " and a.malzemetakipno is not null " + _
                    " and a.malzemetakipno <> '' "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " and a.malzemetakipno in (select y.malzemetakipno " + _
                                                        " from siparis x, sipmodel y " + _
                                                        " where x.kullanicisipno = y.siparisno " + _
                                                        " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '') " + _
                                                        " and (x.planlamaok  = 'E') " + _
                                                        " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '') ) "
            Else
                cSql = cSql + " and a.malzemetakipno in (select malzemetakipno " + _
                                                        " from sipmodel " + _
                                                        " where siparisno = '" + cSiparisno.Trim + "') "
            End If

            cSql = cSql + " order by a.malzemetakipno, b.sira "

            If Not CheckExists(cSql) Then Exit Sub

            nCnt = 0
            ConnYage = OpenConn()
            oReader = GetSQLReader(cSql, ConnYage)

            Do While oReader.Read
                ReDim Preserve aMTF(nCnt)
                aMTF(nCnt).cMTF = SQLReadString(oReader, "malzemetakipno")
                aMTF(nCnt).cDepartman = SQLReadString(oReader, "departman")
                aMTF(nCnt).nSira = SQLReadDouble(oReader, "sira")

                nCnt = nCnt + 1
            Loop
            oReader.Close()
            ConnYage.Close()

            For nCnt = 0 To UBound(aMTF)
                dTarih = #1/1/1950#
                nAvansSuresi = 0

                cSql = "select bitistarihi, kapandi, kapanistarihi, avanssuresi " + _
                        " from mtkfislines " + _
                        " where malzemetakipno = '" + aMTF(nCnt).cMTF + "' " + _
                        " and departman = '" + aMTF(nCnt).cDepartman + "' " + _
                        " order by bitistarihi desc "

                ConnYage = OpenConn()
                oReader = GetSQLReader(cSql, ConnYage)

                If oReader.Read Then
                    nAvansSuresi = SQLReadDouble(oReader, "avanssuresi")
                    If SQLReadString(oReader, "kapandi") = "E" Then
                        ' geçerli bir kapanış tarihi varsa kullan
                        dTarih = SQLReadDate(oReader, "kapanistarihi")
                    Else
                        dTarih = SQLReadDate(oReader, "bitistarihi")
                    End If
                    dTarih = GetNBD(dTarih, nAvansSuresi)
                End If
                oReader.Close()
                ConnYage.Close()

                If dTarih > #1/1/1950# Then
                    cSql = "select distinct uretimtakipno, modelno " + _
                            " from sipmodel " + _
                            " where malzemetakipno = '" + aMTF(nCnt).cMTF + "' " + _
                            " order by uretimtakipno, modelno "

                    If CheckExists(cSql) Then

                        nCnt1 = 0
                        ConnYage = OpenConn()
                        oReader = GetSQLReader(cSql, ConnYage)

                        Do While oReader.Read
                            ReDim Preserve aUTF(nCnt1)
                            aUTF(nCnt1).cUTF = SQLReadString(oReader, "uretimtakipno")
                            aUTF(nCnt1).cModelNo = SQLReadString(oReader, "modelno")

                            nCnt1 = nCnt1 + 1
                        Loop
                        oReader.Close()
                        ConnYage.Close()

                        For nCnt1 = 0 To UBound(aUTF)
                            PlUretRecurse(aUTF(nCnt1).cUTF, aUTF(nCnt1).cModelNo, aMTF(nCnt).cDepartman, dTarih)
                        Next
                    End If
                End If
            Next

            cSql = "select distinct uretimtakipno, modelno, departman, baslamatarihi, sira  " + _
                   " from uretpllines "

            If cSiparisno.Trim = "" Then
                cSql = cSql + " where uretimtakipno in (select y.uretimtakipno " + _
                                                        " from siparis x, sipmodel y " + _
                                                        " where x.kullanicisipno = y.siparisno " + _
                                                        " and y.uretimtakipno is not null " + _
                                                        " and y.uretimtakipno <> '' " + _
                                                        " and (x.dosyakapandi is null or x.dosyakapandi = 'H' or x.dosyakapandi = '') " + _
                                                        " and (x.planlamaok  = 'E') " + _
                                                        " and (x.plkapanis is null or x.plkapanis = 'H' or x.plkapanis = '') ) "
            Else
                cSql = cSql + " where uretimtakipno in (select uretimtakipno " + _
                                                        " from sipmodel " + _
                                                        " where siparisno = '" + cSiparisno.Trim + "' " + _
                                                        " and uretimtakipno is not null " + _
                                                        " and uretimtakipno <> '') "
            End If

            cSql = cSql + " order by uretimtakipno, modelno, sira  "

            If CheckExists(cSql) Then

                nCnt1 = 0
                ConnYage = OpenConn()
                oReader = GetSQLReader(cSql, ConnYage)

                Do While oReader.Read
                    ReDim Preserve aUTF(nCnt1)
                    aUTF(nCnt1).cUTF = SQLReadString(oReader, "uretimtakipno")
                    aUTF(nCnt1).cModelNo = SQLReadString(oReader, "modelno")
                    aUTF(nCnt1).cDepartman = SQLReadString(oReader, "departman")
                    aUTF(nCnt1).dBaslamaTarihi = SQLReadDate(oReader, "baslamatarihi")

                    nCnt1 = nCnt1 + 1
                Loop
                oReader.Close()
                ConnYage.Close()

                For nCnt1 = 0 To UBound(aUTF)
                    PlUretRecurse(aUTF(nCnt1).cUTF, aUTF(nCnt1).cModelNo, aUTF(nCnt1).cDepartman, aUTF(nCnt1).dBaslamaTarihi)
                Next
            End If

            JustForLog("PlMalzemedenUretim finish")

        Catch ex As Exception
            ErrDisp(ex.Message, "PlMalzemedenUretim", cSql)
        End Try
    End Sub

    Private Sub PlUretRecurse(cUTF As String, cModelNo As String, cDept As String, dBasla As Date)
        ' üretim departman planlamasındaki başlama tarihindeki GECIKMEYI hesaplar
        Dim cSql As String = ""
        Dim nSira As Double = 0
        Dim dBitisTarihi As Date = #1/1/1950#
        Dim dBaslamaTarihi As Date = #1/1/1950#
        Dim lUpdate As Boolean = False
        Dim dOKTermin As Date = #1/1/1950#
        Dim nAvansSuresi As Double = 0
        Dim dTermin As Date = #1/1/1950#
        Dim aUTF() As oUTF = Nothing
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nCnt As Integer = 0
        Dim nCnt1 As Integer = 0
        Dim cPrevDept As String = ""
        Dim dPrevBitisTarihi As Date = #1/1/1950#

        Try
            If dBasla = #1/1/1950# Then
                Exit Sub
            End If

            lUpdate = True
            dBaslamaTarihi = dBasla

            cSql = "select baslamatarihi, bitistarihi, teminsuresi, okbilgisi, kapanistarihi, oktipi, departman, pltarihkilitle " + _
                    " from uretpllines " + _
                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                    " and modelno = '" + cModelNo.Trim + "' " + _
                    " and departman = '" + cDept.Trim + "' "

            If Not CheckExists(cSql) Then Exit Sub

            nCnt = 0
            ConnYage = OpenConn()
            oReader = GetSQLReader(cSql, ConnYage)

            Do While oReader.Read
                ReDim Preserve aUTF(nCnt)
                aUTF(nCnt).dBaslamaTarihi = SQLReadDate(oReader, "baslamatarihi")
                aUTF(nCnt).dBitisTarihi = SQLReadDate(oReader, "bitistarihi")
                aUTF(nCnt).nTeminSuresi = SQLReadDouble(oReader, "teminsuresi")
                aUTF(nCnt).cOkBilgisi = SQLReadString(oReader, "okbilgisi")
                aUTF(nCnt).dKapanisTarihi = SQLReadDate(oReader, "kapanistarihi")
                aUTF(nCnt).cOkTipi = SQLReadString(oReader, "oktipi")
                aUTF(nCnt).cDepartman = SQLReadString(oReader, "departman")
                aUTF(nCnt).cPlTarihKilitle = SQLReadString(oReader, "pltarihkilitle")

                nCnt1 = nCnt1 + 1
            Loop
            oReader.Close()
            ConnYage.Close()

            For nCnt = 0 To UBound(aUTF)
                If aUTF(nCnt).cOkBilgisi = "E" Then
                    ' kapanmış UTF satırlarında başlama tarihi değişmez
                    ' kapanmış UTF satırlarında bitiş tarihi kapanış tarihi olur
                    dBaslamaTarihi = aUTF(nCnt).dBaslamaTarihi
                    dBitisTarihi = aUTF(nCnt).dKapanisTarihi
                    lUpdate = False
                ElseIf aUTF(nCnt).cPlTarihKilitle = "E" Then
                    ' kilitlenmiş UTF satırlarında başlama tarihi değişmez
                    ' kilitlenmiş UTF satırlarında bitiş tarihi değişmez
                    dBaslamaTarihi = aUTF(nCnt).dBaslamaTarihi
                    dBitisTarihi = aUTF(nCnt).dBitisTarihi
                    lUpdate = False
                Else
                    ' bir önceki departmanın bitiş tarihi
                    cSql = "select sira " + _
                            " from modeluretim " + _
                            " where modelno = '" + cModelNo.Trim + "' " + _
                            " and departman = '" + cDept.Trim + "' "

                    nSira = SQLGetDouble(cSql)

                    If nSira <> 0 Then
                        cSql = "select departman " + _
                                " from modeluretim " + _
                                " where modelno = '" + cModelNo.Trim + "' " + _
                                " and sira < " + SQLWriteDecimal(nSira) + _
                                " and departman <> '" + cDept.Trim + "' " + _
                                " order by sira desc "

                        cPrevDept = SQLGetString(cSql)

                        If cPrevDept.Trim <> "" Then

                            cSql = "select bitistarihi " + _
                                    " from uretpllines " + _
                                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                    " and modelno = '" + cModelNo.Trim + "' " + _
                                    " and departman = '" + cPrevDept.Trim + "' "

                            dPrevBitisTarihi = SQLGetDate(cSql)

                         End If
                    End If

                    ' bir önceki departmanın bitiş tarihi prosedüre gelen tarihten büyükse onu alalım
                    If dPrevBitisTarihi > dBaslamaTarihi Then
                        dBaslamaTarihi = dPrevBitisTarihi
                    End If

                    ' orjinal başlama tarihi büyükse onu alalım
                    If aUTF(nCnt).dBaslamaTarihi > dBaslamaTarihi Then
                        dBaslamaTarihi = aUTF(nCnt).dBaslamaTarihi
                    End If

                    ' bağlı OKlerin durumuna bakılır
                    dOKTermin = #1/1/1950#

                    cSql = "select pltarihi, oktar2, ok, avanssuresi " + _
                            " from sipok " + _
                            " where oktipi = '" + aUTF(nCnt).cOkTipi + "' " + _
                            " and modelkodu = '" + cModelNo.Trim + "' " + _
                            " and siparisno in (select siparisno " + _
                                                " from sipmodel " + _
                                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                                " and modelno = '" + cModelNo.Trim + "') "

                    ConnYage = OpenConn()
                    oReader = GetSQLReader(cSql, ConnYage)

                    Do While oReader.Read
                        nAvansSuresi = SQLReadDouble(oReader, "avanssuresi")
                        If SQLReadString(oReader, "ok") = "E" Then
                            ' aşama kapandıysa
                            dOKTermin = SQLReadDate(oReader, "oktar2")
                        ElseIf SQLReadDate(oReader, "pltarihi") > #1/1/1950# Then
                            dOKTermin = SQLReadDate(oReader, "pltarihi")
                        End If
                        dOKTermin = GetNBD(dOKTermin, nAvansSuresi)
                    Loop
                    oReader.Close()
                    ConnYage.Close()
                    ' bağlı okeylerin son tarihi büyükse onu alalım
                    If dOKTermin > dBaslamaTarihi Then
                        dBaslamaTarihi = dOKTermin
                    End If

                    ' bağlı malzemelerin durumuna bakılır
                    dTermin = #1/1/1950#

                    cSql = "select bitistarihi, avanssuresi, kapandi, kapanistarihi " + _
                            " from mtkfislines " + _
                            " where departman = '" + aUTF(nCnt).cDepartman + "' " + _
                            " and malzemetakipno in (select malzemetakipno " + _
                                                " from sipmodel " + _
                                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                                " and modelno = '" + cModelNo.Trim + "') " + _
                            " order by bitistarihi "

                    ConnYage = OpenConn()
                    oReader = GetSQLReader(cSql, ConnYage)

                    Do While oReader.Read
                        nAvansSuresi = SQLReadDouble(oReader, "avanssuresi")
                        If SQLReadString(oReader, "kapandi") = "E" Then
                            ' aşama kapandıysa
                            dTermin = SQLReadDate(oReader, "kapanistarihi")
                        ElseIf SQLReadDate(oReader, "bitistarihi") > #1/1/1950# Then
                            dTermin = SQLReadDate(oReader, "bitistarihi")
                        End If
                        dTermin = GetNBD(dTermin, nAvansSuresi)
                        If dTermin > dOKTermin Then
                            dOKTermin = dTermin
                        End If
                    Loop
                    oReader.Close()
                    ConnYage.Close()
                    ' bağlı malzemelerin son tarihi büyükse onu alalım
                    If dTermin > dBaslamaTarihi Then
                        dBaslamaTarihi = dTermin
                    End If
                    ' bitiş tarihini hesaplayalım
                    dBitisTarihi = GetNBD(dBaslamaTarihi, aUTF(nCnt).nTeminSuresi)
                End If
            Next

            If dBaslamaTarihi = #1/1/1950# Or dBitisTarihi = #1/1/1950# Then
                Exit Sub
            End If

            If lUpdate Then
                cSql = "set dateformat dmy " + _
                        " update uretpllines " + _
                        " set baslamatarihi = '" + SQLWriteDate(dBaslamaTarihi) + "', " + _
                        " bitistarihi = '" + SQLWriteDate(dBitisTarihi) + "' " + _
                        " where uretimtakipno = '" + cUTF.Trim + "' " + _
                        " and modelno = '" + cModelNo.Trim + "' " + _
                        " and departman = '" + cDept.Trim + "' "

                ExecuteSQLCommand(cSql)
            End If

            'cSql = "select sira " + _
            '        " from modeluretim " + _
            '        " where modelno = '" + cModelNo.Trim + "' " + _
            '        " and departman = '" + cDept.Trim + "' "

            'nSira = SQLGetDouble(cSql)

            'If nSira = 0 Then
            '    Exit Sub
            'End If

            'cSql = "select departman " + _
            '        " from modeluretim " + _
            '        " where modelno = '" + cModelNo.Trim + "' " + _
            '        " and sira > " + SQLWriteDecimal(nSira) + _
            '        " and departman <> '" + cDept.Trim + "' " + _
            '        " order by sira "

            'cDept = SQLGetString(cSql)

            'If cDept.Trim <> "" Then
            '    PlUretRecurse(cUTF.Trim, cModelNo.Trim, cDept.Trim, dBitisTarihi)
            'End If

        Catch ex As Exception
            ErrDisp(ex.Message, "PlUretRecurse", cSql)
        End Try
    End Sub

    Private Sub UretimPlLinesKapat(nSiraNo As Double, dTarih As Date, Optional cSiparisno As String = "")

        Dim cSql As String = ""
        Dim cDepartman As String = ""
        Dim cModelNo As String = ""
        Dim cOKTipi As String = ""
        Dim cUTF As String = ""
        Dim dGBasla As Date = #1/1/1950#
        Dim dGBitir As Date = #1/1/1950#
        Dim lOK As Boolean = False
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim G_MTKUretCikNoDept As Boolean = False

        Try
            If nSiraNo = 0 Then Exit Sub

            G_MTKUretCikNoDept = (GetSysPar("mtkuretciknodept") = "1")

            ' Bilgilerine ulaş
            cSql = "select uretimtakipno, departman, modelno, oktipi, kapanistarihi  " + _
                    " from uretpllines " + _
                    " where sirano = " + SQLWriteDecimal(nSiraNo)

            ConnYage = OpenConn()
            oReader = GetSQLReader(cSql, ConnYage)

            If oReader.Read Then
                lOK = True
                cUTF = SQLReadString(oReader, "uretimtakipno")
                cDepartman = SQLReadString(oReader, "departman")
                cModelNo = SQLReadString(oReader, "modelno")
                cOKTipi = SQLReadString(oReader, "oktipi")
                dTarih = SQLReadDate(oReader, "kapanistarihi")
            End If
            oReader.Close()
            ConnYage.Close()

            If Not lOK Then Exit Sub

            ConnYage = OpenConn()
            GetUTFGercekTarih(ConnYage, cUTF, cDepartman, cModelNo, dGBasla, dGBitir)
            ConnYage.Close()

            If dGBitir > #1/1/1950# Then
                dTarih = dGBitir
            End If

            ' Kendisini KAPAT
            cSql = "set dateformat dmy " + _
                    " update uretpllines " + _
                    " set okbilgisi = 'E', " + _
                    " kapanistarihi = '" + SQLWriteDate(dTarih) + "' " + _
                    " where sirano = " + SQLWriteDecimal(nSiraNo)

            ExecuteSQLCommand(cSql)

            ' Bağlı MTF satırlarını kapat
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set kapandi = 'E', " + _
                    " kapanistarihi = '" + SQLWriteDate(dTarih) + "' " + _
                    " where departman = '" + cDepartman.Trim + "' " + _
                    " and malzemetakipno in (select malzemetakipno " + _
                                            " from sipmodel " + _
                                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                            " and modelno = '" + cModelNo.Trim + "') "
            ExecuteSQLCommand(cSql)

            ' Bağlı işemri satırlarını kapat
            'cSql = "update isemrilines " + _
            '        " set kapandi = 'E' " + _
            '        " where malzemetakipno in (select malzemetakipno " + _
            '                                " from sipmodel " + _
            '                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
            '                                " and modelno = '" + cModelNo.Trim + "') " + _
            '        " and exists (select malzemetakipno " + _
            '                        " from mtkfislines " + _
            '                        " where malzemetakipno = isemrilines.malzemetakipno " + _
            '                        " and stokno = isemrilines.stokno " + _
            '                        " and renk = isemrilines.renk " + _
            '                        " and beden = isemrilines.beden " + _
            '                        IIf(G_MTKUretCikNoDept, "", " and departman = isemrilines.departman ").ToString + _
            '                        " and kapandi = 'E' )"
            'ExecuteSQLCommand(cSql)

            If cOKTipi.Trim <> "" Then
                ' Üretime bağlı CP adımlarını kapat
                cSql = "set dateformat dmy " + _
                        " update sipok " + _
                        " set ok = 'E', " + _
                        " oktar2 = '" + SQLWriteDate(dTarih) + "' " + _
                        " where oktipi = '" + cOKTipi.Trim + "' " + _
                        " and modelkodu = '" + cModelNo.Trim + "' "

                If cSiparisno.Trim = "" Then
                    cSql = cSql + _
                            " and siparisno in (select siparisno " + _
                                                " from sipmodel " + _
                                                " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                                " and modelno = '" + cModelNo.Trim + "') "
                Else
                    cSql = cSql + _
                            " and siparisno = '" + cSiparisno.Trim + "' "
                End If

                ExecuteSQLCommand(cSql)
            End If
            ' Malzemeye bağlı CP adımlarını kapat
            cSql = "set dateformat dmy " + _
                    " update sipok " + _
                    " set ok = 'E', " + _
                    " oktar2 = '" + SQLWriteDate(dTarih) + "' " + _
                    " where modelkodu = '" + cModelNo.Trim + "' "

            If cSiparisno.Trim = "" Then
                cSql = cSql + _
                        " and siparisno in (select siparisno " + _
                                            " from sipmodel " + _
                                            " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                            " and modelno = '" + cModelNo.Trim + "') "
            Else
                cSql = cSql + _
                        " and siparisno = '" + cSiparisno.Trim + "' "
            End If

            cSql = cSql + _
                " and exists (select oktipi " + _
                            " from mtkfislines " + _
                            " where kapandi = 'E' " + _
                            " and oktipi = sipok.oktipi  " + _
                            " and malzemetakipno in (select malzemetakipno " + _
                                                    " from sipmodel " + _
                                                    " where uretimtakipno = '" + cUTF.Trim + "' " + _
                                                    " and siparisno = sipok.siparisno " + _
                                                    " and modelno = '" + cModelNo.Trim + "') )"
            ExecuteSQLCommand(cSql)

        Catch ex As Exception
            ErrDisp(ex.Message, "UretimPlLinesKapat", cSql)
        End Try
    End Sub

    Private Sub MTKFisLinesKapat(nSiraNo As Double, dTarih As Date, Optional cSiparisno As String = "")

        Dim cSql As String = ""
        Dim cOKTipi As String = ""
        Dim cMTF As String = ""
        Dim dGBasla As Date = #1/1/1950#
        Dim dGBitir As Date = #1/1/1950#
        Dim cStokno As String = ""
        Dim cRenk As String = ""
        Dim cbeden As String = ""
        Dim cDept As String = ""
        Dim lOK As Boolean = False
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim G_MTKUretCikNoDept As Boolean = False

        If nSiraNo = 0 Then Exit Sub

        Try
            G_MTKUretCikNoDept = (GetSysPar("mtkuretciknodept") = "1")

            ' Bilgilerine ulaş
            cSql = "select malzemetakipno, stokno, renk, beden, oktipi, kapanistarihi, departman " + _
                    " from mtkfislines " + _
                    " where sirano = " + SQLWriteDecimal(nSiraNo)

            ConnYage = OpenConn()
            oReader = GetSQLReader(cSql, ConnYage)

            If oReader.Read Then
                lOK = True
                cMTF = SQLReadString(oReader, "malzemetakipno")
                cOKTipi = SQLReadString(oReader, "oktipi")
                dTarih = SQLReadDate(oReader, "kapanistarihi")
                cStokno = SQLReadString(oReader, "stokno")
                cRenk = SQLReadString(oReader, "renk")
                cbeden = SQLReadString(oReader, "beden")
                cDept = SQLReadString(oReader, "departman")
            End If
            oReader.Close()
            ConnYage.Close()

            If Not lOK Then Exit Sub

            ConnYage = OpenConn()
            GetMTFGercekTarih(ConnYage, cMTF, cStokno, cRenk, cbeden, dGBasla, dGBitir)
            ConnYage.Close()

            If dGBitir > #1/1/1950# Then
                dTarih = dGBitir
            End If

            ' Kendisini KAPAT
            cSql = "set dateformat dmy " + _
                    " update mtkfislines " + _
                    " set kapandi = 'E', " + _
                    " kapanistarihi = '" + SQLWriteDate(dTarih) + "' " + _
                    " where sirano = " + SQLWriteDecimal(nSiraNo)

            ExecuteSQLCommand(cSql)

            ' Bağlı işemri satırlarını kapat
            'cSql = "update isemrilines " + _
            '        " set kapandi = 'E' " + _
            '        " where malzemetakipno = '" + cMTF.Trim + "' " + _
            '        " and stokno = '" + cStokno.Trim + "' " + _
            '        " and renk = '" + cRenk.Trim + "' " + _
            '        " and beden = '" + cbeden.Trim + "' " + _
            '        IIf(G_MTKUretCikNoDept Or cDept.Trim = "", "", " and departman = '" + cDept.Trim + "' ").ToString

            'ExecuteSQLCommand(cSql)

            If cOKTipi <> "" Then
                ' Bağlı CP adımlarını kapat
                cSql = "set dateformat dmy " + _
                        " update sipok " + _
                        " set ok = 'E', " + _
                        " oktar2 = '" + SQLWriteDate(dTarih) + "' " + _
                        " where oktipi = '" + cOKTipi.Trim + "' "

                If cSiparisno.Trim = "" Then
                    cSql = cSql + _
                            " and siparisno in (select siparisno " + _
                                                " from sipmodel " + _
                                                " where malzemetakipno = '" + cMTF.Trim + "') "
                Else
                    cSql = cSql + _
                            " and siparisno = '" + cSiparisno.Trim + "' "
                End If

                ExecuteSQLCommand(cSql)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "MTKFisLinesKapat", cSql)
        End Try
    End Sub

    Private Sub SevkPlFisLinesKapat(nSiraNo As Double, dTarih As Date, Optional cSiparisno As String = "")

        Dim cSql As String = ""
        Dim cOKTipi As String = ""
        Dim cSTF As String = ""
        Dim cModelNo As String = ""
        Dim dGerceklesen As Date = #1/1/1950#
        Dim lOK As Boolean = False
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader

        Try
            If nSiraNo = 0 Then Exit Sub

            ' Bilgilerine ulaş
            cSql = "select sevkiyattakipno, siparisno, modelno, oktipi, oktarihi " + _
                    " from SevkPlFisLines " + _
                    " where sirano = " + SQLWriteDecimal(nSiraNo)

            ConnYage = OpenConn()
            oReader = GetSQLReader(cSql, ConnYage)

            If oReader.Read Then
                lOK = True
                cSTF = SQLReadString(oReader, "sevkiyattakipno")
                cSiparisno = SQLReadString(oReader, "siparisno")
                cModelNo = SQLReadString(oReader, "modelno")
                cOKTipi = SQLReadString(oReader, "oktipi")
                dTarih = SQLReadDate(oReader, "oktarihi")
            End If
            oReader.Close()
            ConnYage.Close()

            If Not lOK Then Exit Sub

            ConnYage = OpenConn()
            GetSTFGercekTarih(ConnYage, cSTF, dGerceklesen, cSiparisno, cModelNo)
            ConnYage.Close()

            If dGerceklesen > #1/1/1950# Then
                dTarih = dGerceklesen
            End If

            ' Kendisini KAPAT
            cSql = "set dateformat dmy " + _
                    " update SevkPlFisLines " + _
                    " set ok = 'E', " + _
                    " oktarihi = '" + SQLWriteDate(dTarih) + "' " + _
                    " where sirano = " + SQLWriteDecimal(nSiraNo)

            ExecuteSQLCommand(cSql)

            If cOKTipi.Trim <> "" Then
                ' Bağlı CP adımlarını kapat
                cSql = "set dateformat dmy " + _
                        " update sipok " + _
                        " set ok = 'E', " + _
                        " oktar2 = '" + SQLWriteDate(dTarih) + "' " + _
                        " where oktipi = '" + cOKTipi.Trim + "' " + _
                        " and modelkodu = '" + cModelNo.Trim + "' "

                If cSiparisno.Trim = "" Then
                    cSql = cSql + _
                        " and siparisno in (select siparisno " + _
                                            " from sipmodel " + _
                                            " where sevkiyattakipno = '" + cSTF.Trim + "') "
                Else
                    cSql = cSql + _
                        " and siparisno = '" + cSiparisno.Trim + "' "
                End If

                ExecuteSQLCommand(cSql)
            End If

        Catch ex As Exception
            ErrDisp(ex.Message, "SevkPlFisLinesKapat", cSql)
        End Try
    End Sub

    Public Sub OtomatikKapatmalar(Optional cSiparisno As String = "")

        Dim cSQL As String = ""
        Dim cMTF As String = ""
        Dim cUTF As String = ""
        Dim cSTF As String = ""
        Dim aSiraNo() As Double = Nothing
        Dim ConnYage As SqlConnection
        Dim oReader As SqlDataReader
        Dim nCnt As Integer = 0

        Try
            JustForLog("Otomatik kapatmalar basladi")

            cSTF = GetOpenSTFFromSiparisNo(cSiparisno)
            If cSTF.Trim <> "" Then

                cSQL = "select SiraNo " + _
                        " from sevkplfislines " + _
                        " where (ok is null or ok = '' or ok = 'H') " + _
                        " and sevkiyattakipno in (" + cSTF.Trim + ") " + _
                        " and coalesce(giden,0) >= coalesce(toplam,0) " + _
                        " order by sevkiyattakipno "

                If CheckExists(cSQL) Then
                    nCnt = 0
                    ConnYage = OpenConn()
                    oReader = GetSQLReader(cSQL, ConnYage)

                    Do While oReader.Read
                        ReDim Preserve aSiraNo(nCnt)
                        aSiraNo(nCnt) = SQLReadDouble(oReader, "SiraNo")
                        nCnt = nCnt + 1
                    Loop
                    oReader.Close()
                    ConnYage.Close()

                    For nCnt = 0 To UBound(aSiraNo)
                        SevkPlFisLinesKapat(aSiraNo(nCnt), Now, cSiparisno)
                    Next
                End If
            End If

            cUTF = GetOpenUTFFromSiparisNo(cSiparisno)
            If cUTF.Trim <> "" Then

                cSQL = "select SiraNo " + _
                        " from uretpllines " + _
                        " where (okbilgisi is null or okbilgisi = '' or okbilgisi = 'H') " + _
                        " and uretimtakipno in (" + cUTF.Trim + ") " + _
                        " and coalesce(giden,0) >= coalesce(ToplamAdet,0) " + _
                        " order by uretimtakipno "


                If CheckExists(cSQL) Then
                    nCnt = 0
                    ConnYage = OpenConn()
                    oReader = GetSQLReader(cSQL, ConnYage)

                    Do While oReader.Read
                        ReDim Preserve aSiraNo(nCnt)
                        aSiraNo(nCnt) = SQLReadDouble(oReader, "SiraNo")
                        nCnt = nCnt + 1
                    Loop
                    oReader.Close()
                    ConnYage.Close()

                    For nCnt = 0 To UBound(aSiraNo)
                        UretimPlLinesKapat(aSiraNo(nCnt), Now, cSiparisno)
                    Next
                End If
            End If

            cMTF = GetOpenMTFFromSiparisNo(cSiparisno)
            If cMTF.Trim <> "" Then
                ' Malzeme Planlama Kapanışı
                cSQL = "select SiraNo " + _
                        " from mtkfislines " + _
                        " where (kapandi is null or kapandi = 'H' or kapandi = '') " + _
                        " and MalzemeTakipNo in (" + cMTF.Trim + ") " + _
                        " and (coalesce(uretimicincikis,0) - coalesce(uretimdeniade,0)) >= coalesce(ihtiyac,0) " + _
                        " order by MalzemeTakipNo, StokNo, Renk, Beden "

                If CheckExists(cSQL) Then
                    nCnt = 0
                    ConnYage = OpenConn()
                    oReader = GetSQLReader(cSQL, ConnYage)

                    Do While oReader.Read
                        ReDim Preserve aSiraNo(nCnt)
                        aSiraNo(nCnt) = SQLReadDouble(oReader, "SiraNo")
                        nCnt = nCnt + 1
                    Loop
                    oReader.Close()
                    ConnYage.Close()

                    For nCnt = 0 To UBound(aSiraNo)
                        MTKFisLinesKapat(aSiraNo(nCnt), Now, cSiparisno)
                    Next
                End If
            End If
            JustForLog("Otomatik kapatmalar bitti")

        Catch ex As Exception
            ErrDisp(ex.Message, "OtomatikKapatmalar", cSQL)
        End Try
    End Sub
End Module

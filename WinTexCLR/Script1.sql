use alderswintex
declare @nSonuc int
set @nSonuc = 99
exec dbo.STISonMaliyetUretimCLR ' and (a.dosyakapandi is null or a.dosyakapandi = ''H'' or a.dosyakapandi = '''') ', @nSonuc
select @nSonuc
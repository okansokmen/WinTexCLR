﻿--execute TedarikciDegerlendirme '01.01.2010','31.01.2010','','Kumaş',1
--execute TedarikciDegerlendirme '01.01.2010','31.01.2010','','Aksesuar',1
--execute TedarikciDegerlendirme '01.04.2010','30.04.2010',"a.firma='CENTEKS YIKAMA' and a.uretimtakipno='10-Sip.0404'",'Üretim',1
--execute TedarikciDegerlendirme '01.01.2010','31.01.2010','','Ticari Üretim',1
--execute SipMlzMaliyetList '10-Sip.0861'
--execute PersonelPerformans '01.05.2010','31.05.2010','Melik Gülmez','Ay',1
--execute ValidateStokFis 'validate','0000113037','','','',1
--Execute SiparisSatisIadeUpdate '10-Sip.0057','Cikis','07 Satis'
--execute tekstoktoplam '805-0022.G100.Y001.C','', '', 1
--select * from RaporSiparisDurumu1()
--execute MTFHesaplax '','','','deneme',1
--execute FastMTFBuild '000001196/13WA036', 1
--execute PlanlamaForwardAll '13-0308'
execute BuildMasterPlan '',1
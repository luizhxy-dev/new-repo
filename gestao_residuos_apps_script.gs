// ============================================================
//  NATTO - Gestao de Residuos 2026
//  Google Apps Script - Integracao de Dados
//  Planilha: https://docs.google.com/spreadsheets/d/1NVs-xQtGT6-ZTaAaiAWdCzjVJ3wqP_ndtNKEsuoQTOY
// ============================================================

var SPREADSHEET_ID = '1NVs-xQtGT6-ZTaAaiAWdCzjVJ3wqP_ndtNKEsuoQTOY';

// Dados extraidos do dashboard (273 registros, 4 unidades)
var DADOS = [
  ['Belo Jardim','2026-01-02','Resíduo comum (lixo) (produção e ADM)','',0.0,1170.0,585.0,585.0],
  ['Belo Jardim','2026-01-02','Resíduo da peneira (ETE)','',0.0,2370.0,2370.0,0.0],
  ['Belo Jardim','2026-01-03','Resíduo comum (lixo) (produção e ADM)','',0.0,1180.0,590.0,590.0],
  ['Belo Jardim','2026-01-03','Lodo (ETE)','',0.0,1560.0,1560.0,0.0],
  ['Belo Jardim','2026-01-03','Resíduo da peneira (ETE)','',0.0,2570.0,2570.0,0.0],
  ['Belo Jardim','2026-01-05','Resíduo comum (lixo) (produção e ADM)','',0.0,1940.0,970.0,970.0],
  ['Belo Jardim','2026-01-05','Plástico','',0.0,2100.0,0.0,2100.0],
  ['Belo Jardim','2026-01-05','Resíduo da saúde','',2.0,0.0,2.0,0.0],
  ['Belo Jardim','2026-01-06','Resíduo comum (lixo) (produção e ADM)','',0.0,2070.0,1035.0,1035.0],
  ['Belo Jardim','2026-01-07','Resíduo comum (lixo) (produção e ADM)','',0.0,1730.0,865.0,865.0],
  ['Belo Jardim','2026-01-07','Cinza (caldeira)','',0.0,18810.0,18810.0,0.0],
  ['Belo Jardim','2026-01-07','Resíduo da peneira (ETE)','',0.0,3400.0,3400.0,0.0],
  ['Belo Jardim','2026-01-07','Lodo (ETE)','',0.0,2390.0,2390.0,0.0],
  ['Belo Jardim','2026-01-07','Óleo (frota e manutenção','',2010.0,0.0,0.0,2010.0],
  ['Belo Jardim','2026-01-08','Resíduo comum (lixo)','261028952734',0.0,1620.0,810.0,810.0],
  ['Belo Jardim','2026-01-08','Resíduo da peneira (ETE)','261028959058',0.0,2590.0,2590.0,0.0],
  ['Belo Jardim','2026-01-09','Resíduo comum (lixo) (produção e ADM)','261028973365',0.0,1610.0,805.0,805.0],
  ['Belo Jardim','2026-01-09','Resíduo da peneira (ETE)','261028976816',0.0,2650.0,2650.0,0.0],
  ['Belo Jardim','2026-01-09','Lodo (ETE)','26102897456',0.0,3230.0,3230.0,0.0],
  ['Belo Jardim','2026-01-10','Resíduo comum (lixo) (produção e ADM)','261028987834',0.0,1470.0,735.0,735.0],
  ['Belo Jardim','2026-01-12','Resíduo comum (lixo) (produção e ADM)','261028999697',0.0,1320.0,660.0,660.0],
  ['Belo Jardim','2026-01-13','Resíduo comum (lixo) (produção e ADM)','261029020988',0.0,1920.0,960.0,960.0],
  ['Belo Jardim','2026-01-13','Resíduo da peneira (ETE)','261029026786',0.0,5460.0,5460.0,0.0],
  ['Belo Jardim','2026-01-14','Resíduo comum (lixo) (produção e ADM)','261029039374',0.0,1570.0,785.0,785.0],
  ['Belo Jardim','2026-01-15','Resíduo comum (lixo) (produção e ADM)','261029059723',0.0,1810.0,905.0,905.0],
  ['Belo Jardim','2026-01-15','Resíduo da peneira (ETE)','261029066637',0.0,2370.0,2370.0,0.0],
  ['Belo Jardim','2026-01-16','Resíduo comum (lixo) (produção e ADM)','261029081523',0.0,1540.0,770.0,770.0],
  ['Belo Jardim','2026-01-17','Resíduo comum (lixo) (produção e ADM)','261029096546',0.0,1470.0,735.0,735.0],
  ['Belo Jardim','2026-01-19','Estopa (frota)','261029106959',91.0,0.0,91.0,0.0],
  ['Belo Jardim','2026-01-19','Filtro de óleo (frota)','261029106959',91.0,0.0,91.0,0.0],
  ['Belo Jardim','2026-01-19','Resíduo comum (lixo) (produção e ADM)','261029107872',0.0,1720.0,860.0,860.0],
  ['Belo Jardim','2026-01-19','Lodo (ETE)','261029112760',0.0,4320.0,4320.0,0.0],
  ['Belo Jardim','2026-01-19','Solo contaminado (ETE)','261029113807',880.0,0.0,880.0,0.0],
  ['Belo Jardim','2026-01-20','Resíduo comum (lixo) (produção e ADM)','261029129902',0.0,1510.0,755.0,755.0],
  ['Belo Jardim','2026-01-21','Resíduo comum (lixo) (produção e ADM)','261029152597',0.0,2340.0,1170.0,1170.0],
  ['Belo Jardim','2026-01-22','Resíduo comum (lixo) (produção e ADM)','261029170814',0.0,1400.0,700.0,700.0],
  ['Belo Jardim','2026-01-23','Resíduo comum (lixo) (produção e ADM)','261029189224',0.0,1840.0,920.0,920.0],
  ['Belo Jardim','2026-01-24','Resíduo da peneira (ETE)','261029204414',0.0,3400.0,3400.0,0.0],
  ['Belo Jardim','2026-01-24','Resíduo comum (lixo) (produção e ADM)','261029203977',0.0,1450.0,725.0,725.0],
  ['Belo Jardim','2026-01-26','Plástico','',0.0,1630.0,0.0,1630.0],
  ['Belo Jardim','2026-01-26','Resíduo comum (lixo) (produção e ADM)','',0.0,1180.0,590.0,590.0],
  ['Belo Jardim','2026-01-27','Resíduo comum (lixo) (produção e ADM)','',0.0,2210.0,1105.0,1105.0],
  ['Belo Jardim','2026-01-27','Metralha (resíduo da construção civil)','',0.0,10180.0,10180.0,0.0],
  ['Belo Jardim','2026-01-27','Papelão','',0.0,5370.0,0.0,5370.0],
  ['Belo Jardim','2026-01-28','Resíduo comum (lixo) (produção e ADM)','261029257458',0.0,1590.0,795.0,795.0],
  ['Belo Jardim','2026-01-29','Resíduo comum (lixo) (produção e ADM)','61029277402',0.0,1740.0,870.0,870.0],
  ['Belo Jardim','2026-01-29','Resíduo comum (lixo) (produção e ADM)','261029286976',0.0,8610.0,8610.0,0.0],
  ['Belo Jardim','2026-01-30','Resíduo comum (lixo) (produção e ADM)','261029299313',0.0,1560.0,780.0,780.0],
  ['Belo Jardim','2026-01-30','Resíduo comum (lixo) (produção e ADM)','261029312351',0.0,1560.0,780.0,780.0],
  ['Belo Jardim','2026-02-02','Plástico','261029327877',0.0,580.0,0.0,580.0],
  ['Belo Jardim','2026-02-02','Resíduo comum (lixo) (produção e ADM)','261029328658',0.0,1130.0,565.0,565.0],
  ['Belo Jardim','2026-02-03','Resíduo comum (lixo) (produção e ADM)','261029353458',0.0,2030.0,1015.0,1015.0],
  ['Belo Jardim','2026-02-03','Residuo da UFO (produção e ADM)','261029357757',0.0,11000.0,0.0,11000.0],
  ['Belo Jardim','2026-02-04','Resíduo comum (lixo) (produção e ADM)','',0.0,1440.0,720.0,720.0],
  ['Belo Jardim','2026-02-04','Resíduo da UFO (produção)','261029383585',0.0,11810.0,11810.0,0.0],
  ['Belo Jardim','2026-02-04','Resíduo da UFO (produção)','261029387060',0.0,10740.0,0.0,10740.0],
  ['Belo Jardim','2026-02-05','Resíduo comum (lixo) (produção e ADM)','',0.0,1560.0,780.0,780.0],
  ['Belo Jardim','2026-02-05','Resíduo da peneira (ETE)','261029403355',0.0,9850.0,9850.0,0.0],
  ['Belo Jardim','2026-02-05','Resíduo da UFO (produção)','261029403612',0.0,13790.0,13790.0,0.0],
  ['Belo Jardim','2026-02-05','Resíduo da UFO (produção)','261029408084',0.0,11410.0,0.0,11410.0],
  ['Belo Jardim','2026-02-06','Resíduo comum (lixo) (produção e ADM)','261029430686',0.0,1620.0,810.0,810.0],
  ['Belo Jardim','2026-02-06','Resíduo da UFO (produção)','261029430414',0.0,11240.0,0.0,11240.0],
  ['Belo Jardim','2026-02-06','Resíduo da UFO (produção)','261029430307',0.0,13030.0,13030.0,0.0],
  ['Belo Jardim','2026-02-07','Resíduo comum (lixo) (produção e ADM)','261029430733',0.0,1300.0,650.0,650.0],
  ['Belo Jardim','2026-02-09','Resíduo da UFO (produção)','261029447498',0.0,10950.0,10950.0,0.0],
  ['Belo Jardim','2026-02-09','Resíduo comum (lixo) (produção e ADM)','261029450491',0.0,1630.0,815.0,815.0],
  ['Belo Jardim','2026-02-10','Resíduo comum (lixo) (produção e ADM)','261029465180',0.0,1880.0,940.0,940.0],
  ['Belo Jardim','2026-02-10','Metralha (resíduo da construção civil)','261029466080',0.0,12390.0,12390.0,0.0],
  ['Belo Jardim','2026-02-10','Resíduo da UFO (produção)','261029477121',0.0,17290.0,17290.0,0.0],
  ['Belo Jardim','2026-02-10','Resíduo da UFO (produção)','261029477657',0.0,11920.0,0.0,11920.0],
  ['Belo Jardim','2026-02-11','Resíduo comum (lixo) (produção e ADM)','261029484137',0.0,1840.0,920.0,920.0],
  ['Belo Jardim','2026-02-11','Resíduo da UFO (produção)','261029499637',0.0,16920.0,16920.0,0.0],
  ['Belo Jardim','2026-02-12','Resíduo comum (lixo) (produção e ADM)','',0.0,2040.0,1020.0,1020.0],
  ['Belo Jardim','2026-02-12','Resíduo da peneira (ETE)','261029516254',0.0,8180.0,8180.0,0.0],
  ['Belo Jardim','2026-02-13','Resíduo comum (lixo) (produção e ADM)','261029529556',0.0,1840.0,920.0,920.0],
  ['Belo Jardim','2026-02-13','Resíduo comum (lixo) (produção e ADM)','261029533770',0.0,1800.0,900.0,900.0],
  ['Belo Jardim','2026-02-16','Resíduo comum (lixo) (produção e ADM)','261029553390',0.0,1080.0,540.0,540.0],
  ['Belo Jardim','2026-02-17','Resíduo comum (lixo) (produção e ADM)','',0.0,1320.0,660.0,660.0],
  ['Belo Jardim','2026-02-18','Resíduo da UFO (higienização)','',0.0,8940.0,8940.0,0.0],
  ['Belo Jardim','2026-02-19','Resíduo comum (lixo) (produção e ADM)','261029586804',0.0,1940.0,970.0,970.0],
  ['Belo Jardim','2026-02-19','Resíduo da UFO (produção)','261029594204',0.0,7600.0,7600.0,0.0],
  ['Belo Jardim','2026-02-19','Resíduo da peneira (ETE)','261029600704',0.0,3630.0,3630.0,0.0],
  ['Belo Jardim','2026-02-20','Resíduo comum (lixo) (produção e ADM)','261029607661',0.0,2780.0,1390.0,1390.0],
  ['Belo Jardim','2026-02-23','Resíduo comum (lixo) (produção e ADM)','261029636511',0.0,1540.0,770.0,770.0],
  ['Belo Jardim','2026-02-24','Resíduo comum (lixo) (produção e ADM)','261029660756',0.0,2320.0,1160.0,1160.0],
  ['Belo Jardim','2026-02-24','Resíduo da UFO (produção)','261029666463',0.0,14960.0,14960.0,0.0],
  ['Belo Jardim','2026-02-25','Metralha (resíduo da construção civil)','',0.0,7970.0,7970.0,0.0],
  ['Belo Jardim','2026-02-25','Resíduo comum (lixo) (produção e ADM)','',0.0,2080.0,1040.0,1040.0],
  ['Belo Jardim','2026-02-25','Resíduo da UFO (produção)','',0.0,10070.0,10070.0,0.0],
  ['Belo Jardim','2026-02-26','Resíduo comum (lixo) (produção e ADM)','',0.0,1520.0,760.0,760.0],
  ['Belo Jardim','2026-02-26','Cinza (caldeira)','',0.0,13000.0,13000.0,0.0],
  ['Belo Jardim','2026-02-26','Resíduo da peneira (ETE)','',0.0,9220.0,9220.0,0.0],
  ['Belo Jardim','2026-02-27','Resíduo comum (lixo) (produção e ADM)','',0.0,1700.0,850.0,850.0],
  ['Belo Jardim','2026-02-27','Metralha (resíduo da construção civil)','',0.0,6510.0,6510.0,0.0],
  ['Belo Jardim','2026-02-27','Resíduo da UFO (produção)','',0.0,6980.0,6980.0,0.0],
  ['Belo Jardim','2026-02-28','Resíduo comum (lixo) (produção e ADM)','',0.0,1820.0,910.0,910.0],
  ['Belo Jardim','2026-02-28','Resíduo da UFO (produção)','',0.0,9770.0,9770.0,0.0],
  ['Belo Jardim','2026-02-28','Resíduo da UFO (produção)','',0.0,7790.0,7790.0,0.0],
  ['Belo Jardim','2026-03-02','Metralha (resíduo da construção civil)','261029743881',0.0,12650.0,12650.0,0.0],
  ['Belo Jardim','2026-03-02','Resíduo comum (lixo) (produção e ADM)','261029746086',0.0,1370.0,685.0,685.0],
  ['Belo Jardim','2026-03-02','Resíduo da UFO (produção)','',0.0,10580.0,10580.0,0.0],
  ['Belo Jardim','2026-03-02','Papelão','',0.0,2630.0,0.0,2630.0],
  ['Belo Jardim','2026-03-02','Resíduo da UFO (produção)','',0.0,3950.0,3950.0,0.0],
  ['Belo Jardim','2026-03-03','Resíduo comum (lixo) (produção e ADM)','',0.0,2560.0,1280.0,1280.0],
  ['Belo Jardim','2026-03-03','Metralha (resíduo da construção civil)','',0.0,15860.0,15860.0,0.0],
  ['Belo Jardim','2026-03-03','Resíduo da UFO (produção)','',0.0,7660.0,7660.0,0.0],
  ['Belo Jardim','2026-03-03','Resíduo da UFO (produção)','',0.0,9450.0,9450.0,0.0],
  ['Belo Jardim','2026-03-04','Resíduo da UFO (produção)','261029793146',0.0,10380.0,10380.0,0.0],
  ['Belo Jardim','2026-03-04','Resíduo da UFO (produção)','261029793995',0.0,14350.0,14350.0,0.0],
  ['Belo Jardim','2026-03-04','Resíduo da UFO (produção)','',0.0,6140.0,6140.0,0.0],
  ['Belo Jardim','2026-03-04','Resíduo comum (lixo) (produção e ADM)','261029796195',0.0,2590.0,1295.0,1295.0],
  ['Belo Jardim','2026-03-05','Resíduo da UFO (produção)','',0.0,10650.0,10650.0,0.0],
  ['Belo Jardim','2026-03-05','Resíduo comum (lixo) (produção e ADM)','261029836322',0.0,2490.0,1245.0,1245.0],
  ['Belo Jardim','2026-03-06','Resíduo da UFO (produção)','',0.0,15410.0,15410.0,0.0],
  ['Belo Jardim','2026-03-06','Resíduo comum (lixo) (produção e ADM)','261029843062',0.0,1900.0,950.0,950.0],
  ['Belo Jardim','2026-03-07','Resíduo da UFO (produção)','',0.0,18640.0,18640.0,0.0],
  ['Belo Jardim','2026-03-09','Resíduo comum (lixo) (produção e ADM)','261029864143',0.0,1180.0,590.0,590.0],
  ['Belo Jardim','2026-03-09','Metralha (resíduo da construção civil)','261029875114',0.0,13980.0,13980.0,0.0],
  ['Belo Jardim','2026-03-10','Resíduo comum (lixo) (produção e ADM)','261029886362',0.0,2240.0,1120.0,1120.0],
  ['Belo Jardim','2026-03-10','Plástico','261029884201',0.0,2470.0,0.0,2470.0],
  ['Belo Jardim','2026-03-10','Plástico','261029891418',0.0,3080.0,0.0,3080.0],
  ['Belo Jardim','2026-03-10','Plástico','261029891464',0.0,830.0,0.0,830.0],
  ['Belo Jardim','2026-03-10','Resíduo da UFO (produção)','261029897074',0.0,14000.0,14000.0,0.0],
  ['Belo Jardim','2026-03-11','Resíduo comum (lixo) (produção e ADM)','261029909089',0.0,2260.0,1130.0,1130.0],
  ['Belo Jardim','2026-03-12','Resíduo comum (lixo) (produção e ADM)','261029924142',0.0,1750.0,875.0,875.0],
  ['Belo Jardim','2026-03-12','Resíduo da UFO (produção)','261029926276',0.0,8490.0,8490.0,0.0],
  ['Belo Jardim','2026-03-12','Resíduo da peneira (ETE)','261029930907',0.0,7590.0,7590.0,0.0],
  ['Belo Jardim','2026-03-12','Papelão','261029936160',0.0,1870.0,0.0,1870.0],
  ['Belo Jardim','2026-03-13','Resíduo comum (lixo) (produção e ADM)','261029946025',0.0,1940.0,970.0,970.0],
  ['Belo Jardim','2026-03-13','Resíduo da UFO (produção)','',0.0,8560.0,8560.0,0.0],
  ['Belo Jardim','2026-03-16','Resíduo comum (lixo) (produção e ADM)','261029971906',0.0,1570.0,785.0,785.0],
  ['Belo Jardim','2026-03-16','Resíduo da saúde','261029982604',11000.0,0.0,11000.0,0.0],
  ['Belo Jardim','2026-03-17','Resíduo da UFO (produção)','261029992341',0.0,7700.0,7700.0,0.0],
  ['Belo Jardim','2026-03-17','Resíduo comum (lixo) (produção e ADM)','261029993815',0.0,2030.0,1015.0,1015.0],
  ['Belo Jardim','2026-03-18','Cinza (caldeira)','261030013737',0.0,11750.0,11750.0,0.0],
  ['Belo Jardim','2026-03-18','Resíduo comum (lixo) (produção e ADM)','261030015867',0.0,1860.0,930.0,930.0],
  ['Belo Jardim','2026-03-19','Resíduo comum (lixo) (produção e ADM)','261030034136',0.0,1640.0,820.0,820.0],
  ['Belo Jardim','2026-03-19','Resíduo da UFO (produção)','',0.0,5890.0,5890.0,0.0],
  ['Belo Jardim','2026-03-19','Resíduo da UFO (produção)','',0.0,8080.0,8080.0,0.0],
  ['Belo Jardim','2026-03-20','Resíduo da peneira (ETE)','261030039943',0.0,8610.0,8610.0,0.0],
  ['Belo Jardim','2026-03-20','Resíduo da UFO (produção)','',0.0,7180.0,7180.0,0.0],
  ['Belo Jardim','2026-03-20','Resíduo da UFO (produção)','',0.0,10550.0,10550.0,0.0],
  ['Belo Jardim','2026-03-21','Resíduo da UFO (produção)','261030073887',0.0,11480.0,11480.0,0.0],
  ['Belo Jardim','2026-03-21','Resíduo da UFO (produção)','',0.0,10120.0,10120.0,0.0],
  ['Belo Jardim','2026-03-21','Resíduo da UFO (produção)','',0.0,6330.0,6330.0,0.0],
  ['Belo Jardim','2026-03-22','Resíduo da UFO (produção)','',0.0,19780.0,19780.0,0.0],
  ['Belo Jardim','2026-03-23','Resíduo comum (lixo) (produção e ADM)','261030083101',0.0,2050.0,1025.0,1025.0],
  ['Belo Jardim','2026-03-23','Resíduo da UFO (produção)','',0.0,14110.0,14110.0,0.0],
  ['Belo Jardim','2026-03-23','Resíduo da UFO (produção)','',0.0,10250.0,10250.0,0.0],
  ['Belo Jardim','2026-03-24','Plástico','261030104630',0.0,1260.0,0.0,1260.0],
  ['Belo Jardim','2026-03-24','Resíduo comum (lixo) (produção e ADM)','261030107018',0.0,1940.0,970.0,970.0],
  ['Belo Jardim','2026-03-24','Resíduo da UFO (produção)','',0.0,17570.0,17570.0,0.0],
  ['Belo Jardim','2026-03-24','Resíduo da UFO (produção)','',0.0,16990.0,16990.0,0.0],
  ['Belo Jardim','2026-03-25','Resíduo comum (lixo) (produção e ADM)','261030127338',0.0,1910.0,955.0,955.0],
  ['Belo Jardim','2026-03-25','Resíduo da UFO (produção)','261030127058',0.0,11520.0,11520.0,0.0],
  ['Belo Jardim','2026-03-25','Resíduo da UFO (produção)','261030134607',0.0,9950.0,9950.0,0.0],
  ['Belo Jardim','2026-03-25','Papelão','261030139392',0.0,3210.0,0.0,3210.0],
  ['Belo Jardim','2026-03-26','Resíduo comum (lixo) (produção e ADM)','261030147084',0.0,2060.0,1030.0,1030.0],
  ['Belo Jardim','2026-03-26','Resíduo da UFO (produção)','261030146756',0.0,10990.0,10990.0,0.0],
  ['Belo Jardim','2026-03-26','Resíduo da UFO (produção)','261030152301',0.0,9950.0,9950.0,0.0],
  ['Belo Jardim','2026-03-27','Resíduo comum (lixo) (produção e ADM)','261030166842',0.0,1630.0,815.0,815.0],
  ['Belo Jardim','2026-03-27','Plástico','261030171038',0.0,140.0,0.0,140.0],
  ['Belo Jardim','2026-03-27','Resíduo da peneira (ETE)','261030173332',0.0,8360.0,8360.0,0.0],
  ['Belo Jardim','2026-03-28','Resíduo da UFO (produção)','',0.0,13010.0,13010.0,0.0],
  ['Belo Jardim','2026-03-30','Resíduo comum (lixo) (produção e ADM)','',0.0,1810.0,905.0,905.0],
  ['Belo Jardim','2026-03-31','Resíduo da UFO (produção)','',0.0,12630.0,12630.0,0.0],
  ['Belo Jardim','2026-03-31','Resíduo comum (lixo) (produção e ADM)','',0.0,1830.0,915.0,915.0],
  ['Belo Jardim','2026-04-01','Resíduo da UFO (produção)','261030249604',0.0,8740.0,8740.0,0.0],
  ['Belo Jardim','2026-04-01','Resíduo comum (lixo) (produção e ADM)','261030243246',0.0,2270.0,1135.0,1135.0],
  ['Belo Jardim','2026-04-01','Resíduo da UFO (produção)','261030262259',0.0,15980.0,15980.0,0.0],
  ['Belo Jardim','2026-04-02','Resíduo comum (lixo) (produção e ADM)','261030265855',0.0,1870.0,935.0,935.0],
  ['Belo Jardim','2026-04-02','Resíduo da peneira (ETE)','261030270955',0.0,5620.0,5620.0,0.0],
  ['Belo Jardim','2026-04-06','Resíduo comum (lixo) (produção e ADM)','261030296225',0.0,1550.0,775.0,775.0],
  ['Belo Jardim','2026-04-06','Resíduo da UFO (produção)','261030302243',0.0,8360.0,8360.0,0.0],
  ['Belo Jardim','2026-04-06','Resíduo da saúde','',14200.0,0.0,14200.0,0.0],
  ['Belo Jardim','2026-04-06','Metralha (resíduo da construção civil)','261030307899',0.0,10380.0,10380.0,0.0],
  ['Belo Jardim','2026-04-07','Resíduo da UFO (farinha)','',0.0,12920.0,12920.0,0.0],
  ['Belo Jardim','2026-04-07','Resíduo comum (lixo) (produção e ADM)','',0.0,1400.0,700.0,700.0],
  ['Belo Jardim','2026-04-08','Resíduo comum (lixo) (produção e ADM)','',0.0,2420.0,1210.0,1210.0],
  ['Belo Jardim','2026-04-09','Resíduo comum (lixo) (produção e ADM)','261030358455',0.0,1580.0,790.0,790.0],
  ['Belo Jardim','2026-04-09','Resíduo da peneira (ETE)','261030364847',0.0,4005.0,4005.0,0.0],
  ['Belo Jardim','2026-04-09','Lodo (ETE)','261030364766',0.0,4005.0,4005.0,0.0],
  ['Belo Jardim','2026-04-10','Resíduo comum (lixo) (produção e ADM)','',0.0,1910.0,955.0,955.0],
  ['Belo Jardim','2026-04-10','Plástico (baldes)','261030385957',0.0,840.0,0.0,840.0],
  ['Belo Jardim','2026-04-10','Plástico','261030393333',0.0,640.0,0.0,640.0],
  ['Belo Jardim','2026-04-10','Resíduo comum (lixo) (produção e ADM)','261030393571',0.0,1400.0,700.0,700.0],
  ['Belo Jardim','2026-04-13','Resíduo comum (lixo) (produção e ADM)','',0.0,1340.0,670.0,670.0],
  ['Belo Jardim','2026-04-14','Resíduo comum (lixo) (produção e ADM)','261030431566',0.0,2200.0,1100.0,1100.0],
  ['Belo Jardim','2026-04-14','Plástico','261030434079',0.0,1570.0,0.0,1570.0],
  ['Belo Jardim','2026-04-15','Resíduo comum (lixo) (produção e ADM)','261030451538',0.0,1650.0,825.0,825.0],
  ['Belo Jardim','2026-04-16','Resíduo comum (lixo) (produção e ADM)','261030472444',0.0,1660.0,830.0,830.0],
  ['Belo Jardim','2026-04-16','Resíduo da peneira (ETE)','261030477890',0.0,8780.0,8780.0,0.0],
  ['Belo Jardim','2026-04-17','Resíduo comum (lixo) (produção e ADM)','261030494550',0.0,1580.0,790.0,790.0],
  ['Belo Jardim','2026-04-17','Resíduo comum (lixo) (produção e ADM)','261030512237',0.0,1470.0,735.0,735.0],
  ['Belo Jardim','2026-04-20','Resíduo comum (lixo) (produção e ADM)','261030522947',0.0,1330.0,665.0,665.0],
  ['Belo Jardim','2026-04-20','Óleo queimado','261030530557',1260.0,0.0,0.0,1260.0],
  ['Belo Jardim','2026-04-21','Resíduo da saúde','261030540904',2.1,0.0,2.1,0.0],
  ['Belo Jardim','2026-04-21','Resíduo comum (lixo) (produção e ADM)','261030540607',0.0,2570.0,1285.0,1285.0],
  ['Belo Jardim','2026-04-21','Resíduo UFO','261030541568',0.0,6580.0,6580.0,0.0],
  ['Belo Jardim','2026-04-22','Resíduo comum (lixo) (produção e ADM)','261030546377',0.0,1680.0,840.0,840.0],
  ['Belo Jardim','2026-04-22','Papelão','261030560262',0.0,5640.0,0.0,5640.0],
  ['Belo Jardim','2026-04-23','Resíduo comum (lixo) (produção e ADM)','261030569333',0.0,2150.0,1075.0,1075.0],
  ['Belo Jardim','2026-04-23','Resíduo da peneira (ETE)','261030573923',0.0,9580.0,9580.0,0.0],
  ['Belo Jardim','2026-04-24','Resíduo comum (lixo) (produção e ADM)','261030589928',0.0,1760.0,880.0,880.0],
  ['Belo Jardim','2026-04-25','Resíduo comum (lixo) (produção e ADM)','261030607280',0.0,1470.0,735.0,735.0],
  ['Belo Jardim','2026-04-27','Resíduo comum (lixo) (produção e ADM)','261030620104',0.0,1460.0,730.0,730.0],
  ['Belo Jardim','2026-04-27','Plástico','261030623082',0.0,1240.0,0.0,1240.0],
  ['Belo Jardim','2026-04-28','Resíduo comum (lixo) (produção e ADM)','261030643882',0.0,2070.0,1035.0,1035.0],
  ['Belo Jardim','2026-04-28','Lodo (ETE)','261030646612',0.0,5420.0,5420.0,0.0],
  ['Belo Jardim','2026-04-29','Resíduo comum (lixo) (produção e ADM)','261030666525',0.0,1700.0,850.0,850.0],
  ['Belo Jardim','2026-04-30','Resíduo comum (lixo) (produção e ADM)','261030689568',0.0,1550.0,775.0,775.0],
  ['Belo Jardim','2026-04-30','Resíduo da peneira (ETE)','261030697815',0.0,8450.0,8450.0,0.0],
  ['Belo Jardim','2026-05-02','Resíduo comum (lixo) (produção e ADM)','261030709357',0.0,1520.0,760.0,760.0],
  ['Belo Jardim','2026-05-04','Resíduo comum (lixo) (produção e ADM)','261030725453',0.0,1440.0,720.0,720.0],
  ['Belo Jardim','2026-05-04','Cinza (caldeira)','261030731630',0.0,14190.0,14190.0,0.0],
  ['Belo Jardim','2026-05-05','Resíduo comum (lixo) (produção e ADM)','261030752638',0.0,1810.0,905.0,905.0],
  ['Belo Jardim','2026-05-05','Resíduo da UFO (produção)','261030758608',0.0,7790.0,7790.0,0.0],
  ['Belo Jardim','2026-05-06','Resíduo comum (lixo) (produção e ADM)','261030774720',0.0,1890.0,945.0,945.0],
  ['Belo Jardim','2026-05-07','Resíduo comum (lixo) (produção e ADM)','261030795178',0.0,1740.0,870.0,870.0],
  ['Belo Jardim','2026-05-07','Resíduo da peneira (ETE)','261030801352',0.0,4850.0,4850.0,0.0],
  ['Belo Jardim','2026-05-08','Resíduo comum (lixo) (produção e ADM)','261030815786',0.0,2090.0,1045.0,1045.0],
  ['Belo Jardim','2026-05-11','Resíduo comum (lixo) (produção e ADM)','261030843302',0.0,1480.0,740.0,740.0],
  ['Belo Jardim','2026-05-12','Resíduo comum (lixo) (produção e ADM)','261030864721',0.0,1860.0,930.0,930.0],
  ['Pesqueira','2026-01-05','Resíduo comum (lixo) (produção e ADM)','261028904026',0.0,1790.0,895.0,895.0],
  ['Pesqueira','2026-01-20','Resíduo comum (lixo) (produção e ADM)','261029131842',0.0,1780.0,890.0,890.0],
  ['Pesqueira','2026-02-02','Óleo queimado (manutenção)','261029343303',370.0,0.0,0.0,370.0],
  ['Pesqueira','2026-02-05','Resíduo comum (lixo) (produção e ADM)','261029409766',0.0,1850.0,925.0,925.0],
  ['Pesqueira','2026-02-25','Resíduo comum (lixo) (produção e ADM)','',0.0,1610.0,805.0,805.0],
  ['Pesqueira','2026-03-12','Resíduo comum (lixo) (produção e ADM)','261029929562',0.0,1650.0,825.0,825.0],
  ['Pesqueira','2026-04-30','Resíduo comum (lixo) (produção e ADM)','261030692175',0.0,2740.0,1370.0,1370.0],
  ['Garanhuns','2026-01-05','Resíduo comum (lixo) (produção e ADM)','261028904026',0.0,1790.0,895.0,895.0],
  ['Garanhuns','2026-01-20','Resíduo comum (lixo) (produção e ADM)','261029131842',0.0,1780.0,890.0,890.0],
  ['Itapissuma','2026-01-05','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-01-09','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-01-14','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-01-17','Lodo (ETE)','',0.0,18000.0,18000.0,0.0],
  ['Itapissuma','2026-01-17','Plástico','',0.0,980.0,0.0,980.0],
  ['Itapissuma','2026-01-17','Papelão','',0.0,999.0,0.0,999.0],
  ['Itapissuma','2026-01-17','Ferro e aço','',0.0,995.0,0.0,995.0],
  ['Itapissuma','2026-01-19','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-01-23','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-01-23','Cinza (caldeira)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-01-27','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-01-30','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-03','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-06','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-06','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-10','Papelão','',0.0,975.0,0.0,975.0],
  ['Itapissuma','2026-02-10','Plástico','',0.0,990.0,0.0,990.0],
  ['Itapissuma','2026-02-10','Papelão','',0.0,985.0,0.0,985.0],
  ['Itapissuma','2026-02-10','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-13','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-19','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-19','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-26','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-02-28','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-02','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-05','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-08','Cinza (caldeira)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-09','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-11','Plástico','',0.0,995.0,0.0,995.0],
  ['Itapissuma','2026-03-11','Papelão','',0.0,998.0,0.0,998.0],
  ['Itapissuma','2026-03-11','Ferro e aço','',0.0,785.0,0.0,785.0],
  ['Itapissuma','2026-03-13','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-17','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-20','Lodo (ETE)','',0.0,18000.0,18000.0,0.0],
  ['Itapissuma','2026-03-20','Plástico','',0.0,998.0,0.0,998.0],
  ['Itapissuma','2026-03-20','Papelão','',0.0,998.0,0.0,998.0],
  ['Itapissuma','2026-03-20','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-24','Cinza (caldeira)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-24','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-27','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0],
  ['Itapissuma','2026-03-31','Resíduo comum (lixo) (produção e ADM)','',0.0,3000.0,3000.0,0.0]
];

var CABECALHO = ['Unidade','Data','Tipo de Residuo','MTR','Classe I (kg)','Classe II (kg)','Aterro (kg)','Reciclagem (kg)'];

// ============================================================
//  WEB APP — GET
//  .../exec                → importa dados hardcoded (carga inicial)
//  .../exec?acao=dados     → retorna JSON com todos os registros da planilha
//  .../exec?acao=status    → retorna JSON de status
// ============================================================
function doGet(e) {
  var acao = (e && e.parameter && e.parameter.acao) ? e.parameter.acao : 'atualizar';

  // ── Retorna todos os registros da aba "Dados Brutos" como JSON ──
  if (acao === 'dados') {
    try {
      var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
      var aba     = ss.getSheetByName('Dados Brutos');
      var records = [];
      if (aba) {
        var vals = aba.getDataRange().getValues();
        for (var i = 1; i < vals.length; i++) {  // pula cabeçalho
          var r = vals[i];
          if (!r[0]) continue;                    // linha vazia
          // Formata data como YYYY-MM-DD independente do tipo retornado pelo Sheets
          var dv = r[1];
          var ds = '';
          if (dv instanceof Date) {
            var yy = dv.getFullYear();
            var mm = String(dv.getMonth() + 1).padStart(2, '0');
            var dd = String(dv.getDate()).padStart(2, '0');
            ds = yy + '-' + mm + '-' + dd;
          } else {
            // Se já for string "YYYY-MM-DD", mantém; senão tenta converter
            ds = String(dv).substring(0, 10);
          }
          records.push({
            Unidade:    String(r[0]).trim(),
            Data:       ds,
            Tipo:       String(r[2]).trim(),
            MTR:        String(r[3]).trim(),
            Classe_I:   parseFloat(r[4])  || 0,
            Classe_II:  parseFloat(r[5])  || 0,
            Aterro:     parseFloat(r[6])  || 0,
            Reciclagem: parseFloat(r[7])  || 0
          });
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, registros: records }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, erro: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── Status ──
  if (acao === 'status') {
    var ss2  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba2 = ss2.getSheetByName('Dados Brutos');
    var n    = aba2 ? Math.max(0, aba2.getLastRow() - 1) : 0;
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, registros: n, atualizado: new Date().toLocaleString('pt-BR') }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── Ação padrão: importa dados hardcoded e exibe página ──
  try {
    importarDados();
    var html = '<html><body style="font-family:sans-serif;padding:32px;max-width:500px">'
      + '<h2 style="color:#1B7B3A">&#x2705; Dados atualizados!</h2>'
      + '<p><b>' + DADOS.length + ' registros</b> gravados na planilha.</p>'
      + '<ul><li>Dados Brutos</li><li>Belo Jardim &middot; Pesqueira &middot; Garanhuns &middot; Itapissuma</li>'
      + '<li>Resumo Mensal &middot; Resumo por Tipo</li></ul>'
      + '<p style="color:#888;font-size:.85em">Atualizado em: ' + new Date().toLocaleString('pt-BR') + '</p>'
      + '<p><a href="https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '" target="_blank">Abrir planilha &rarr;</a></p>'
      + '</body></html>';
    return HtmlService.createHtmlOutput(html);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:32px">'
      + '<h2 style="color:#C8102E">&#x274C; Erro</h2><pre>' + err.toString() + '</pre></body></html>'
    );
  }
}

// ============================================================
//  WEB APP — POST
//  Recebe um novo registro do frontend e grava na planilha
//  Body JSON: { Unidade, Data, Tipo, MTR, Classe_I, Classe_II, Aterro, Reciclagem }
// ============================================================
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);

    var nomeUnidade = (dados.Unidade || '').trim();
    var novaLinha = [
      nomeUnidade,
      (dados.Data      || '').trim(),
      (dados.Tipo      || '').trim(),
      (dados.MTR       || '').trim(),
      parseFloat(dados.Classe_I)   || 0,
      parseFloat(dados.Classe_II)  || 0,
      parseFloat(dados.Aterro)     || 0,
      parseFloat(dados.Reciclagem) || 0
    ];

    // Grava em "Dados Brutos"
    var abaBrutos = ss.getSheetByName('Dados Brutos');
    if (!abaBrutos) abaBrutos = ss.insertSheet('Dados Brutos');
    abaBrutos.appendRow(novaLinha);

    // Grava na aba da unidade (sem a coluna Unidade) de forma case-insensitive
    var abaUnidade = ss.getSheetByName(nomeUnidade);
    if (!abaUnidade) {
      var abas = ss.getSheets();
      for (var i = 0; i < abas.length; i++) {
        if (abas[i].getName().trim().toLowerCase() === nomeUnidade.toLowerCase()) {
          abaUnidade = abas[i];
          break;
        }
      }
    }
    if (abaUnidade) abaUnidade.appendRow(novaLinha.slice(1));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, mensagem: 'Registro gravado com sucesso.' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, erro: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
//  FUNÇÃO PRINCIPAL — pode ser chamada pelo editor OU pelo doGet
// ============================================================
function importarDados() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  criarAbaDadosBrutos(ss);
  criarAbaPorUnidade(ss, 'Belo Jardim');
  criarAbaPorUnidade(ss, 'Pesqueira');
  criarAbaPorUnidade(ss, 'Garanhuns');
  criarAbaPorUnidade(ss, 'Itapissuma');
  criarAbaResumo(ss);
  criarAbaResumoPorTipo(ss);

  Logger.log('Importacao concluida: ' + DADOS.length + ' registros em ' + new Date().toLocaleString('pt-BR'));
}

// ============================================================
//  ABA: DADOS BRUTOS (todos os registros)
// ============================================================
function criarAbaDadosBrutos(ss) {
  var nome = 'Dados Brutos';
  var aba = obterOuCriarAba(ss, nome);
  aba.clearContents();

  var linhas = [CABECALHO].concat(DADOS);
  aba.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);

  formatarCabecalho(aba, 1, CABECALHO.length, '#1B7B3A');
  formatarTabelaCompleta(aba, linhas.length, CABECALHO.length);
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, CABECALHO.length);

  var rangeNum = aba.getRange(2, 5, DADOS.length, 4);
  rangeNum.setNumberFormat('#,##0.00');

  Logger.log('Aba "' + nome + '" criada com ' + DADOS.length + ' registros.');
}

// ============================================================
//  ABA: POR UNIDADE
// ============================================================
function criarAbaPorUnidade(ss, unidade) {
  var aba = obterOuCriarAba(ss, unidade);
  aba.clearContents();

  var filtrado = DADOS.filter(function(r) { return r[0] === unidade; });
  var cabSemUnidade = CABECALHO.slice(1);
  var dadosSemUnidade = filtrado.map(function(r) { return r.slice(1); });

  var cor = corUnidade(unidade);
  var linhas = [cabSemUnidade].concat(dadosSemUnidade);
  aba.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);

  formatarCabecalho(aba, 1, cabSemUnidade.length, cor);
  formatarTabelaCompleta(aba, linhas.length, cabSemUnidade.length);
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, cabSemUnidade.length);

  if (dadosSemUnidade.length > 0) {
    var rangeNum = aba.getRange(2, 4, dadosSemUnidade.length, 4);
    rangeNum.setNumberFormat('#,##0.00');

    // Linha de totais
    var ultimaLinha = linhas.length + 2;
    aba.getRange(ultimaLinha, 1).setValue('TOTAL').setFontWeight('bold');
    for (var col = 4; col <= 7; col++) {
      var letra = colunaLetra(col);
      aba.getRange(ultimaLinha, col)
        .setFormula('=SUM(' + letra + '2:' + letra + linhas.length + ')')
        .setNumberFormat('#,##0.00')
        .setFontWeight('bold')
        .setBackground('#E8F5EE');
    }
  }

  Logger.log('Aba "' + unidade + '" criada com ' + filtrado.length + ' registros.');
}

// ============================================================
//  ABA: RESUMO MENSAL (por unidade x mês)
// ============================================================
function criarAbaResumo(ss) {
  var nome = 'Resumo Mensal';
  var aba = obterOuCriarAba(ss, nome);
  aba.clearContents();

  var agregado = {};
  DADOS.forEach(function(r) {
    var unidade = r[0];
    var mes     = r[1].substring(0, 7); // "2026-01"
    var key     = unidade + '|' + mes;
    if (!agregado[key]) {
      agregado[key] = { unidade: unidade, mes: mes, classeI: 0, classeII: 0, aterro: 0, reciclagem: 0, registros: 0 };
    }
    agregado[key].classeI    += r[4];
    agregado[key].classeII   += r[5];
    agregado[key].aterro     += r[6];
    agregado[key].reciclagem += r[7];
    agregado[key].registros  += 1;
  });

  var cabResumo = ['Unidade','Mês','Classe I (kg)','Classe II (kg)','Aterro (kg)','Reciclagem (kg)','Total Gerado (kg)','% Reciclagem','Qtd Registros'];
  var linhasResumo = [cabResumo];

  Object.keys(agregado).sort().forEach(function(key) {
    var d = agregado[key];
    var totalGerado = d.classeI + d.classeII;
    var percRec = totalGerado > 0 ? (d.reciclagem / totalGerado * 100) : 0;
    linhasResumo.push([d.unidade, d.mes, d.classeI, d.classeII, d.aterro, d.reciclagem, totalGerado, percRec, d.registros]);
  });

  aba.getRange(1, 1, linhasResumo.length, linhasResumo[0].length).setValues(linhasResumo);
  formatarCabecalho(aba, 1, cabResumo.length, '#1D6FA4');
  formatarTabelaCompleta(aba, linhasResumo.length, cabResumo.length);
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, cabResumo.length);

  if (linhasResumo.length > 1) {
    aba.getRange(2, 3, linhasResumo.length - 1, 5).setNumberFormat('#,##0.00');
    aba.getRange(2, 8, linhasResumo.length - 1, 1).setNumberFormat('0.00"%"');
    aba.getRange(2, 9, linhasResumo.length - 1, 1).setNumberFormat('#,##0');
  }

  Logger.log('Aba "' + nome + '" criada com ' + (linhasResumo.length - 1) + ' linhas.');
}

// ============================================================
//  ABA: RESUMO POR TIPO DE RESÍDUO
// ============================================================
function criarAbaResumoPorTipo(ss) {
  var nome = 'Resumo por Tipo';
  var aba = obterOuCriarAba(ss, nome);
  aba.clearContents();

  var agregado = {};
  DADOS.forEach(function(r) {
    var tipo = r[2];
    if (!agregado[tipo]) {
      agregado[tipo] = { tipo: tipo, classeI: 0, classeII: 0, aterro: 0, reciclagem: 0, registros: 0 };
    }
    agregado[tipo].classeI    += r[4];
    agregado[tipo].classeII   += r[5];
    agregado[tipo].aterro     += r[6];
    agregado[tipo].reciclagem += r[7];
    agregado[tipo].registros  += 1;
  });

  var cabTipo = ['Tipo de Resíduo','Classe I (kg)','Classe II (kg)','Aterro (kg)','Reciclagem (kg)','Total Gerado (kg)','% Reciclagem','Qtd Registros'];
  var linhasTipo = [cabTipo];

  var sorted = Object.values(agregado).sort(function(a, b) {
    return (b.classeI + b.classeII) - (a.classeI + a.classeII);
  });

  sorted.forEach(function(d) {
    var totalGerado = d.classeI + d.classeII;
    var percRec = totalGerado > 0 ? (d.reciclagem / totalGerado * 100) : 0;
    linhasTipo.push([d.tipo, d.classeI, d.classeII, d.aterro, d.reciclagem, totalGerado, percRec, d.registros]);
  });

  aba.getRange(1, 1, linhasTipo.length, linhasTipo[0].length).setValues(linhasTipo);
  formatarCabecalho(aba, 1, cabTipo.length, '#D97706');
  formatarTabelaCompleta(aba, linhasTipo.length, cabTipo.length);
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, cabTipo.length);

  if (linhasTipo.length > 1) {
    aba.getRange(2, 2, linhasTipo.length - 1, 5).setNumberFormat('#,##0.00');
    aba.getRange(2, 7, linhasTipo.length - 1, 1).setNumberFormat('0.00"%"');
    aba.getRange(2, 8, linhasTipo.length - 1, 1).setNumberFormat('#,##0');
  }

  Logger.log('Aba "' + nome + '" criada com ' + sorted.length + ' tipos de resíduo.');
}

// ============================================================
//  UTILITÁRIOS
// ============================================================
function obterOuCriarAba(ss, nome) {
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
  }
  return aba;
}

function formatarCabecalho(aba, linha, numColunas, corHex) {
  var range = aba.getRange(linha, 1, 1, numColunas);
  range.setBackground(corHex)
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setFontSize(10)
       .setHorizontalAlignment('center')
       .setBorder(true, true, true, true, true, true);
}

function formatarTabelaCompleta(aba, numLinhas, numColunas) {
  if (numLinhas <= 1) return;
  var range = aba.getRange(2, 1, numLinhas - 1, numColunas);
  range.setBorder(true, true, true, true, true, true,
    '#E0DDD8', SpreadsheetApp.BorderStyle.SOLID);

  // Zebrado
  for (var i = 2; i <= numLinhas; i++) {
    var cor = (i % 2 === 0) ? '#F7F6F2' : '#FFFFFF';
    aba.getRange(i, 1, 1, numColunas).setBackground(cor);
  }
}

function corUnidade(unidade) {
  var cores = {
    'Belo Jardim': '#1B7B3A',
    'Pesqueira':   '#C8102E',
    'Garanhuns':   '#1D6FA4',
    'Itapissuma':  '#D97706'
  };
  return cores[unidade] || '#555555';
}

function colunaLetra(n) {
  var letra = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    letra = String.fromCharCode(65 + rem) + letra;
    n = Math.floor((n - 1) / 26);
  }
  return letra;
}

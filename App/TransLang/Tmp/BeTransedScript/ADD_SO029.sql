/*
  建立權限架構表 SO029
  Date: 2000.07.21
*/
set heading off
variable str varchar2(80)
exec :str := '<< 建立權限架構表 SO029 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool p:\gird\Log\Add_SO029.log
print str

   DELETE FROM SO029;

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO0000', null, '冠崴有線電視CMIS', '', '' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1000', null, '客服管理', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1100', null, '客戶資料/客服管理', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11001', null, '資料新增', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11002', null, '資料修改', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO110021', null, '修改基本資料', '', 'SO11002', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO110022', null, '修改帳戶資料', '', 'SO11002', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO110023', null, '修改其他資料', '', 'SO11002', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11003', null, '資料註銷', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11005', null, '資料列印', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11006', null, '斷線設定', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1111', null, '裝/加/復機服務', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11111', null, '資料新增', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11112', null, '資料修改', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11113', null, '資料刪除', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11114', null, '資料查看', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11115', null, '資料列印', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11116', null, '派工結案', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1112', null, '維修服務', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11121', null, '資料新增', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11122', null, '資料修改', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11123', null, '資料刪除', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11124', null, '資料查看', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11125', null, '資料列印', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11126', null, '派工結案', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1113', null, '停/拆/移機服務', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11131', null, '資料新增', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11132', null, '資料修改', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11133', null, '資料刪除', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11134', null, '資料查看', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11135', null, '資料列印', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11136', null, '派工結案', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1120', null, '設備項目', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11201', null, '資料新增', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11202', null, '資料修改', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11203', null, '資料刪除', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11204', null, '資料查看', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1131', null, '收費項目設定', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11311', null, '資料新增', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11312', null, '資料修改', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11313', null, '資料刪除', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11314', null, '資料查看', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1132', null, '收費資料瀏覽', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11321', null, '資料新增', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11322', null, '資料修改', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11323', null, '資料作廢', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11324', null, '資料查看', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11325', null, '資料列印', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11326', null, '派工單補單', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1133', null, '歷史收費資料瀏覽', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1141', null, '解碼器設定', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1142', null, '付費頻道設定', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11421', null, '資料新增', '', 'SO1133', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11422', null, '資料修改', '', 'SO1133', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11423', null, '資料刪除', '', 'SO1133', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1150', null, 'PPV', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1161', null, '移機記錄瀏覽', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1162', null, '客戶服務申告', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11621', null, '資料新增', '', 'SO1162', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11622', null, '資料修改', '', 'SO1162', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1200', null, '地址資料管理', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12001', null, '資料新增', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12002', null, '資料修改', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12004', null, '資料查看', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12006', null, '歷史住戶', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300', null, '大樓資料管理', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13001', null, '資料新增', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13002', null, '資料修改', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13004', null, '資料查看', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13005', null, '資料列印', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13006', null, '重新計算大樓戶', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13007', null, '客戶類別調整', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13008', null, '預設合約內容設定', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13009', null, '瀏覽客戶資料', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300A', null, '簽約/續約', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300B', null, '前次合約', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300C', null, '統收改個收', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300D', null, '個收改統收', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E', null, '歷任委員管理', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E1', null, '資料新增', '', 'SO1300E', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E2', null, '資料修改', '', 'SO1300E', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E3', null, '資料刪除', '', 'SO1300E', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F', null, '大樓費率設定', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F1', null, '資料新增', '', 'SO1300F', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F2', null, '資料修改', '', 'SO1300F', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F3', null, '資料刪除', '', 'SO1300F', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1400', null, '銷售點資料管理', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14001', null, '資料新增', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14002', null, '資料修改', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14004', null, '資料查看', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14005', null, '資料列印', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1500', null, '供電戶資料管理', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15001', null, '資料新增', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15002', null, '資料修改', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15003', null, '資料刪除', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15004', null, '資料查看', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15005', null, '資料列印', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15006', null, '計算公式', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO150061', null, '資料新增', '', 'SO15006', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO150062', null, '資料修改', '', 'SO15006', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO150065', null, '資料列印', '', 'SO15006', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15007', null, '列印付款單', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1600', null, '定址系統管理', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1610', null, '解碼器管理', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16101', null, '資料新增', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16102', null, '資料修改', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16103', null, '資料刪除', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16104', null, '資料查看', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1620', null, '解碼器批次設定', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1630', null, 'IPPV使用點數讀取', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1640', null, 'PPV節目表', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1700', null, '服務申告管理', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO17002', null, '資料修改', '', 'SO1700', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO17004', null, '資料查看', '', 'SO1700', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2000', null, '工務管理', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2100', null, '派工單登錄', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2200', null, '派工單查詢/瀏覽', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22001', null, '資料新增', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22002', null, '資料修改', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22003', null, '資料刪除', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22004', null, '資料查看', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22005', null, '資料列印', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2300', null, '派工單/報表列印', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2310', null, '派工單列印', '', 'SO2300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2320', null, '派工報表列印', '', 'SO2300', 'SO2310B.RPT'); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2330', null, '派工執行明細報表列印', '', 'SO2300', 'SO2310C.RPT'); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2400', null, '派工單日結', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2500', null, '故障區域設定', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25001', null, '資料新增', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25002', null, '資料修改', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25003', null, '資料刪除', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25004', null, '資料查看', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25006', null, '維修整批回單', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2600', null, '斷線修復管理', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2700', null, '以拆機單結待收月費單', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3000', null, '計費管理', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3100', null, '出單管理', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3110', null, '本期收費資料產生', '', 'SO3100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3120', null, '本期應收資料查詢列印', '', 'SO3100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3121', null, '本期應收明細', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3122', null, '本期應收彙總', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3123', null, '本期應收區域及道路分析', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3124', null, '本期應收期數及金額分析', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3125', null, '本期出單數及金額查詢', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3130', null, '換單調整', '', 'SO3100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3200', null, '收費單管理', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3210', null, '收費單過入', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3220', null, '印單序號重編', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3230', null, '收費報表查詢', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3240', null, '收費資料查詢/編輯', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32401', null, '資料新增', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32402', null, '資料修改', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32403', null, '資料作廢', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32404', null, '資料查看', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32405', null, '資料列印', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32406', null, '派工單補單', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3250', null, '整批調整', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3251', null, '印單金額調整', '', 'SO3250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3252', null, '收費單整批作廢', '', 'SO3250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3253', null, '收費單換單調整', '', 'SO3250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3260', null, '單據列印', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3261', null, '整批收費單據列印', '', 'SO3260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3262', null, '單筆收費單據列印', '', 'SO3260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3263', null, '收費單據重新列印', '', 'SO3260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3270', null, '轉扣帳磁片資料產生', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3271', null, '轉帳資料(自訂格式)產生', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3272', null, '信用卡扣帳資料產生', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3273', null, '臨櫃代收資料產生', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3274', null, '金資格式轉帳資料產生', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3275', null, 'ATM入帳折讓資料產生', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3300', null, '收款管理', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3310', null, '收費資料登錄', '', 'SO3300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3311', null, '收費單/劃撥單登錄', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3312', null, '轉帳成功磁片(自訂格式)登錄', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3313', null, '信用卡入帳登錄', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3314', null, '臨櫃代收入帳登錄', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3315', null, '金資格式入帳登錄', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3316', null, 'ATM入帳登錄', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO331A', null, '派工單所附收費資料登錄', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3320', null, '收款報表查詢', '', 'SO3300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3400', null, '未收款管理', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3410', null, '收費單未收原因登錄', '', 'SO3400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3420', null, '拆機戶整批作業', '', 'SO3400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4000', null, '帳務管理', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4100', null, '過帳處理', '', 'SO4000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4110', null, '日結帳作業1(含報表)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4120', null, '日結帳作業2(含報表)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4130', null, '發票拋檔(含報表)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4140', null, '快速拋帳作業', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4150', null, 'PPV計費月結(含報表)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4200', null, '整批處理', '', 'SO4000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4210', null, '提列壞帳(含報表)', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4220', null, '搬移過帳已收資料至歷史檔', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4230', null, '刪除拆機戶週期性收費項目資料', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4240', null, '週期性收費項目金額調整', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5000', null, '報表管理', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5100', null, '人員績效統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5200', null, '收費資料統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5300', null, '收視頻道統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5400', null, '客戶設備統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5500', null, '客戶資料統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5600', null, '派工單統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5700', null, '各種原因統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5800', null, '客戶服務統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5900', null, '郵遞標籤統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5A00', null, '輔助分析報表', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5B00', null, '資料檢核報表', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5V00', null, 'PPV統計', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6000', null, '代碼管理', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6100', null, '區組代碼類', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6110', null, '行政區代碼', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61101', null, '資料新增', '', 'SO6110', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61102', null, '資料修改', '', 'SO6110', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61105', null, '資料列印', '', 'SO6110', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6120', null, '服務區代碼', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61201', null, '資料新增', '', 'SO6120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61202', null, '資料修改', '', 'SO6120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61205', null, '資料列印', '', 'SO6120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6130', null, '工程組代碼', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61301', null, '資料新增', '', 'SO6130', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61302', null, '資料修改', '', 'SO6130', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61305', null, '資料列印', '', 'SO6130', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6140', null, '收費區代碼', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61401', null, '資料新增', '', 'SO6140', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61402', null, '資料修改', '', 'SO6140', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61405', null, '資料列印', '', 'SO6140', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6200', null, '類別代碼類', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6210', null, '客戶類別代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62101', null, '資料新增', '', 'SO6210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62102', null, '資料修改', '', 'SO6210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62105', null, '資料列印', '', 'SO6210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6220', null, '裝機類別代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62201', null, '資料新增', '', 'SO6220', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62202', null, '資料修改', '', 'SO6220', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62205', null, '資料列印', '', 'SO6220', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6230', null, '維修申告類別代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62301', null, '資料新增', '', 'SO6230', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62302', null, '資料修改', '', 'SO6230', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62305', null, '資料列印', '', 'SO6230', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6240', null, '停拆移機類別代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62401', null, '資料新增', '', 'SO6240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62402', null, '資料修改', '', 'SO6240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62405', null, '資料列印', '', 'SO6240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6250', null, '電話申告類別代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62501', null, '資料新增', '', 'SO6250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62502', null, '資料修改', '', 'SO6250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62505', null, '資料列印', '', 'SO6250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6260', null, '媒體介紹類別代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62601', null, '資料新增', '', 'SO6260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62602', null, '資料修改', '', 'SO6260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62605', null, '資料列印', '', 'SO6260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6270', null, '銷售點行業類別代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62701', null, '資料新增', '', 'SO6270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62702', null, '資料修改', '', 'SO6270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62705', null, '資料列印', '', 'SO6270', null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6280', null, '促銷代碼', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62801', null, '資料新增', '', 'SO6280', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62802', null, '資料修改', '', 'SO6280', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62805', null, '資料列印', '', 'SO6280', null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6300', null, '原因代碼類', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6310', null, '故障原因代碼', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63101', null, '資料新增', '', 'SO6310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63102', null, '資料修改', '', 'SO6310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63105', null, '資料列印', '', 'SO6310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6320', null, '註銷原因代碼', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63201', null, '資料新增', '', 'SO6320', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63202', null, '資料修改', '', 'SO6320', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63205', null, '資料列印', '', 'SO6320', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6330', null, '未繳費原因代碼', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63301', null, '資料新增', '', 'SO6330', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63302', null, '資料修改', '', 'SO6330', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63305', null, '資料列印', '', 'SO6330', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6340', null, '停拆機原因代碼', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63401', null, '資料新增', '', 'SO6340', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63402', null, '資料修改', '', 'SO6340', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63405', null, '資料列印', '', 'SO6340', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6350', null, '退單原因代碼', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63501', null, '資料新增', '', 'SO6350', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63502', null, '資料修改', '', 'SO6350', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63505', null, '資料列印', '', 'SO6350', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6360', null, '短收原因代碼', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63601', null, '資料新增', '', 'SO6360', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63602', null, '資料修改', '', 'SO6360', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63605', null, '資料列印', '', 'SO6360', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6400', null, '編號代碼類', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6410', null, '街道編號代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64101', null, '資料新增', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64102', null, '資料修改', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64104', null, '資料查看', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64105', null, '資料列印', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64106', null, '街道移位', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64107', null, '街道對調', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64108', null, '街道統計', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6420', null, '銀行編號代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64201', null, '資料新增', '', 'SO6420', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64202', null, '資料修改', '', 'SO6420', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64205', null, '資料列印', '', 'SO6420', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6430', null, '收費項目代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64301', null, '資料新增', '', 'SO6430', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64302', null, '資料修改', '', 'SO6430', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64305', null, '資料列印', '', 'SO6430', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6440', null, '付費意願代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64401', null, '資料新增', '', 'SO6440', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64402', null, '資料修改', '', 'SO6440', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64405', null, '資料列印', '', 'SO6440', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6450', null, '建築型態代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64501', null, '資料新增', '', 'SO6450', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64502', null, '資料修改', '', 'SO6450', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64505', null, '資料列印', '', 'SO6450', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6460', null, '品名編號代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64601', null, '資料新增', '', 'SO6460', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64602', null, '資料修改', '', 'SO6460', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64605', null, '資料列印', '', 'SO6460', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6470', null, '巷弄編號代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64701', null, '資料新增', '', 'SO6470', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64702', null, '資料修改', '', 'SO6470', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64705', null, '資料列印', '', 'SO6470', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6480', null, '頻道編號代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64801', null, '資料新增', '', 'SO6480', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64802', null, '資料修改', '', 'SO6480', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64805', null, '資料列印', '', 'SO6480', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6490', null, '客戶來源代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64901', null, '資料新增', '', 'SO6490', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64902', null, '資料修改', '', 'SO6490', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64905', null, '資料列印', '', 'SO6490', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A0', null, '服務滿意度代碼', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A01', null, '資料新增', '', 'SO64A0', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A02', null, '資料修改', '', 'SO64A0', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A05', null, '資料列印', '', 'SO64A0', null); 
/*
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6500', null, '其他內控類', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6510', null, '收費方式代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6520', null, '付款種類代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6530', null, '稅率代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6540', null, '買賣方式代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6550', null, '客戶狀態代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6560', null, '派工狀態代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6570', null, '信用卡種類代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6580', null, '公司別代碼', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6590', null, '資料庫代碼', '', 'SO6500', null); 
*/
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6600', null, '共用小檔維護', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6610', null, '會計科目小檔', '', 'SO6600', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66101', null, '資料新增', '', 'SO6610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66102', null, '資料修改', '', 'SO6610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6620', null, '物料編號小檔', '', 'SO6600', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66201', null, '資料新增', '', 'SO6620', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66202', null, '資料修改', '', 'SO6620', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6630', null, '人事資料小檔', '', 'SO6600', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66301', null, '資料新增', '', 'SO6630', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66302', null, '資料修改', '', 'SO6630', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7000', null, '系統管理', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7200', null, '使用者管理', '', 'SO7000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7210', null, '使用者新增刪改', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72101', null, '資料新增', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72102', null, '資料修改', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72103', null, '資料刪除', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72105', null, '資料列印', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7220', null, '使用者權限設定', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7230', null, '使用者權限複製', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7240', null, '使用者權限報表列印', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7300', null, '系統參數設定', '', 'SO7000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7310', null, '業者參數設定', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7320', null, '派工參數設定', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7330', null, '預設收費參數設定', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7340', null, '其他預設參數設定', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7350', null, '輸出參數設定', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7360', null, '預約派工時段設定', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7400', null, '地址對應資料管理', '', 'SO7000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74001', null, '資料新增', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74002', null, '資料修改', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74003', null, '資料刪除', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74004', null, '資料查看', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74005', null, '資料列印', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74006', null, '整批設定', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7500', null, '批次作業', '', 'SO7000', null); 
/*
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7600', null, '報表格式檔設定', '', 'SO7000', null); 
*/
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8000', null, '其他', '', 'SO0000', null);
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8100', null, '公佈欄', '', 'SO8000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81001', null, '資料新增', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81002', null, '資料修改', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81003', null, '資料刪除', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81005', null, '資料列印', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8200', null, '切換資料庫', '', 'SO8000', null); 
@\gird\script\add_so029_1;
   update so029 set Group1=1;

-- ************************
-- 新增一筆空白資料於SO039 收費單備註檔
insert into SO039 (Memo1, Memo2) values (null, null);

COMMIT;

spool off
set heading on

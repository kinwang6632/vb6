/*
	建立權限資料
	SELECT MID FROM SO029 WHERE GROUP=user之群組 And COMPCODE=公司別
	若 So18C01 在此Dataset 中,表示有權限
*/
====================== 新的 ====================
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8B00', null, '佣金管理', '', 'SO8000' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8B10', null, '佣金計算之抽佣比例/金額檔', '', 'SO8B00' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8B20', null, '佣金計算', '', 'SO8B00' , null);
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8B30', null, '佣金之特殊參數', '', 'SO8B00' , null);
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8B40', null, '鎖帳', '', 'SO8B00' , null);


====================== 舊的 ====================

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18B0', null, '佣金管理', '', '' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18B1', null, '佣金計算之抽佣比例/金額檔', '', 'SO18B0' , null); 


   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18B2', null, '佣金計算', '', 'SO18B0' , null); 







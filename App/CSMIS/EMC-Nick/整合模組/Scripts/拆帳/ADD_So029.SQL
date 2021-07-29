/*
	建立權限資料
	SELECT MID FROM SO029 WHERE GROUP=user之群組 And COMPCODE=公司別
	若 So18C01 在此Dataset 中,表示有權限
*/

====================== 新的 ====================
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8A00', null, '拆帳管理', '', 'SO8000' , null);
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8A10', null, '拆帳公式代碼', '', 'SO8A00' , null);
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8A20', null, '頻道商代碼', '', 'SO8A00' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8A30', null, '頻道拆帳公式設定', '', 'SO8A00' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8A40', null, '拆帳金額計算', '', 'SO8A00' , null);


====================== 舊的 ====================
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18A0', null, '拆帳管理', '', 'SO1800' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18A1', null, '拆帳公式代碼', '', 'SO18A0' , null); 


   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18A2', null, '頻道商代碼', '', 'SO18A0' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18A3', null, '頻道拆帳公式設定', '', 'SO18A0' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18A4', null, '拆帳金額計算', '', 'SO18A0' , null); 






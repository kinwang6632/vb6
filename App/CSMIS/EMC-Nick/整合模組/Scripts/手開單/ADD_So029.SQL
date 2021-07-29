/*
	建立權限資料
	SELECT MID FROM SO029 WHERE GROUP=user之群組 And COMPCODE=公司別
	若 So18C01 在此Dataset 中,表示有權限
*/
====================== 新的 ====================
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8C00', null, '手開單', '', 'SO8000' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8C10', null, '手開單-領單作業', '', 'SO8C00' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8C20', null, '手開單-單據做廢', '', 'SO8C00' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8C30', null, '手開單-報表列印', '', 'SO8C00' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8C40', null, '手開單-修改手開單號(已日結)', '', 'SO8C00' , null);




====================== 舊的 ====================

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18C0', null, '手開單', '', 'SO1800' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18C01', null, '手開單-領單作業', '', 'SO18C0' , null); 


   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18C02', null, '手開單-單據做廢', '', 'SO18C0' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18C03', null, '手開單-報表列印', '', 'SO18C0' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18C04', null, '手開單-修改手開單號(已日結)', '', 'SO18C0' , null); 






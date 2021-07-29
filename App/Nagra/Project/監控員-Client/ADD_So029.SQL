/*
	建立權限資料
	SELECT MID FROM SO029 WHERE GROUP=user之群組 And COMPCODE=公司別
	若 SO18D01 在此Dataset 中,表示有權限
*/
====================== 新的 ====================
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8D00', null, 'CA監控前端程式', '', 'SO8000' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8D10', null, 'CA監控前端程式-偵測CA', '', 'SO8D00' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8D20', null, 'CA監控前端程式-停止/啟動CA', '', 'SO8D00' , null);


====================== 舊的 ====================
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18D0', null, 'CA監控前端程式', '', 'SO1800' , null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18D01', null, 'CA監控前端程式-偵測CA', '', 'SO18D0' , null); 


   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO18D02', null, 'CA監控前端程式-停止/啟動CA', '', 'SO18D0' , null); 







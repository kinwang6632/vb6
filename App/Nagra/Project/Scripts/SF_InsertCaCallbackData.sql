CREATE OR REPLACE FUNCTION SF_InsertCaCallbackData
  (p_CallbackData Varchar2, p_RetCode Out Number, p_RetMsg Out Varchar2)
  RETURN  number AS
/*
--@D:\APP\Nagra\Project\Scripts\SF_InsertCaCallbackData.sql;
VARIABLE RetMsg VARCHAR2(4000)
VARIABLE RetCode number
VARIABLE ReturnNo NUMBER
EXEC :ReturnNo := SF_InsertCaCallbackData('600011756040257000344289200302181160052756020500000000065536123456          999999999999999999999999999999990000000055528509',:RetCode ,:RetMsg)
PRINT ReturnNo
PRINT RetCode
PRINT RetMsg


  檔名: SF_InsertCaCallbackData.sql

  說明: 將 ca command gateway 收到的 callback data 存到 table 中


  IN  
	p_CallbackData			Varchar2 : Callback data

  OUT 
	p_RetMsg        		varchar2: 結果訊息 (至少200 bytes)
	p_RetCode       		number: 結果代碼
  Return 
	number: 處理結果代碼, 對應訊息存於p_RetMsg中
              0: p_RetMsg='insert callback data 完成'
             -1: p_RetMsg='參數錯誤'
            -99: p_RetMsg='其他錯誤'


    By: Howard
    Date: 2003.02.20
*/

  s_TransactionNum Varchar2(9);

BEGIN
--檢查所有參數皆要有值 => 所有參數應該在呼叫端就檢查

    if p_CallbackData Is Null then
      p_RetMsg := '參數錯誤:指令序號不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;



    begin
      s_TransactionNum := SubStr(p_CallbackData,1,9);

      insert into CaCallbackData(TRANSACTIONNUM,CALLBACKDATA,PROCESSINGSTATE,UPDTIME) values(s_TransactionNum,p_CallbackData,'N',sysdate); 	

    EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'INSERT CaCallbackData 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert callback data 完成';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
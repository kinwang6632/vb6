CREATE OR REPLACE FUNCTION SF_InsertCmdResponseLog
  (p_FullCmdResponseText Varchar2, p_RetCode Out Number, p_RetMsg Out Varchar2)
  RETURN  number AS
/*
--@D:\APP\Nagra\Project\Scripts\SF_InsertCmdResponseLog.sql;
VARIABLE RetMsg VARCHAR2(4000)
VARIABLE RetCode number
VARIABLE ReturnNo NUMBER
EXEC :ReturnNo := SF_InsertCmdResponseLog('000000001050257000344289200302101000020000001000000000000000000000000',:RetCode ,:RetMsg)
PRINT ReturnNo
PRINT RetCode
PRINT RetMsg


  檔名: SF_InsertCmdResponseLog.sql

  說明: 將 ca command gateway 要傳給 Nagra 的 Cmd Text 存到 table 中


  IN  

	p_FullCmdResponseText				Varchar2 : 傳給 Nagra 的, 完整的指令字串	

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

  s_NagraTransactionNum Varchar2(9);
  s_TransactionNum Varchar2(9);
  s_CmdResponseID Varchar2(4);

BEGIN
--檢查所有參數皆要有值 => 所有參數應該在呼叫端就檢查

    if p_FullCmdResponseText Is Null then
      p_RetMsg := '參數錯誤:指令內容不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;


    begin

      s_NagraTransactionNum := SubStr(p_FullCmdResponseText,1,9);
      s_CmdResponseID := SubStr(p_FullCmdResponseText,33,4);
      s_TransactionNum := SubStr(p_FullCmdResponseText,37,9);


      insert into CaCmdResponseLog(NAGRATRANSACTIONNUM,TRANSACTIONNUM,FULLRESPONSETEXT,CMDRESPONSEID,UPDTIME) values(s_NagraTransactionNum,s_TransactionNum,p_FullCmdResponseText,s_CmdResponseID,sysdate); 	

    EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'INSERT CaCmdResponseLog 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert command response log 完成';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
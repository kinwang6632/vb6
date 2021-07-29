CREATE OR REPLACE FUNCTION SF_InsertSendCmdLog
  (p_HighLevelCmdID Varchar2, p_FullCmdText Varchar2, p_Operator Varchar2, p_RetCode Out Number, p_RetMsg Out Varchar2)
  RETURN  number AS
/*
--@D:\APP\Nagra\Project\Scripts\SF_InsertSendCmdLog.sql;
VARIABLE RetMsg VARCHAR2(4000)
VARIABLE RetCode number
VARIABLE ReturnNo NUMBER
EXEC :ReturnNo := SF_InsertSendCmdLog('A1','099001001','howard',:RetCode ,:RetMsg)
PRINT ReturnNo
PRINT RetCode
PRINT RetMsg


  檔名: SF_InsertSendCmdLog.sql

  說明: 將 ca command gateway 要傳給 Nagra 的 Cmd Text 存到 table 中


  IN  
	p_HighLevelCmdID			Varchar2 : 高階指令代碼檔
	p_FullCmdText				Varchar2 : 傳給 Nagra 的, 完整的指令字串	

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

    if p_HighLevelCmdID Is Null then
      p_RetMsg := '參數錯誤:高階指令代號號不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;

    if p_FullCmdText Is Null then
      p_RetMsg := '參數錯誤:指令內容不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;


    begin
      s_TransactionNum := SubStr(p_FullCmdText,1,9);


      insert into CaSendCmdLog(HIGH_LEVEL_CMD_ID,TRANSACTIONNUM,FULLCMDTEXT,OPERATOR,UPDTIME) values(p_HighLevelCmdID,s_TransactionNum,p_FullCmdText,p_Operator,sysdate); 	

    EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'INSERT CaSendCmdLog 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert send command log 完成';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
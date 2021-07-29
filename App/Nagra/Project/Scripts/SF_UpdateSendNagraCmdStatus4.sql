CREATE OR REPLACE FUNCTION SF_UpdateSendNagraCmdStatus4
  (p_CmdStatus Varchar2, p_ErrorCode Varchar2,p_ErrorMsg Varchar2,
   p_SeqNo Varchar2, p_RetCode Out Number, p_RetMsg Out Varchar2)
  RETURN  number AS
/*
--@D:\APP\Nagra\Project\Scripts\SF_UpdateSendNagraCmdStatus4.sql;
VARIABLE RetMsg VARCHAR2(4000)
VARIABLE RetCode number
VARIABLE ReturnNo NUMBER
EXEC :ReturnNo := SF_UpdateSendNagraCmdStatus4('E','err-code','err-msg','2',:RetCode ,:RetMsg)
EXEC :ReturnNo := SF_UpdateSendNagraCmdStatus4('C','','','2',:RetCode ,:RetMsg)
EXEC :ReturnNo := SF_UpdateSendNagraCmdStatus4('P','','','2',:RetCode ,:RetMsg)
PRINT ReturnNo
PRINT RetCode
PRINT RetMsg


  更新 Send_Nagra 的指令處理狀態
  檔名: SF_UpdateSendNagraCmdStatus4.sql

  說明:更新 Send_Nagra 的指令處理狀態


  IN  
	p_CmdStatus          		Varchar2 : 指令處理狀態
	p_ErrorCode			Varchar2 : 錯誤訊息代碼
	p_ErrorMsg 			Varchar2 : 錯誤訊息
	p_SeqNo		 		Varchar2 : 指令序號
  OUT 
	p_RetMsg        		varchar2: 結果訊息 (至少200 bytes)
	p_RetCode       		number: 結果代碼
  Return 
	number: 處理結果代碼, 對應訊息存於p_RetMsg中
              0: p_RetMsg='更新指令處理狀態完畢'
             -1: p_RetMsg='參數錯誤'
            -99: p_RetMsg='其他錯誤'


    By: Howard
    Date: 2003.02.15
*/

  s_SQL VARCHAR2(4000);



BEGIN
--檢查所有參數皆要有值 => 所有參數應該在呼叫端就檢查
    if p_CmdStatus Is Null then
      p_RetMsg := '參數錯誤:指令處理狀態不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;


    if p_SeqNo Is Null then
      p_RetMsg := '參數錯誤:指令序號不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;


    if (p_CmdStatus='E')   then
      begin
  	s_SQL := 'UPDATE Send_Nagra SET CMD_STATUS =' || p_CmdStatus || ', ERR_CODE =' || p_ErrorCode|| ', ERR_MSG =' || p_ErrorMsg || ' where SEQNO=' || p_SeqNo;
        UPDATE Send_Nagra SET CMD_STATUS = p_CmdStatus, ERR_CODE = p_ErrorCode, ERR_MSG= p_ErrorMsg WHERE SEQNO = p_SeqNo;
      EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'UPDATE SEND_NAGRA 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM || ' SQL=>' || s_SQL;
          p_RetCode := -99;
          RETURN -99;
      end;
    elsif (p_CmdStatus='P')   then   
      begin
  	s_SQL := 'UPDATE Send_Nagra SET CMD_STATUS =' || '''' || p_CmdStatus || '''' || ', RESENTTIMES=NVL(RESENTTIMES,0)+1  where SEQNO =' || p_SeqNo || ' and CMD_STATUS not in (' || '''' || 'C' || '''' || ',' || '''' || 'E' || '''' || ')';
        EXECUTE IMMEDIATE s_SQL;
      EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'UPDATE SEND_NAGRA 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM || ' SQL=>' || s_SQL;
          p_RetCode := -99;
          RETURN -99;
      end;
    else
      begin
  	s_SQL := 'UPDATE Send_Nagra SET CMD_STATUS =' || p_CmdStatus || ' where SEQNO=' || p_SeqNo;
        UPDATE Send_Nagra SET CMD_STATUS = p_CmdStatus WHERE SEQNO = p_SeqNo and CMD_STATUS<>'E';
      EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'UPDATE SEND_NAGRA 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM || ' SQL=>' || s_SQL;
          p_RetCode := -99;
          RETURN -99;
      end;	
    end if;


    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := '更新指令處理狀態完畢';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
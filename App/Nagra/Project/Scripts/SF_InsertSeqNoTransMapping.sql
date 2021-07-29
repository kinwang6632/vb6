CREATE OR REPLACE FUNCTION SF_InsertSeqNoTransMapping
  (p_SourceID Varchar2, p_SeqNo Varchar2, p_TransNumber Varchar2, p_RetCode Out Number, p_RetMsg Out Varchar2)
  RETURN  number AS
/*
--@D:\APP\Nagra\Project\Scripts\SF_InsertSeqNoTransMapping.sql;
VARIABLE RetMsg VARCHAR2(4000)
VARIABLE RetCode number
VARIABLE ReturnNo NUMBER
EXEC :ReturnNo := SF_InsertSeqNoTransMapping('1','123','009123456',:RetCode ,:RetMsg)
PRINT ReturnNo
PRINT RetCode
PRINT RetMsg


  檔名: SF_InsertSeqNoTransMapping.sql

  說明: 將 ca command gateway 的 SeqNo 與 Transaction Number 的對應關係存到 table 中


  IN  
	p_SeqNo			Varchar2 : sequence number
	p_TransNumber		Varchar2 : transaction number

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

BEGIN
--檢查所有參數皆要有值 => 所有參數應該在呼叫端就檢查

    if p_SeqNo Is Null then
      p_RetMsg := '參數錯誤:Seq No不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;


    if p_TransNumber Is Null then
      p_RetMsg := '參數錯誤:Transaction Number不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;

    if p_SourceID Is Null then
      p_RetMsg := '參數錯誤:Source ID不能為空值';
      p_RetCode := -1; 
      Return -1;
    end if;

    begin
      delete CaSeqNoTransNumMappingData where TRANSACTIONNUM=p_TransNumber;	 


      insert into CaSeqNoTransNumMappingData(SOURCEID,SEQNO,TRANSACTIONNUM,UPDTIME) values(p_SourceID, p_SeqNo,p_TransNumber,sysdate); 	

    EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'Insert CaSeqNoTransNumMappingData 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert CaSeqNoTransNumMappingData 完成';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
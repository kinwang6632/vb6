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


  �ɦW: SF_InsertSendCmdLog.sql

  ����: �N ca command gateway �n�ǵ� Nagra �� Cmd Text �s�� table ��


  IN  
	p_HighLevelCmdID			Varchar2 : �������O�N�X��
	p_FullCmdText				Varchar2 : �ǵ� Nagra ��, ���㪺���O�r��	

  OUT 
	p_RetMsg        		varchar2: ���G�T�� (�ܤ�200 bytes)
	p_RetCode       		number: ���G�N�X
  Return 
	number: �B�z���G�N�X, �����T���s��p_RetMsg��
              0: p_RetMsg='insert callback data ����'
             -1: p_RetMsg='�Ѽƿ��~'
            -99: p_RetMsg='��L���~'


    By: Howard
    Date: 2003.02.20
*/


  s_TransactionNum Varchar2(9);

BEGIN
--�ˬd�Ҧ��ѼƬҭn���� => �Ҧ��Ѽ����Ӧb�I�s�ݴN�ˬd

    if p_HighLevelCmdID Is Null then
      p_RetMsg := '�Ѽƿ��~:�������O�N�������ର�ŭ�';
      p_RetCode := -1; 
      Return -1;
    end if;

    if p_FullCmdText Is Null then
      p_RetMsg := '�Ѽƿ��~:���O���e���ର�ŭ�';
      p_RetCode := -1; 
      Return -1;
    end if;


    begin
      s_TransactionNum := SubStr(p_FullCmdText,1,9);


      insert into CaSendCmdLog(HIGH_LEVEL_CMD_ID,TRANSACTIONNUM,FULLCMDTEXT,OPERATOR,UPDTIME) values(p_HighLevelCmdID,s_TransactionNum,p_FullCmdText,p_Operator,sysdate); 	

    EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'INSERT CaSendCmdLog �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert send command log ����';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
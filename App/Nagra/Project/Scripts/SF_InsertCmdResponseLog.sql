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


  �ɦW: SF_InsertCmdResponseLog.sql

  ����: �N ca command gateway �n�ǵ� Nagra �� Cmd Text �s�� table ��


  IN  

	p_FullCmdResponseText				Varchar2 : �ǵ� Nagra ��, ���㪺���O�r��	

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

  s_NagraTransactionNum Varchar2(9);
  s_TransactionNum Varchar2(9);
  s_CmdResponseID Varchar2(4);

BEGIN
--�ˬd�Ҧ��ѼƬҭn���� => �Ҧ��Ѽ����Ӧb�I�s�ݴN�ˬd

    if p_FullCmdResponseText Is Null then
      p_RetMsg := '�Ѽƿ��~:���O���e���ର�ŭ�';
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
          p_RetMsg := 'INSERT CaCmdResponseLog �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert command response log ����';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
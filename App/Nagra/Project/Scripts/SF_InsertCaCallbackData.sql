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


  �ɦW: SF_InsertCaCallbackData.sql

  ����: �N ca command gateway ���쪺 callback data �s�� table ��


  IN  
	p_CallbackData			Varchar2 : Callback data

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

    if p_CallbackData Is Null then
      p_RetMsg := '�Ѽƿ��~:���O�Ǹ����ର�ŭ�';
      p_RetCode := -1; 
      Return -1;
    end if;



    begin
      s_TransactionNum := SubStr(p_CallbackData,1,9);

      insert into CaCallbackData(TRANSACTIONNUM,CALLBACKDATA,PROCESSINGSTATE,UPDTIME) values(s_TransactionNum,p_CallbackData,'N',sysdate); 	

    EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'INSERT CaCallbackData �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert callback data ����';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
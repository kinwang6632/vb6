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


  �ɦW: SF_InsertSeqNoTransMapping.sql

  ����: �N ca command gateway �� SeqNo �P Transaction Number ���������Y�s�� table ��


  IN  
	p_SeqNo			Varchar2 : sequence number
	p_TransNumber		Varchar2 : transaction number

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

BEGIN
--�ˬd�Ҧ��ѼƬҭn���� => �Ҧ��Ѽ����Ӧb�I�s�ݴN�ˬd

    if p_SeqNo Is Null then
      p_RetMsg := '�Ѽƿ��~:Seq No���ର�ŭ�';
      p_RetCode := -1; 
      Return -1;
    end if;


    if p_TransNumber Is Null then
      p_RetMsg := '�Ѽƿ��~:Transaction Number���ର�ŭ�';
      p_RetCode := -1; 
      Return -1;
    end if;

    if p_SourceID Is Null then
      p_RetMsg := '�Ѽƿ��~:Source ID���ର�ŭ�';
      p_RetCode := -1; 
      Return -1;
    end if;

    begin
      delete CaSeqNoTransNumMappingData where TRANSACTIONNUM=p_TransNumber;	 


      insert into CaSeqNoTransNumMappingData(SOURCEID,SEQNO,TRANSACTIONNUM,UPDTIME) values(p_SourceID, p_SeqNo,p_TransNumber,sysdate); 	

    EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetMsg := 'Insert CaSeqNoTransNumMappingData �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
          p_RetCode := -99;
          RETURN -99;
    end;

    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := 'insert CaSeqNoTransNumMappingData ����';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
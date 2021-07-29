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


  ��s Send_Nagra �����O�B�z���A
  �ɦW: SF_UpdateSendNagraCmdStatus4.sql

  ����:��s Send_Nagra �����O�B�z���A


  IN  
	p_CmdStatus          		Varchar2 : ���O�B�z���A
	p_ErrorCode			Varchar2 : ���~�T���N�X
	p_ErrorMsg 			Varchar2 : ���~�T��
	p_SeqNo		 		Varchar2 : ���O�Ǹ�
  OUT 
	p_RetMsg        		varchar2: ���G�T�� (�ܤ�200 bytes)
	p_RetCode       		number: ���G�N�X
  Return 
	number: �B�z���G�N�X, �����T���s��p_RetMsg��
              0: p_RetMsg='��s���O�B�z���A����'
             -1: p_RetMsg='�Ѽƿ��~'
            -99: p_RetMsg='��L���~'


    By: Howard
    Date: 2003.02.15
*/

  s_SQL VARCHAR2(4000);



BEGIN
--�ˬd�Ҧ��ѼƬҭn���� => �Ҧ��Ѽ����Ӧb�I�s�ݴN�ˬd
    if p_CmdStatus Is Null then
      p_RetMsg := '�Ѽƿ��~:���O�B�z���A���ର�ŭ�';
      p_RetCode := -1; 
      Return -1;
    end if;


    if p_SeqNo Is Null then
      p_RetMsg := '�Ѽƿ��~:���O�Ǹ����ର�ŭ�';
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
          p_RetMsg := 'UPDATE SEND_NAGRA �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM || ' SQL=>' || s_SQL;
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
          p_RetMsg := 'UPDATE SEND_NAGRA �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM || ' SQL=>' || s_SQL;
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
          p_RetMsg := 'UPDATE SEND_NAGRA �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM || ' SQL=>' || s_SQL;
          p_RetCode := -99;
          RETURN -99;
      end;	
    end if;


    COMMIT;
--    ROLLBACK; -- for testing
    p_RetCode := 0;
    p_RetMsg := '��s���O�B�z���A����';
    RETURN 0;
EXCEPTION
    WHEN others THEN
      ROLLBACK;
      p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
      RETURN -99;
END;
/
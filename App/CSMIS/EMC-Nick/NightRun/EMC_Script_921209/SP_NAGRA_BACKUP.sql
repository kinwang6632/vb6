CREATE OR REPLACE PROCEDURE SP_NAGRA_BACKUP(p_RetCode Out Number, p_RetMsg Out Varchar2) IS

/*
  --@C:\SP_NAGRA_BACKUP.sql
    variable p_RetCode number
    variable p_RetMsg varchar2(100)
    exec  SP_NAGRA_BACKUP(:p_RetCode,:p_RetMsg);
    print P_Retcode
    print p_RetMsg

  �ت�:�NSEND_NAGRA��CMD_STATUS='E'�ӥB�T�ѫe����Ƴƥ���SEND_NAGRA_LOG
  �ɦW: SP_NAGRA_BACKUP.sql

  OUT p_RetMsg  varchar2: �B�z���G�ۦ�

      p_RetCode number: �B�z���G�N�X, �����T���s��p_RetMsg��

              0: p_RetMsg='SEND_NAGRA�ƥ�����'
             -1: p_RetMsg='INSERT �ɵo�Ϳ��~'
             -2: P_RetMsg='Delete�ɵo�Ϳ��~'
            -99: p_RetMsg='��L���~'
  �y�{:
	Loop
     �s�W�@����Ʀܦ@�θ�ưϪ�Send_Nagra_LOG��

     �R��SEND_NAGRA�즳�����
 
  Date:2003/03/25
  By:Morris    	       
*/

V_TIME    DATE;
V_COUNT   NUMBER:=0;

Cursor cr_1 is

  select * from send_nagra where cmd_status='E' and (SYSDATE-UPDTIME)>=4;
  
BEGIN

  V_TIME:=SYSDATE;
  
  for cr1_record in cr_1 loop
    begin      
      insert into send_nagra_log(HIGH_LEVEL_CMD_ID,ICC_NO,STB_NO,
                        SUBSCRIPTION_BEGIN_DATE ,SUBSCRIPTION_END_DATE ,
                        NOTES,CMD_STATUS,OPERATOR,UPDTIME,ERR_CODE,ERR_MSG,
                        IMS_PRODUCT_ID,ZIP_CODE,CREDIT_MODE,CREDIT,CREDIT_LIMIT,
                        THRESHOLD_CREDIT,EVENT_NAME,PRICE,CC_NUMBER_1,IP_ADDR,
                        CC_PORT,CALLBACK_DATE,CALLBACK_TIME,CALLBACK_FREQUENCY,
                        FIRST_CALLBACK_DATE,PHONE_NUM_1,PHONE_NUM_2,PHONE_NUM_3,
                        MIS_IRD_CMD_ID,MIS_IRD_CMD_DATA,SEQNO,COMPCODE,RESENTTIMES,LOGDATE)
                          values(cr1_record.HIGH_LEVEL_CMD_ID,cr1_record.ICC_NO,
                                 cr1_record.STB_NO,cr1_record. SUBSCRIPTION_BEGIN_DATE,
                                 cr1_record.SUBSCRIPTION_END_DATE  ,cr1_record.NOTES,
                                 cr1_record.CMD_STATUS,cr1_record.OPERATOR,
                                 cr1_record.UPDTIME,cr1_record.ERR_CODE,cr1_record.ERR_MSG,
                                 cr1_record.IMS_PRODUCT_ID,cr1_record.ZIP_CODE,
                                 cr1_record.CREDIT_MODE,cr1_record.CREDIT,
                                 cr1_record.CREDIT_LIMIT,cr1_record.THRESHOLD_CREDIT,
                                 cr1_record.EVENT_NAME,cr1_record.PRICE,cr1_record.CC_NUMBER_1,
                                 cr1_record.IP_ADDR,cr1_record.CC_PORT,cr1_record.CALLBACK_DATE,
                                 cr1_record.CALLBACK_TIME,cr1_record.CALLBACK_FREQUENCY,
                                 cr1_record.FIRST_CALLBACK_DATE,cr1_record.PHONE_NUM_1,
                                 cr1_record.PHONE_NUM_2,cr1_record.PHONE_NUM_3,
                                 cr1_record. MIS_IRD_CMD_ID,cr1_record.MIS_IRD_CMD_DATA,
                                 cr1_record.SEQNO,cr1_record.COMPCODE,
                                 cr1_record.RESENTTIMES,V_TIME);
      
      V_COUNT:=V_COUNT+1;
      
    exception
      When OTHERS Then
        DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
        p_RetMsg := 'INSERT�ɵo�Ϳ��~: '||SQLERRM;
        p_RetCode := -1;                                
    end;      
    
    begin
      Delete send_nagra where CMD_STATUS='E' and SEQNO=cr1_record.SEQNO;
      
    exception
      When Others Then
        DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
        p_RetMsg := 'Delete�ɵo�Ϳ��~: '||SQLERRM;
        p_RetCode := -2;
    end;
  end loop;
  
  commit;
    p_RetMsg:='SEND_NAGRA�ƥ�����,SEND_NAGRA_LOG �@���� ' || V_count ||' ���ƥ���� ';
    p_RetCode := 0;  
     
EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := '��L���~: '||SQLERRM;
    p_RetCode := -99;
    rollback;
END; -- Procedure

/

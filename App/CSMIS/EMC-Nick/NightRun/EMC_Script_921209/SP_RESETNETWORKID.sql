CREATE OR REPLACE PROCEDURE SP_RESETNETWORKID
   ( p_RetCode OUT NUMBER,
     P_RetMsg  OUT Varchar2) IS
/*
  --@C:\gird\csmis\script\SP_RESETNETWORKID
    variable p_RetCode number
    variable p_RetMsg varchar2(100)
    exec  SP_RESETNETWORKID(:p_RetCode,:p_RetMsg);
    print P_Retcode
    print p_RetMsg
    
    目的: 將新唐城現有使用中的STB, 重新設定其Network id.
    檔名: SP_RESETNETWORKID.sql
    
    OUT p_RetMsg  varchar2: 處理結果自串

    p_RetCode number: 處理結果代碼, 對應訊息存於p_RetMsg中

              0: p_RetMsg='RESTNETWORKID執行完畢'
             -1: p_RetMsg='INSERT至COM.SEND_NAGRA 時發生錯誤'
             -2: P_RetMsg='INSERT至SOAC0202時發生錯誤'                       
            -99: p_RetMsg='其他錯誤'
   
   流程:
   Loop所有使用中的STB 
   新增一筆定時執行的Set network id指令至Com.Send_nagra
   log一筆至STB/ICC設定記錄檔(SOAC0202)
   
Create by Morris
Create Date:2003/04/23           

*/     
     
V_COMPCODE           NUMBER;
V_HIGH_LEVEL_CMD_ID  VARCHAR2(4);
V_CMD_STATUS         CHAR(1);
V_MIS_IRD_CMD_ID     NUMBER;
V_PROCESSINGDATE     DATE;
V_OPERATOR           VARCHAR2(20);
V_UPDTIME            DATE;
V_SEQNO              NUMBER;
V_CNT                NUMBER;
     

CURSOR C1 IS
select CustId, FaciSNo, SmartCardNo
from SO004
where FaciCode=201 and InstDate is not null and PrDate is null;

BEGIN
  V_CNT:=0;
  V_COMPCODE:=13;
  V_HIGH_LEVEL_CMD_ID:='E4';
  V_CMD_STATUS:='W';
  V_MIS_IRD_CMD_ID:=4;
  V_PROCESSINGDATE :=to_date('20030424140000', 'YYYYMMDDHH24MISS');
  V_OPERATOR:='批次作業';
  V_UPDTIME:=SYSDATE;
  
  FOR CR1 IN C1 LOOP
    begin
      SELECT COM.S_SENDNAGRA_SEQNO.NEXTVAL INTO V_SEQNO FROM DUAL;
      
      begin
        INSERT INTO COM.SEND_NAGRA(HIGH_LEVEL_CMD_ID,ICC_NO,CMD_STATUS,OPERATOR,
                                   UPDTIME,IMS_PRODUCT_ID,SEQNO,COMPCODE,PROCESSINGDATE)
                            VALUES(V_HIGH_LEVEL_CMD_ID,CR1.SMARTCARDNO,V_CMD_STATUS,
                                   V_OPERATOR,V_UPDTIME,V_MIS_IRD_CMD_ID,V_SEQNO,V_COMPCODE,
                                   V_PROCESSINGDATE);
        V_CNT:=V_CNT+1;                                          
      exception
        WHEN OTHERS THEN
        DBMS_OUTPUT.put_line('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
        P_RetMsg:='Insert至send_nagra時發生錯誤:'||SQLERRM;
        p_RetCode:=-1;         
      end;
      
      begin
        INSERT INTO SOAC0202(COMPCODE,CUSTID,STBSNO,SMARTCARDNO,MODETYPE,IPPVCREDIT,
                             UPDTIME,UPDEN)
                      VALUES(V_COMPCODE,CR1.CUSTID,CR1.FACISNO,CR1.SMARTCARDNO,V_HIGH_LEVEL_CMD_ID,
                             0,GiPackage.GetDTString2(V_PROCESSINGDATE),V_OPERATOR);                                    
      exception
      WHEN OTHERS THEN
      DBMS_OUTPUT.put_line('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      P_RetMsg:='Insert至SOAC0202時發生錯誤:'||SQLERRM;
      p_RetCode:=-2;     
      end;  
    end;
  END LOOP;
  
  COMMIT;
  
  P_RetMsg:='共新增'||V_CNT||' 筆至SENDNAGRA';
  p_RetCode:=0;   
EXCEPTION
    WHEN OTHERS THEN
       DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM); 
       P_RetMsg:='其他錯誤:'||SQLERRM;
       p_RetCode:=-99;
       rollback;
END; -- Procedure
/

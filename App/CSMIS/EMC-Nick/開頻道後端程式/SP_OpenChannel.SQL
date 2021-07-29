/*
select distinct ICC_NO from com.send_nagra  where operator='批次作業';
select count(*) from com.send_nagra where operator='批次作業';
select cmd_Status, count(*) from com.send_nagra where operator='批次作業' group by cmd_Status;

  delete from com.send_nagra where operator='批次作業';
  --@D:\App\EMC\開頻道後端程式\SP_OpenChannel
  -- @c:\Emc_Script\SP_OpenChannel

    -- 測試指令
    variable p_RetCode number
    variable p_RetMsg varchar2(500)
    exec  SP_OpenChannel(:p_RetCode,:p_RetMsg);
    print P_Retcode
    print p_RetMsg
    
    目的: 為系統台開頻道 (境外頻道: Product Id. = 60)
    檔名: SP_OpenChannel1.sql
    
    OUT p_RetMsg  varchar2: 處理結果自串
	p_RetCode number: 處理結果代碼, 對應訊息存於p_RetMsg中
              0: p_RetMsg='OpenChannel執行完畢'
             -1: p_RetMsg='INSERT至COM.SEND_NAGRA 時發生錯誤'
	     -2: p_RetMsg='取公司代碼時發生錯誤'||SQLERRM
	     -3: p_RetMsg='取頻道代碼/名稱時發生錯誤'||SQLERRM
            -99: p_RetMsg='其他錯誤'
   
   流程:
   Loop所有使用中的STB 
	. 新增一筆定時執行的開頻道指令資料至Com.Send_nagra
	. 若SO005無該筆頻道資料, 則新增一筆
	. log一筆資料至SOAC0202

  Create by Howard
  Create Date:2003/08/01
	2003.08.04 Pierre改: 加上處理SO005, SOAC0202的段落
*/     
CREATE OR REPLACE PROCEDURE SP_OpenChannel
   (p_RetCode OUT NUMBER, P_RetMsg OUT Varchar2) IS
  V_COMPCODE           NUMBER;
  v_ServiceType		char(1);
  V_HIGH_LEVEL_CMD_ID  VARCHAR2(4);
  V_CMD_STATUS         CHAR(1);

  V_PROCESSINGDATE     DATE;
  V_OPERATOR           VARCHAR2(20);
  V_UPDTIME            DATE;
  V_SEQNO              NUMBER;
  V_CNT                NUMBER;
  v_Flag	number;
  v_ChCode	varchar2(3);
  v_ChName	varchar2(20);
  v_SO005Cnt	number;
  v_BeginDate	date;
  v_DueDate	date;
     
  CURSOR C1 IS 
	select CustId, FaciSNo, SmartCardNo from SO004 where FaciCode=201 and SmartCardNo is not null and 
	  (PRDate is null and InstDate is not null);

BEGIN
  -- ************************************************
  -- 設定批次作業參數
  -- ************************************************
  V_CNT:=0;
  V_HIGH_LEVEL_CMD_ID:='B1';
  V_CMD_STATUS:='W';
  V_OPERATOR:='批次作業';
  V_UPDTIME:=SYSDATE;
  v_BeginDate := trunc(V_UPDTIME);			-- Subscription begin date
  v_DueDate := to_date('20101231','YYYYMMDD');		-- Subscription end date

  -- 預定此批次作業執行之當晚23:30:00來開頻道
  V_PROCESSINGDATE := to_date(to_char(trunc(SYSDATE),'YYYYMMDD')||'233000', 'YYYYMMDDHH24MISS');

  -- 取公司代碼, 服務別
  begin 
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501a where rownum<=1;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      p_RetMsg := '取公司代碼時發生錯誤'||SQLERRM;
      p_RetCode := -2;
      commit;
      return;
  end;

  -- 取頻道代碼/名稱
  begin
    select CodeNo, Description into v_ChCode, v_ChName from CD024 where ChannelId='60';
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      p_RetMsg := '取頻道代碼/名稱時發生錯誤'||SQLERRM;
      p_RetCode := -3;
      commit;
      return;
  end;

  -- ************************************************
  -- Loop每一筆正常STB之卡號, 進行以下步驟
  --	1. 新增一筆定時執行的開頻道指令資料至Com.Send_nagra
  --	2. 若SO005無該筆頻道資料, 則新增一筆
  --	3. log一筆資料至SOAC0202
  -- ************************************************
  FOR CR1 IN C1 LOOP
    v_Flag := 0;		-- ok旗標
    v_SO005Cnt := 0;		-- 該卡於SO005中資料筆數

    -- 1. 新增一筆定時執行的開頻道指令資料至Com.Send_nagra
    begin
      insert into com.send_nagra(HIGH_LEVEL_CMD_ID, ICC_NO, CMD_STATUS, NOTES, OPERATOR, SEQNO, COMPCODE, 
	SUBSCRIPTION_BEGIN_DATE, SUBSCRIPTION_END_DATE, PROCESSINGDATE, Stb_No)
	values(V_HIGH_LEVEL_CMD_ID, CR1.SmartCardNo, V_CMD_STATUS, 'A~60', V_OPERATOR, com.S_SENDNAGRA_SEQNO.NEXTVAL,
	V_Compcode, v_BeginDate, v_DueDate, V_PROCESSINGDATE, cr1.FaciSNo);
      V_CNT := V_CNT + 1;
      v_Flag := 1;
    exception
        WHEN OTHERS THEN
        DBMS_OUTPUT.put_line('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
        P_RetMsg:='Insert至send_nagra時發生錯誤:'||SQLERRM;
        p_RetCode:=-1;         
    end;

    if v_Flag = 1 then		-- 若新增指令至Com.Send_nagraok, 才做以下部份
      -- 2. 若SO005無該筆頻道資料, 則新增一筆
      select count(*) into v_SO005Cnt from SO005 where SmartCardNo=cr1.SmartCardNo;
      if v_SO005Cnt = 0 then
	begin
	  Insert into SO005 (CustId, ChCode, Status, SetTime, CvtSNo, DueDate, SetEn, SetName, ServiceType, 
	    CompCode, SmartCardNo) Values 
	    (cr1.CustId, v_ChCode, 1, V_PROCESSINGDATE, cr1.FaciSNo, v_DueDate, V_OPERATOR, V_OPERATOR, v_ServiceType, 
	    v_CompCode, cr1.SmartCardNo);
	exception
	  WHEN OTHERS THEN
	    DBMS_OUTPUT.put_line('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	end;
      end if;

      -- 3. log一筆資料至SOAC0202
      begin
	Insert into SOAC0202 (CompCode, CustId, STBSNo, SmartCardNo, ModeType, ChCode, ChName, 
	  UpdTime, UpdEn, AuthorStartDate, AuthorStopDate) Values 
	  (v_CompCode, cr1.CustId, cr1.FaciSNo, cr1.SmartCardNo, V_HIGH_LEVEL_CMD_ID, v_ChCode, v_ChName, 
	  GiPackage.GetDTString2(V_PROCESSINGDATE), V_OPERATOR, v_BeginDate, v_DueDate);
      exception
	WHEN OTHERS THEN
	  DBMS_OUTPUT.put_line('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      end;
    end if;
  END LOOP;
  COMMIT;

  -- ************************************************
  -- 設定傳回值
  -- ************************************************
  P_RetMsg:='完畢! 共新增'||V_CNT||' 筆至SENDNAGRA';
  p_RetCode:=0;   

EXCEPTION
    WHEN OTHERS THEN
       DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM); 
       P_RetMsg:='其他錯誤:'||SQLERRM;
       p_RetCode:=-99;
       rollback;
END; -- Procedure
/

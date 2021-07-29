/*
	EMC V3.0 Night-run 機制

@C:\Gird\v400\Script\SP_JDailyOp1

exec SP_JDailyOp1

  批次作業__營運日報表: 每日統計客戶裝機/拆機數, MTD, YTD	***** V3.0 *****
  檔名: SF_JDailyOp1.sql

  IN	p_CompCode number:	公司別
	p_ServiceType char(1):	服務類別, 例: 'C', 'I'
	p_StopDate  varchar2:	報表日期(統計截止日), 西曆'YYYYMMDD'
	p_Operator varchar2:	操作者名稱

  Return p_RetMsg VARCHAR2： 結果訊息, 第一個','為return code, 呼叫端變數至少需100 bytes
        0: 正常完畢
	-1: p_RetMsg='參數錯誤'
	-3: p_RetMsg='SQL1錯誤: ' || v_SQL1
        -99：p_RetMsg='其他錯誤'

  By: Pierre
  Date: 2001.08.16
	2001.09.26 改: 傳回值改為字串, 包含return code and return message
	2001.11.19 改: bug, SO093.StartDate --> StopDate
	2001.12.19 改: 自動將產生之資料複製到共用檔(SO501X)中, 以便MSO統計各系統資料
        2001.12.20 改: GICMIS.SO501X相關之語法為EXECUTE IMMEDIATE by Lawrence
	2002.03.14 改: 加上呼叫SP_JDailyOp2 營運日報表3(行政區)
	2002.10.03 改: 加上呼叫SP_JDailyOp4 營運日報表4(行政區/服務區)
	2003.05.19 改: 因應多服務別機制, Night-run參數檔(SO501A)會有多筆資料而做的修改

  By: Lawrence
  Date: 2003.11.18 Mark delete用法,改為Truncate
*/

create or replace procedure SP_JDailyOp1
as
  v_CompCode number;
  v_ServiceType char(1);
  v_StopDate varchar2(8);
  v_Operator varchar2(10);
  v_PRFlag number;
  v_ReturnSQL varchar2(200);
  v_JobId number;

  v_TmpTableName varchar2(20);
  c39	char(1) := chr(39);
  v_RetCode number;
  v_RetMsg varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;
  v_SQL varchar2(1000);

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- 開始執行時間

  -- use so501a to save these values
  --v_CompCode 		:= 3;
  --v_ServiceType 	:= 'C';
  --v_StopDate 		:= '20011117';
  --v_Operator 		:= 'MAINTAIN';
  --v_PRFlag 		:= 1;
  --v_ReturnSQL		:= null;

  v_JobId		:= 999998;
  v_FuncName		:= '營運日報表--資料產生作業';
  v_PrgName		:= 'SP_JDailyOp1';


  -- Loop Night-run參數檔每一筆來執行
  for cr1 in cc1 loop
    DBMS_OUTPUT.PUT_LINE('(1) 開始執行...');
    v_ExecuteCount	:= v_ExecuteCount + 1;	-- 累加執行次數
    v_CompCode		:= cr1.CompCode;
    v_ServiceType	:= cr1.ServiceType;
    v_Operator		:= cr1.Operator;
    v_PRFlag		:= cr1.PRFlag;
    v_ReturnSQL		:= cr1.ReturnSQL;
    v_StopDate		:= cr1.NextDataDay;

    -- ******************************
    -- (1) 呼叫正式的後端程式
    -- ******************************
    v_RetCode := SF_DailyOp1(v_CompCode, v_ServiceType, v_StopDate, to_char(Sysdate, 'YYYYMMDDHH24MI'), 
		v_Operator, v_PRFlag, v_ReturnSQL, v_TmpTableName, v_RetMsg);

    -- ******************************
    -- (2) 根據結果來處理
    -- ******************************
    if v_RetCode >= 0 then 		-- ok
      DBMS_OUTPUT.PUT_LINE('(2) 執行完畢, 儲存執行記錄...');
      -- 先刪除同日的資料(本資料區)
      delete from SO501 where CompCode=v_CompCode and ServiceType=v_ServiceType and 
	RptDate=to_date(v_StopDate, 'YYYYMMDD');
/*
-- *********************************************************************************
-- 若在EMC, 則此段需comment

      -- 先刪除同日的資料(共用資料區)	2001.12.19 added
--    delete from gicmis.SO501X where gicmis.CompCode=v_CompCode and gicmis.ServiceType=v_ServiceType and 
--      gicmis.RptDate=to_date(v_StopDate, 'YYYYMMDD');
      begin
         v_SQL := 'delete from gicmis.SO501X where CompCode='||to_char(v_CompCode)||' and ServiceType='||c39||v_ServiceType||c39||' and RptDate='||
             'to_date('||c39||v_StopDate||c39||', '||c39||'YYYYMMDD'||c39||')';
         EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;
-- *********************************************************************************
*/
      -- 搬移結果至統計檔(本資料區)
      insert into SO501 (CompCode, ServiceType, RptDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime) 
	(select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime from SO093);
/*
-- *********************************************************************************
-- 若在EMC, 則此段需comment
      -- 搬移結果至統計檔(共用資料區)	2001.12.19 added
      begin
        v_SQL :='insert into gicmis.SO501X (CompCode, ServiceType, RptDate, ReportTime, Operator, 
   	   ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	   Value04, Value05, Value06, Value07, Value08, CutoffTime) 
	   (select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	   ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	   Value04, Value05, Value06, Value07, Value08, CutoffTime from SO093)';
        EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;
-- *********************************************************************************
*/
      -- 刪除暫存檔內容
      -- delete from SO093;
      -- 2003.11.18 by Lawrence Mark delete用法,改為Truncate
      DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO093');
    end if;

    -- ******************************
    -- (3) Log into operation log file: so030a
    -- ******************************
    v_StopExecTime := sysdate;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, v_RetCode, v_RetMsg);

    DBMS_OUTPUT.PUT_LINE('(3) 執行完畢, Return message: ' || v_RetMsg);
    commit;


    -- ***************************************************************************
    -- (4) 其他Night-run程式外掛區
    -- ***************************************************************************
    SP_JDailyOp2;		-- 呼叫: 營運日報表2(行政區)
    SP_JDailyOp4;		-- 呼叫: 營運日報表4(各種戶數統計)
    SP_GetSnapShotTableCounts   -- 呼叫: 計算要傳至 SnapShot Table 的筆數
    -- ***************************************************************************


    -- ******************************
    -- (5) 修改下次執行參數 (修改NextDataDay為次一日)
    -- ******************************
    update SO501A set NextDataDay = to_char(to_date(v_StopDate,'YYYYMMDD')+1,'YYYYMMDD')
	where CompCode=v_CompCode and ServiceType=v_ServiceType;

    commit;
  end loop;

  if v_ExecuteCount = 0 then	-- SO501A: 執行參數未設定ok
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, null, null, v_StartExecTime, sysdate, 
	null, v_PrgName, v_FuncName, -99, 'SO501A: 執行參數未設定ok');
      commit;
  end if;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;

    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, -99, '其他錯誤');
    commit;

    --p_RetMsg := SQLERRM;
    --return '-99,'||SQLERRM;
end;
/
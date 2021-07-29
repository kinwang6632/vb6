/*
@C:\Gird\v400\Script\SP_JDailyOp2

  批次作業__營運日報表(行政區): 每日統計客戶裝機/拆機數, MTD, YTD
  檔名: SF_JDailyOp2.sql

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
  Date: 2002.03.12
	2002.04.04 改: 加入行政區名稱欄位
	2003.05.19 改: 因應多服務別機制, Night-run參數檔(SO501A)會有多筆資料而做的修改

  By: Lawrence
  Date: 2003.11.18 Mark delete用法,改為Truncate
*/

create or replace procedure SP_JDailyOp2
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
  v_JobId	:= 999997;
  v_FuncName	:= '營運日報表(2)--資料產生作業';
  v_PrgName	:= 'SP_JDailyOp2';

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
    v_RetCode := SF_DailyOp2(v_CompCode, v_ServiceType, v_StopDate, to_char(Sysdate, 'YYYYMMDDHH24MI'), 
		v_Operator, v_PRFlag, v_ReturnSQL, v_TmpTableName, v_RetMsg);

    -- ******************************
    -- (2) 根據結果來處理
    -- ******************************
    if v_RetCode >= 0 then 		-- ok
      DBMS_OUTPUT.PUT_LINE('(2) 執行完畢, 儲存執行記錄...');

      -- 先刪除同日的資料(本資料區)
      delete from SO502 where CompCode=v_CompCode and ServiceType=v_ServiceType and 
	RptDate=to_date(v_StopDate, 'YYYYMMDD');

      -- 先刪除同日的資料(共用資料區)
      begin
        v_SQL := 'delete from gicmis.SO502X where CompCode='||v_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
		' and RptDate=to_date('||c39||v_StopDate||c39||','||c39||'YYYYMMDD'||c39||')';
        EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;

      -- 搬移結果至統計檔(本資料區)
      insert into SO502 (CompCode, ServiceType, RptDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName) 
	(select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName from SO102);

      -- 搬移結果至統計檔(共用資料區)
      begin
        v_SQL := 'insert into gicmis.SO502X (CompCode, ServiceType, RptDate, ReportTime, Operator, 
	  ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	  Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName) 
	  (select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	  ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	  Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName from SO102)';
        EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;

      -- delete from SO102;
      -- 2003.11.18 by Lawrence Mark delete用法,改為Truncate
      DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO102');
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
/*
@C:\EMC_Script\SP_JDailyOp5

exec SP_JDailyOp5

  批次作業__每月自動做工單日結機制(每月1日02:00一次, 自動呼叫SF_ClsWorkOrder)

  檔名: SP_JDailyOp5.sql 

  By: Pierre
  Date: 2003.04.08 注意: Lawrence已修改此支SF_ClsWorkOrder 成V3.5版, 新版參數未變, 若有變動,則需修改.
	2003.05.05 改: 改成自動結到上月最後一日, 再將SO088日結日加一天
	2003.05.19 改: 因應多服務別機制, Night-run參數檔(SO501A)會有多筆資料而做的修改
	2003.06.09 改: 將原先"再將SO088日結日加一天"還原

*/

create or replace procedure SP_JDailyOp5
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
  v_RetCode number := 0;
  v_RetMsg varchar2(1000) := '無結果';
  v_StartExecTime date;
  v_StopExecTime date;
  v_SQL varchar2(1000);
  v_Now date;

  v_LastClsDate1 date;		-- 上次裝機單日結截止日
  v_LastClsDate2 date;		-- 上次維修單日結截止日
  v_LastClsDate3 date;		-- 上次停拆移機單日結截止日
  v_StartDate	 date;		-- 本次工單日結資料起始日 (date格式)
  v_ClsStartDate varchar2(8);	-- 本次工單日結資料起始日 (varchar2格式)
  v_ClsStopDate  varchar2(8);	-- 本次工單日結資料截止日

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_JobId	:= 999995;
  v_FuncName	:= '(5) 每月自動工單日結機制';
  v_PrgName	:= 'SP_JDailyOp5';

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

    -- ********************************************************
    -- (1) 取執行參數
    -- ********************************************************
    -- 1. 取三種工單的前次日結資料截止日
    begin
      select InstStopTime, MaintainStopTime, PrStopTime into v_LastClsDate1, v_LastClsDate2, 
	v_LastClsDate3 from SO088 where CompCode=v_CompCode and ServiceType=v_ServiceType;
    exception
      when NO_DATA_FOUND then	-- 無日結過, 故以200210101為起始日
        v_ClsStartDate := '20020101';
      when others then	-- 有問題, 故不做此"自動做工單日結機制", 直接跳過
        DBMS_OUTPUT.PUT_LINE(v_FuncName||'--取日結參數錯誤: ' || SQLERRM);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	  v_FuncName, -1, '取日結參數錯誤');
        goto GO_NEXT1;
    end;

    -- 2. 取最早的日期當做此次日結資料起始日
    v_StartDate := least(nvl(v_LastClsDate1,to_date('20020101','YYYYMMDD')), 
		nvl(v_LastClsDate2,to_date('20020101','YYYYMMDD')), nvl(v_LastClsDate3,to_date('20020101','YYYYMMDD')));
    v_ClsStartDate := to_char(v_StartDate, 'YYYYMMDD');	-- 轉成varchar2格式

    -- 3. 取本次工單日結資料截止日 (執行日前月之最後一日)
    --v_ClsStopDate := to_char(sysdate, 'YYYYMM')||'01';	-- (執行日該月之第一日)
    v_ClsStopDate := to_char(last_day(add_months(sysdate,-1)), 'YYYYMMDD');

    -- ********************************************************
    -- (2) 呼叫Lawrence的正式日結函式 SF_ClsWorkOrder:  傳回錯誤碼或者資料批號
    -- ********************************************************
    v_RetCode := SF_ClsWorkOrder(v_ClsStartDate, v_ClsStopDate, 'Night-run', 'Night-run', 
			1, 1, 1, v_CompCode, v_ServiceType, v_RetMsg);

    -- ********************************************************
    -- (3) 處理執行結果, 若日結成功, 再將SO088日結日加一天 (92.06.09取消此項)
    -- ********************************************************
    v_StopExecTime := sysdate;
    if v_RetCode < 0 then
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'--執行SF_ClsWorkOrder錯誤, Return code: ' || v_RetCode);
    else
      -- 若日結成功, 再將SO088日結日加一天 (92.06.09取消此項)
      -- 若日結成功, 更新三種工單日結日期
      update SO088 set InstStopTime=to_date(v_ClsStopDate,'YYYYMMDD'), MaintainStopTime=to_date(v_ClsStopDate,'YYYYMMDD'),
	PrStopTime=to_date(v_ClsStopDate,'YYYYMMDD') where CompCode=v_CompCode and ServiceType=v_ServiceType;
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'--執行SF_ClsWorkOrder完畢, 資料批號: ' || v_RetCode);
    end if;

    -- ********************************************************
    -- (4) Log into operation log file: so030a
    -- ********************************************************
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, v_RetCode, v_RetMsg);

    <<GO_NEXT1>>
    commit;
    NULL;
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
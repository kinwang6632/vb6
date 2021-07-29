/*
@C:\EMC_Script\SP_JDailyOp4

exec SP_JDailyOp4

  批次作業__營運日報表4(行政區): 每日統計各類客戶數
  檔名: SF_JDailyOp4.sql

  營運日報表4統計資料檔: SO504
	SeqNo		number(3)
	CompCode	number(3)
	ServiceType	char(1)
	RptDate		Date
	RptTime		Date
	DataType	number(2)
	ServAreaCode	varchar2(3)
	DataCode	number(3)
	CustCnt		number(8)

  By: Pierre
  Date: 2002.10.03
	2002.10.12 改: EMC改成呼叫3.0版的方式
	2003.03.31 改: 加上呼叫 SF_AdjCustCount()
	2003.04.07 改: 加強程式註解
	2003.05.19 改: 因應多服務別機制, Night-run參數檔(SO501A)會有多筆資料而做的修改
	2003.08.11 改: 拿掉呼叫 SF_AdjCustCount()
*/

create or replace procedure SP_JDailyOp4
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

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_JobId	:= 999996;
  v_FuncName	:= '營運日報表(4)--資料產生作業';
  v_PrgName	:= 'SP_JDailyOp4';

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
    -- (1) 先刪除同日的資料
    -- ******************************
    v_Now := trunc(sysdate);
    delete from SO504 where CompCode=v_CompCode and ServiceType=v_ServiceType and RptDate=v_Now;

    -- ******************************
    -- (2) 營運日報表(4)--資料產生作業, 呼叫正式的後端程式SF_DailyOp4_30
    -- ******************************
    --  v_RetCode := SF_DailyOp4_20(v_CompCode, v_ServiceType, v_Operator, 1, v_RetMsg);
    v_RetCode := SF_DailyOp4_30(v_CompCode, v_ServiceType, v_Operator, 1, v_RetMsg);

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

/*
    -- ********************************************************
    -- (4) 執行Debby的大樓戶統收戶數自動計算程式 SF_AdjCustCount()	-- 2003.03.31
    -- ********************************************************
    v_RetCode := SF_AdjCustCount(v_CompCode);
    if v_RetCode < 0 then
      DBMS_OUTPUT.PUT_LINE('(4) 執行SF_AdjCustCount錯誤, Return code: ' || v_RetCode);
    else
      DBMS_OUTPUT.PUT_LINE('(4) 執行SF_AdjCustCount完畢, 處理筆數: ' || v_RetCode);
    end if;
    commit;
*/
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
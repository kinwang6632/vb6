/*
@c:\emc_script\SP_AdjSO001ToSO002

  批次作業__???
  檔名: SP_AdjSO001ToSO002.sql

  By: Pierre
  Date: 2002.03.12
*/

create or replace procedure SP_AdjSO001ToSO002
as
  v_CompCode number;
  v_ServiceType char(1);
  v_JobId number;

  c39	char(1) := chr(39);
  v_RetCode number;
  v_RetMsg varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_JobId	:= 0;
  v_FuncName	:= '???';
  v_PrgName	:= 'SP_???';

  -- Loop Night-run參數檔每一筆來執行
  for cr1 in cc1 loop
    DBMS_OUTPUT.PUT_LINE('(1) 開始執行...');
    v_ExecuteCount	:= v_ExecuteCount + 1;	-- 累加執行次數
    v_CompCode		:= cr1.CompCode;
    v_ServiceType	:= cr1.ServiceType;

    -- ******************************
    -- (1) 呼叫正式的後端程式
    -- ******************************
    v_RetCode := SF_AdjSO001ToSO002(v_CompCode, v_ServiceType, 1, v_RetMsg);

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
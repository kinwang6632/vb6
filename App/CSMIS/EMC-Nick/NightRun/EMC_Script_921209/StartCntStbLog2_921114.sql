/*
  StartCntStbLog2_921114.sql
  統計每日STB開機次數/成功率/失敗率之Night-run 設定程式 (每日00:30執行)

  注意: 
	1. 若Night-run執行時間不是如上述, 請修改v_Time的時間值, 再重新執行此檔即可

  Date: 2003.08.28
	2003.11.14 修改Night-run執行時間為每日00:30執行

  	By: Pierre
*/

set serveroutput on

VARIABLE jobno number;

declare
  c39 char(1) := chr(39);
  v_Time varchar2(20);
  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);
  v_FuncName varchar2(200);
  v_NextTime varchar2(200);

BEGIN
  -- ***********************************************************************
  -- 設定NIght-run參數
  -- ***********************************************************************
  v_FuncName := '[統計每日STB開機次數/成功率/失敗率(SP_CntStbLog2.sql) Night-run]';
  v_Time := '0030';		-- 執行時間
  v_ExecTime := to_date(to_char(sysdate+1,'YYYYMMDD')||v_Time, 'YYYYMMDDHH24MI'); -- 當時之次日v_Time為第一次執行	
  v_NextTime := 'to_date(to_char(trunc(sysdate)+1,'||c39||'YYYYMMDD'||c39||')||'||chr(39)||v_Time||chr(39)||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  v_What := 'SP_CNTSTBLOG2;';	-- 待執行的指令名稱

  -- 移除同名的night-run job
  begin		
    select Job into v_Job from User_jobs where upper(What)=upper(v_What);
    DBMS_JOB.REMOVE(v_Job); 
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'設定失敗! 錯誤訊息: '||SQLERRM); 
      goto GO_NULL;
  end;

  -- 設定night-run job	[此種寫法確保程式固定是每次的v_Time執行, 而不會有時間偏差]
  DBMS_JOB.SUBMIT(:jobno, upper(v_What), v_ExecTime, v_NextTime);
  --DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, 'sysdate+1');	-- 此種寫法執行時間會偏差

  commit;
  DBMS_OUTPUT.PUT_LINE(v_FuncName||'設定完成');			

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/

/*
  StartMonthlyRptA1_921114.sql
  定義每月自動產生加值頻道單項產品分析資料night-run (呼叫SF_RptA1或SF_RptA1_2) 每月1日01:30執行

  注意: 
	1. 若Night-run執行時間各家不同, 請修改v_Time的時間值, 再重新執行此檔即可

  By: Pierre
  Date: 2003.05.05
	2003.11.14 (1) 改成不會使執行時間偏差的寫法, (2) 修改Night-run執行時間為每月1日01:30

*/
set serveroutput on

VARIABLE jobno number;

declare
  c39 char(1) := chr(39);
  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);
  v_Time varchar2(20);
  v_FuncName varchar2(200);
  v_NextTime varchar2(200);

BEGIN
  -- **************************************************************
  -- 設定參數
  -- **************************************************************
  v_FuncName := '每月自動產生加值頻道單項產品分析資料(呼叫SF_RptA1或SF_RptA1_2) Night-run';
  v_Time := '010130';		-- 執行時間為每月1日01:30
  v_ExecTime := to_date(to_char(add_months(sysdate,1), 'YYYYMM')||v_Time, 'YYYYMMDDHH24MI');	-- 第一次執行時間
  v_NextTime := 'to_date(to_char(add_months(trunc(sysdate), 1),'||c39||'YYYYMM'||c39||')||'||c39||v_Time||c39||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  v_What := 'SP_MONTHLYRPTA1;'; 				-- 待執行程式

  -- 先移除同名之job
  begin		
    select Job into v_Job from User_jobs where upper(What)=upper(v_What);
    DBMS_JOB.REMOVE(v_Job);
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE('設定失敗! 錯誤訊息: '||SQLERRM); 
      goto GO_NULL;
  end;

  -- 定義此job
  DBMS_JOB.SUBMIT(:jobno, upper(v_What), v_ExecTime, v_NextTime);

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



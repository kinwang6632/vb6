/*
  @C:\EMC_Script\StartDailyOp5_921114
  定義EMC每月執行的自動工單日結批次作業 (呼叫SP_JDailyOp5) 每月6日01:00執行
  以月為單位的night-run可參考此檔做法.

  注意: 
	1. 若Night-run執行時間各家不同, 請修改v_Time的時間值, 再重新執行此檔即可

  By: Pierre
  Date: 2003.04.08
	2003.11.14 (1) 改成不會使執行時間偏差的寫法, (2) 修改Night-run執行時間為每月6日01:00
*/

set serveroutput on

VARIABLE jobno number;

declare
  c39 char(1) := chr(39);
  v_CompCode number;			-- 公司別
  v_ServiceType char(1);		-- 服務別
  v_PRFlag number := 1;			-- for CATV SO
  v_NextDataDay varchar2(8);		-- 下次資料參考日
  v_Cnt number;
  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);
  v_Today number;
  v_Time varchar2(20);
  v_FuncName varchar2(200);
  v_NextTime varchar2(200);

BEGIN

  -- **************************************************************
  -- 定義: (5) 批次作業_自動工單日結
  -- *************************************************************
  -- 1. 設定執行參數
  v_FuncName := '自動工單日結批次作業(呼叫SP_JDailyOp5) Night-run';
  v_Time := '060100';	-- 該次執行時間, 格式: 日時分 (若需改變該次執行時間, 請修此變數內容)
  v_ExecTime := to_date(to_char(add_months(sysdate,1), 'YYYYMM')||v_Time, 'YYYYMMDDHH24MI');	-- 第一次執行時間
  v_NextTime := 'to_date(to_char(add_months(trunc(sysdate), 1),'||c39||'YYYYMM'||c39||')||'||c39||v_Time||c39||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  v_What := 'SP_JDailyOp5;';		-- 要night-run的程式名稱

  -- 取目前時間之日, 以便決定是本月執行, 或是在次月執行
  v_Today := to_number(to_char(sysdate, 'DD'));
  if v_Today>=1 and v_Today<=5 then		-- 尚在本月初, 故設定在本月執行
    v_ExecTime := to_date(to_char(sysdate, 'YYYYMM')||v_Time, 'YYYYMMDDHH24MI');
  else						-- 已過月初, 故設定在次月執行
    v_NextTime := 'to_date(to_char(add_months(trunc(sysdate), 1),'||c39||'YYYYMM'||c39||')||'||c39||v_Time||c39||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  end if;

  -- 2. 移除同名之舊night-run程式
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

  -- 3. 定義此night-run程式
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



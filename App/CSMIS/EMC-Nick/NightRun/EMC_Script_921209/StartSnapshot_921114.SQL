/*
 @C:\EMC_Script\StartSnapshot_921114.SQL

  定義EMC每日清晨執行的Snapshot, 呼叫SP_Refresh (各系統台分別於每日02:00~07:00執行)

  注意: 
	1. 若Night-run執行時間各家不同, 請修改v_Time的時間值, 再重新執行此檔即可

  By: Pierre
  Date: 2003.06.20
	2003.11.14 (1) 改成不會使執行時間偏差的寫法, (2) 修改Night-run執行時間為02:00~07:00
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
  -- 定義: EMC每日清晨執行的Snapshot
  -- **************************************************************

  -- 1. 設定執行參數
  v_FuncName := 'EMC每日清晨執行的Snapshot (SP_Refresh) Night-run';
  v_Time := '0400';	-- 該次執行時間, 格式: 時分 (若需改變該次執行時間, 請修此變數內容)
  v_ExecTime := to_date(to_char(sysdate+1, 'YYYYMMDD')||v_Time, 'YYYYMMDDHH24MI');
  v_NextTime := 'to_date(to_char(trunc(sysdate)+1,'||c39||'YYYYMMDD'||c39||')||'||chr(39)||v_Time||chr(39)||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  v_What := 'SP_Refresh;';	-- 要night-run的程式名稱

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



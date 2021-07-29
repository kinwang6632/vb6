/* 
  StartCheckTransform_921225.sql
  92.12.25 23:00執行一次檢查資料庫中代碼異常使用的狀況

  注意: 
	1. 若Night-run執行時間各家不同, 請修改v_Time的時間值, 再重新執行此檔即可

  By: Pierre
  Date: 2003.12.24

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
  -- 定義Night-run參數
  -- *************************************************************
  -- 1. 設定執行參數
  v_FuncName := '檢查資料庫中代碼異常使用的狀況(呼叫SP_CHECKTRANSFORMDATA) Night-run';
  v_Time := '200312252300';		-- 該次執行時間, 格式: 年月日時分 (若需改變該次執行時間, 請修此變數內容)
  v_ExecTime := to_date(v_Time, 'YYYYMMDDHH24MI');	-- 執行時間
  v_NextTime := null;					-- null表示只執行一次
  v_What := 'SP_CHECKTRANSFORMDATA;';			-- 要night-run的程式名稱
	
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

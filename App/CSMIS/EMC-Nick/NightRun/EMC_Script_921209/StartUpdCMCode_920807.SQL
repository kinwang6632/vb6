/*
 @C:\EMC_Script\StartUpdCmCode_920807

  定義EMC收費方式代碼檔(CD031)統一化之資料整批調整程式 StartUpdCmCode_920807.sql
  只於92.08.10 23:30 執行一次

  By: Pierre
  Date: 2003.08.07
*/

set serveroutput on

VARIABLE jobno number;

declare
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

BEGIN

  -- **************************************************************
  -- 定義: 執行參數
  -- **************************************************************
  -- 1. 設定執行參數
  -- 該次執行時間, 格式: 年月日時分 (若需改變該次執行時間, 請修此變數內容)
  v_Time := '200308102300';
  v_ExecTime := to_date(v_Time, 'YYYYMMDDHH24MI');
  v_What := 'SP_UpdCMCode;';		-- 要night-run的程式名稱

  -- 2. 移除同名之舊night-run程式
  begin		
    select Job into v_Job from User_jobs where upper(What)=upper(v_What);
    DBMS_JOB.REMOVE(v_Job); 
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE('定義: 收費方式代碼檔(CD031)統一化之資料整批調整程式失敗! 錯誤訊息: '||SQLERRM);
      goto GO_NULL;
  end;

  -- 3. 定義此night-run程式
  DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, null);	-- null = 只執行一次
  commit;
  DBMS_OUTPUT.PUT_LINE('Night run環境設定完成: 收費方式代碼檔(CD031)統一化之資料整批調整程式');

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/



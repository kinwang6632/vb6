/*
@C:\Gird\csmis\Script\TBC_DailyOp\StartDailyOp_921114

  定義TBC/EMC二支每日晚上執行批次作業: StartDailyOp_911003.sql
	1. 批次作業_刪除暫時性的table及view		00:10執行
	2. 批次作業_產生營運日報表1/2/3/4統計資料	00:20執行

  注意: 
	1. 若Night-run執行時間不是如上述, 請修改兩行v_Time的時間值, 再重新執行此檔即可

  By: Pierre
  Date: 2002.10.03
	2003.11.14 (1) 改成不會使執行時間偏差的寫法, (2) 修改Night-run執行時間為00:10與00:20
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
  v_FuncName varchar2(200);
  v_Time varchar2(20);
  v_NextTime varchar2(200);

BEGIN
  -- ***********************************************************************
  -- 取該系統台之公司別/服務別
  -- ***********************************************************************
  begin
    select CompCode into v_CompCode from SO041 where Rownum=1;
    select substr(ServiceType,1,1) into v_ServiceType from SO041 where Rownum=1;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('讀取SO041公司別/服務別失敗! 錯誤訊息: '||SQLERRM);
      goto GO_NULL;
  end;

  -- ***********************************************************************
  -- 設定產生營運日報表1/2/3/4統計資料所需的程式參數 (存於SO501A)
  -- ***********************************************************************
  v_NextDataDay := to_char(sysdate-1, 'YYYYMMDD');
  begin
    select count(*) into v_Cnt from SO501A;
    if v_Cnt = 0 then
      insert into SO501A (CompCode, ServiceType, Operator, PRFlag, NextDataDay) values 
	(v_CompCode, v_ServiceType, 'MAINTAIN', v_PRFlag, v_NextDataDay);
    else		-- 已有參數檔, 故只更新NextDataDay欄位
      update SO501A set NextDataDay=v_NextDataDay;
    end if;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('設定產生營運日報表1/2/3/4統計資料所需的程式參數失敗! 錯誤訊息: '||SQLERRM);
      goto GO_NULL;
  end;

  -- ***********************************************************************
  -- 定義: 1.批次作業_刪除暫時性的table及view
  -- ***********************************************************************
  v_FuncName := '批次作業_刪除暫時性的table及view(SP_DropTmpTV) Night-run';
  v_Time := '0010';
  v_ExecTime := to_date(to_char(sysdate+1, 'YYYYMMDD')||v_Time, 'YYYYMMDDHH24MI');
  v_NextTime := 'to_date(to_char(trunc(sysdate)+1,'||c39||'YYYYMMDD'||c39||')||'||chr(39)||v_Time||chr(39)||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  v_What := 'SP_DropTmpTV;';

  begin					-- remove old job first
    select Job into v_Job from User_jobs where upper(What)=upper(v_What);
    DBMS_JOB.REMOVE(v_Job); 
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'設定失敗! 錯誤訊息: '||SQLERRM);
      goto GO_NULL;
  end;
  DBMS_JOB.SUBMIT(:jobno, upper(v_What), v_ExecTime, v_NextTime);
  DBMS_OUTPUT.PUT_LINE(v_FuncName||'設定完成');

  -- ***********************************************************************
  -- 定義: 2.批次作業_產生營運日報表1/2/3/4統計資料
  -- ***********************************************************************
  v_FuncName := '批次作業_產生營運日報表1/2/3/4統計資料(SP_JDailyOp1) Night-run';
  v_Time := '0020';
  v_ExecTime := to_date(to_char(sysdate+1, 'YYYYMMDD')||v_Time, 'YYYYMMDDHH24MI');
  v_NextTime := 'to_date(to_char(trunc(sysdate)+1,'||c39||'YYYYMMDD'||c39||')||'||chr(39)||v_Time||chr(39)||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  v_What := 'SP_JDailyOp1;';

  begin					-- remove old job first
    select Job into v_Job from User_jobs where upper(What)=upper(v_What);
    DBMS_JOB.REMOVE(v_Job); 
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'設定失敗! 錯誤訊息: '||SQLERRM);
      goto GO_NULL;
  end;
  DBMS_JOB.SUBMIT(:jobno, upper(v_What), v_ExecTime, v_NextTime);
  DBMS_OUTPUT.PUT_LINE(v_FuncName||'設定完成');

  commit;
  --DBMS_OUTPUT.PUT_LINE('批次作業Night run環境設定完成');

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/



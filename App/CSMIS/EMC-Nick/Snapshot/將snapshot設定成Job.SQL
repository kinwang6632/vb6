/*
@D:\App\EMC\Snapshot\將snapshot設定成Job.SQL

1. 若要停止此 job, 請執行 SQL>execute dbms_job.remove(<job number>);

2. 若要改變此 job 的執行時間點, 請執行 SQL>execute dbms_job.interval(<job number>,'SYSDATE+(1/48)');
   以上表示每半小時執行一次

3. 若要強迫執行(馬上執行)此 job, 請執行 SQL> execute dbms_job.run(<job number>);

*/
set serveroutput on

VARIABLE jobno number;

declare

  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);

BEGIN
/*
	20021001052000 => 表示第一次執行此job 是在 2002/10/01 05:20:00
*/
  v_ExecTime := to_date('20030620020000', 'YYYYMMDDHH24MISS'); 	-- 要修改
  v_What := 'SP_RefreshGroupPackage;'; 				-- 要修改
  begin		-- remove old job first
    select Job into v_Job from User_jobs where What=v_What;
    DBMS_JOB.REMOVE(v_Job); 
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE('snapshot night run環境設定失敗! 錯誤訊息: '||SQLERRM); -- 要修改
      goto GO_NULL;
  end;
/*
	sysdate+7 => 表示以後每 7 天執行一次此 job
*/
  DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, 'sysdate+1'); 			-- 要修改



  commit;
  DBMS_OUTPUT.PUT_LINE('snapshot night run環境設定完成');			-- 要修改

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/



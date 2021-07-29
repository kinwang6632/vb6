set serveroutput on

VARIABLE jobno number;

declare

  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);

BEGIN

  v_ExecTime := to_date('200305010300', 'YYYYMMDDHH24MI'); 	-- 執行時間
  v_What := 'SP_GETSTBDATA;'; 					-- 待執行程式

  begin		
    select Job into v_Job from User_jobs where What=v_What;
    DBMS_JOB.REMOVE(v_Job); 		-- 先移除同名之job
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE('設定失敗! 錯誤訊息: '||SQLERRM); 
      goto GO_NULL;
  end;

  -- 定義此job
  DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, null);	-- 只當時執行一次

  commit;
  DBMS_OUTPUT.PUT_LINE('設定完成');			

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/



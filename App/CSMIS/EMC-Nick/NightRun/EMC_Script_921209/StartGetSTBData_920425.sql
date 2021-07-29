set serveroutput on

VARIABLE jobno number;

declare

  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);

BEGIN

  v_ExecTime := to_date('200305010300', 'YYYYMMDDHH24MI'); 	-- ����ɶ�
  v_What := 'SP_GETSTBDATA;'; 					-- �ݰ���{��

  begin		
    select Job into v_Job from User_jobs where What=v_What;
    DBMS_JOB.REMOVE(v_Job); 		-- �������P�W��job
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE('�]�w����! ���~�T��: '||SQLERRM); 
      goto GO_NULL;
  end;

  -- �w�q��job
  DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, null);	-- �u��ɰ���@��

  commit;
  DBMS_OUTPUT.PUT_LINE('�]�w����');			

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/



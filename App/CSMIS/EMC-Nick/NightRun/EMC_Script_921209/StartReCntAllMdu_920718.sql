-- ����Ҧ��j�ӦU�ؤ�����(SP_ReCntAllMdu.sql) Night-run �]�w�{��
-- Date: 2003.07.18	By: Pierre

set serveroutput on

VARIABLE jobno number;

declare

  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);

BEGIN

  v_ExecTime := to_date('20030718230000', 'YYYYMMDDHH24MISS'); 	
  v_What := 'SP_RECNTALLMDU;'; 				
  begin		
    select Job into v_Job from User_jobs where What=v_What;
    DBMS_JOB.REMOVE(v_Job); 
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE('[����Ҧ��j�Ӧ��Ĥ�(SP_ReCntAllMdu.sql) Night-run] �]�w����! ���~�T��: '||SQLERRM); 
      goto GO_NULL;
  end;

  DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, 'ADD_MONTHS(sysdate,1)');

  commit;
  DBMS_OUTPUT.PUT_LINE('[����Ҧ��j�Ӧ��Ĥ�(SP_ReCntAllMdu.sql) Night-run] �]�w����');			

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/



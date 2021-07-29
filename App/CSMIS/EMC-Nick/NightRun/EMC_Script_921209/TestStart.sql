-- ���ըC��T�w�ɶ��Ұ�night-run�{�����s�g�k
-- �έp�C��STB�}������/���\�v/���Ѳv��Night-run �]�w�{�� (�C02:00����)
-- 92.11.14	By: Pierre

set serveroutput on

VARIABLE jobno number;

declare
  c39 char(1) := chr(39);
  v_Time varchar2(20);
  v_ExecTime date;
  v_NextTime varchar2(200);
  v_Job number;
  v_What varchar2(4000);
  v_FuncName varchar2(200);

BEGIN
  -- �]�wNIght-run�Ѽ�
  v_FuncName := '[�έp�C��STB�}������/���\�v/���Ѳv(SP_CntStbLog2.sql) Night-run]';
  v_Time := '060200';	-- ����ɶ����C��6��02:00
  --v_Time := '0200';		-- �C��02:00
  v_ExecTime := to_date('200311140000', 'YYYYMMDDHH24MI'); -- ��ɥߧY����(�w�L�ɶ�)
  v_What := 'SP_CNTSTBLOG2;';	-- �ݰ��檺���O�W��

  -- �����P�W��night-run job
  begin		
    select Job into v_Job from User_jobs where upper(What)=upper(v_What);
    DBMS_JOB.REMOVE(v_Job); 
  exception
    when no_data_found then
      null;
    when others then
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'�]�w����! ���~�T��: '||SQLERRM); 
      goto GO_NULL;
  end;

  -- �]�wnight-run job [���ؼg�k�T�O�{���T�w�O�C����v_Time����, �Ӥ��|���ɶ����t]
  v_NextTime := 'to_date(to_char(add_months(trunc(sysdate), 1),'||c39||'YYYYMM'||c39||')||'||c39||v_Time||c39||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  --v_NextTime := 'to_date(to_char(trunc(sysdate)+1,'||c39||'YYYYMMDD'||c39||')||'||chr(39)||v_Time||chr(39)||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  --v_NextTime := 'sysdate+1';	-- ���ؼg�k����ɶ��|���t
  DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, v_NextTime);

  commit;
  DBMS_OUTPUT.PUT_LINE(v_FuncName||'�]�w����');			

  <<GO_NULL>>
  NULL;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/

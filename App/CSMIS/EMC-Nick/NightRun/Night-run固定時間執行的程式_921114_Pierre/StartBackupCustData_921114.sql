/* 
  StartBackupCustData_921114.sql
  EMC�C��ƥ����n���(�I�sSP_BackupCustData) �C��6��01:30����

  �`�N: 
	1. �YNight-run����ɶ��U�a���P, �Эק�v_Time���ɶ���, �A���s���榹�ɧY�i

  By: Pierre
  Date: 2003.11.14 (1) �令���|�ϰ���ɶ����t���g�k, (2) �ק�Night-run����ɶ����C��6��01:30

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
  -- �w�qNight-run�Ѽ�
  -- *************************************************************
  -- 1. �]�w����Ѽ�
  v_FuncName := 'EMC�C��ƥ����n���(�I�sSP_BackupCustData) Night-run';
  v_Time := '060130';	-- �Ӧ�����ɶ�, �榡: ��ɤ� (�Y�ݧ��ܸӦ�����ɶ�, �Эצ��ܼƤ��e)
  v_ExecTime := to_date(to_char(add_months(sysdate,1), 'YYYYMM')||v_Time, 'YYYYMMDDHH24MI');	-- �Ĥ@������ɶ�
  v_NextTime := 'to_date(to_char(add_months(trunc(sysdate), 1),'||c39||'YYYYMM'||c39||')||'||c39||v_Time||c39||', '||c39||'YYYYMMDDHH24MI'||c39||')';
  v_What := 'SP_BACKUPCUSTDATA;';		-- �nnight-run���{���W��
	
  -- 2. �����P�W����night-run�{��			
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

  -- 3. �w�q��night-run�{��
  DBMS_JOB.SUBMIT(:jobno, upper(v_What), v_ExecTime, v_NextTime);
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



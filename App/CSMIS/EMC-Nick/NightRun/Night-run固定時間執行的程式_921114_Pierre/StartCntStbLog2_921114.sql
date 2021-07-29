/*
  StartCntStbLog2_921114.sql
  �έp�C��STB�}������/���\�v/���Ѳv��Night-run �]�w�{�� (�C��00:30����)

  �`�N: 
	1. �YNight-run����ɶ����O�p�W�z, �Эק�v_Time���ɶ���, �A���s���榹�ɧY�i

  Date: 2003.08.28
	2003.11.14 �ק�Night-run����ɶ����C��00:30����

  	By: Pierre
*/

set serveroutput on

VARIABLE jobno number;

declare
  c39 char(1) := chr(39);
  v_Time varchar2(20);
  v_ExecTime date;
  v_Job number;
  v_What varchar2(4000);
  v_FuncName varchar2(200);
  v_NextTime varchar2(200);

BEGIN
  -- ***********************************************************************
  -- �]�wNIght-run�Ѽ�
  -- ***********************************************************************
  v_FuncName := '[�έp�C��STB�}������/���\�v/���Ѳv(SP_CntStbLog2.sql) Night-run]';
  v_Time := '0030';		-- ����ɶ�
  v_ExecTime := to_date(to_char(sysdate+1,'YYYYMMDD')||v_Time, 'YYYYMMDDHH24MI'); -- ��ɤ�����v_Time���Ĥ@������	
  v_NextTime := 'to_date(to_char(trunc(sysdate)+1,'||c39||'YYYYMMDD'||c39||')||'||chr(39)||v_Time||chr(39)||', '||c39||'YYYYMMDDHH24MI'||c39||')';
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

  -- �]�wnight-run job	[���ؼg�k�T�O�{���T�w�O�C����v_Time����, �Ӥ��|���ɶ����t]
  DBMS_JOB.SUBMIT(:jobno, upper(v_What), v_ExecTime, v_NextTime);
  --DBMS_JOB.SUBMIT(:jobno, v_What, v_ExecTime, 'sysdate+1');	-- ���ؼg�k����ɶ��|���t

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

/*
  -- compile�P���ի��O
  @c:\emc_script\SP_UpdCMCode
  set serveroutput on
  exec SP_UpdCMCode;

  EMC���O�覡�N�X��(CD031)�Τ@�Ƥ���ƾ��վ�{��, ���{�����ӥu����@��, ��StartUpdCmCode_920807.sql
  �ӱ��WOracle night-run, ����ɶ��w�w��2003.08.10 23:30

  �ɦW: SP_UpdCMCode.sql

  ����: �ھ�EMC����n���Ѥ����O�覡�N�X�s�¹�Ӫ�, �N��Ʈw�����Ψ즬�O�覡�N�X���ɮפ��e
	�վ㦨�s�N�X�P�s�W��, �̫�A�վ㦬�O�覡�N�X��(CD031)���e. 
	�ҽվ㪺�����list, �аѦҦ��q��[�վ㤺�e]����.

  OUT	p_RetCode number: ���G�N�X
	p_RetMsg VARCHAR2�G ���G�T��, �I�s���ܼƦܤֻ�300 bytes

        >0: ��Ƨ帹, ��ܥ��`����
	-1: p_RetMsg='�Ѽƿ��~'
	-2: p_RetMsg='�����q�N�X�ɵo�Ϳ��~'||SQLERRM
        -99�Gp_RetMsg='��L���~'

  *** �ݽ�SO 
	1=�[�@, 2=�̫n, 3=�n��, 5=�s�W�D, 6=�׷�, 7=���D, 8=���p, 9=�����s, 
	10=�s�x�_, 11=���W�D, 12=�j�w��s, 13=�s��, 14=�j�s��, 16=�_���

  *** �վ㤺�e ***
     1. �վ��ɮ�list (�H�U���t���CMCode, CMName)
	SO002, SO003, SO033, SO033DEBT, SO034, SO035, SO036, SO050, SO051, SO061, SO106
     2. ���վ��ɮ�list (���FSO044�u�tCMCode, �H�U���t���CMCode, CMName)
	SO032, SO044, SO067, SO074, SO077
     3. CD031 ���O�覡�N�X�ɤ��e����SO�Τ@, ���s�𫰻P�j�s���O�dCodeNo=9���ӵ����

  *** �`�N: 
     1. �Y����{����Oracle 9i���x��SQL*plus�W����, �]���õLRBS27��rollback segment,
	�h�|�X�{�H�U�����~�T��, �����v�T�{�������`����, ���|�վ���, �Щ����o�ǰT��:
	SQLCODE=-1534 SQLERRM=ORA-01534: �˦^�Ϭq 'RBS27' ���s�b
	SQLCODE=-1453 SQLERRM=ORA-01453: SET TRANSACTION �����O��������Ĥ@�ӱԭz�y
	SQLCODE=-1534 SQLERRM=ORA-01534: �˦^�Ϭq 'RBS27' ���s�b

  By: Pierre
  Date: 2002.08.07
*/

create or replace procedure SP_UpdCMCode
as
  p_RetCode number;
  p_RetMsg varchar2(500);

  v_StartExecTime	date;
  v_EndExecTime		date;
  c39			char(1) := chr(39);
  v_CompCode		number;
  v_ServiceType 	char(1);
  v_TranCnt		number := 0;
  j number;
  v_UpdTime		varchar2(20);
  v_FuncName		varchar2(100);
  v_PrgName		varchar2(100);
  v_JobID	number;

  v_SQL			varchar2(1000);
  v_User		varchar2(100);
  v_TmpRBS		varchar2(100);

  TYPE NumberAry IS TABLE OF number(3) INDEX BY BINARY_INTEGER;
  OldCodeNo NumberAry;		-- �¥N�X array
  NewCodeNo NumberAry;		-- �s�N�X array

  TYPE VC2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  NewCodeName VC2Ary;		-- �s�N�X�W�� array

begin
  v_JobID	:= 0;
  v_FuncName	:= 'EMC���O�覡�N�X��(CD031)�Τ@�Ƥ���ƾ��վ�{��';
  v_PrgName	:= 'SP_UPDCMCODE';
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_UpdTime	:= GiPackage.GetDTString2(v_StartExecTime);
  v_TmpRBS	:= 'RBS27';

  -- ************************************************
  -- �����q�N�X, �A�ȧO
  -- ************************************************
  begin 
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501a where rownum<=1;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      p_RetMsg := '�����q�N�X�ɵo�Ϳ��~'||SQLERRM;
      p_RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(v_EndExecTime-v_StartExecTime)), v_PrgName, v_FuncName, 
	p_RetCode, p_RetMsg);
      commit;
      return;
  end;

  -- ************************************************
  -- �ھڤ��q�O�N�X, �]�w�s�¥N�X������Yarray
  -- ************************************************
  /* 1=�[�@, 2=�̫n, 3=�n��, 5=�s�W�D, 6=�׷�, 7=���D, 8=���p, 9=�����s, 
     10=�s�x�_, 11=���W�D, 12=�j�w��s, 13=�s��, 14=�j�s��, 16=�_��� */
  if v_CompCode=1 or v_CompCode=2 then		-- 1=�[�@, 2=�̫n
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '���H����';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 50;	NewCodeName(2) := '�N��';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 20;	NewCodeName(3) := '�H�Υd';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 45;	NewCodeName(4) := '�˸m��';
    OldCodeNo(5) := 9;	NewCodeNo(5) := 41;	NewCodeName(5) := '��~����';
    OldCodeNo(6) := 10;	NewCodeNo(6) := 49;	NewCodeName(6) := '�����u��';
    OldCodeNo(7) := 11;	NewCodeNo(7) := 44;	NewCodeName(7) := '�u�{�˸m';
    OldCodeNo(8) := 12;	NewCodeNo(8) := 42;	NewCodeName(8) := '�u�{�N��';
    OldCodeNo(9) := 13;	NewCodeNo(9) := 48;	NewCodeName(9) := '�d�r�p��';
    v_TranCnt := 9;
   
  elsif v_CompCode=3 then	-- 3=�n��
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '�N��';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '�H�Υd';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 200;	NewCodeName(3) := '����';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 42;	NewCodeName(4) := '�u�{�N��';
    v_TranCnt := 4;

  elsif v_CompCode=5 then	-- 5=�s�W�D
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '�N��';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '�H�Υd';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 205;	NewCodeName(3) := '���ذӻ�';
    v_TranCnt := 3;

   elsif v_CompCode=6 then	-- 6=�׷�
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '�N��';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '�H�Υd';
    v_TranCnt := 2;

  elsif v_CompCode=7 then	-- 7=���D
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '���H����';
    OldCodeNo(2) := 7;	NewCodeNo(2) := 32;	NewCodeName(2) := '�C�鹺��';
    OldCodeNo(3) := 8;	NewCodeNo(3) := 20;	NewCodeName(3) := '�H�Υd';
    v_TranCnt := 3;

  elsif v_CompCode=8 then	-- 8=���p
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '���H����';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '�H�Υd';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 42;	NewCodeName(3) := '�u�{�N��';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 46;	NewCodeName(4) := '��}���';
    OldCodeNo(5) := 9;	NewCodeNo(5) := 47;	NewCodeName(5) := '�M��';
    OldCodeNo(6) := 10;	NewCodeNo(6) := 43;	NewCodeName(6) := '��������';
    OldCodeNo(7) := 11;	NewCodeNo(7) := 101;	NewCodeName(7) := '7-11';
    OldCodeNo(8) := 12;	NewCodeNo(8) := 102;	NewCodeName(8) := '�ܺ��I';
    OldCodeNo(9) := 13;	NewCodeNo(9) := 103;	NewCodeName(9) := '���a';
    OldCodeNo(10) := 14; NewCodeNo(10) := 201;	NewCodeName(10) := '�l��';
    OldCodeNo(11) := 15; NewCodeNo(11) := 202;	NewCodeName(11) := '�Ĥ@�Ȧ�';
    OldCodeNo(12) := 16; NewCodeNo(12) := 203;	NewCodeName(12) := '�X�@���w';
    OldCodeNo(13) := 17; NewCodeNo(13) := 204;	NewCodeName(13) := '�ɤs�Ȧ�';
    OldCodeNo(14) := 18; NewCodeNo(14) := 104;	NewCodeName(14) := 'OK';
    OldCodeNo(15) := 19; NewCodeNo(15) := 105;	NewCodeName(15) := '�֫Ȧh';
    v_TranCnt := 15;

  elsif v_CompCode=9 or v_CompCode=10 or v_CompCode=11 or v_CompCode=12 then	-- �_���x9~12
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '���H����';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 101;	NewCodeName(2) := '7-11';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 32;	NewCodeName(3) := '�C�鹺��';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 20;	NewCodeName(4) := '�H�Υd';
    v_TranCnt := 4;

  elsif v_CompCode=13 or v_CompCode=14 then	-- 13=�s��, 14=�j�s��
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '�N��';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '�H�Υd';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 100;	NewCodeName(3) := '�K�Q�ө��N��';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 200;	NewCodeName(4) := '�Ȧ�N��';
    OldCodeNo(5) := 10;	NewCodeNo(5) := 21;	NewCodeName(5) := '�H�d���v';
    v_TranCnt := 5;

  elsif v_CompCode=16 then	-- 16=�_���
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '�N��';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '�H�Υd';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 32;	NewCodeName(3) := '�C�鹺��';
    v_TranCnt := 3;

  else
    p_RetMsg := '�L�����q�O: ' || v_CompCode;
    p_RetCode := -2;
    v_TranCnt := 0;
    return;
  end if;


  -- ************************************************************
  -- ����alter rollback segment �v��
  begin
    select user into v_User from dual;
    --DBMS_OUTPUT.PUT_LINE('User = ' || v_User);
    v_SQL := 'GRANT ALTER ROLLBACK SEGMENT TO ' || v_User;
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;

  -- enable rollback segment RBS27
  begin
    v_SQL := 'alter rollback segment '||v_TmpRBS||' online';
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;

  -- ���wtransaction�ܸ�rollback segment
  begin
    set transaction use rollback segment rbS6;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      p_RetMsg := '�L�k�ҥ�RBS27: '||SQLERRM;
  end;
  -- ************************************************************


  -- ************************************************************
  -- �N��Ʈw�����Ψ즬�O�覡�N�X���ɮפ��e, �վ㦨�s�N�X�P�s�W��
  -- ************************************************************
  for j in 1..v_TranCnt loop
    update SO002 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO003 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO033 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO033Debt set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO034 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO035 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO036 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO050 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO051 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO051 set CMCodeB=NewCodeNo(j), CMNameB=NewCodeName(j) where CMCodeB=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO061 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO106 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;


  -- ************************************************************
  -- �վ㦬�O�覡�N�X��(CD031)���e
  --	1. �R�����Ϊ��N�X�q��
  --	2. �s�W�s�@�ΥN�X
  -- ************************************************************
  --	1. �R�����Ϊ��N�X�q��
  if v_CompCode=13 or v_CompCode=14 then	-- 13=�s��, 14=�j�s��
    delete from CD031 where CodeNo>=5 and CodeNo!=9;
  else
    delete from CD031 where CodeNo>=5;
  end if;

  --	2. �s�WCD031�s�@�ΥN�X
  begin
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(20, '�H�Υd', 4, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(21, '�H�d���v', 4, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(30, '�䲼', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(31, '���H����', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(32, '�C�鹺��', null, v_UpdTime, '�t�Φ۰�', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(41, '��~����', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(42, '�u�{�N��', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(43, '��������', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(44, '�u�{�˸m', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(45, '�˸m��', null, v_UpdTime, '�t�Φ۰�', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(46, '��}���', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(47, '�M��', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(48, '�d�r�p��', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(49, '�����u��', null, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(50, '�N��', null, v_UpdTime, '�t�Φ۰�', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(100, '�K�Q�ө��N��', 3, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(101, '7-11', 3, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(102, '�ܺ��I', 3, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(103, '���a', 3, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(104, 'OK', 3, v_UpdTime, '�t�Φ۰�', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(105, '�֫Ȧh', 3, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(200, '�Ȧ�N��', 2, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(201, '�l��', 2, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(202, '�Ĥ@�Ȧ�', 2, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(203, '�X�@���w', 2, v_UpdTime, '�t�Φ۰�', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(204, '�ɤs�Ȧ�', 2, v_UpdTime, '�t�Φ۰�', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(205, '���ذӻ�', 2, v_UpdTime, '�t�Φ۰�', 'C', 0);
  
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      p_RetMsg := '�s�WCD031�s�@�ΥN�X���~: '||SQLERRM;
      p_RetCode := -2;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(v_EndExecTime-v_StartExecTime)), v_PrgName, v_FuncName, 
	p_RetCode, p_RetMsg);
      commit;
      --return; 	-- ���nreturn
  end;


  -- ************************************************************
  v_EndExecTime := sysdate;		-- ���浲���ɶ�
  p_RetMsg := '�վ㧹��, �@��O'||to_char(trunc(86400*(v_EndExecTime-v_StartExecTime)))||'��';
  p_RetCode := 0;
  commit;


  -- ************************************************************
  -- disable��rollback segment
  begin
    v_SQL := 'alter rollback segment '||v_TmpRBS||' offline';
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;

  -- ������alter rollback segment �v��
  begin
    v_SQL := 'REVOKE ALTER ROLLBACK SEGMENT FROM ' || v_User;
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;
  -- ************************************************************


  -- ************************************************************
  -- Log�@�����榨�\�O����SO030A
  -- ************************************************************
  begin
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, v_EndExecTime, 
	trunc(86400*(v_EndExecTime-v_StartExecTime)), v_PrgName, v_FuncName, 
	p_RetCode, p_RetMsg); 
    commit;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      rollback;
      p_RetMsg := 'Log��SO030A���~: '||SQLERRM;
      p_RetCode := -98;
  end;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    p_RetMsg := '��L���~: '||SQLERRM;
    p_RetCode := -99;
end;
/
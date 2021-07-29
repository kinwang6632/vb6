/*
@C:\Gird\v400\Script\SP_JDailyOp2

  �妸�@�~__��B�����(��F��): �C��έp�Ȥ�˾�/�����, MTD, YTD
  �ɦW: SF_JDailyOp2.sql

  IN	p_CompCode number:	���q�O
	p_ServiceType char(1):	�A�����O, ��: 'C', 'I'
	p_StopDate  varchar2:	������(�έp�I���), ���'YYYYMMDD'
	p_Operator varchar2:	�ާ@�̦W��

  Return p_RetMsg VARCHAR2�G ���G�T��, �Ĥ@��','��return code, �I�s���ܼƦܤֻ�100 bytes
        0: ���`����
	-1: p_RetMsg='�Ѽƿ��~'
	-3: p_RetMsg='SQL1���~: ' || v_SQL1
        -99�Gp_RetMsg='��L���~'

  By: Pierre
  Date: 2002.03.12
	2002.04.04 ��: �[�J��F�ϦW�����
	2003.05.19 ��: �]���h�A�ȧO����, Night-run�Ѽ���(SO501A)�|���h����ƦӰ����ק�

  By: Lawrence
  Date: 2003.11.18 Mark delete�Ϊk,�אּTruncate
*/

create or replace procedure SP_JDailyOp2
as
  v_CompCode number;
  v_ServiceType char(1);
  v_StopDate varchar2(8);
  v_Operator varchar2(10);
  v_PRFlag number;
  v_ReturnSQL varchar2(200);
  v_JobId number;

  v_TmpTableName varchar2(20);
  c39	char(1) := chr(39);
  v_RetCode number;
  v_RetMsg varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;
  v_SQL varchar2(1000);

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_JobId	:= 999997;
  v_FuncName	:= '��B�����(2)--��Ʋ��ͧ@�~';
  v_PrgName	:= 'SP_JDailyOp2';

  -- Loop Night-run�Ѽ��ɨC�@���Ӱ���
  for cr1 in cc1 loop
    DBMS_OUTPUT.PUT_LINE('(1) �}�l����...');
    v_ExecuteCount	:= v_ExecuteCount + 1;	-- �֥[���榸��
    v_CompCode		:= cr1.CompCode;
    v_ServiceType	:= cr1.ServiceType;
    v_Operator		:= cr1.Operator;
    v_PRFlag		:= cr1.PRFlag;
    v_ReturnSQL		:= cr1.ReturnSQL;
    v_StopDate		:= cr1.NextDataDay;

    -- ******************************
    -- (1) �I�s��������ݵ{��
    -- ******************************
    v_RetCode := SF_DailyOp2(v_CompCode, v_ServiceType, v_StopDate, to_char(Sysdate, 'YYYYMMDDHH24MI'), 
		v_Operator, v_PRFlag, v_ReturnSQL, v_TmpTableName, v_RetMsg);

    -- ******************************
    -- (2) �ھڵ��G�ӳB�z
    -- ******************************
    if v_RetCode >= 0 then 		-- ok
      DBMS_OUTPUT.PUT_LINE('(2) ���槹��, �x�s����O��...');

      -- ���R���P�骺���(����ư�)
      delete from SO502 where CompCode=v_CompCode and ServiceType=v_ServiceType and 
	RptDate=to_date(v_StopDate, 'YYYYMMDD');

      -- ���R���P�骺���(�@�θ�ư�)
      begin
        v_SQL := 'delete from gicmis.SO502X where CompCode='||v_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
		' and RptDate=to_date('||c39||v_StopDate||c39||','||c39||'YYYYMMDD'||c39||')';
        EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;

      -- �h�����G�ܲέp��(����ư�)
      insert into SO502 (CompCode, ServiceType, RptDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName) 
	(select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName from SO102);

      -- �h�����G�ܲέp��(�@�θ�ư�)
      begin
        v_SQL := 'insert into gicmis.SO502X (CompCode, ServiceType, RptDate, ReportTime, Operator, 
	  ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	  Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName) 
	  (select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	  ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	  Value04, Value05, Value06, Value07, Value08, CutoffTime, AreaCode, AreaName from SO102)';
        EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;

      -- delete from SO102;
      -- 2003.11.18 by Lawrence Mark delete�Ϊk,�אּTruncate
      DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO102');
    end if;

    -- ******************************
    -- (3) Log into operation log file: so030a
    -- ******************************
    v_StopExecTime := sysdate;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, v_RetCode, v_RetMsg);

    DBMS_OUTPUT.PUT_LINE('(3) ���槹��, Return message: ' || v_RetMsg);
    commit;
  end loop;

  if v_ExecuteCount = 0 then	-- SO501A: ����Ѽƥ��]�wok
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, null, null, v_StartExecTime, sysdate, 
	null, v_PrgName, v_FuncName, -99, 'SO501A: ����Ѽƥ��]�wok');
      commit;
  end if;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;

    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, -99, '��L���~');
    commit;

    --p_RetMsg := SQLERRM;
    --return '-99,'||SQLERRM;
end;
/
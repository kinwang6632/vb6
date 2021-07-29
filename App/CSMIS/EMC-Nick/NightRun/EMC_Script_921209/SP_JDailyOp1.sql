/*
	EMC V3.0 Night-run ����

@C:\Gird\v400\Script\SP_JDailyOp1

exec SP_JDailyOp1

  �妸�@�~__��B�����: �C��έp�Ȥ�˾�/�����, MTD, YTD	***** V3.0 *****
  �ɦW: SF_JDailyOp1.sql

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
  Date: 2001.08.16
	2001.09.26 ��: �Ǧ^�ȧאּ�r��, �]�treturn code and return message
	2001.11.19 ��: bug, SO093.StartDate --> StopDate
	2001.12.19 ��: �۰ʱN���ͤ���ƽƻs��@����(SO501X)��, �H�KMSO�έp�U�t�θ��
        2001.12.20 ��: GICMIS.SO501X�������y�k��EXECUTE IMMEDIATE by Lawrence
	2002.03.14 ��: �[�W�I�sSP_JDailyOp2 ��B�����3(��F��)
	2002.10.03 ��: �[�W�I�sSP_JDailyOp4 ��B�����4(��F��/�A�Ȱ�)
	2003.05.19 ��: �]���h�A�ȧO����, Night-run�Ѽ���(SO501A)�|���h����ƦӰ����ק�

  By: Lawrence
  Date: 2003.11.18 Mark delete�Ϊk,�אּTruncate
*/

create or replace procedure SP_JDailyOp1
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

  -- use so501a to save these values
  --v_CompCode 		:= 3;
  --v_ServiceType 	:= 'C';
  --v_StopDate 		:= '20011117';
  --v_Operator 		:= 'MAINTAIN';
  --v_PRFlag 		:= 1;
  --v_ReturnSQL		:= null;

  v_JobId		:= 999998;
  v_FuncName		:= '��B�����--��Ʋ��ͧ@�~';
  v_PrgName		:= 'SP_JDailyOp1';


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
    v_RetCode := SF_DailyOp1(v_CompCode, v_ServiceType, v_StopDate, to_char(Sysdate, 'YYYYMMDDHH24MI'), 
		v_Operator, v_PRFlag, v_ReturnSQL, v_TmpTableName, v_RetMsg);

    -- ******************************
    -- (2) �ھڵ��G�ӳB�z
    -- ******************************
    if v_RetCode >= 0 then 		-- ok
      DBMS_OUTPUT.PUT_LINE('(2) ���槹��, �x�s����O��...');
      -- ���R���P�骺���(����ư�)
      delete from SO501 where CompCode=v_CompCode and ServiceType=v_ServiceType and 
	RptDate=to_date(v_StopDate, 'YYYYMMDD');
/*
-- *********************************************************************************
-- �Y�bEMC, �h���q��comment

      -- ���R���P�骺���(�@�θ�ư�)	2001.12.19 added
--    delete from gicmis.SO501X where gicmis.CompCode=v_CompCode and gicmis.ServiceType=v_ServiceType and 
--      gicmis.RptDate=to_date(v_StopDate, 'YYYYMMDD');
      begin
         v_SQL := 'delete from gicmis.SO501X where CompCode='||to_char(v_CompCode)||' and ServiceType='||c39||v_ServiceType||c39||' and RptDate='||
             'to_date('||c39||v_StopDate||c39||', '||c39||'YYYYMMDD'||c39||')';
         EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;
-- *********************************************************************************
*/
      -- �h�����G�ܲέp��(����ư�)
      insert into SO501 (CompCode, ServiceType, RptDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime) 
	(select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	Value04, Value05, Value06, Value07, Value08, CutoffTime from SO093);
/*
-- *********************************************************************************
-- �Y�bEMC, �h���q��comment
      -- �h�����G�ܲέp��(�@�θ�ư�)	2001.12.19 added
      begin
        v_SQL :='insert into gicmis.SO501X (CompCode, ServiceType, RptDate, ReportTime, Operator, 
   	   ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	   Value04, Value05, Value06, Value07, Value08, CutoffTime) 
	   (select CompCode, ServiceType, StopDate, ReportTime, Operator, 
	   ItemCode1, ItemName1, ItemCode2, ItemName2, Value01, Value02, Value03, 
	   Value04, Value05, Value06, Value07, Value08, CutoffTime from SO093)';
        EXECUTE IMMEDIATE v_SQL;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
          DBMS_OUTPUT.PUT_LINE(v_SQL);
          rollback;
      end;
-- *********************************************************************************
*/
      -- �R���Ȧs�ɤ��e
      -- delete from SO093;
      -- 2003.11.18 by Lawrence Mark delete�Ϊk,�אּTruncate
      DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO093');
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


    -- ***************************************************************************
    -- (4) ��LNight-run�{���~����
    -- ***************************************************************************
    SP_JDailyOp2;		-- �I�s: ��B�����2(��F��)
    SP_JDailyOp4;		-- �I�s: ��B�����4(�U�ؤ�Ʋέp)
    SP_GetSnapShotTableCounts   -- �I�s: �p��n�Ǧ� SnapShot Table ������
    -- ***************************************************************************


    -- ******************************
    -- (5) �ק�U������Ѽ� (�ק�NextDataDay�����@��)
    -- ******************************
    update SO501A set NextDataDay = to_char(to_date(v_StopDate,'YYYYMMDD')+1,'YYYYMMDD')
	where CompCode=v_CompCode and ServiceType=v_ServiceType;

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
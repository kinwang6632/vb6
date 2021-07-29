/*
@C:\EMC_Script\SP_JDailyOp4

exec SP_JDailyOp4

  �妸�@�~__��B�����4(��F��): �C��έp�U���Ȥ��
  �ɦW: SF_JDailyOp4.sql

  ��B�����4�έp�����: SO504
	SeqNo		number(3)
	CompCode	number(3)
	ServiceType	char(1)
	RptDate		Date
	RptTime		Date
	DataType	number(2)
	ServAreaCode	varchar2(3)
	DataCode	number(3)
	CustCnt		number(8)

  By: Pierre
  Date: 2002.10.03
	2002.10.12 ��: EMC�令�I�s3.0�����覡
	2003.03.31 ��: �[�W�I�s SF_AdjCustCount()
	2003.04.07 ��: �[�j�{������
	2003.05.19 ��: �]���h�A�ȧO����, Night-run�Ѽ���(SO501A)�|���h����ƦӰ����ק�
	2003.08.11 ��: �����I�s SF_AdjCustCount()
*/

create or replace procedure SP_JDailyOp4
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
  v_RetCode number := 0;
  v_RetMsg varchar2(1000) := '�L���G';
  v_StartExecTime date;
  v_StopExecTime date;
  v_SQL varchar2(1000);
  v_Now date;

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_JobId	:= 999996;
  v_FuncName	:= '��B�����(4)--��Ʋ��ͧ@�~';
  v_PrgName	:= 'SP_JDailyOp4';

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
    -- (1) ���R���P�骺���
    -- ******************************
    v_Now := trunc(sysdate);
    delete from SO504 where CompCode=v_CompCode and ServiceType=v_ServiceType and RptDate=v_Now;

    -- ******************************
    -- (2) ��B�����(4)--��Ʋ��ͧ@�~, �I�s��������ݵ{��SF_DailyOp4_30
    -- ******************************
    --  v_RetCode := SF_DailyOp4_20(v_CompCode, v_ServiceType, v_Operator, 1, v_RetMsg);
    v_RetCode := SF_DailyOp4_30(v_CompCode, v_ServiceType, v_Operator, 1, v_RetMsg);

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

/*
    -- ********************************************************
    -- (4) ����Debby���j�Ӥ�Φ���Ʀ۰ʭp��{�� SF_AdjCustCount()	-- 2003.03.31
    -- ********************************************************
    v_RetCode := SF_AdjCustCount(v_CompCode);
    if v_RetCode < 0 then
      DBMS_OUTPUT.PUT_LINE('(4) ����SF_AdjCustCount���~, Return code: ' || v_RetCode);
    else
      DBMS_OUTPUT.PUT_LINE('(4) ����SF_AdjCustCount����, �B�z����: ' || v_RetCode);
    end if;
    commit;
*/
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
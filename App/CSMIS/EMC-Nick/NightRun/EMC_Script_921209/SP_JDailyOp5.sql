/*
@C:\EMC_Script\SP_JDailyOp5

exec SP_JDailyOp5

  �妸�@�~__�C��۰ʰ��u��鵲����(�C��1��02:00�@��, �۰ʩI�sSF_ClsWorkOrder)

  �ɦW: SP_JDailyOp5.sql 

  By: Pierre
  Date: 2003.04.08 �`�N: Lawrence�w�ק惡��SF_ClsWorkOrder ��V3.5��, �s���Ѽƥ���, �Y���ܰ�,�h�ݭק�.
	2003.05.05 ��: �令�۰ʵ���W��̫�@��, �A�NSO088�鵲��[�@��
	2003.05.19 ��: �]���h�A�ȧO����, Night-run�Ѽ���(SO501A)�|���h����ƦӰ����ק�
	2003.06.09 ��: �N���"�A�NSO088�鵲��[�@��"�٭�

*/

create or replace procedure SP_JDailyOp5
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

  v_LastClsDate1 date;		-- �W���˾���鵲�I���
  v_LastClsDate2 date;		-- �W�����׳�鵲�I���
  v_LastClsDate3 date;		-- �W���������鵲�I���
  v_StartDate	 date;		-- �����u��鵲��ư_�l�� (date�榡)
  v_ClsStartDate varchar2(8);	-- �����u��鵲��ư_�l�� (varchar2�榡)
  v_ClsStopDate  varchar2(8);	-- �����u��鵲��ƺI���

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_JobId	:= 999995;
  v_FuncName	:= '(5) �C��۰ʤu��鵲����';
  v_PrgName	:= 'SP_JDailyOp5';

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

    -- ********************************************************
    -- (1) ������Ѽ�
    -- ********************************************************
    -- 1. ���T�ؤu�檺�e���鵲��ƺI���
    begin
      select InstStopTime, MaintainStopTime, PrStopTime into v_LastClsDate1, v_LastClsDate2, 
	v_LastClsDate3 from SO088 where CompCode=v_CompCode and ServiceType=v_ServiceType;
    exception
      when NO_DATA_FOUND then	-- �L�鵲�L, �G�H200210101���_�l��
        v_ClsStartDate := '20020101';
      when others then	-- �����D, �G������"�۰ʰ��u��鵲����", �������L
        DBMS_OUTPUT.PUT_LINE(v_FuncName||'--���鵲�Ѽƿ��~: ' || SQLERRM);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	  v_FuncName, -1, '���鵲�Ѽƿ��~');
        goto GO_NEXT1;
    end;

    -- 2. ���̦�������������鵲��ư_�l��
    v_StartDate := least(nvl(v_LastClsDate1,to_date('20020101','YYYYMMDD')), 
		nvl(v_LastClsDate2,to_date('20020101','YYYYMMDD')), nvl(v_LastClsDate3,to_date('20020101','YYYYMMDD')));
    v_ClsStartDate := to_char(v_StartDate, 'YYYYMMDD');	-- �নvarchar2�榡

    -- 3. �������u��鵲��ƺI��� (�����e�뤧�̫�@��)
    --v_ClsStopDate := to_char(sysdate, 'YYYYMM')||'01';	-- (�����Ӥ뤧�Ĥ@��)
    v_ClsStopDate := to_char(last_day(add_months(sysdate,-1)), 'YYYYMMDD');

    -- ********************************************************
    -- (2) �I�sLawrence�������鵲�禡 SF_ClsWorkOrder:  �Ǧ^���~�X�Ϊ̸�Ƨ帹
    -- ********************************************************
    v_RetCode := SF_ClsWorkOrder(v_ClsStartDate, v_ClsStopDate, 'Night-run', 'Night-run', 
			1, 1, 1, v_CompCode, v_ServiceType, v_RetMsg);

    -- ********************************************************
    -- (3) �B�z���浲�G, �Y�鵲���\, �A�NSO088�鵲��[�@�� (92.06.09��������)
    -- ********************************************************
    v_StopExecTime := sysdate;
    if v_RetCode < 0 then
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'--����SF_ClsWorkOrder���~, Return code: ' || v_RetCode);
    else
      -- �Y�鵲���\, �A�NSO088�鵲��[�@�� (92.06.09��������)
      -- �Y�鵲���\, ��s�T�ؤu��鵲���
      update SO088 set InstStopTime=to_date(v_ClsStopDate,'YYYYMMDD'), MaintainStopTime=to_date(v_ClsStopDate,'YYYYMMDD'),
	PrStopTime=to_date(v_ClsStopDate,'YYYYMMDD') where CompCode=v_CompCode and ServiceType=v_ServiceType;
      DBMS_OUTPUT.PUT_LINE(v_FuncName||'--����SF_ClsWorkOrder����, ��Ƨ帹: ' || v_RetCode);
    end if;

    -- ********************************************************
    -- (4) Log into operation log file: so030a
    -- ********************************************************
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, v_RetCode, v_RetMsg);

    <<GO_NEXT1>>
    commit;
    NULL;
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
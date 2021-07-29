/*
@c:\emc_script\SP_AdjSO001ToSO002

  �妸�@�~__???
  �ɦW: SP_AdjSO001ToSO002.sql

  By: Pierre
  Date: 2002.03.12
*/

create or replace procedure SP_AdjSO001ToSO002
as
  v_CompCode number;
  v_ServiceType char(1);
  v_JobId number;

  c39	char(1) := chr(39);
  v_RetCode number;
  v_RetMsg varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;

  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  cursor cc1 is 
    select CompCode, ServiceType, Operator, PRFlag, ReturnSQL, NextDataDay from so501A;

begin
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_JobId	:= 0;
  v_FuncName	:= '???';
  v_PrgName	:= 'SP_???';

  -- Loop Night-run�Ѽ��ɨC�@���Ӱ���
  for cr1 in cc1 loop
    DBMS_OUTPUT.PUT_LINE('(1) �}�l����...');
    v_ExecuteCount	:= v_ExecuteCount + 1;	-- �֥[���榸��
    v_CompCode		:= cr1.CompCode;
    v_ServiceType	:= cr1.ServiceType;

    -- ******************************
    -- (1) �I�s��������ݵ{��
    -- ******************************
    v_RetCode := SF_AdjSO001ToSO002(v_CompCode, v_ServiceType, 1, v_RetMsg);

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
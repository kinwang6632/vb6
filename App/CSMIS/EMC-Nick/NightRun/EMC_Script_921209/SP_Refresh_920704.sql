/*
@C:\EMC_Script\SP_Refresh
set serveroutput on
exec SP_Refresh

  Night-run: ��Y�@��ưϰ���snapshot
  �ɦW: SP_Refresh.sql

  * �p���ˬd��Night-run���浲�G?
	select * from so030a where jobid=999993

  * Snapshot �]�w��: (�������ʮ�, �ݭק�{���í��scompile)
	���q�O	OwnerName GroupName
	======	========= =========
	1	EMCKS	EMCKS1~5
	2	EMCPN	EMCPN1~5
	3	EMCNT	EMCNT1~5
	5	EMCNCC	EMCNCC1~5
	6	EMCFM	EMCFM1~5
	7	EMCCT	EMCCT1~5
	8	EMCUC	EMCUC1~5
	9	EMCYMS 	EMCYMS1~5
	10	EMCNTP	EMCNTP1~5
	11	EMCKP	EMCKP1~5
	12	EMCDW	EMCDW1~5
	13	EMCTC	EMCTC1~5

  By: Pierre
  Date: 2003.06.20
	2003.07.04 ��: �C���@��loop�Nlog�@��, ���ަ��L���~
*/

create or replace procedure SP_Refresh
as
  v_CompCode number;
  v_ServiceType char(1);
  v_JobId number;
  c39	char(1) := chr(39);
  v_RetCode number := 0;
  v_RetMsg varchar2(1000) := '�L���G';
  v_StartExecTime date;
  v_StopExecTime date;
  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_ExecuteCount number := 0;
  v_GroupCount	number;
  i 		number;
  v_OwnerName	varchar2(40);
  v_TmpStr	varchar2(200);
  v_Flag	number := 0;

begin
  -- ********************************************************
  -- 1. �]�wNight-run�򥻰Ѽ�
  -- ********************************************************
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_JobId	:= 999993;
  v_FuncName	:= '����Snapshot';
  v_PrgName	:= 'SP_Refresh';
  v_GroupCount	:= 5;			-- �ثe�w�q5��snapshot group

  -- ********************************************************
  -- 2. ��Night-run��L�����Ѽ�: ���q�O, �A�ȧO(�����Ȥ����n)
  -- ********************************************************
  begin
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501A where Rownum<=1;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, -99, 'SO501A: ����Ѽƥ��]�wok');
      commit;
      return;
  end;

  -- ********************************************************
  -- 3. �ھڤ��q�O, �M�w�Ӹ�ưϪ�Owner name
  -- ********************************************************
  if v_CompCode = 1 then
    v_OwnerName := 'EMCKS';
  elsif v_CompCode = 2 then
    v_OwnerName := 'EMCPN';
  elsif v_CompCode = 3 then
    v_OwnerName := 'EMCNT';
  elsif v_CompCode = 4 then
    v_OwnerName := '';
  elsif v_CompCode = 5 then
    v_OwnerName := 'EMCNC';
  elsif v_CompCode = 6 then
    v_OwnerName := 'EMCFM';
  elsif v_CompCode = 7 then
    v_OwnerName := 'EMCCT';
  elsif v_CompCode = 8 then
    v_OwnerName := 'EMCUC';
  elsif v_CompCode = 9 then
    v_OwnerName := 'EMCYMS';
  elsif v_CompCode = 10 then
    v_OwnerName := 'EMCNTP';
  elsif v_CompCode = 11 then
    v_OwnerName := 'EMCKP';
  elsif v_CompCode = 12 then
    v_OwnerName := 'EMCDW';
  elsif v_CompCode = 13 then
    v_OwnerName := 'EMCTC';
  elsif v_CompCode = 14 then
    v_OwnerName := '';
  elsif v_CompCode = 15 then
    v_OwnerName := '';
  elsif v_CompCode = 16 then
    v_OwnerName := '';
  end if;
  if v_OwnerName is null then
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, -98, '�L�����q�O�POwner name���������Y, �L�k��Snapshot! ���q�O='||v_CompCode);
    commit;
    return;
  end if;

  -- **************************************************************************
  -- 4. Loop�Ҧ��w�w�q�n��group��snapshot(�e��: group name�W�h��'emcpn.emcpn1')
  -- **************************************************************************
  for i in 1..v_GroupCount loop
    v_TmpStr := v_OwnerName||'.'||v_OwnerName||ltrim(to_char(i));
    begin
      DBMS_REFRESH.REFRESH(v_TmpStr);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, 0, 'Snapshot���槹��('||i||')');
      commit;
    exception
      when others then
	v_RetMsg := '('||i||')���~�T��:'||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
    	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	  v_FuncName, -97, v_RetMsg);
	v_Flag := v_Flag + 1;			-- ���~flag
	commit;
    end;
  end loop;
/*
  -- ********************************************************
  -- 5. Log���浲�G
  -- ********************************************************
  if v_Flag=0 then
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, 0, 'Snapshot���槹��');
    commit;
  end if;
*/
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
end;
/
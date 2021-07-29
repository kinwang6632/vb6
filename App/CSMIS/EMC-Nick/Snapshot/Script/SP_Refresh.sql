create or replace procedure SP_Refresh
as
/*
@C:\EMC_Script\SP_Refresh
set serveroutput on
exec SP_Refresh

  Night-run: 於某一資料區執行snapshot
  檔名: SP_Refresh.sql

  * 如何檢查此Night-run執行結果? select * from so030a where jobid=999993

  * Snapshot 設定表: (此表有異動時, 需修改程式並重新compile)
  * GroupName 是用 DBMS_REFRESH.MAKE( ) 先建好, 當使用 DBMS_REFRESH.REFRESH() 時就可一一次更心新該 Group
  
	公司別	OwnerName GroupName
	======	========= =========
	1	EMCKS	EMCKS1~32
	2	EMCPN	EMCPN1~32
	3	EMCNT	EMCNT1~32
	5	EMCNCC	EMCNCC1~32
	6	EMCFM	EMCFM1~32
	7	EMCCT	EMCCT1~32
	8	EMCUC	EMCUC1~32
	9	EMCYMS 	EMCYMS1~32
	10	EMCNTP	EMCNTP1~32
	11	EMCKP	EMCKP1~32
	12	EMCDW	EMCDW1~32
	13	EMCTC	EMCTC1~32
	14      EMCST   EMCST1~32 
	16      EMCNTY  EMCNTY1~32       
	

  By: Pierre
  Date: 2003.06.20
	2003.07.04 改: 每做一個loop就log一次, 不管有無錯誤
        2003.11.07 改: 動態產生group name  [Jackal]
*/
  v_CompCode number;
  v_ServiceType char(1);
  v_JobId number;
  c39	char(1) := chr(39);
  v_RetCode number := 0;
  v_RetMsg varchar2(1000) := '無結果';
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

  cursor cc1(p_ownername varchar2) is
    select rname from all_refresh where rowner=p_ownername;

begin
  -- ********************************************************
  -- 1. 設定Night-run基本參數
  -- ********************************************************
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_JobId	:= 999993;
  v_FuncName	:= '執行Snapshot';
  v_PrgName	:= 'SP_Refresh';
  v_GroupCount	:= 5;			-- 目前定義5個snapshot group

  -- ********************************************************
  -- 2. 取Night-run其他相關參數: 公司別, 服務別(但此值不重要)
  -- ********************************************************
  begin
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501A where Rownum<=1;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, -99, 'SO501A: 執行參數未設定ok');
      commit;
      return;
  end;

  -- ********************************************************
  -- 3. 根據公司別, 決定該資料區的Owner name
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
	v_FuncName, -98, '無此公司別與Owner name的對應關係, 無法做Snapshot! 公司別='||v_CompCode);
    commit;
    return;
  end if;

  -- **************************************************************************
  -- 4. Loop所有已定義好的group做snapshot(前提: group n	ame規則為'emcpn.emcpn1')
  -- **************************************************************************
  --for i in 1..v_GroupCount loop
  --  v_TmpStr := v_OwnerName||'.'||v_OwnerName||ltrim(to_char(i));
    for cr in cc1(v_OwnerName) loop
      v_TmpStr := v_OwnerName||'.'||cr.rname;
    begin
      DBMS_REFRESH.REFRESH(v_TmpStr);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	--v_FuncName, 0, 'Snapshot執行完畢('||i||')');
        v_FuncName, 0, 'Snapshot執行完畢('||v_TmpStr||')');
      commit;
    exception
      when others then
	--v_RetMsg := '('||i||')錯誤訊息:'||SQLERRM;
	v_RetMsg := '('||v_TmpStr||')錯誤訊息:'||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
    	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	  v_FuncName, -97, v_RetMsg);
	v_Flag := v_Flag + 1;			-- 錯誤flag
	commit;
    end;
  end loop;
/*
  -- ********************************************************
  -- 5. Log執行結果
  -- ********************************************************
  if v_Flag=0 then
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, 
	v_FuncName, 0, 'Snapshot執行完畢');
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
	v_FuncName, -99, '其他錯誤');
    commit;
end;
/
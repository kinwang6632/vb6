/*
@C:\emc_script\SP_GetSTBData

set serveroutput on
exec SP_GetSTBData

  (批次作業)取STB相關資料, 供物料系統核對, 將資料存於暫存檔MSSData1/2/3/4
  檔名: SP_GetSTBData.sql

  MSSDATA1 = 資料暫存檔1(有用過的盒子)
  MSSDATA2 = 資料暫存檔2(有用過的IC卡)
  MSSDATA3 = 資料暫存檔3(在庫的盒子)
  MSSDATA4 = 資料暫存檔4(在庫的ICC)

  By: Pierre
  Date: 2003.04.25
	2003.07.25 改: 因應全面將設備序號欄長統一為varchar2(20)的修改

*/

create or replace procedure SP_GetSTBData
as
  v_CompCode number;
  v_ServiceType char(1);
  v_StopDate varchar2(8);
  v_Operator varchar2(10);
  v_ReturnSQL varchar2(200);
  v_JobId number;

  v_TmpTableName1 varchar2(20);
  v_TmpTableName2 varchar2(20);
  c39	char(1) := chr(39);
  v_RetCode number := 0;
  v_RetMsg varchar2(1000) := '無結果';
  v_StartExecTime date;
  v_StopExecTime date;
  v_SQL varchar2(4000);
  v_Now date;
  v_Cnt1 number;
  v_FinTime date;

  x_CompCode number;
  x_FaciSno varchar2(20);
  x_MaterialNo varchar2(15);
  x_BuyName varchar2(20);
  x_Deposit number;
  x_SmartCardNo varchar2(20);
  x_Amount number;
  x_CustId number;
  x_CustName varchar2(50);
  x_Tel1 varchar2(20);
  x_InstDate date;
  x_Sno varchar2(15);
  x_PrDate date;
  x_PrSno varchar2(15);
  x_Note varchar2(250);
  x_UpdTime varchar2(19);


  -- 有用過的盒子 
  cursor cc1 is 
    select a.compcode,a.facisno,c.materialno,a.buyname,a.deposit,a.smartcardno,a.amount,b.custid,b.custname,
	b.tel1,a.instdate,a.sno,a.prdate,a.prsno,a.note,a.updtime 
	from so004 a, so001 b,soac0201a c 
	where a.custid=b.custid and c.facisno=a.facisno and a.facicode=201 and a.facisno is not null and a.prdate is null 
	union
	select a.compcode,a.facisno,c.materialno,'拆' BuyName,a.deposit,a.smartcardno,a.amount,b.custid,b.custname,
	b.tel1,a.instdate,a.sno,a.prdate,a.prsno,a.note,a.updtime 
	from so004 a, so001 b,soac0201a c 
	where a.custid=b.custid and c.facisno=a.facisno and a.facicode=201 and a.prdate >= a.instdate
	order by 2, 16;
	--order by facisno,updtime;

  -- 有用過的ICC
  cursor cc2 is 
    select a.compcode,a.facisno,c.materialno,a.buyname,a.deposit,a.smartcardno,a.amount,b.custid,b.custname,
	b.tel1,a.instdate,a.sno,a.prdate,a.prsno,a.note,a.updtime 
	from so004 a, so001 b,soac0201b c 
	where a.custid=b.custid and c.facisno=a.facisno and a.facicode=202 and a.facisno is not null and a.prdate is null 
	union
	select a.compcode,a.facisno,c.materialno,'拆' BuyName,a.deposit,a.smartcardno,a.amount,b.custid,b.custname,
	b.tel1,a.instdate,a.sno,a.prdate,a.prsno,a.note,a.updtime 
	from so004 a, so001 b,soac0201b c 
	where a.custid=b.custid and c.facisno=a.facisno and a.facicode=201 and a.prdate >= a.instdate
	order by 2, 16;

/*
  -- soac0201a 在庫的盒子 
  cursor cc3 is 
    select * from soac0201a where facisno not in 
	(select distinct(facisno) from (select a.facisno from so004 a, so001 b 
	where a.custid=b.custid and a.facicode=201 and a.facisno is not null and a.prdate is null 
	union
	select a.facisno from so004 a, so001 b 
	where a.custid=b.custid and a.facicode=201 and a.prdate >= a.instdate )) ;

  -- soac0201b 在庫的ICC
  cursor cc4 is 
    select * from soac0201b where facisno not in 
	(select distinct(facisno) from (select a.facisno from so004 a, so001 b 
	where a.custid=b.custid and a.facicode=202 and a.facisno is not null and a.prdate is null 
	union
	select a.facisno from so004 a, so001 b 
	where a.custid=b.custid and a.facicode=202 and a.prdate >= a.instdate )) ;
*/
begin


  v_StartExecTime := sysdate;		-- 開始執行時間
  v_JobId := 999991;			-- 開博Night-run job id.

  DBMS_OUTPUT.PUT_LINE('(0) 開始執行...');

  -- 取公司別...等參數
  begin
    select CompCode, ServiceType, Operator, ReturnSQL, NextDataDay
	into v_CompCode, v_ServiceType, v_Operator, v_ReturnSQL, v_StopDate
	from so501A where Rownum<=1;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	'取STB相關資料', -99, 'SO501A: 執行參數未設定ok');
      --commit;
      --return;
      v_CompCode := 0;
      v_ServiceType := 'C';
  end;

  -- *******************************************************************
  DBMS_OUTPUT.PUT_LINE('(1) MSSDATA1 = 資料暫存檔1(有用過的盒子)...');
  v_TmpTableName1 := 'MSSDATA1';		-- 資料暫存檔1(有用過的盒子)

  -- 刪除同名之資料暫存檔
  select count(*) into v_Cnt1 from user_tables where table_name=v_TmpTableName1;
  if v_cnt1 > 0 then 
    v_SQL := 'drop table ' || v_TmpTableName1;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法刪除');
	commit;
	return;
    end;
  end if; 

  -- 建立資料暫存檔
  v_SQL := 'CREATE TABLE '|| v_TmpTableName1 || ' (' ||
		'CompCode number(3),FaciSno varchar2(20),MaterialNo varchar2(15),BuyName varchar2(20),' ||
		'Deposit number(8),SmartCardNo varchar2(20),Amount number(8),CustId number(8),'||
		'CustName varchar2(50),Tel1 varchar2(20),InstDate date,Sno varchar2(15),PrDate date,'||
		'PrSno varchar2(15),Note varchar2(250),UpdTime varchar2(19))';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL); 
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法建立');
      commit;
      return;
  end;
 
  v_Cnt1 := 0;
  for cr1 in cc1 loop
    begin
      -- 若此筆之來源單號為'M'單且有完工時間, 但對應之SO008尚未完工, 則屬於線上回報, 但尚未工單登錄 ==> 故不計入
      if substr(cr1.SNo,7,1)='M' and cr1.InstDate is not null then
	begin
	  select FinTime into v_FinTime from SO008 where SNo=cr1.SNo;
	  if v_FinTime is null then
	    goto GO_NEXT1;
	  end if;
	exception 
	  when others then		-- something wrong: no such work order
	    goto GO_NEXT1;
	end;
      end if;

      v_Cnt1 := v_Cnt1 + 1;
      v_SQL := 'Insert into '||v_TmpTableName1||' (COMPCODE,FACISNO,MATERIALNO,BUYNAME,DEPOSIT,' ||
		'SMARTCARDNO,AMOUNT,CUSTID,CUSTNAME,TEL1,INSTDATE,SNO,PRDATE,PRSNO,NOTE,UPDTIME) ' ||
		'values (';
      v_SQL := v_SQL||cr1.CompCode||','||c39||cr1.FaciSNo||c39||','||c39||cr1.MaterialNo||c39||','||c39||cr1.BuyName||c39||','||
		nvl(cr1.Deposit,0)||','||c39||cr1.SmartCardNo||c39||','||nvl(cr1.Amount,0)||','||cr1.CustId||','||
		c39||cr1.CustName||c39||','||c39||nvl(cr1.Tel1,'')||c39||',';
      if cr1.InstDate is null then
	v_SQL := v_SQL||'null,';
      else
	v_SQL := v_SQL||'to_date('||c39||to_char(cr1.InstDate,'YYYYMMDD')||c39||','||c39||'YYYYMMDD'||c39||'),';
      end if;
      v_SQL := v_SQL||c39||nvl(cr1.SNO,'')||c39||',';
      if cr1.PrDate is null then
	v_SQL := v_SQL||'null,';
      else
	v_SQL := v_SQL||'to_date('||c39||to_char(cr1.PrDate,'YYYYMMDD')||c39||','||c39||'YYYYMMDD'||c39||'),';
      end if;
      v_SQL := v_SQL||c39||nvl(cr1.PrSNO,'')||c39||','||c39||cr1.Note||c39||','||c39||cr1.UpdTime||c39||')';

      EXECUTE IMMEDIATE (v_SQL); 
              
    exception
      when others then
        v_RetMsg := 'INSERT時發生錯誤: '||SQLERRM;
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法建立');
	-- commit;
	-- return;		有錯誤繼續做 !!!
    end;

    <<GO_NEXT1>>
    NULL;
  end loop;


  -- *******************************************************************
  DBMS_OUTPUT.PUT_LINE('(2) MSSDATA2 = 資料暫存檔2(有用過的IC卡)...');
  v_TmpTableName1 := 'MSSDATA2';		-- 資料暫存檔2(有用過的IC卡)

  -- 刪除同名之資料暫存檔
  select count(*) into v_Cnt1 from user_tables where table_name=v_TmpTableName1;
  if v_cnt1 > 0 then 
    v_SQL := 'drop table ' || v_TmpTableName1;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法刪除');
	commit;
	return;
    end;
  end if; 

  -- 建立資料暫存檔
  v_SQL := 'CREATE TABLE '|| v_TmpTableName1 || ' (' ||
		'CompCode number(3),FaciSno varchar2(20),MaterialNo varchar2(15),BuyName varchar2(20),' ||
		'Deposit number(8),SmartCardNo varchar2(20),Amount number(8),CustId number(8),'||
		'CustName varchar2(50),Tel1 varchar2(20),InstDate date,Sno varchar2(15),PrDate date,'||
		'PrSno varchar2(15),Note varchar2(250),UpdTime varchar2(19))';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL); 
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法建立');
      commit;
      return;
  end;
 
  v_Cnt1 := 0;
  for cr1 in cc2 loop
    begin
      -- 若此筆之來源單號為'M'單且有完工時間, 但對應之SO008尚未完工, 則屬於線上回報, 但尚未工單登錄 ==> 故不計入
      if substr(cr1.SNo,7,1)='M' and cr1.InstDate is not null then
	begin
	  select FinTime into v_FinTime from SO008 where SNo=cr1.SNo;
	  if v_FinTime is null then
	    goto GO_NEXT2;
	  end if;
	exception 
	  when others then		-- something wrong: no such work order
	    goto GO_NEXT2;
	end;
      end if;

      v_Cnt1 := v_Cnt1 + 1;
      v_SQL := 'Insert into '||v_TmpTableName1||' (COMPCODE,FACISNO,MATERIALNO,BUYNAME,DEPOSIT,' ||
		'SMARTCARDNO,AMOUNT,CUSTID,CUSTNAME,TEL1,INSTDATE,SNO,PRDATE,PRSNO,NOTE,UPDTIME) ' ||
		'values (';
      v_SQL := v_SQL||cr1.CompCode||','||c39||cr1.FaciSNo||c39||','||c39||cr1.MaterialNo||c39||','||c39||cr1.BuyName||c39||','||
		nvl(cr1.Deposit,0)||','||c39||cr1.SmartCardNo||c39||','||nvl(cr1.Amount,0)||','||cr1.CustId||','||
		c39||cr1.CustName||c39||','||c39||nvl(cr1.Tel1,'')||c39||',';
      if cr1.InstDate is null then
	v_SQL := v_SQL||'null,';
      else
	v_SQL := v_SQL||'to_date('||c39||to_char(cr1.InstDate,'YYYYMMDD')||c39||','||c39||'YYYYMMDD'||c39||'),';
      end if;
      v_SQL := v_SQL||c39||nvl(cr1.SNO,'')||c39||',';
      if cr1.PrDate is null then
	v_SQL := v_SQL||'null,';
      else
	v_SQL := v_SQL||'to_date('||c39||to_char(cr1.PrDate,'YYYYMMDD')||c39||','||c39||'YYYYMMDD'||c39||'),';
      end if;
      v_SQL := v_SQL||c39||nvl(cr1.PrSNO,'')||c39||','||c39||cr1.Note||c39||','||c39||cr1.UpdTime||c39||')';

      EXECUTE IMMEDIATE (v_SQL); 
              
    exception
      when others then
        v_RetMsg := 'INSERT時發生錯誤: '||SQLERRM;
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法建立');
	-- commit;
	-- return;		有錯誤繼續做 !!!
    end;

    <<GO_NEXT2>>
    NULL;
  end loop;

  -- *******************************************************************
  DBMS_OUTPUT.PUT_LINE('(3) MSSDATA3 = 資料暫存檔3(在庫的盒子)...');
  v_TmpTableName1 := 'MSSDATA3';		-- 資料暫存檔3(在庫的盒子)

  -- 刪除同名之資料暫存檔
  select count(*) into v_Cnt1 from user_tables where table_name=v_TmpTableName1;
  if v_cnt1 > 0 then 
    v_SQL := 'drop table ' || v_TmpTableName1;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法刪除');
	commit;
	return;
    end;
  end if; 

  -- 建立資料暫存檔
  v_SQL := 'CREATE TABLE '|| v_TmpTableName1 || ' as (' ||
	'select * from soac0201a where facisno not in (select distinct(facisno) from ' || 
	'(select a.facisno from so004 a, so001 b where a.custid=b.custid and a.facicode=201 ' ||
	'and a.facisno is not null and a.prdate is null union select a.facisno from so004 a, so001 b ' ||
	'where a.custid=b.custid and a.facicode=201 and a.prdate >= a.instdate )))';

  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL); 
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法建立');
      commit;
      return;
  end;

  -- *******************************************************************
  DBMS_OUTPUT.PUT_LINE('(4) MSSDATA4 = 資料暫存檔4(在庫的ICC)...');
  v_TmpTableName1 := 'MSSDATA4';		-- 資料暫存檔4(在庫的ICC)

  -- 刪除同名之資料暫存檔
  select count(*) into v_Cnt1 from user_tables where table_name=v_TmpTableName1;
  if v_cnt1 > 0 then 
    v_SQL := 'drop table ' || v_TmpTableName1;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法刪除');
	commit;
	return;
    end;
  end if; 

  -- 建立資料暫存檔
  v_SQL := 'CREATE TABLE '|| v_TmpTableName1 || ' as (' ||
	'select * from soac0201b where facisno not in (select distinct(facisno) from ' || 
	'(select a.facisno from so004 a, so001 b where a.custid=b.custid and a.facicode=202 ' ||
	'and a.facisno is not null and a.prdate is null union select a.facisno from so004 a, so001 b ' ||
	'where a.custid=b.custid and a.facicode=202 and a.prdate >= a.instdate )))';

  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL); 
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	  '取STB相關資料', -99, v_TmpTableName1||'無法建立');
      commit;
      return;
  end;

  -- ***************************************************************************************
  -- Log into operation log file: so030a
  v_StopExecTime := sysdate;
  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	'取STB相關資料', v_RetCode, '完成');

  DBMS_OUTPUT.PUT_LINE('(9) 執行完畢, Return message: 完成');

  commit;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;

    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), 'SP_GetSTBData', 
	'取STB相關資料', -99, '其他錯誤');
    commit;
end;
/
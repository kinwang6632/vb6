/*
@c:\EMC_script\SP_BackupCustData
set serveroutput on
exec SP_BackupCustData;  

  EMC每個月備份重要資料: 
  1. 客戶重要資料檔(so033.34中收費項目=301, 有實收日期.實收金額,沒有未收原因,沒有短收原因或短收原因的參考號<>1
	客戶狀態正常) [檔名範例: EMCYMS_SO001ALL_200305]
  2. 客戶設備資料檔(FaciCode之參考號為3,4者) [檔名範例: EMCYMS_SO004_200305]
  3. 大樓客戶數檔: 每月將各SO的SO017中有關客戶數與合約期限資料等部份欄位備份起來 [檔名範例: EMCYMS_SO017_200305]

  檔名: SP_BackupCustData.sql

  流程: 
  . 先取的公司名稱,與SYSDATE串成TABLENAME
  . 依TABLENAME,CREATE TABLE
  . 在將資料新增至新CREATE的TABLE

  傳回碼代碼說明v_RetCode:
	0:'備份資料完成',
	-1:'TABLE無法刪除',
	-2:'TABLE無法建立',
	-3:'新增時發生錯誤',
	-4:'取公司代碼發生錯誤',
	-5:'取公司名稱發生錯誤',
	-99:'其他錯誤'

  Date: 2003.05.13 By:Morris
	2003.05.30 改: 因應以下規則異動
	   1. 檔案命名原則修改, 以便備份其他檔案時, 可以取不同的名稱. 各種備份檔案命名原則為:
		系統台名稱_SO檔名_西曆年月 ==> 例: EMCYMS_SO001ALL_200305,  EMCYMS_SO004_200305

	   2. 原備份之客戶重要資料檔名稱因此由EMCYMS200305, 改成EMCYMS_SO001ALL_200305

	   3. 新增兩種重要資料的備份:
		. 客戶設備資料檔: 每月將各SO的SO004(FaciCode之參考號為3,4者)備份起來, 例: EMCYMS_SO004_200305
		. 大樓客戶數檔: 每月將各SO的SO017中有關客戶數與合約期限資料等部份欄位備份起來, 例: EMCYMS_SO017_200305
	2003.06.06 改: 因為北桃園目前的收費項目代碼尚未與EMC其他系統台統一, 其使用11, 故將程式改成若該系統台有301項目,
			則用該代碼值統計, 若無301代碼, 則用11來統計.
	2003.07.18 改: SO017的有效戶數拆為個收有效戶數(PerInstCnt)及統收有效戶數(InstCnt), 備份時也要將新欄位一起備份

	2003.08.20 改: 加上備份SO033/34的Classcode/CustCount二個欄位
	2003.08.26 改: 更正PerInstCnt, PerInstCnt1的欄位名稱誤輸入bug

	2003.09.04 改: 加上呼叫SP_Gen5DRptData1: EMC營運管理報表資料(SO5D01/02/03/04/07)產生作業
			目的是希望能於備份完畢後, 立刻做營運管理報表資料產生作業

  1. 客戶重要資料檔: 
	CustId 		NUMBER(8)
	CustName	VARCHAR2(50)
	Tel1 		VARCHAR2(20)
	Tel2		VARCHAR2(20)
	Tel3		VARCHAR2(20)
	InstTime 	DATE
	StopTime 	DATE
	CustStatusCode	NUMBER(3)
	CustStatusName	VARCHAR2(12)
	ViewLevel	CHAR(1)
	InstCount	NUMBER(4)
	ClassCode1	NUMBER(3)
	ClassName1	VARCHAR2(20)
	CMCode		NUMBER(3)
	CMName		VARCHAR2(20)
	CitemCode	NUMBER(3)
	CitemName	VARCHAR2(20)
	Period 		NUMBER(2)
	Amount 		NUMBER(8)
	StartDate 	DATE
	StopDate 	DATE
	ClctDate 	DATE
	InstAddrNo 	NUMBER(8)
	InstAddress 	VARCHAR2(60)
	UPDDate 	DATE
	CompCode 	NUMBER(3)
	CompName 	VARCHAR2(20)
	ServiceType 	VARCHAR2(50)
	ServiceName 	VARCHAR2(150)
	Billno 		VARCHAR(15)
	Realstartdate 	DATE
	Realstopdate 	DATE
	Realdate 	DATE
	Realperiod 	NUMBER(2)
	RealAmt 	NUMBER(8)
	CClassCode	number(3)
	CustCount	number(4)

  2. 客戶設備資料檔: 結構完全同SO004

  3. 大樓客戶數檔:
	CompCode	number(3)	公司別
	MduId		varchar2(8)	大樓編號
	Name		varchar2(20)	大樓名稱
	ClctMethod	number(1)	計費方式
	MainCustId	number(8)	統收戶客戶編號
	MainCustName	varchar2(26)	統收客戶名稱, vc2(50)
	BCAmount	number(8)	基本台統收金額
	BCPeriod	number(2)	基本台期數
	TotalCnt	number(4)	總戶數
	AcceptCnt	number(4)	建檔戶數
	InstCnt		number(4)	有效戶數
	UnInstCnt	number(4)	停拆/待拆戶數
	UnPRCnt		number(4)	待拆戶數
	AcceptCnt1	number(4)	CAISP建檔數
	InstCnt1	number(4)	CAISP有效戶數
	UnInstCnt1	number(4)	CAISP停拆/待拆戶數
	Flag		varchar2(6)	狀態
	CStartDate	date		合約起始日
	CDueDate	date		合約截止日
	CSignDate	date		簽約日
	CSignEn		varchar2(10)	簽約人員代號
	CSignName	varchar2(20)	簽約人員姓名
	ContractStatus	number(1)	合約狀態
*/

CREATE OR REPLACE PROCEDURE SP_BackupCustData  IS
  v_NewTableName varchar2(100);      -- New Table Name
  v_SQL		varchar2(4000);      -- for SQL statement
  v_Cnt1	number:=0;	     -- counter
  v_RcdCount	number:=0;
  v_Time	Date;
  v_CompCode	number;
  v_CompName	varchar2(20);
  v_JobID	number;
  v_StartExecTime   Date;
  v_ServiceType     char(1);
  v_RetMsg	varchar2(100);
  v_RetCode	number;
  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_CitemCode   number:=0;

BEGIN
  -- *************************************************************
  -- 0. 設定參數初值
  -- *************************************************************
  v_StartExecTime := sysdate;	
  v_JobID	:= 999992;
  v_StartExecTime := sysdate;
  v_ServiceType := 'C';
  v_FuncName	:= '客戶資料備份';
  v_PrgName	:= 'SP_BACKUPCUSTDATA';

  -- 若該系統台有301項目, 則用該代碼值統計, 若無301代碼, 則用11來統計	2003.06.06
  select count(*) into v_cnt1 from CD019 where CodeNo=301;
  if v_cnt1>0 then
    v_CitemCode := 301;
  else
    v_CitemCode := 11;
  end if;

  -- 取公司代碼
  begin 
    select compcode into v_compcode from so501a where rownum<=1;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      v_RetMsg := '取公司代碼時發生錯誤'||SQLERRM;
      v_RetCode := -4;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;


  -- *************************************************************
  -- 1. 備份: 客戶相關重要資料
  -- *************************************************************
  ---------------------------------建立暫存檔backupso5d01---------------------------- 
  select count(*) into v_cnt1 from user_tables where table_name='BACKUPSO5D01';
  if v_cnt1 > 0 then 		-- 先刪除同名檔案
    v_SQL := 'drop table BACKUPSO5D01';
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
	DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	v_RetMsg := v_NewTableName||'無法刪除'||SQLERRM;
	v_RetCode := -1;
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
	commit;
	return;
    end;
  end if;  

  -- 建立暫存檔
  v_SQL:='create table BACKUPSO5D01 as (    
    select custid,billno,realstartdate,realstopdate,realdate,realperiod,realamt,ClassCode,CustCount from so033
	where citemcode='||v_CitemCode||' and realamt>0 and realdate is not null and 
	(uccode=0 or uccode is null) and cancelflag=0 and 
	(stcode not in (select codeno from cd016 where refno=1) or stcode is null) and 
	custid in (select custid from so002 where custstatuscode=1)
    union all
    select custid,billno,realstartdate,realstopdate,realdate,realperiod,realamt,ClassCode,CustCount from so034
	where citemcode='||v_CitemCode||' and realamt>0 and realdate is not null and 
	(uccode=0 or uccode is null) and cancelflag=0 and  
	(stcode not in (select codeno from cd016 where refno=1) or stcode is null) and 
	custid in (select custid from so002 where custstatuscode=1))';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT (v_SQL); 
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      v_RetMsg := 'TABLE無法建立,請稍後再試'||SQLERRM;
      v_RetCode := -2;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  -------------------------------建立暫存檔backupso5d01a---------------------------  
  select count(*) into v_cnt1 from user_tables where table_name='BACKUPSO5D01A';
  if v_cnt1 > 0 then 		-- 先刪除同名檔案
    v_SQL := 'drop table BACKUPSO5D01A';
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
	DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	v_RetMsg := v_NewTableName||'無法刪除'||SQLERRM;
	v_RetCode := -1;
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
	commit;
	return;
    end;
  end if;

  -- 建立暫存檔
  v_SQL:='create table BACKUPSO5D01A as(
    select a.custid,a.billno,a.realstartdate,a.realstopdate,a.realdate,a.realperiod,a.realamt, 
	a.ClassCode,a.CustCount 
	from backupso5d01 a, (select custid,max(billno) billno from backupso5d01 group by custid ) b
	where a.custid=b.custid and a.billno=b.billno)';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT (v_SQL); 
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      v_RetMsg := 'TABLE無法建立,請稍後再試'||SQLERRM;
      v_RetCode := -2;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;
  
  ----------------------------creat備份資料檔-----------------------------------------------------
  -- 取系統台代號(一定要大寫), 例: 'EMCYMS', 'EMCKP'
  begin
    select upper(a.description) into v_compname from so507 a,so501a b where a.codeno=v_compcode;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      v_RetMsg := '取公司名稱時發生錯誤'||SQLERRM;
      v_RetCode := -5;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
    return;
  end;  

  -- 組成客戶資料檔備份檔名, 例: EMCYMS_SO001ALL_200305
  v_time:=ADD_MONTHS(SYSDATE,-1);
  v_NewTableName:=v_compname||'_SO001ALL_'||to_char(v_time,'YYYYMM');

  -- 先刪除相同檔名的table
  select count(*) into v_cnt1 from user_tables where table_name=v_NewTableName;
  if v_cnt1 > 0 then 
    v_SQL := 'drop table  ' || v_NewTableName;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
	DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	v_RetMsg := v_NewTableName||'無法刪除'||SQLERRM;
	v_RetCode := -1;
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
	commit;
	return;
    end;
  end if; 

  -- 建立備份資料檔
  v_SQL := 'CREATE TABLE  '|| v_NewTableName || ' ( ' ||
	'CustId NUMBER(8),CustName VARCHAR2(50) NOT NULL,Tel1 VARCHAR2(20),
	Tel2 VARCHAR2(20),Tel3 VARCHAR2(20),InstTime DATE, StopTime DATE,
	CustStatusCode NUMBER(3) DEFAULT 1 NOT NULL,CustStatusName VARCHAR2(12),
	ViewLevel CHAR(1), InstCount NUMBER(4) DEFAULT 0,ClassCode1 NUMBER(3),
	ClassName1 VARCHAR2(20),CMCode NUMBER(3) DEFAULT 1 NOT NULL,
	CMName VARCHAR2(20),CitemCode NUMBER(3),CitemName VARCHAR2(20) NOT NULL,
	Period NUMBER(2),Amount NUMBER(8), StartDate DATE, StopDate DATE,
	ClctDate DATE, InstAddrNo NUMBER(8),InstAddress VARCHAR2(60),
	UPDDate DATE,CompCode NUMBER(3) DEFAULT 1,CompName VARCHAR2(20),
	ServiceType VARCHAR2(50),ServiceName VARCHAR2(150),Billno VARCHAR(15),
	Realstartdate DATE,Realstopdate DATE,Realdate DATE,Realperiod NUMBER(2),
	RealAmt NUMBER(8), CClassCode NUMBER(3), CustCount NUMBER(4))';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT (v_SQL); 
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      v_RetMsg := 'TABLE無法建立,請稍後再試'||SQLERRM;
      v_RetCode := -2;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  -- 將資料備份至此檔
  begin
    v_SQL:='Insert into '||v_NewTableName||' ( '||
	'CustId ,CustName ,Tel1 ,Tel2 ,Tel3 ,InstTime ,StopTime,
	CustStatusCode ,CustStatusName , ViewLevel ,InstCount ,ClassCode1 ,
	ClassName1,CMCode,CMName,CitemCode,CitemName,Period,Amount,StartDate,
	StopDate,ClctDate ,InstAddrNo ,InstAddress ,UPDDate ,CompCode ,CompName ,
	ServiceType,ServiceName,Billno,Realstartdate,Realstopdate,Realdate,
	Realperiod,RealAmt,CClassCode,CustCount) (select a.custid,a.custname,a.tel1,a.tel2,
	a.tel3,b.insttime,b.stoptime,custstatuscode,custstatusname,a.viewlevel,
	a.instcount,a.classcode1,a.classname1,b.cmcode,b.cmname,c.citemcode,c.citemname,
	c.period,c.amount,c.startdate,c.stopdate,c.clctdate,a.instaddrno,a.instaddress,
	a.nextbcdate,a.compcode,a.compname,a.servicetype,a.servicename,d.Billno,
	d.Realstartdate,d.Realstopdate,d.Realdate,d.Realperiod,d.RealAmt,d.ClassCode,d.CustCount
	from so001 a,so002 b,so003 c,backupso5d01a d
	where a.custid=b.custid and b.custstatuscode=1 and a.custid=c.custid 
	and c.citemcode='||v_CitemCode||' and a.custid=d.custid(+))';
    EXECUTE IMMEDIATE v_SQL;
    v_RcdCount := v_RcdCount + sql%rowcount;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      v_RetMsg := 'INSERT時發生錯誤: '||SQLERRM;
      v_RetCode := -3;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg); 
      commit;
      return;     
  end;

  ----------------------------drop table backupso5d01,backupso5d01a---------------------------
  v_SQL := 'drop table BACKUPSO5D01';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      v_RetMsg := v_NewTableName||'無法刪除'||SQLERRM;
      v_RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;
   
  v_SQL := 'drop table BACKUPSO5D01A';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      v_RetMsg := v_NewTableName||'無法刪除'||SQLERRM;
      v_RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;
  ----------------------------------------------------------------------------------------


  -- *************************************************************
  -- 2. 備份: 客戶設備資料檔(FaciCode之參考號為3,4者)
  -- *************************************************************
  -- 組成客戶設備資料檔備份檔名, 例: EMCYMS_SO004_200305
  v_NewTableName:=v_compname||'_SO004_'||to_char(v_time,'YYYYMM');

  -- 先刪除相同檔名的table
  select count(*) into v_cnt1 from user_tables where table_name=v_NewTableName;
  if v_cnt1 > 0 then 
    v_SQL := 'drop table  ' || v_NewTableName;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
	DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	v_RetMsg := v_NewTableName||'無法刪除'||SQLERRM;
	v_RetCode := -1;
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
	commit;
	return;
    end;
  end if; 

  -- 將資料備份至此檔
  begin
    v_SQL := 'create table '||v_NewTableName||' as (select a.* from so004 a, cd022 b
	where a.Facicode=b.CodeNo and b.RefNo in (3,4))';
    EXECUTE IMMEDIATE v_SQL;
    v_RcdCount := v_RcdCount + sql%rowcount;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      v_RetMsg := 'INSERT時發生錯誤: '||SQLERRM;
      v_RetCode := -3;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg); 
      commit;
      return;     
  end;
  ----------------------------------------------------------------------------------------

  -- *************************************************************
  -- 3. 備份: 大樓客戶數檔
  -- *************************************************************
  -- 組成大樓客戶數檔備份檔名, 例: EMCYMS_SO017_200305
  v_NewTableName:=v_compname||'_SO017_'||to_char(v_time,'YYYYMM');

  -- 先刪除相同檔名的table
  select count(*) into v_cnt1 from user_tables where table_name=v_NewTableName;
  if v_cnt1 > 0 then 
    v_SQL := 'drop table  ' || v_NewTableName;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
	DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
	v_RetMsg := v_NewTableName||'無法刪除'||SQLERRM;
	v_RetCode := -1;
	insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
	commit;
	return;
    end;
  end if;

  -- 將資料備份至此檔: 	2003.07.18 加備份欄位 PerInstCnt, PerInstCnt1
  begin
    v_SQL := 'create table '||v_NewTableName||' as (select CompCode, MduId, Name, ClctMethod, 
	MainCustId, MainCustName, BCAmount, BCPeriod, TotalCnt, AcceptCnt, InstCnt, UnInstCnt, 
	UnPRCnt, AcceptCnt1, InstCnt1, UnInstCnt1, Flag, CStartDate, CDueDate, CSignDate, 
	CSignEn, CSignName, ContractStatus, PerInstCnt, PerInstCnt1 from SO017)';
    EXECUTE IMMEDIATE v_SQL;
    v_RcdCount := v_RcdCount + sql%rowcount;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      v_RetMsg := 'INSERT時發生錯誤: '||SQLERRM;
      v_RetCode := -3;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg); 
      commit;
      return;     
  end;
  ----------------------------------------------------------------------------------------

  -- *************************************************************
  -- Log執行結果
  -- *************************************************************
  v_RetMsg :='共新增'||v_RcdCount||'筆, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'秒';
  v_RetCode :=0;
  DBMS_OUTPUT.PUT_LINE('執行完畢'|| v_RetMsg); 
  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg); 
  COMMIT;

  -- *************************************************************
  -- 呼叫SP_Gen5DRptData1: EMC營運管理報表資料(SO5D01/02/03/04/07)產生作業	
  -- 目的: 是希望能於備份完畢後, 立刻做營運管理報表資料產生作業
  -- 注意: 不於此處log此函式執行結果, 由該支程式內部處理
  -- Date: 2003.09.04
  -- *************************************************************
  SP_Gen5DRptData1;


EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code='||SQLCODE||' '|| 'Error Message='||SQLERRM);
    v_RetMsg := '其他錯誤 :'||SQLERRM;
    v_RetCode := -99;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
    return;
END;
/




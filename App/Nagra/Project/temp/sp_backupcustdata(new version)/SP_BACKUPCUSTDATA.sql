CREATE OR REPLACE PROCEDURE SP_BACKUPCUSTDATA  IS
v_NewTableName    varchar2(100);       -- New Table Name
v_SQL             varchar2(4000);      -- for SQL statement
v_Cnt1            number:=0;	       -- counter
v_Cnt2            number:=0;              --�p��s�W����
v_Time            Date;
v_CompCode        number;
v_CompName        varchar2(20);
v_JobID           number;
v_StartExecTime   Date;
v_ServiceType     char(1);
RetMsg            varchar2(100);
RetCode           number;
/*
--@c:\gird\V300\script\SP_BACKUPCUSTDATA
exec SP_BACKUPCUSTDATA;  
  �C�Ӥ�ƥ��ŦX����(so033.34�����O����=301
  ���ꦬ���.�ꦬ���B,�S��������],�S���u����]�εu����]���ѦҸ�<>1
  �Ȥ᪬�A���`)���Ȥ���
  �ɦW: SP_BACKUPCUSTDATA
  ���������q�W��,�PSYSDATE�ꦨTABLENAME,��TABLENAME,CREATE TABLE
  �b�N��Ʒs�W�ܷsCREATE��TABLE.
�N�X����
    RetCode: 0:'�ƥ���Ƨ���',
             -1:'TABLE�L�k�R��',
             -2:'TABLE�L�k�إ�',
             -3:'�s�W�ɵo�Ϳ��~',
             -4:'�����q�N�X�o�Ϳ��~',
             -5:'�����q�W�ٵo�Ϳ��~',
             -99:'��L���~'
Date: 2003/04/21
By:Morris
*/
   
BEGIN
  v_JobID:=999992;
  v_StartExecTime:=sysdate;
  v_ServiceType:='C';
---------------------------------creat�Ȧs��backupso5d01---------------------------- 
  select count(*) into v_cnt1 from user_tables where table_name='backupso5d01';
   if v_cnt1 > 0 then 
    v_SQL := 'drop table backupso5d01';
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
        RetMsg := v_NewTableName||'�L�k�R��'||SQLERRM;
	RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg);
      commit;
      return;
    end;
   end if;  
    
   v_SQL:='create table backupso5d01 as (    
	select custid,billno,realstartdate,realstopdate,realdate,realperiod,realamt from so033
      where citemcode=301 and realamt>0 and realdate is not null and 
            (uccode=0 or uccode is null) and cancelflag=0 and 
    		(stcode not in (select codeno from cd016 where refno=1) or stcode is null) and 
    		custid in (select custid from so002 where custstatuscode=1)
    union all
    select custid,billno,realstartdate,realstopdate,realdate,realperiod,realamt from so034
      where citemcode=301 and realamt>0 and realdate is not null and 
            (uccode=0 or uccode is null) and cancelflag=0 and  
    		(stcode not in (select codeno from cd016 where refno=1) or stcode is null) and 
    		custid in (select custid from so002 where custstatuscode=1))'
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT (v_SQL); 
  exception
  when others then
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    RetMsg := 'TABLE�L�k�إ�,�еy��A��'||SQLERRM;
    RetCode := -2;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	'�Ȥ��Ƴƥ�', RetCode, RetMsg);
  commit;
  return;
  end;
-------------------------------create�Ȧs��backupso5d01a---------------------------  
 select count(*) into v_cnt1 from user_tables where table_name='backupso5d01a';
   if v_cnt1 > 0 then 
    v_SQL := 'drop table backupso5d01a';
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
        RetMsg := v_NewTableName||'�L�k�R��'||SQLERRM;
	RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg);
      commit;
      return;
    end;
   end if;
   
   v_SQL:='create table backupso5d01a as(
   select a.custid,a.billno,a.realstartdate,a.realstopdate,a.realdate,a.realperiod,a.realamt 
     from backupso5d01 a,  
   (select custid,max(billno) billno from backupso5d01 group by custid ) b
   where a.custid=b.custid and a.billno=b.billno)'
  
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT (v_SQL); 
  exception
  when others then
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    RetMsg := 'TABLE�L�k�إ�,�еy��A��'||SQLERRM;
    RetCode := -2;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	'�Ȥ��Ƴƥ�', RetCode, RetMsg);
  commit;
  return;
  end;
        
  
----------------------------creat�ƥ������-----------------------------------------------------
  begin 
    select compcode into v_compcode from so501a where rownum<=1;
  exception
    when others then
    DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
    RetMsg := '�����q�N�X�ɵo�Ϳ��~'||SQLERRM;
    RetCode :=-4;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg);
    commit;
    return;
  end;
  begin
  select a.description into v_compname from so507 a,so501a b
   where a.codeno=v_compcode;
  exception
    when others then
    DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
    RetMsg:='�����q�W�ٮɵo�Ϳ��~'||SQLERRM;
    RetCode :=-5;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	 TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	 values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	 trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	 '�Ȥ��Ƴƥ�', RetCode, RetMsg);
    commit;
    return;
  end;  
   v_time:=SYSDATE;
   v_NewTableName:=v_compname||to_char(v_time,'YYYYMM');
   
   
   select count(*) into v_cnt1 from user_tables where table_name=v_NewTableName;
   if v_cnt1 > 0 then 
    v_SQL := 'drop table  ' || v_NewTableName;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
        RetMsg := v_NewTableName||'�L�k�R��'||SQLERRM;
	RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg);
      commit;
      return;
    end;
   end if; 
   v_SQL := 'CREATE TABLE  '|| v_NewTableName || ' ( ' ||
            'CustId NUMBER(8),CustName VARCHAR2(50) NOT NULL,Tel1 VARCHAR2(20),
	         Tel2 VARCHAR2(20),Tel3 VARCHAR2(20),InstTime DATE,	StopTime DATE,
             CustStatusCode NUMBER(3) DEFAULT 1 NOT NULL,CustStatusName VARCHAR2(12),
	         ViewLevel CHAR(1),	InstCount NUMBER(4) DEFAULT 0,ClassCode1 NUMBER(3),
	         ClassName1 VARCHAR2(20),CMCode NUMBER(3) DEFAULT 1 NOT NULL,
	         CMName VARCHAR2(20),CitemCode NUMBER(3),CitemName VARCHAR2(20) NOT NULL,
	         Period NUMBER(2),Amount NUMBER(8),	StartDate DATE,	StopDate DATE,
	         ClctDate DATE,	InstAddrNo NUMBER(8),InstAddress VARCHAR2(60),
	         UPDDate DATE,CompCode NUMBER(3) DEFAULT 1,CompName VARCHAR2(20),
	         ServiceType VARCHAR2(50),ServiceName VARCHAR2(150),Billno VARCHAR(15),
             Realstartdate DATE,Realstopdate DATE,Realdate DATE,Realperiod NUMBER(2),
             RealAmt NUMBER(8))';
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT (v_SQL); 
      exception
      when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      RetMsg := 'TABLE�L�k�إ�,�еy��A��'||SQLERRM;
      RetCode := -2;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg);
      commit;
      return;
    end;
    begin
      v_SQL:='Insert into '||v_NewTableName||' ( '||
      'CustId ,CustName ,Tel1 ,Tel2 ,Tel3 ,InstTime ,StopTime,
       CustStatusCode ,CustStatusName , ViewLevel ,InstCount ,ClassCode1 ,
       ClassName1,CMCode,CMName,CitemCode,CitemName,Period,Amount,StartDate,
       StopDate,ClctDate ,InstAddrNo ,InstAddress ,UPDDate ,CompCode ,CompName ,
       ServiceType,ServiceName,Billno,Realstartdate,Realstopdate,Realdate,
       Realperiod,RealAmt'||') '||' ('||'select a.custid,a.custname,a.tel1,a.tel2,
	   a.tel3,b.insttime,b.stoptime,custstatuscode,custstatusname,a.viewlevel,
	   a.instcount,a.classcode1,a.classname1,b.cmcode,b.cmname,c.citemcode,c.citemname,
	   c.period,c.amount,c.startdate,c.stopdate,c.clctdate,a.instaddrno,a.instaddress,
	   a.nextbcdate,a.compcode,a.compname,a.servicetype,a.servicename,d.Billno,
       d.Realstartdate,d.Realstopdate,d.Realdate,d.Realperiod,d.RealAmt
       from so001 a,so002 b,so003 c,backupso5d01a d
       where a.custid=b.custid and b.custstatuscode=1 and a.custid=c.custid 
       and c.citemcode=301 and a.custid=d.custid(+)'||')';
       
      EXECUTE IMMEDIATE v_SQL;
    
      RetMsg :='�s�W����';
      RetCode :=0;

      exception
      when others then
      DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      RetMsg := 'INSERT�ɵo�Ϳ��~: '||SQLERRM;
      RetCode := -3;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg); 
      commit;
      return;     
    end;
----------------------------drop table backupso5d01,backupso5d01a---------------------------
  
    v_SQL := 'drop table backupso5d01';
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
        RetMsg := v_NewTableName||'�L�k�R��'||SQLERRM;
	RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg);
      commit;
      return;
    end;
   
    v_SQL := 'drop table backupso5d01a';
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
        RetMsg := v_NewTableName||'�L�k�R��'||SQLERRM;
	RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	  '�Ȥ��Ƴƥ�', RetCode, RetMsg);
      commit;
      return;
    end;

----------------------------------------------------------------------------------------
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	'�Ȥ��Ƴƥ�', RetCode, RetMsg);    
    COMMIT;
EXCEPTION
    WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code='||SQLCODE||' '|| 'Error Message='||SQLERRM);
    RetMsg := '��L���~ :'||SQLERRM;
    RetCode := -99;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), 'SP_BACKUPCUSTDATA', 
	'�Ȥ��Ƴƥ�', RetCode, RetMsg);
    commit;
    return;
END; -- Procedure

/

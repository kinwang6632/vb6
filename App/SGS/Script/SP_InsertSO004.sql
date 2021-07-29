/*
@D:\SGS\SP_InsertSO004

set transaction use rollback segment RBS27;
set serveroutput on
exec SP_InsertSO004

  For 3,000,000 test data: 
	(1) Insert 3,000,000 ICC and 3,000,000 STB records into SO004
	(2) Insert 3,000,000 ICC records into SOAC0201B
	(3) Insert 3,000,000 STB records into SOAC0201A

  File name: SP_InsertSO004.sql
  
  Note:
  	   (1) specify certain Oracle Rollback segment for this program
	   (2) drop some index before execution of this program

  By: Pierre
  Date: 2004.04.05
	2004.04.06 
	2004.04.07 改更新SO451的做法, 使用流水號SerialNo來每次取50筆資料

*/

create or replace procedure SP_InsertSO004
as
  v_StartExecTime date;			-- Start time
  v_StopExecTime date;			-- Stop time
  v_TtlSecond number;			-- Total execution time (in seconds)
  v_Hour number;				-- Hour part of total execution time
  v_Minute number;				-- Minute part of total execution time
  v_Second number;				-- Second part of total execution time
  v_IccTotalCnt number := 0;	-- Total ICC count, start from 0, max=3,000,000 
  v_IccEachCnt number := 0;		-- ICC count for each customer, max=50
  TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  RowIdList Varchar2Ary;	    -- Array for rowid
  IccNoList Varchar2Ary;	    -- Array for ICC number
  TYPE DateAry IS TABLE OF date INDEX BY BINARY_INTEGER;
  InstDateList DateAry;	   		-- Array for ICC installation date
  i number;  					-- for loop index
  j number;  					-- for loop index
  v_FaciCode1 number;			-- ICC code no in CD024
  v_FaciName1 varchar2(20);		-- ICC name in CD024
  v_FaciCode2 number;			-- STB code no in CD024
  v_FaciName2 varchar2(20);		-- STB name in CD024
  v_BuyCode number;				-- Buy code for saling STB/ICC
  v_BuyName varchar2(20);		-- Buy name for saling STB/ICC
  v_InitPlaceNo number;			-- Location code for the STB/ICC
  v_ServiceType char(1);		-- Service type code
  v_CompCode number;			-- Company code
  v_NextVal number;				-- next available sequence number
  v_SeqNo1 varchar2(15);		-- Serial number for ICC
  v_SeqNo2 varchar2(15);		-- Serial number for STB
  v_UpdEn varchar2(20);			-- Update EN name
  v_Now date;					-- Current time
  v_BatchNo1 varchar2(15);		-- Batch number for ICC records
  v_BatchNo2 varchar2(15);		-- Batch number for STB records
  v_MaterialNo1 varchar2(15);	-- Material number for ICC records
  v_MaterialNo2 varchar2(15);	-- Material number for STB records
  v_FetchCnt number := 50;	-- Fetch count for SO451

  -- all O002 Normal customers
  cursor cc1 is
  	select CustId from SO002 where CustStatusCode=1 and ServiceType='C';

  -- next 50 ICC records
  --cursor cc2 is
  --	select RowId,IcCardNo,RealDateTime from SO451 where nvl(Flag,0)!=1 and RowNum<=50;
  cursor cc2(p_S1 number, p_S2 number) is
  	select RowId,IcCardNo,RealDateTime from SO451 where SerialNo between p_S1 and p_S2;

begin
  -- *******************************************
  -- Initialization
  -- *******************************************
  DBMS_OUTPUT.PUT_LINE('*** SP_InsertSO004 ***');
  v_StartExecTime := sysdate;		-- Start time
  v_Now := v_StartExecTime;			-- Current time
  DBMS_OUTPUT.PUT_LINE('Start time = '||to_char(v_StartExecTime,'YYYY.MM.DD HH24:MI:SS'));
  
  v_InitPlaceNo := 1;		  -- location code for the STB/ICC
  v_ServiceType := 'C';		  -- Service type code
  v_CompCode := 12;	   		  -- Company code
  v_UpdEn := 'SP_INSERTSO004'; -- Update EN name
  v_BatchNo1 := '1';		   -- Batch number for ICC records
  v_BatchNo2 := '2';		   -- Batch number for STB records
  v_MaterialNo1 := 'A123';     -- Material number for ICC records
  v_MaterialNo2 := 'A456';     -- Material number for STB records

  -- read ICC codeno and name for later use
  begin
    select CodeNo, Description into v_FaciCode1, v_FaciName1 from CD022 where RefNo=4 and RowNum<=1;
  EXCEPTION
    WHEN OTHERS THEN
	  DBMS_OUTPUT.PUT_LINE('No ICC record in CD022');
	  return;
  end;
  
  -- read STB codeno and name for later use
  begin
    select CodeNo, Description into v_FaciCode2, v_FaciName2 from CD022 where RefNo=3 and RowNum<=1;
  EXCEPTION
    WHEN OTHERS THEN
	  DBMS_OUTPUT.PUT_LINE('No STB record in CD022');
	  return;
  end;
  
  -- read buy code and name for later use
  begin
    select CodeNo, Description into v_BuyCode, v_BuyName from CD034 where CodeNo=1;
  EXCEPTION
    WHEN OTHERS THEN
	  DBMS_OUTPUT.PUT_LINE('No BuyCode record in CD034');
	  return;
  end;
 
 
  -- *******************************************
  -- loop each normal customer to add max 50 ICC and 50 STB records
  -- *******************************************
  for cr1 in cc1 loop
    if v_IccTotalCnt <= 3000000 then  -- no matter how many customers, only handles 3000000 ICC records
      i := 0;
	  for cr2 in cc2(v_IccTotalCnt+1,v_IccTotalCnt+v_FetchCnt) loop
	    i := i + 1;	  	 			   	 -- increment index count for current customer
	    RowIdList(i) := cr2.RowId;				 -- save RowId in array
	    IccNoList(i) := cr2.IcCardNo;			 -- save ICC number in array
	    InstDateList(i) := trunc(cr2.RealDateTime); -- save ICC installtion date in array
	  end loop;
	  v_IccTotalCnt := v_IccTotalCnt + i;	 	 -- increment total ICC record count
	  
	  --for j in 1..i loop						 -- mark handle flag for these ICC records
	  --  update SO451 set Flag=1 where RowId=RowIdList(j);
	  --end loop;
	  
	  for j in 1..i loop 		  				-- insert those data into SO004, SOAC0201B, SOAC0201B
	    -- Insert those ICC data into SO004
	    select S_SO004.NextVal into v_NextVal from dual;
	    v_SeqNo1 := to_char(InstDateList(j),'YYYYMMDD')||ltrim(to_char(v_NextVal,'0999999'));
	    insert into SO004 (CustId,FaciCode,FaciName,FaciSNo,BuyCode,BuyName,InstDate,Quantity,SmartCardNo,
			   ServiceType,CompCode,InitPlaceNo,SeqNo,FaciStatusCode,UpdEn) values
			   (cr1.CustId,v_FaciCode1,v_FaciName1,IccNoList(j),v_BuyCode,v_BuyName,InstDateList(j),1,null,
			   v_ServiceType,v_CompCode,v_InitPlaceNo,v_SeqNo1,1,v_UpdEn);

	    -- Insert those STB data into SO004
	    -- STB sno is right 12 bytes of ICC card number
	    select S_SO004.NextVal into v_NextVal from dual;
	    v_SeqNo2 := to_char(InstDateList(j),'YYYYMMDD')||ltrim(to_char(v_NextVal,'0999999'));
	    insert into SO004 (CustId,FaciCode,FaciName,FaciSNo,BuyCode,BuyName,InstDate,Quantity,SmartCardNo,
			   ServiceType,CompCode,InitPlaceNo,SeqNo,FaciStatusCode,UpdEn) values
			   (cr1.CustId,v_FaciCode2,v_FaciName2,substr(IccNoList(j),5),v_BuyCode,v_BuyName,InstDateList(j),1,IccNoList(j),
			   v_ServiceType,v_CompCode,v_InitPlaceNo,v_SeqNo2,1,v_UpdEn);

	    -- Insert those ICC data into SOAC0201B
	    insert into SOAC0201B (CompCode,BatchNo,ExecTime,MaterialNo,FaciSNo,Description,UseFlag,UpdEn) values 
			   (v_CompCode,v_BatchNo1,v_Now,v_MaterialNo1,v_SeqNo1,'ICC',0,v_UpdEn);

	    -- Insert those STB data into SOAC0201A
	    insert into SOAC0201A (CompCode,BatchNo,ExecTime,MaterialNo,FaciSNo,Description,UseFlag,UpdEn) values 
			   (v_CompCode,v_BatchNo2,v_Now,v_MaterialNo2,v_SeqNo2,'STB',0,v_UpdEn);

	  end loop;
	end if;
  end loop;

  
  commit;
  -- *******************************************
  -- End
  -- *******************************************
  -- Count execution time
  v_StopExecTime := sysdate;
  v_TtlSecond := round(86400*(v_StopExecTime-v_StartExecTime));
  v_Hour := trunc(v_TtlSecond/3600);
  v_Second := Mod(v_TtlSecond,60);
  v_Minute := trunc((v_TtlSecond - v_Hour*3600)/60);
  DBMS_OUTPUT.PUT_LINE('Stop time  = '||to_char(v_StopExecTime, 'YYYY.MM.DD HH24:MI:SS'));
  DBMS_OUTPUT.PUT_LINE('*** Total time = '||v_Hour||':'||v_Minute||':'||v_Second);

exception
  when others then
    rollback;
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||','||'SQLERRM='||SQLERRM);
    v_StopExecTime := sysdate;
    v_TtlSecond := round(86400*(v_StopExecTime-v_StartExecTime));
    v_Hour := trunc(v_TtlSecond/3600);
    v_Second := Mod(v_TtlSecond,60);
    v_Minute := trunc((v_TtlSecond - v_Hour*3600)/60);
    DBMS_OUTPUT.PUT_LINE('Stop time  = '||to_char(v_StopExecTime, 'YYYY.MM.DD HH24:MI:SS'));
    DBMS_OUTPUT.PUT_LINE('*** Total time = '||v_Hour||':'||v_Minute||':'||v_Second);
end;
/

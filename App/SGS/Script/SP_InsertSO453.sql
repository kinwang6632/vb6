CREATE OR REPLACE PROCEDURE SP_INSERTSO453
as

  -- Current SO453 records [should have 20 records]
  cursor cc1 is
  	select RowId, CompCode, HandleEDateTime, RealDateTime, ExecDateTime, ProductCode, ProductName,
		CasID, Provider, ChannelName from SO453;

  TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  RowIdList Varchar2Ary;	    -- Array for rowid
  v_CompCode number;
  v_FirstDate date;
  v_NextDate date;
  v_Now date;
  i number := 0;
  j number := 0;
  h number := 0;
  v_HandleEDateTime Varchar2(20) := ' ';

begin
  -- *******************************************
  -- Initialization
  -- *******************************************
  DBMS_OUTPUT.PUT_LINE('*** SP_InsertSO453 ***');
  v_CompCode := 12;					-- Company code
  v_FirstDate := to_date('20030818','YYYYMMDD');	-- First date
  v_Now := sysdate;

  -- Delete duplicated records
  delete from SO453 where HandleEDateTime>v_FirstDate+1;

  -- Update old 20 records
  --update SO453 set HandleEDateTime=v_FirstDate;
  update SO453 set HandleEDateTime=TO_DATE(TO_CHAR(v_FirstDate,'YYYY/MM/DD') || ' 23:59:59','YYYY/MM/DD HH24:MI:SS');

  -- Read RowId of current 20 records into array RowIdList
  for cr1 in cc1 loop
    i := i + 1;
    RowIdList(i) := cr1.RowId;
  end loop;

  -- Loop 99 times to insert 1800 records into SO453
  for cr1 in cc1 loop
    v_NextDate := v_FirstDate;
    
    
    for j in 1 .. 99 loop
      v_NextDate := v_NextDate + 1;
      
      v_NextDate := TO_DATE(TO_CHAR(v_NextDate,'YYYY/MM/DD') || ' 23:59:59','YYYY/MM/DD HH24:MI:SS');
      
      insert into SO453 (CompCode, HandleEDateTime, RealDateTime, ExecDateTime, ProductCode, ProductName,
	CasID, Provider, ChannelName) values (
	v_CompCode, v_NextDate, cr1.RealDateTime, v_Now, cr1.ProductCode, cr1.ProductName,
	cr1.CasID, cr1.Provider, cr1.ChannelName);
    end loop;
  end loop;

  commit;
  -- *******************************************
  -- End
  -- *******************************************
  DBMS_OUTPUT.PUT_LINE('*** OK');

exception
  when others then
    rollback;
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||','||'SQLERRM='||SQLERRM);
end;
/
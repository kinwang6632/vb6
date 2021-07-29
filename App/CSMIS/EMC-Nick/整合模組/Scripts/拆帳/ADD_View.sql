/*
  ADD_VIEW.SQL: Create views
  Naming rule: V_xxxxxx
  Date: 2002.08.27
*/
set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool C:\Temp
print str 



prompt V_ChargeInfoForTally: mis客戶繳費資料
drop view V_ChargeInfoForTally;
create view V_ChargeInfoForTally as
  select 33 Source, a.CompCode, a.CustId, a.BillNo, a.Item,
	a.CitemCode, a.CitemName, a.ShouldDate, 
	a.ShouldAmt,a.RealDate,a.RealAmt,a.RealPeriod,
	a.RealStartDate,a.RealStopDate,a.ClctName
 	from so033 a
	where a.CancelFlag=0
  union
  select 34 Source, a.CompCode, a.CustId, a.BillNo, a.Item,
	a.CitemCode, a.CitemName, a.ShouldDate, 
	a.ShouldAmt,a.RealDate,a.RealAmt,a.RealPeriod,
	a.RealStartDate,a.RealStopDate,a.ClctName
 	from so034 a
	where a.CancelFlag=0;


/* 用法: select * from V_ChargeInfoForTally */

CREATE INDEX I_V_ChargeInfoForTally_CompCodeReadDate ON V_ChargeInfoForTally (CompCode,RealDate);
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 300K MINEXTENTS 1 MAXEXTENTS 200) TABLESPACE GICMIS_NDX;
spool off
set heading on




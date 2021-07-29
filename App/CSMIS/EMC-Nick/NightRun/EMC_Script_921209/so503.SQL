drop table SO503 cascade constraint;

create table SO503 (
	SeqNo		number(8),
	CompCode 	number(3),
	CompName 	varchar2(20),
	ServiceType 	char(1),
	AcceptEn 	varchar2(10),
	AcceptName 	varchar2(26),
	Cnt1		number(8) default 0,
	Cnt2		number(8) default 0,
	Cnt3		number(8) default 0,
	Cnt4		number(8) default 0,
	Cnt5		number(8) default 0,
	CntTotal	number(8) default 0,
	RptDate		date,
	ReportTime	date,
	Operator	varchar2(20),
	MediaCode	number(3),
	MediaName	varchar2(20));

/*
create index I_SO503_SeqNo on SO503 (SeqNo);
create index I_SO503_RptDate1 on SO503 (SeqNo, RptDate, AcceptEn);
create index I_SO503_RptDate2 on SO503 (SeqNo, RptDate, AcceptEn, CompCode);
create index I_SO503_AcceptEn on SO503 (SeqNo, AcceptEn, RptDate, CompCode);
*/

drop view V_SO503Charge;
create view V_SO503Charge as
  select BillNo, Item, CitemCode, RealAmt
    from SO033 where RealDate is not null and UCCode is null and CancelFlag=0
  union
  select BillNo, Item, CitemCode, RealAmt
    from SO034 where RealDate is not null and UCCode is null and CancelFlag=0;


/*
create index I_SO503_SeqNo on SO503 (SeqNo)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE YMS_NDX;
create index I_SO503_RptDate1 on SO503 (SeqNo, RptDate, AcceptEn)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE YMS_NDX;
create index I_SO503_RptDate2 on SO503 (SeqNo, RptDate, AcceptEn, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE YMS_NDX;
create index I_SO503_AcceptEn on SO503 (SeqNo, AcceptEn, RptDate, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE YMS_NDX;

create index I_SO503_SeqNo on SO503 (SeqNo)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE LK_NDX;
create index I_SO503_RptDate1 on SO503 (SeqNo, RptDate, AcceptEn)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE LK_NDX;
create index I_SO503_RptDate2 on SO503 (SeqNo, RptDate, AcceptEn, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE LK_NDX;
create index I_SO503_AcceptEn on SO503 (SeqNo, AcceptEn, RptDate, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE LK_NDX;

create index I_SO503_SeqNo on SO503 (SeqNo)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE KP_NDX;
create index I_SO503_RptDate1 on SO503 (SeqNo, RptDate, AcceptEn)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE KP_NDX;
create index I_SO503_RptDate2 on SO503 (SeqNo, RptDate, AcceptEn, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE KP_NDX;
create index I_SO503_AcceptEn on SO503 (SeqNo, AcceptEn, RptDate, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE KP_NDX;

create index I_SO503_SeqNo on SO503 (SeqNo)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE NTP_NDX;
create index I_SO503_RptDate1 on SO503 (SeqNo, RptDate, AcceptEn)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE NTP_NDX;
create index I_SO503_RptDate2 on SO503 (SeqNo, RptDate, AcceptEn, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE NTP_NDX;
create index I_SO503_AcceptEn on SO503 (SeqNo, AcceptEn, RptDate, CompCode)
  PCTFREE 10 STORAGE (INITIAL 100K NEXT 200K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE NTP_NDX;
*/

-- 以下權限指令請使用適當的User連線至各資料區, 並執行對應的指令
/*
-- connect to EMCYMS
grant SELECT on CD019 to "EMCKP";
grant SELECT on CD039 to "EMCKP";
grant SELECT on CM003 to "EMCKP";
grant SELECT on SO007 to "EMCKP";
grant SELECT on SO008 to "EMCKP";
grant SELECT on SO009 to "EMCKP";
grant SELECT on V_SO503Charge to "EMCKP";

grant SELECT on CD019 to "EMCLK";
grant SELECT on CD039 to "EMCLK";
grant SELECT on CM003 to "EMCLK";
grant SELECT on SO007 to "EMCLK";
grant SELECT on SO008 to "EMCLK";
grant SELECT on SO009 to "EMCLK";
grant SELECT on V_SO503Charge to "EMCLK";

grant SELECT on CD019 to "EMCNTP";
grant SELECT on CD039 to "EMCNTP";
grant SELECT on CM003 to "EMCNTP";
grant SELECT on SO007 to "EMCNTP";
grant SELECT on SO008 to "EMCNTP";
grant SELECT on SO009 to "EMCNTP";
grant SELECT on V_SO503Charge to "EMCNTP";
*/

/*
-- connect to EMCLK
grant SELECT on CD019 to "EMCYMS";
grant SELECT on CD039 to "EMCYMS";
grant SELECT on CM003 to "EMCYMS";
grant SELECT on SO007 to "EMCYMS";
grant SELECT on SO008 to "EMCYMS";
grant SELECT on SO009 to "EMCYMS";
grant SELECT on V_SO503Charge to "EMCYMS";

grant SELECT on CD019 to "EMCKP";
grant SELECT on CD039 to "EMCKP";
grant SELECT on CM003 to "EMCKP";
grant SELECT on SO007 to "EMCKP";
grant SELECT on SO008 to "EMCKP";
grant SELECT on SO009 to "EMCKP";
grant SELECT on V_SO503Charge to "EMCKP";

grant SELECT on CD019 to "EMCNTP";
grant SELECT on CD039 to "EMCNTP";
grant SELECT on CM003 to "EMCNTP";
grant SELECT on SO007 to "EMCNTP";
grant SELECT on SO008 to "EMCNTP";
grant SELECT on SO009 to "EMCNTP";
grant SELECT on V_SO503Charge to "EMCNTP";
*/

/*
-- connect to EMCKP
grant SELECT on CD019 to "EMCYMS";
grant SELECT on CD039 to "EMCYMS";
grant SELECT on CM003 to "EMCYMS";
grant SELECT on SO007 to "EMCYMS";
grant SELECT on SO008 to "EMCYMS";
grant SELECT on SO009 to "EMCYMS";
grant SELECT on V_SO503Charge to "EMCYMS";

grant SELECT on CD019 to "EMCLK";
grant SELECT on CD039 to "EMCLK";
grant SELECT on CM003 to "EMCLK";
grant SELECT on SO007 to "EMCLK";
grant SELECT on SO008 to "EMCLK";
grant SELECT on SO009 to "EMCLK";
grant SELECT on V_SO503Charge to "EMCLK";

grant SELECT on CD019 to "EMCNTP";
grant SELECT on CD039 to "EMCNTP";
grant SELECT on CM003 to "EMCNTP";
grant SELECT on SO007 to "EMCNTP";
grant SELECT on SO008 to "EMCNTP";
grant SELECT on SO009 to "EMCNTP";
grant SELECT on V_SO503Charge to "EMCNTP";
*/

/*
-- connect to EMCNTP
grant SELECT on CD019 to "EMCYMS";
grant SELECT on CD039 to "EMCYMS";
grant SELECT on CM003 to "EMCYMS";
grant SELECT on SO007 to "EMCYMS";
grant SELECT on SO008 to "EMCYMS";
grant SELECT on SO009 to "EMCYMS";
grant SELECT on V_SO503Charge to "EMCYMS";

grant SELECT on CD019 to "EMCKP";
grant SELECT on CD039 to "EMCKP";
grant SELECT on CM003 to "EMCKP";
grant SELECT on SO007 to "EMCKP";
grant SELECT on SO008 to "EMCKP";
grant SELECT on SO009 to "EMCKP";
grant SELECT on V_SO503Charge to "EMCKP";

grant SELECT on CD019 to "EMCLK";
grant SELECT on CD039 to "EMCLK";
grant SELECT on CM003 to "EMCLK";
grant SELECT on SO007 to "EMCLK";
grant SELECT on SO008 to "EMCLK";
grant SELECT on SO009 to "EMCLK";
grant SELECT on V_SO503Charge to "EMCLK";
*/



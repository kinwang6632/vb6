/*
  Create indexes after test data been loaded

  Note:
	1. Please change index table space name before execution of this program

  By: Pierre
  Date: 2004.04.07
	2004.04.13 Add 1 new index for SO004
	2004.04.16 Add 1 new index for SO004, add index for SO005 & SOAC0202
*/

set heading off
variable str varchar2(80)
variable v_Now varchar2(20)
variable v_Hour number
variable v_Minute number
variable v_Second number
variable v_TtlSec1 number
variable v_TtlSec2 number

-- Save starting time in v_TtlSec1
begin
    :v_Now    := to_char(sysdate, 'HH24MISS');
    :v_Hour   := substr(:v_Now,1,2);
    :v_Minute := substr(:v_Now,3,2);
    :v_Second := substr(:v_Now,5,2);
    :v_TtlSec1 := :v_Hour*3600+:v_Minute*60+:v_Second;
end;
/

PROMPT *** Create Indexes ***
exec :str := 'Start time = ' || to_char(sysdate, 'YYYY.MM.DD HH24:MI:SS');
print str
-- ***********************************************************************

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_OrderNo
CREATE INDEX I_SO004_OrderNo ON SO004 (OrderNo)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_SSBFP
CREATE INDEX I_SO004_SSBFP ON SO004 (SNo,SmartCardNo,BuyCode,FaciCode,PRSNO)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_CustId_FaciCode
CREATE INDEX I_SO004_CustId_FaciCode ON SO004 (CustId,FaciCode)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_CvtId
CREATE INDEX I_SO004_CvtId ON SO004 (CvtId)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_FaciSNo
CREATE INDEX I_SO004_FaciSNo ON SO004 (FaciSNo)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_CSS
CREATE INDEX I_SO004_CSS ON SO004 (CompCode,ServiceType,SNo)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_MediaCode
CREATE INDEX I_SO004_MediaCode ON SO004 (MediaCode)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_PRDate
CREATE INDEX I_SO004_PRDate ON SO004 (PRDate)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_FaciCode
CREATE INDEX I_SO004_FaciCode ON SO004 (FaciCode)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_BuyCode
CREATE INDEX I_SO004_BuyCode ON SO004 (BuyCode)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_InstDate
CREATE INDEX I_SO004_InstDate ON SO004 (InstDate)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_PRSNO
CREATE INDEX I_SO004_PRSNO ON SO004 (PRSNO )
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_CheckNo
CREATE INDEX I_SO004_CheckNo ON SO004 (CheckNo)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_MACAddress
CREATE INDEX I_SO004_MACAddress ON SO004 (MACAddress)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_SmartCardNo
CREATE INDEX I_SO004_SmartCardNo ON SO004 (SmartCardNo)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO004_SGS_1
CREATE INDEX I_SO004_SGS_1 ON SO004 (CompCode,PRDate,InstDate,SmartCardNo)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;


PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO005_CustId
CREATE INDEX I_SO005_CustId ON SO005 (CustId)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO005_CustIdCvtIdChCode
CREATE INDEX I_SO005_CustIdCvtIdChCode ON SO005 (CustId,CvtId,ChCode)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SO005_CSC
CREATE INDEX I_SO005_CSC ON SO005 (CompCode,ServiceType,CustId)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;


PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SOAC0202_ICCSNO
CREATE INDEX I_SOAC0202_ICCSNO ON SOAC0202 (CompCode,CustId,SmartCardNo,UpdTime)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 300K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX TABLESPACE : SIMPLE_NDX     INDEX NAME : I_SOAC0202_STBSNO
CREATE INDEX I_SOAC0202_STBSNO ON SOAC0202 (CompCode,CustId,STBSNo,UpdTime)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 300K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;



PROMPT INDEX NAME : I_SO451_HandleTime
CREATE INDEX I_SO451_HandleTime on SO451 (CompCode,HandleEDateTime)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX NAME : I_SO453_HandleTime
CREATE INDEX I_SO453_HandleTime on SO453 (CompCode,HandleEDateTime)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX NAME : I_SO455_HandleTime
CREATE INDEX I_SO455_HandleTime on SO455 (CompCode,HandleEDateTime)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX NAME : I_SO457_HandleTime
CREATE INDEX I_SO457_HandleTime on SO457 (CompCode,HandleEDateTime)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

-- ***********************************************************************
exec :str := 'Stop time  = ' || to_char(sysdate, 'YYYY.MM.DD HH24:MI:SS');
print str

-- Get stop time in v_TtlSec2 and count total execution time in str
begin
    :v_Now    := to_char(sysdate, 'HH24MISS');
    :v_Hour   := substr(:v_Now,1,2);
    :v_Minute := substr(:v_Now,3,2);
    :v_Second := substr(:v_Now,5,2);
    :v_TtlSec2 := :v_Hour*3600+:v_Minute*60+:v_Second;

    :v_Hour   := trunc((:v_TtlSec2-:v_TtlSec1)/3600);
    :v_Second := mod(:v_TtlSec2-:v_TtlSec1,60);
    :v_Minute := trunc(((:v_TtlSec2-:v_TtlSec1) - :v_Hour*3600)/60);
    :str := '*** Total execution time = ' || :v_Hour || ':' || :v_Minute || ':' || :v_Second;
end;
/
print str

PROMPT OK !!!
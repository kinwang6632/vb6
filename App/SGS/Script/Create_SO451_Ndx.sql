/*
    Create index for SO451. This index is for SP_InsertSO004.sql to use.
    By: Pierre
    Date: 2004.04.07
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

PROMPT *** Create SO451 Index ***
exec :str := 'Start time = ' || to_char(sysdate, 'YYYY.MM.DD HH24:MI:SS');
print str
-- ***********************************************************************
Prompt Add column 'SerialNo'
alter table SO451 add (SerialNo number(8));

Prompt Update SO451.SerialNo
declare
  v_SerialNo number := 0;
  cursor cc is
	select RowId from SO451;

begin
  for cr in cc loop
	v_SerialNo := v_SerialNo + 1;
	update SO451 set SerialNo=v_SerialNo where RowId=cr.RowId;
  end loop;
  commit;

exception
  when others then
    rollback;
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||','||'SQLERRM='||SQLERRM);
end;
/

--PROMPT INDEX NAME : I_SO451_FLAG
--CREATE INDEX I_SO451_FLAG ON SO451 (FLAG)
--	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) TABLESPACE SIMPLE_NDX;

PROMPT INDEX NAME : I_SO451_SERIALNO
CREATE INDEX I_SO451_SERIALNO ON SO451 (SERIALNO)
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

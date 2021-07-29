/*
@c:\emc_script\Try2

variable str varchar2(200)
variable nn number



*/
set serveroutput on
declare
  v_Description varchar2(40) := 'Default value';

begin
  begin
    select Description into v_Description from CD039 where CodeNo = 7; 
    DBMS_OUTPUT.PUT_LINE(v_Description);

  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('¿ù»~°T®§='||SQLERRM);
  end;

  DBMS_OUTPUT.PUT_LINE(v_Description);
 
exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);

end;
/

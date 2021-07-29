/*
@c:\emc_script\SF_Try1

exec :nn := sf_try1();
print nn


*/

create or replace Function SF_Try1 return number
as
  v_SQL3 varchar2(500);
  type CurTyp is ref cursor;  --¦Û­qcursor«¬ºA
  v_DyCursor1 CurTyp;          --¨Ñdynamic SQL

  v_CompCode number;
  v_CompName varchar2(20);

begin

    begin
      v_SQL3 := 'select CodeNo, Description from emclk.CD039 where RowNum=1';
      open v_DyCursor1 for v_SQL3;
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
        rollback;
        return -3;
    end;
    fetch v_DyCursor1 INTO v_CompCode, v_CompName; 
    close v_DyCursor1;
    -- DBMS_OUTPUT.PUT_LINE(v_CompCode||','||v_CompName);

   return 0;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    return -99;
end;
/

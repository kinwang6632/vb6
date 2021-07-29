/*
  Disable all night-run 

  By: Pierre
  Date: 2003.12.25
*/

set serveroutput on

VARIABLE jobno number;

declare
  cursor cc is
	select * from User_Jobs;

BEGIN
  for cr1 in cc loop
    DBMS_JOB.BROKEN(cr1.Job, TRUE); 
  end loop;

  commit;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/



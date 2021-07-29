/*
@c:\gird\updbill

variable nn number
exec :nn := updbill();

*/



create or replace function UpdBill return number
  as
  cursor cc is 
    select distinct a.sno, a.custid from so007 a, so033 b
	where a.custid=b.custid and substr(a.sno, 9)=substr(b.billno, 9)
	and a.sno!=b.billno
	and a.accepttime>=to_date('20020801','YYYYMMDD');

  v_cnt number := 0;
  x varchar2(15);

begin
  for cr1 in cc loop
    x := cr1.sno;
    v_cnt := v_cnt + 1;
    update SO033 set billno = cr1.sno, item=10+item where custid=cr1.custid and 
	substr(billno, 9)=substr(cr1.sno, 9) and billno!=cr1.sno;

  end loop;

  return v_cnt;

exception
  when others then
    rollback;
    DBMS_OUTPUT.PUT_LINE(x);
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    return -99;
end;
/

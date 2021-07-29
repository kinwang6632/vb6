/*
@c:\gird\GetSec

variable nn number
exec :nn := GetSec('5��4��');
exec :nn := GetSec('54��');
exec :nn := GetSec('2��54��');
exec :nn := GetSec('12��');
exec :nn := GetSec('1��5��4��');

print nn

*/
create or replace function GetSec(p_TimeStr varchar2) return number
as
  v_Sec number := 0;
  x number;
  y varchar2(100);
  z varchar2(50);

begin
  y := p_TimeStr;

  x := instr(y, '��');
  if x > 0 then
    z := substr(y, 1, x-1);
    v_Sec := v_Sec + to_number(z)*3600;
    y := substr(y,x+1);
  end if;

  x := instr(y, '��');
  if x > 0 then
    z := substr(y, 1, x-1);
    v_Sec := v_Sec + to_number(z)*60;
    y := substr(y,x+1);
  end if;

  x := instr(y, '��');
  if x > 0 then
    z := substr(y, 1, x-1);
    v_Sec := v_Sec + to_number(z);
    y := substr(y,x+1);
  end if;

  return v_Sec;

exception 
  when others then
    return -99;
end;
/
/*
@c:\gird\GetSec

variable nn number
exec :nn := GetSec('5分4秒');
exec :nn := GetSec('54秒');
exec :nn := GetSec('2時54秒');
exec :nn := GetSec('12分');
exec :nn := GetSec('1時5分4秒');

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

  x := instr(y, '時');
  if x > 0 then
    z := substr(y, 1, x-1);
    v_Sec := v_Sec + to_number(z)*3600;
    y := substr(y,x+1);
  end if;

  x := instr(y, '分');
  if x > 0 then
    z := substr(y, 1, x-1);
    v_Sec := v_Sec + to_number(z)*60;
    y := substr(y,x+1);
  end if;

  x := instr(y, '秒');
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
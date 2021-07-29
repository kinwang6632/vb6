/*
@c:\emc_script\Try

variable str varchar2(200)
variable nn number

exec :nn := SF_Try('IN (11,12,13,14,115)', 11);
exec :nn := SF_Try('NOT IN (111,112,113,114,115)', 11);
exec :nn := SF_Try(' = 100', 100);
exec :nn := SF_Try(' != 200', 300);
print nn
*/

create or replace function SF_Try
  (p_CitemSQL varchar2, p_CitemCode number) 
  return number as

  v_TmpCitem varchar2(500);	-- for週期收費項目處理用
  I number;
  v_Index  number;
  v_CitemCodeCnt number := 0;
  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER;
  CitemCode NumberAry;


begin

 -- 處理收費項目SQL條件字串, 將其轉成array, 以便後續判斷負項資料所對應之正項是否落於此條件內 2003.04.24
  v_CitemCodeCnt := 0;
  if p_CitemSQL is not null then
    v_TmpCitem := rtrim(ltrim(p_CitemSQL));
    if substr(v_TmpCitem, 1, 2) = 'IN' then			-- 'IN (...,..,..)'
      v_TmpCitem := substr(v_TmpCitem, 5);
      v_TmpCitem := substr(v_TmpCitem, 1, length(v_TmpCitem)-1);
    elsif substr(v_TmpCitem, 1, 6) = 'NOT IN' then		-- 'NOT IN (...,..,..)'
      v_TmpCitem := substr(v_TmpCitem, 9);
      v_TmpCitem := substr(v_TmpCitem, 1, length(v_TmpCitem)-1);
    elsif substr(v_TmpCitem, 1, 1) = '=' then			-- '=...'
      v_TmpCitem := ltrim(substr(v_TmpCitem, 2));
    elsif substr(v_TmpCitem, 1, 2) = '!=' then			-- '!=...'
      v_TmpCitem := ltrim(substr(v_TmpCitem, 3));
    end if;

    v_Index := 1;
    I := 1;
    while v_TmpCitem is not null loop
      v_Index := instr(v_TmpCitem, ',');
      if v_Index > 0 then
	begin
	  CitemCode(I) := ltrim(rtrim(substr(v_TmpCitem, 1, v_Index-1)));
	  v_TmpCitem := substr(v_TmpCitem, v_Index+1);
	  I:=I+1;     
	exception
	  when others then
	    v_TmpCitem := null;
	end;   
      else
	CitemCode(I) := to_number(rtrim(ltrim(v_TmpCitem)));
	v_TmpCitem := null;     
      end if;         
    end loop;
    v_CitemCodeCnt := I;

  end if;

  for I in 1..v_CitemCodeCnt loop
    DBMS_OUTPUT.PUT_LINE(CitemCode(I));
    if CitemCode(I)=p_CitemCode then	-- 查到
      DBMS_OUTPUT.PUT_LINE('*** ' || CitemCode(I));
    end if;
  end loop;

  return v_CitemCodeCnt;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    return -99;
end;
/

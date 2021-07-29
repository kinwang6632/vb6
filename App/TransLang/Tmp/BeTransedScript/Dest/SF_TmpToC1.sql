 
create or replace function SF_TmpToC1
  (p_UserId varchar2, p_RetMsg out varchar2) return number
  as
  v_RcdCount number := 0;
begin
/*
  -- Check parameters
  if p_UserId is null then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
*/
 
  -- ∑h≤æ¶‹•ø¶°¿…
  insert into SO033 (select * from SO032);
  v_RcdCount := sql%rowcount;
  delete from SO032;
 
  -- Log
 
  commit;
  p_RetMsg := 'Execution OK, Rcd count=   '||v_RcdCount;
  return v_RcdCount;
 
exception
  when others then
    rollback;
    p_RetMsg := 'Other error.';
    return -99;
end;
/

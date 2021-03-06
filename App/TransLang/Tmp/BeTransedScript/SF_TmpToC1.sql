
create or replace function SF_TmpToC1
  (p_UserId varchar2, p_RetMsg out varchar2) return number
  as
  v_RcdCount number := 0;
begin
/*
  -- Check parameters
  if p_UserId is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;
*/

  -- 搬移至正式檔
  insert into SO033 (select * from SO032);
  v_RcdCount := sql%rowcount;
  delete from SO032;

  -- Log
  
  commit;
  p_RetMsg := '執行完畢, 處理及搬移筆數為 '||v_RcdCount;
  return v_RcdCount;

exception
  when others then
    rollback;
    p_RetMsg := '其他錯誤';
    return -99;
end;
/
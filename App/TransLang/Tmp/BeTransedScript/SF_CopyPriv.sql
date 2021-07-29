
create or replace function SF_CopyPriv
  (p_SourGId in number, p_DestGId in number, p_RetMsg out varchar2) return number
  as
/*
  v_Count number;
  cursor cc1 is 
    select Mid from SO027 where GroupId=p_SourGId;
*/

  v_Cursor	NUMBER;			-- 供dynamic SQL
  v_RetCode	NUMBER;
  v_Command	VARCHAR2(2000);

begin
  -- 檢查參數
  if (p_SourGId is null) or (p_DestGId is null) then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;
/*
  -- 刪除目標組別舊有權限資料
  delete from SO027 where GroupId=p_DestGId;

  -- 為目標組別新增相關權限資料
  -- 註: 下2行為使用subquery, 會錯誤
  -- insert into SO027 (GroupId, Mid) 
  --   values (select p_DestGId GroupId, Mid from SO027 where GrouId=p_SourGId);
  for cr1 in cc1 loop
    insert into SO027 (GroupId, Mid) values (p_DestGId, cr1.Mid);
  end loop;
*/

  v_Command := 'update SO029 set Group'||to_char(p_DestGId)||'=Group'||to_char(p_SourGId);

  begin
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_Command, DBMS_SQL.V7);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
    DBMS_SQL.Close_cursor(v_Cursor);    
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor);
      p_RetMsg := 'SQL指令錯誤: ' || v_Command;
      return -2;
  end;

  commit;
  p_RetMsg := '執行完畢';
  return 0;

exception
  when others then
    DBMS_SQL.Close_cursor(v_Cursor);
    rollback;
    p_RetMsg := '其他錯誤';
    return -99;
end;
/
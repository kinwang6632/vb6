
create or replace function SF_CopyPriv
  (p_SourGId in number, p_DestGId in number, p_RetMsg out varchar2) return number
  as
/*
  v_Count number;
  cursor cc1 is 
    select Mid from SO027 where GroupId=p_SourGId;
*/

  v_Cursor	NUMBER;			-- ��dynamic SQL
  v_RetCode	NUMBER;
  v_Command	VARCHAR2(2000);

begin
  -- �ˬd�Ѽ�
  if (p_SourGId is null) or (p_DestGId is null) then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;
/*
  -- �R���ؼвէO�¦��v�����
  delete from SO027 where GroupId=p_DestGId;

  -- ���ؼвէO�s�W�����v�����
  -- ��: �U2�欰�ϥ�subquery, �|���~
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
      p_RetMsg := 'SQL���O���~: ' || v_Command;
      return -2;
  end;

  commit;
  p_RetMsg := '���槹��';
  return 0;

exception
  when others then
    DBMS_SQL.Close_cursor(v_Cursor);
    rollback;
    p_RetMsg := '��L���~';
    return -99;
end;
/
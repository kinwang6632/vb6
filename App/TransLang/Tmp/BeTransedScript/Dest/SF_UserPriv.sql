CREATE OR REPLACE FUNCTION SF_UserPriv
  (p_GroupId IN NUMBER, p_FunctionId IN VARCHAR2) RETURN VARCHAR2
  AS
 
  v_Cursor	NUMBER;			-- ��dynamic SQL
  v_RetCode	NUMBER;
  v_Command	VARCHAR2(2000);
 
  v_MidString   varchar2(2000) := '';
  v_Mid 	varchar2(15);
  v_Count	number := 0;
 
begin
  -- �ˬd�Ѽ�
  if p_GroupId is null or p_FunctionId is null then
    return '';
  end if;
 
  -- �զ��d�ߦr��
  v_Command := 'select Mid from SO029 where Group'||to_char(p_GroupId)||
	'=1 and Link='||chr(39)||p_FunctionId||chr(39)||' order by Mid';
 
  v_Cursor := DBMS_SQL.Open_cursor;
  DBMS_SQL.Parse(v_Cursor, v_Command, DBMS_SQL.V7);
  DBMS_SQL.Define_Column(v_Cursor, 1, v_Mid, 15);	-- �w�q���X�����
  v_RetCode := DBMS_SQL.Execute(v_Cursor);
 
  -- Loop�C�@��, �NMid��_��
  loop
    if DBMS_SQL.Fetch_rows(v_Cursor) > 0 then		-- Fetch 1 row
      DBMS_SQL.Column_Value(v_Cursor, 1, v_Mid);	-- ��Mid����
      v_MidString := v_MidString||'('||v_Mid||')';	-- �ꦨ(Mid)�榡
      v_Count := v_Count + 1;
    else					-- No more rows to fetch
      exit;
    end if;
  end loop;
 
  DBMS_SQL.Close_cursor(v_Cursor);    
 
  return v_MidString;
 
exception
  when others then
    DBMS_SQL.Close_cursor(v_Cursor);
    return '';
end;
/

 
create or replace  function SF_GetCodeByRef(p_TableName in varchar2, p_RefNo in number, 
	p_Description out varchar2) return number is
    v_Cursor	number;			-- 供dynamic SQL
    v_RetCode	number;
    v_Command	varchar2(2000);
    v_CodeNo	number;
    v_Description varchar2(40);
 
  begin
    -- Check parameters
    p_Description := '';
    if p_TableName is null or p_RefNo is null then
      return null;
    end if;
 
    v_Command := 'select CodeNo, Description from '||p_TableName||
	' where RefNo='||to_char(p_RefNo);
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_Command, DBMS_SQL.V7);
    DBMS_SQL.Define_Column(v_Cursor, 1, v_CodeNo);	-- 定義取出的欄位
    DBMS_SQL.Define_Column(v_Cursor, 2, v_Description, 40);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
 
    if DBMS_SQL.Fetch_rows(v_Cursor) > 0 then		-- Fetch 1 row
      DBMS_SQL.Column_Value(v_Cursor, 1, v_CodeNo);	-- 取CodeNo欄位值
      DBMS_SQL.Column_Value(v_Cursor, 2, v_Description);-- 取Description欄位值
      p_Description := v_Description;
    end if;
 
    DBMS_SQL.Close_cursor(v_Cursor);    
 
    return v_CodeNo;
 
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor); 
      return null;
  end;
/

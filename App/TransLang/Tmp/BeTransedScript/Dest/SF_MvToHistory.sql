 
 
create or replace function SF_MvToHistory(p_ShouldDate1 varchar2, p_ShouldDate2 varchar2,
	p_RealDate1 varchar2, p_RealDate2 varchar2, p_ResvCount number, 
	p_RetMsg out varchar2) return number
  AS
  v_Cursor	NUMBER;			-- 供dynamic SQL
  v_RetCode	NUMBER;
  v_Command	VARCHAR2(2000);
  c39 char(1) := chr(39);
 
begin
 
  -- 參數檢查 (用and)
  if p_ShouldDate1 is null and p_ShouldDate2 is null and 
    p_RealDate1 is null and p_RealDate2 is null then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  -- Lock應收資料檔
  begin
    lock table SO034 in exclusive mode nowait;
  exception
    when others then
      p_RetMsg := 'Fail to lock, data may use by others.     ';
      return -3;
  end;
 
  -- 串起SQL指令
  v_Command := 'insert into SO036 (select * from SO034 where ';
  if p_ShouldDate1 is not null then
    v_Command := v_Command || 'ShouldDate between to_date('||c39||p_ShouldDate1||c39||','||
	c39||'YYYYMMDD'||c39||') and to_date('||c39||p_ShouldDate2||c39||','||
	c39||'YYYYMMDD'||c39||')';
  end if;
  if p_RealDate1 is not null then
    v_Command := v_Command || ' and RealDate between to_date('||c39||p_RealDate1||c39||','||
	c39||'YYYYMMDD'||c39||') and to_date('||c39||p_RealDate2||c39||','||
	c39||'YYYYMMDD'||c39||')';
  end if;
  v_Command := v_Command ||')';
  -- DBMS_OUTPUT.PUT_LINE(v_Command);
 
  -- 執行SQL
  begin
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_Command, DBMS_SQL.V7);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
    DBMS_SQL.Close_cursor(v_Cursor);    
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor);
      rollback;
      p_RetMsg := 'SQL錯誤: ' || v_Command;
      return -2;
  end;
 
  commit;
 
  p_RetMsg := 'Execution OK';
  return 0;
 
exception
  when others then
    DBMS_SQL.Close_cursor(v_Cursor);
    rollback;
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := 'Other error.';
    return -99;
end;
/

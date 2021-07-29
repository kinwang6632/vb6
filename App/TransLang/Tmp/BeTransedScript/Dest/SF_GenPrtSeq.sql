create or replace function SF_GenPrtSeq
  (p_YM in number) return number
  AS
  v_SeqName	varchar2(20);		-- Sequence object name
  v_Cursor	number;			-- ¨Ñdynamic SQL
  v_RetCode	number;
  v_Command	varchar2(2000);
  v_Count	number;
 
begin
  v_SeqName := 'S_PRTSNO_' || ltrim(to_char(p_YM,'099999'));
  select count(*) into v_Count from USER_OBJECTS where Object_Name=v_SeqName;
 
  if v_Count = 0 then
    v_Command := 'create sequence '||v_SeqName||' MINVALUE 1 MAXVALUE 99999999 INCREMENT BY 1 START WITH 1 NOCACHE NOCYCLE';
 
    v_Cursor := DBMS_SQL.Open_cursor;
 
    DBMS_SQL.Parse(v_Cursor, v_Command, DBMS_SQL.V7);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
 
    DBMS_SQL.Close_cursor(v_Cursor);    
  end if;
 
  return 0;
 
exception
  when others then
    DBMS_SQL.Close_cursor(v_Cursor);
    return -99;
end;
/

CREATE OR REPLACE PROCEDURE SP_UPDCODE
   (  p_ColumnName  IN Varchar2 ,   p_OldCodeNo   IN Number,
     p_NewCodeNo  IN Number)
   IS
/*
@c:\gird\V300\script\SP_UpdCode
set serveroutput on
exec SP_SP_UpdCode('BANKCODE', 3, 21);

  調整目前資料庫各Table中的某種數字型代碼欄位內容, 將舊值改為新值
  檔名: SP_UpdCode

  IN p_ColumnName varchar2: 代碼欄位名稱, 需大寫, 例 'BANKCODE'
     p_OldCodeNo number: 舊代碼值
     p_NewCodeNo number: 新代碼值

Date: 2003/03/03
By:Morris

*/
 
 
  
  
v_Cnt number := 0;
v_Flag number := 0;
v_Cmd varchar2(200);
v_Where varchar2(50);


cursor cTable is
	select Table_name from User_Tables  order by Table_name;

BEGIN

  v_Where := ' where'|| ' ' || upper(p_ColumnName) ||  ' = ' || p_OldCodeNo ;

 
  for cr1 in cTable loop
    v_Cnt := v_Cnt + 1;
    select count(*) into v_Flag from User_tab_columns 
	where Table_name = cr1.Table_name and Column_name=upper(p_ColumnName);
    if v_Flag > 0 then
      v_Cmd := 'update ' || cr1.Table_name || ' set ' || upper(p_ColumnName)|| ' = ' || p_NewCodeNo ||' '|| v_Where;

      dbms_output.put_line('*** ' || cr1.Table_name);

      begin
        EXECUTE IMMEDIATE v_Cmd;
        commit;
      exception
        when others then
        rollback;
        dbms_output.put_line('SQL1錯誤: ' || v_Cmd);
        return;
      end;
    end if;
  end loop;

  commit;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/

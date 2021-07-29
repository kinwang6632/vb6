/*
   調整目前資料庫各Table中的公司別(CompCode)欄位內容
   調整方式: 
	v_Option = 1 強迫全部改為新值
	v_Option = 2 空白的才補上新值
	v_Option = 3 由原公司別改成新公司別
   調整值: 修改程式中有*************的部分 , 可以改為不同值(欲轉為何公司別)

   Date: 2001.10.25
   By: Lawrence
@c:\gird\script\FillCompCode
*/

set serveroutput on

DECLARE
  v_CompCode number;
  v_SourceCompCode number;
  v_Cnt number := 0;
  v_Flag number := 0;
  v_Cmd varchar2(200);
  v_Option number;
  v_Where varchar2(50);

  cursor cTable is
	select Table_name from User_tables order by Table_name;

BEGIN

  /**********************************************************************************/
  -- 改此行
  v_Option := 1;           -- 參數選擇
  v_SourceCompCode := 7;   -- 原公司別
  v_CompCode := 6;         -- 新公司別
  /**********************************************************************************/

  if v_Option = 1 then
    v_Where := null;
  elsif v_Option = 2 then
    v_Where := ' where CompCode is null';
  else
    v_Where := ' where CompCode = ' || v_SourceCompCode;
  end if;
 
  for cr1 in cTable loop
    v_Cnt := v_Cnt + 1;
    select count(*) into v_Flag from User_tab_columns 
	where Table_name = cr1.Table_name and Column_name='COMPCODE';
    if v_Flag > 0 then
      v_Cmd := 'update ' || cr1.Table_name || ' set CompCode= ' || v_CompCode||' '|| v_Where;

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
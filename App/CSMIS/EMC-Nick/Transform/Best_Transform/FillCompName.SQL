/*
   調整目前資料庫各Table中的公司別(CompName)欄位內容
   調整方式: 
	v_Option = 1 強迫全部改為新值
	v_Option = 2 空白的才補上新值
	v_Option = 3 由原公司別改成新公司別
   調整值: 修改程式中有*************的部分 , 可以改為不同值(欲轉為何公司別)

   Date: 2002.08.17
   By: Lawrence
@c:\gird\csmis\script\FillCompName
*/

set serveroutput on

DECLARE
  v_CompName varchar2(20);
  v_SourceCompName varchar2(20);
  v_Cnt number := 0;
  v_Flag number := 0;
  v_Cmd varchar2(200);
  v_Option number;
  v_Where varchar2(50);
  c39 char(1) := chr(39);

  cursor cTable is
	select Table_name from User_tables order by Table_name;

BEGIN

  /**********************************************************************************/
  -- 改此行
  v_Option := 1;           -- 參數選擇
  v_SourceCompName := '5';   -- 原公司別
  v_CompName :='北視光速通訊';         -- 新公司別
  /**********************************************************************************/

  if v_Option = 1 then
    v_Where := null;
  elsif v_Option = 2 then
    v_Where := ' where CompName is null';
  else
    v_Where := ' where CompName = ' || v_SourceCompName;
  end if;
 
  for cr1 in cTable loop
    v_Cnt := v_Cnt + 1;
    select count(*) into v_Flag from User_tab_columns 
	where Table_name = cr1.Table_name and Column_name='COMPNAME';
    if v_Flag > 0 then
      v_Cmd := 'update ' || cr1.Table_name || ' set CompName= ' ||c39||v_CompName||c39||' '|| v_Where;

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
  WHEN OTH RS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/
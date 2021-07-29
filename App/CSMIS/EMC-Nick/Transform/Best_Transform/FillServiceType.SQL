/*
   調整目前資料庫各Table中的服務類別(ServiceType)欄位內容
   調整方式: 
	v_Option = 1 強迫全部改為新值
	v_Option = 2 空白的才補上新值
	v_Option = 3 由'C'改為'I', 或'I'改為'C'
   調整值: 修改程式中有****************的部份, 可以改為不同值

   Date: 2001.03.01
   By: Pierre
@c:\gird\script\FillServiceType
*/

set serveroutput on

DECLARE
  v_SType1 char(1);
  v_SName1 varchar2(10);
  v_SType2 char(2);
  v_SName2 varchar2(10);
  v_Cnt number := 0;
  v_Flag number := 0;
  v_Cmd varchar2(200);
  v_Option number;
  v_Where varchar2(50);
  v_SourceType char(1);

  cursor cTable is
	select Table_name from User_tables order by Table_name;

BEGIN

  /**********************************************************************************/
  -- 改此行, 可將後續的ServiceType皆設為新值 例: ('I', 'CAISP') or ('C', 'CATV')
  -- 若由'C'改為'I', 則v_SourceType:='C', 否則相反
  v_Option := 1;
  v_SType1 := 'I';
  v_SName1 := 'CAISP';
  v_SourceType := 'I';
  /**********************************************************************************/

  if v_Option = 1 then
    v_Where := null;
  elsif v_Option = 2 then
    v_Where := ' where ServiceType is null';
  else
    v_Where := ' where ServiceType ='||chr(39)||v_SourceType||chr(39);
  end if;
 
  for cr1 in cTable loop
    v_Cnt := v_Cnt + 1;
    select count(*) into v_Flag from User_tab_columns 
	where Table_name = cr1.Table_name and Column_name='SERVICETYPE';

    if v_Flag > 0 then
      v_Cmd := 'update ' || cr1.Table_name || ' set ServiceType=' || 
		chr(39) || v_SType1 || chr(39) || v_Where;

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

  update SO001 set ServiceName = v_SName1 where ServiceName is null;
  commit;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    ROLLBACK;
END;
/
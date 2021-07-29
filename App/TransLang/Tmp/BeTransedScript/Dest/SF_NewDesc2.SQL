 
CREATE OR REPLACE FUNCTION SF_NewDesc2
  (p_TableName VARCHAR2, p_CodeNo VARCHAR2, p_OldDesc VARCHAR2, p_NewDesc VARCHAR2, 
  p_RetMsg OUT VARCHAR2) RETURN Number
 
  as
 
begin
  -- 檢查參數
  if p_TableName is null or p_CodeNo is null or p_OldDesc is null or p_NewDesc is null then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
/*
  dbms_output.put_line('A1');
  dbms_output.put_line('# of rcd updated = ' || to_char(SQL%ROWCOUNT));
  dbms_output.put_line('A2');
*/
 
  /* 
    行政區代碼檔 CD001: 待改檔: SO014, SO016, SO019
    服務區代碼檔 CD002: 待改檔: SO001, SO014, SO016, SO019
    工程組代碼檔 CD003: 待改檔: SO007, SO008, SO009
    頻道代碼檔   CD024: 待改檔: SOAC0102
    收費區代碼檔 CD040: 待改檔: SO001, SO014, SO016
  */
  if p_TableName = 'CD001' then
    update SO014 set AreaName=p_NewDesc where AreaCode=p_CodeNo;
    update SO016 set AreaName=p_NewDesc where AreaCode=p_CodeNo;
    update SO019 set AreaName=p_NewDesc where AreaCode=p_CodeNo;
 
  elsif p_TableName = 'CD002' then
    update SO001 set ServArea=p_NewDesc where ServCode=p_CodeNo;
    update SO014 set ServName=p_NewDesc where ServCode=p_CodeNo;
    update SO016 set ServName=p_NewDesc where ServCode=p_CodeNo;
    update SO019 set AreaName=p_NewDesc where AreaCode=p_CodeNo;
 
  elsif p_TableName = 'CD003' then
    update SO007 set GroupName=p_NewDesc where GroupCode=p_CodeNo;
    update SO008 set GroupName=p_NewDesc where GroupCode=p_CodeNo;
    update SO009 set GroupName=p_NewDesc where GroupCode=p_CodeNo;
 
  elsif p_TableName = 'CD024' then
    update SOAC0102 set ChName=p_NewDesc where ChCode=p_CodeNo;
 
  elsif p_TableName = 'CD040' then
    update SO001 set ClctAreaName=p_NewDesc where ClctAreaCode=p_CodeNo;
    update SO014 set ClctAreaName=p_NewDesc where ClctAreaCode=p_CodeNo;
    update SO016 set ClctAreaName=p_NewDesc where ClctAreaCode=p_CodeNo;
 
 
  end if;
 
 
--  commit;
  p_RetMsg := 'Execution OK';
  return 0;
 
exception
  when others then
    rollback;
    p_RetMsg := 'Other error.';
    return -99;
end;
/


CREATE OR REPLACE FUNCTION SF_GetDataArea
  (p_UserID IN VARCHAR2) RETURN VARCHAR2
  AS
  v_TotalAreaCnt number;
  v_UserAreaCnt number;
  v_Tmp varchar2(100) := '';
  v_Flag boolean;

  -- 該操作者可使用的資料區
  cursor CC1 is 
      select ServCode from SO028 where UserId=p_UserID;

begin
  -- 參數檢查
  if p_UserId is null then
    return '執行錯誤';
  end if;

  -- 取該員可使用區之資料筆數
  select count(*) into v_UserAreaCnt from SO028 where UserId=p_UserId;

  -- 若該員可使用區之資料筆數=0, 則可使用所有資料區
  if v_UserAreaCnt = 0 then
    return '';
  end if;

  -- 若該員可使用區之資料筆數=1, 則可使用某一區
  if v_UserAreaCnt = 1 then
    select ServCode into v_Tmp from SO028 where UserId=p_userId;
    return 'ServCode=' || chr(39) || v_Tmp || chr(39);
  end if;

  -- 取所有資料區之資料筆數
  select count(*) into v_TotalAreaCnt from CD002;

  -- 若該員可使用區之資料筆數比所有資料區筆數差1筆, 則不可使用某一區
  if (v_TotalAreaCnt - v_UserAreaCnt) = 1 then
    select CodeNo into v_Tmp from CD002 where CodeNo not in
      (select ServCode from SO028 where UserId=p_UserId);
    return 'ServCode!=' || chr(39) || v_Tmp || chr(39);
  end if;

  -- 可使用某幾區
  v_Tmp := 'ServCode in (';
  for cr1 in CC1 loop
    v_Tmp := v_Tmp || chr(39) || cr1.ServCode || chr(39) || ',';
  end loop;
  v_Tmp := rtrim(v_Tmp,',') || ')';
  return v_Tmp;

exception
  when others then
    return '資料區錯誤';
end;
/
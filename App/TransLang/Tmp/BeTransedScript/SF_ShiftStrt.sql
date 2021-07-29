
create or replace function SF_ShiftStrt
  (p_OldStrtCode in number, p_NewStrtCode in number, p_RetMsg out varchar2) return number
  as
  v_Count number;
  v_NewStr varchar2(10);

begin
  -- 檢查參數
  if (p_OldStrtCode is null) or (p_NewStrtCode is null) then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

  -- 檢查新代碼是否已被使用
  select count(*) into v_Count from CD017 where CodeNo=p_NewStrtCode;
  if v_Count > 0 then
    p_RetMsg := '新街道代碼已被使用';
    return -2;
  end if;

  -- 檢查舊新代碼是否尚未使用
  select count(*) into v_Count from CD017 where CodeNo=p_OldStrtCode;
  if v_Count = 0 then
    p_RetMsg := '舊街道代碼尚未使用';
    return -3;
  end if;

  v_NewStr := substr(to_char(p_NewStrtCode, '999999'),2); 

  -- 下列檔案中街道代碼=原街道代碼的資料, 將街道代碼欄位值改為新街道代碼, 檔案表列:
  -- SO007, SO008, SO009, SO014, SO016, SO023, SO032, SO033, SO034, SO035, SO036, 
  -- SO067免: 應其為暫存檔
  -- 下列檔案中街道代碼=原街道代碼的資料, 將地址排序字串中的街道代碼改為新值, 檔案表列:
  -- SO014, SO016, SO023
  update SO007 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;
  update SO008 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;
  update SO009 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;

  update SO014 set StrtCode=p_NewStrtCode, AddrSort=v_NewStr||substr(AddrSort,7)
	where StrtCode=p_OldStrtCode;

  update SO016 set StrtCode=p_NewStrtCode, AddrSortA=v_NewStr||substr(AddrSortA,7),
	AddrSortB=v_NewStr||substr(AddrSortB,7) where StrtCode=p_OldStrtCode;

  update SO023 set StrtCode=p_NewStrtCode, AddrSortA=v_NewStr||substr(AddrSortA,7),
	AddrSortB=v_NewStr||substr(AddrSortB,7) where StrtCode=p_OldStrtCode;

  update SO032 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;
  update SO033 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;
  update SO034 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;
  update SO035 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;
  update SO036 set StrtCode=p_NewStrtCode where StrtCode=p_OldStrtCode;

  -- 將街道代碼檔中代碼值更新
  update CD017 set CodeNo=p_NewStrtCode	where CodeNo=p_OldStrtCode;

  commit;
  p_RetMsg := '執行完畢';
  return 0;

exception
  when others then
    rollback;
    p_RetMsg := '其他錯誤';
    return -99;
end;
/
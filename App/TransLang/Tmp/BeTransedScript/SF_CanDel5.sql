
create or replace function SF_CanDel5
  (p_Type number, p_BillNo varchar2, p_Item number, p_RetMsg out varchar2) return number
  as

  v_CitemName varchar2(12);
  v_CancelFlag number;
  v_STCode number;

  v_ReturnCode number;
  v_Count number;

begin
  -- Check parameters
  if p_BillNo is null or p_Item is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

  p_RetMsg := '';

  --if p_Type is not null and p_Type != 1 and p_Type!=0 then
  if p_Type = 2 or p_Type = 3 then
    p_RetMsg := '已日結資料, 不可作廢';
    return -3;
  end if;

  begin
    select CitemName, CancelFlag, STCode into v_CitemName, v_CancelFlag, v_STCode
	from SO033 where BillNo=p_BillNo and Item=p_Item;
    if v_CancelFlag = 1 then
      p_RetMsg := '已作廢資料, 不可再作廢';
      return -4;
    end if;
    if v_STCode is not null then
      p_RetMsg := '已設定為短收資料, 不可再作廢';
      return -5;
    end if;
  exception
    when others then
      if v_CitemName is null then
        p_RetMsg := '無此收費單號: '||p_BillNo||', 或該單據無項次='||to_char(p_Item)||'的資料';
        return -2;
      end if;
  end;

  return 1;

exception
  when others then
    p_RetMsg := '其他錯誤';
    return -99;
end;
/
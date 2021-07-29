 
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
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  p_RetMsg := '';
 
  --if p_Type is not null and p_Type != 1 and p_Type!=0 then
  if p_Type = 2 or p_Type = 3 then
    p_RetMsg := 'Cannot cancel closed data.';
    return -3;
  end if;
 
  begin
    select CitemName, CancelFlag, STCode into v_CitemName, v_CancelFlag, v_STCode
	from SO033 where BillNo=p_BillNo and Item=p_Item;
    if v_CancelFlag = 1 then
      p_RetMsg := 'Data already been cancelled.';
      return -4;
    end if;
    if v_STCode is not null then
      p_RetMsg := 'Cannot cancel a discount charge.';
      return -5;
    end if;
  exception
    when others then
      if v_CitemName is null then
        p_RetMsg := 'No such invoice no.'||p_BillNo||',or this bill has no item ='||to_char(p_Item)||'ªº¸ê®Æ';
        return -2;
      end if;
  end;
 
  return 1;
 
exception
  when others then
    p_RetMsg := 'Other error.';
    return -99;
end;
/


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
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  p_RetMsg := '';

  --if p_Type is not null and p_Type != 1 and p_Type!=0 then
  if p_Type = 2 or p_Type = 3 then
    p_RetMsg := '�w�鵲���, ���i�@�o';
    return -3;
  end if;

  begin
    select CitemName, CancelFlag, STCode into v_CitemName, v_CancelFlag, v_STCode
	from SO033 where BillNo=p_BillNo and Item=p_Item;
    if v_CancelFlag = 1 then
      p_RetMsg := '�w�@�o���, ���i�A�@�o';
      return -4;
    end if;
    if v_STCode is not null then
      p_RetMsg := '�w�]�w���u�����, ���i�A�@�o';
      return -5;
    end if;
  exception
    when others then
      if v_CitemName is null then
        p_RetMsg := '�L�����O�渹: '||p_BillNo||', �θӳ�ڵL����='||to_char(p_Item)||'�����';
        return -2;
      end if;
  end;

  return 1;

exception
  when others then
    p_RetMsg := '��L���~';
    return -99;
end;
/
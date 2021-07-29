 
create or replace function SF_CanClose1
  (p_SNo in varchar2, p_RetMsg out varchar2) return number
  as
  v_CustId number;
  v_FinTime date;
  v_ReturnCode number;
  v_Count number;
 
begin
  -- Check parameters
  if p_SNo is null then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  p_RetMsg := '';
 
  begin
    select CustId, FinTime, ReturnCode into v_CustId, v_FinTime, v_ReturnCode
	from SO007 where SNo=p_SNo;
  exception
    when others then
      if v_CustId is null then
        p_RetMsg := 'No such w/o s/n.:'||p_SNo;
        return -2;
      end if;
  end;
 
  if v_FinTime is not null then
    p_RetMsg := 'Cannot return a closed w/o.   ';
    return -3;
  end if;
 
  if v_ReturnCode is not null then
    p_RetMsg := 'Cannot return again.          ';
    return -4;
  end if;
 
  select count(*) into v_Count from SO004 
    where CustId=v_CustId and SNo=p_SNo and InstDate is not null;
  if v_Count > 0 then
    p_RetMsg := 'Cannot return.                                ';
    return -5;
  end if;
 
  -- no need: 'CustId=v_CustId and '
  select count(*) into v_Count from SO033
    where BillNo=p_SNo and RealDate is not null;
  if v_Count > 0 then
    p_RetMsg := 'Cannot return, since associated charge data is entered or cancelled.';
    return -6;
  end if;
 
  select count(*) into v_Count from SO034
    where BillNo=p_SNo;
  if v_Count > 0 then
    p_RetMsg := 'Cannot return.                                ';
    return -7;
  end if;
 
  select count(*) into v_Count from SO035
    where BillNo=p_SNo;
  if v_Count > 0 then
    p_RetMsg := 'Cannot return.                                      ';
    return -8;
  end if;
 
  return 1;
 
exception
  when others then
    p_RetMsg := 'Other error.';
    return -99;
end;
/

 
CREATE OR REPLACE FUNCTION SF_UserPriv2
  (p_GroupId IN NUMBER, p_FunctionId IN VARCHAR2) RETURN NUMBER
  AS
  v_Count number;
begin
  -- Check parameters
  if p_GroupId is null or p_FunctionId is null then
    return 0;
  end if;
 
  select count(*) into v_Count from SO027 where GroupId=p_GroupId and Mid=p_FunctionId;
  if v_Count > 0 then
    return 1;
  else
    return 0;
  end if;
 
exception
  when others then
    return 0;
end;
/

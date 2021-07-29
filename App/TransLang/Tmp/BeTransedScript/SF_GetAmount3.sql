
create or replace function SF_GetAmount3
  (p_CustId number, p_CitemCode number, p_Period number) 
  return number as

/*  v_Para3 number;
  v_Para10 number; */

  v_ClassCode1 number;
  v_MduId varchar2(10);
  v_ChargeType number;

  v_Amount number;
  v_Amount2 number;
  v_FindFlag number := 0;

begin

  -- Check parameters
  if p_CustId is null or p_CitemCode is null or p_Period is null then
    -- p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  --p_RetMsg := '';
/*
  -- �����O�Ѽ��ɬ����Ѽ�: Para3: �����_�Ϥ�P�����I���, Para10: �ϥζO�v��
  begin
    select Para3, Para10 into v_Para3, v_Para10 from SO043;
    if v_Para3=1 then
      v_Para3 := 0;
    else
      v_Para3 := 1;
    end if;
  exception
    when others then
      --p_RetMsg := '���O�Ѽ��ɤ��e���~';
      return -2;
  end;
*/

  -- �q�g���ʶ��س]�w�ɧ�������, �H�M�w�Ǧ^�����
  begin
    select nvl(Amount,0) into v_Amount2 from SO003 
      where CustId=p_CustId and CitemCode=p_CitemCode;
  exception
    when others then
      v_FindFlag := 0;
  end;

  -- ���ӫȤ�������: �Ȥ����O�@, �j�ӽs��, ���O�ݩ�
  begin
    select ClassCode1, MduId, ChargeType 
      into v_ClassCode1, v_MduId, v_ChargeType
      from SO001 where CustId=p_CustId;
  end;

  if v_MduId is not null and v_ChargeType>1 then -- ��j�ӶO�v��, 1=�@��
    begin 
      select nvl(Amount,0) into v_Amount
        from CD019SO017 where MduId=v_MduId and CitemCode=p_CitemCode 
	and Period=p_Period;
      v_FindFlag := 1;		-- ���Ӥj�ӶO�v��
    exception 			-- �L�Ӥj�ӶO�v��, �G��D�S�w�j�ӶO�v��
      when others then
	begin
	  select nvl(Amount,0) into v_Amount
	    from CD019SO017 where MduId is null and CitemCode=p_CitemCode and Period=p_Period;
	  v_FindFlag := 1;
        exception		-- �]�L�D�S�w�j�ӶO�v��, �N�A�d�Ȥ����O�O�v��
	  when others then
	    v_FindFlag := 0;
	end;
    end;
  end if;

  if v_FindFlag != 1 then	-- ��Ȥ����O�O�v��
    begin 
      select nvl(Amount,0) into v_Amount
        from CD019CD004 where CitemCode=p_CitemCode and Period=p_Period 
	and ClassCode=v_ClassCode1;
      v_FindFlag := 1;		-- ���ӫȤ����O�O�v��
    exception 			-- �L�ӫȤ����O�O�v��, �G��D�S�w�Ȥ����O�O�v��
      when others then
	begin
	  select nvl(Amount,0) into v_Amount
	    from CD019CD004 where ClassCode is null and CitemCode=p_CitemCode and Period=p_Period;
	  v_FindFlag := 1;
        exception	-- �]�L�D�S�w�Ȥ����O�O�v��, �h�u�έ즬�O���إN�X���w�]��
	  when others then
	    v_FindFlag := 0;
	end;
    end;
  end if;

  if v_FindFlag = 1 then
    return v_Amount;
  else
    return v_Amount2;
  end if;

exception
  when others then
     DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    --p_RetMsg := '��L���~';
    return -99;
end;
/

create or replace function SF_ShiftStrt
  (p_OldStrtCode in number, p_NewStrtCode in number, p_RetMsg out varchar2) return number
  as
  v_Count number;
  v_NewStr varchar2(10);

begin
  -- �ˬd�Ѽ�
  if (p_OldStrtCode is null) or (p_NewStrtCode is null) then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  -- �ˬd�s�N�X�O�_�w�Q�ϥ�
  select count(*) into v_Count from CD017 where CodeNo=p_NewStrtCode;
  if v_Count > 0 then
    p_RetMsg := '�s��D�N�X�w�Q�ϥ�';
    return -2;
  end if;

  -- �ˬd�·s�N�X�O�_�|���ϥ�
  select count(*) into v_Count from CD017 where CodeNo=p_OldStrtCode;
  if v_Count = 0 then
    p_RetMsg := '�µ�D�N�X�|���ϥ�';
    return -3;
  end if;

  v_NewStr := substr(to_char(p_NewStrtCode, '999999'),2); 

  -- �U�C�ɮפ���D�N�X=���D�N�X�����, �N��D�N�X���ȧאּ�s��D�N�X, �ɮת�C:
  -- SO007, SO008, SO009, SO014, SO016, SO023, SO032, SO033, SO034, SO035, SO036, 
  -- SO067�K: ���䬰�Ȧs��
  -- �U�C�ɮפ���D�N�X=���D�N�X�����, �N�a�}�ƧǦr�ꤤ����D�N�X�אּ�s��, �ɮת�C:
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

  -- �N��D�N�X�ɤ��N�X�ȧ�s
  update CD017 set CodeNo=p_NewStrtCode	where CodeNo=p_OldStrtCode;

  commit;
  p_RetMsg := '���槹��';
  return 0;

exception
  when others then
    rollback;
    p_RetMsg := '��L���~';
    return -99;
end;
/
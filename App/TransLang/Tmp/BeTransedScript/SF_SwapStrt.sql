
create or replace function SF_SwapStrt
  (p_AStrtCode in number, p_BStrtCode in number, p_RetMsg out varchar2) 
   return number
AS 
  v_TmpStrtCode Number := -1;
  v_StrA varchar2(10);
  v_StrB varchar2(10);

begin
  -- �ˬd�Ѽ�
  if (p_AStrtCode is null) or (p_BStrtCode is null) then
     p_RetMsg := '�Ѽƿ��~';
     return -1;
  end if;

--  v_TmpStrtCode := p_AStrtCode   
  v_StrA := substr(to_char(p_AStrtCode, '999999'),2);
  v_StrB := substr(to_char(p_BStrtCode, '999999'),2);

  -- �U�C�ɮפ���D�N�X=���D�N�X�����, �N��D�N�X���ȹ��, �ɮת�C:
  -- SO007, SO008, SO009, SO014, SO016, SO023, SO032, SO033, SO034, SO035, SO036,
  -- SO067�K: ���䬰�Ȧs��
  -- �U�C�ɮפ���D�N�X=���D�N�X�����, �N�a�}�ƧǦr�ꤤ����D�N�X���, �ɮת�C:
  -- SO014, SO016, SO023
  update SO007 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO007 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO007 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  update SO008 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO008 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO008 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  update SO009 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO009 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO009 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  update SO014 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO014 set StrtCode=P_AStrtCode, AddrSort=v_StrA||substr(AddrSort,7)
	where StrtCode=p_BStrtCode;
  update SO014 set StrtCode=P_BStrtCode, AddrSort=v_StrB||substr(AddrSort,7)
	where StrtCode=V_TmpStrtCode;

  update SO016 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO016 set StrtCode=P_AStrtCode, AddrSortA=v_StrA||substr(AddrSortA,7),
	AddrSortB=v_StrA||substr(AddrSortB,7) where StrtCode=p_BStrtCode;
  update SO016 set StrtCode=P_BStrtCode, AddrSortA=v_StrB||substr(AddrSortA,7),
	AddrSortB=v_StrB||substr(AddrSortB,7) where StrtCode=V_TmpStrtCode;

  update SO023 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO023 set StrtCode=P_AStrtCode, AddrSortA=v_StrA||substr(AddrSortA,7),
	AddrSortB=v_StrA||substr(AddrSortB,7) where StrtCode=p_BStrtCode;
  update SO023 set StrtCode=P_BStrtCode, AddrSortA=v_StrB||substr(AddrSortA,7),
	AddrSortB=v_StrB||substr(AddrSortB,7) where StrtCode=V_TmpStrtCode;

  update SO032 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO032 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO032 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  update SO033 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO033 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO033 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  update SO034 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO034 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO034 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  update SO035 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO035 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO035 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  update SO036 set StrtCode=v_TmpStrtCode where StrtCode=p_AStrtCode;
  update SO036 set StrtCode=P_AStrtCode where StrtCode=p_BStrtCode;
  update SO036 set StrtCode=P_BStrtCode where StrtCode=V_TmpStrtCode;

  -- �N��D�N�X�ɤ��N�X�ȧ�s
  update CD017 set CodeNo=v_TmpStrtCode where CodeNo=p_AStrtCode;
  update CD017 set CodeNo=p_AStrtCode where CodeNo=p_BStrtCode;
  update CD017 set CodeNo=p_BStrtCode where CodeNo=V_TmpStrtCode;

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
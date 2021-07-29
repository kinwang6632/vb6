create or replace function SF_StrtCompute
  (p_RetMsg out varchar2) return number
  as
  v_Count number;

  cursor cc1 is 
    select CodeNo from CD017 order by CodeNo;

begin

  -- ���Ҧ����īȤᤧ�Ȥ�s���P�˾��a�}����D�N�X, �s�ܼȦs�ɤ�
  delete from TMP0001;
  insert into TMP0001 (StrtCode, CustId) 
    (select B.StrtCode, A.CustId from SO001 A, SO014 B 
	    where A.InstAddrNo=B.AddrNo and A.CustStatusCode=1);

  -- Loop �C�@��D
  --   �p��ӵ�D�����Ĥ��
  --   ��s��D�N�X�ɤ������Ĥ�����
  for cr1 in cc1 loop
    select count(*) into v_Count from TMP0001 where StrtCode=cr1.CodeNo;
    update CD017 set InstCnt=v_Count where CodeNo=cr1.CodeNo;
  end loop;

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
 
create or replace function SF_NewWip1
  (p_CustId in number) return number
  as
  v_NewWipCode1 number := 0;
  v_NewWipName1 varchar2(12);
  v_RefNo number;
 
  -- �ӫȤ��˾������ulog(SO071)�ɤ�,�̪񪺤@��log���
  cursor cc1 is
    select RefNo from SO071 where CustId=p_CustId order by AcceptTime desc;
 
begin
  -- Check parameters
  if p_CustId is null then
    return -1;
  end if;
 
  -- �����u���A���O�N�X��(CD036)�N�X=0���W��
  begin
    select Description into v_NewWipName1 from CD036 where CodeNo=0;
  exception 
    when others then
      return -2;
  end;
 
  -- ���ӫȤ��˾������ulog(SO071)�ɤ�,�̪񪺤@��log��ƪ��˾����O�N�X���ѦҸ�
  -- �Y�����, �P�_�ӰѦҸ�, �ӨM�w�N�ӫȤ᪺���u���A�@�אּ���
  -- �Y�L���, �h�N���u���A�@�אּ0
  v_RefNo := null;
  open cc1;
    fetch cc1 into v_RefNo;
  close cc1;
 
  if v_RefNo is null then
    v_NewWipCode1 := 0;		-- �L
  elsif v_RefNo=0 then
    v_NewWipCode1 := 5;		-- ��L(�Ǥ�����,���u)
  elsif v_RefNo=1 then
    v_NewWipCode1 := 1;		-- �˾���
  elsif v_RefNo=5 or v_RefNo=7 then
    v_NewWipCode1 := 2;		-- �_����
  elsif v_RefNo=6 then
    v_NewWipCode1 := 4;		-- ���
  else
    v_NewWipCode1 := 3;		-- �]�ƥ[��
  end if;
  if v_RefNo is not null then
    select Description into v_NewWipName1 from CD036 where CodeNo=v_NewWipCode1; -- ���W��
  end if;
  update SO001 set WipCode1=v_NewWipCode1, WipName1=v_NewWipName1 where CustId=p_CustId;
 
  commit;
  return 0;
 
exception
  when others then
    rollback;
    return -99;
end;
/

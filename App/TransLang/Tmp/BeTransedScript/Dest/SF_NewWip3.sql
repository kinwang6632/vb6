 
create or replace function SF_NewWip3
  (p_CustId in number) return number
  as
  --v_AcceptTime date;
  --v_PrCode number;
  v_NewWipCode3 number := 0;
  v_NewWipName3 varchar2(12);
  v_IgnoreCode number;
  v_RefNo number;
 
  -- �ӫȤ�󰱩�������ulog(SO073)�ɤ�,�̪񪺤@��log���
  cursor cc1(vc_IgnoreCode number) is
    select RefNo from SO073 where CustId=p_CustId and PRCode!=vc_IgnoreCode
	order by AcceptTime desc;
 
begin
  -- Check parameters
  if p_CustId is null then
    return -1;
  end if;
 
  -- ���ˬd���u���O��'D/M '��(RefNo=4), �G����'D/M '�����O�N�X
  begin
    select CodeNo into v_IgnoreCode from CD007 where RefNo=4;
  exception
    when others then
      return -3;
  end;
 
  -- �����u���A���O�N�X��(CD036)�N�X=0���W��
  begin
    select Description into v_NewWipName3 from CD036 where CodeNo=0;
  exception 
    when others then
      return -2;
  end;
 
  -- ���ӫȤ�󰱩�������ulog(SO073)�ɤ�,�̪񪺤@��log��ƪ���������O�N�X���ѦҸ�
  -- �Y�����, �P�_�ӰѦҸ�, �ӨM�w�N�ӫȤ᪺���u���A�T�אּ���
  -- �Y�L���, �h�N���u���A�T�אּ0
  open cc1(v_IgnoreCode);
  fetch cc1 into v_RefNo;
  if cc1%notfound then
    v_RefNo := 0;
  end if;
  close cc1;
 
  if v_RefNo=0 then
    v_NewWipCode3 := 0;		-- �L
  elsif v_RefNo=3 then
    v_NewWipCode3 := 11;	-- ������
  elsif v_RefNo=2 or v_RefNo=5 then
    v_NewWipCode3 := 12;	-- �����
  elsif v_RefNo=1 then
    v_NewWipCode3 := 13;	-- ������
  elsif v_RefNo=4 then
    v_NewWipCode3 := 14;	-- ���
  end if;
  if v_RefNo!=0 then
    select Description into v_NewWipName3 from CD036 where CodeNo=v_NewWipCode3; -- ���W��
  end if;
  update SO001 set WipCode3=v_NewWipCode3, WipName3=v_NewWipName3 where CustId=p_CustId;
 
  commit;
  return 0;
 
exception
  when others then
    rollback;
    return -99;
end;
/

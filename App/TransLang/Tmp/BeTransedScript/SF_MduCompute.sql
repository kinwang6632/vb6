
create or replace function SF_MduCompute
  (p_MduId in Varchar2, p_RetMsg out Varchar2) return number
  As
  v_CountA number;
  v_CountB number;
  v_CountC number;

begin
  -- �ˬd�Ѽ�
  if (p_MduId is null) then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  -- ���ɤ��[A]: �Ȥ����ɤ��j�ӽs��=<�Ӥj��>, 
  -- �B�Ȥ᪬�A!=�P�P��(5), ���P(4)
  select count(*) into v_CountA from SO001 where MduId= p_MduId 
     and CustStatusCode!=4 and CustStatusCode!=5;

  -- ���Ĥ��[B]: �Ȥ����ɤ��j�ӽs��=<�Ӥj��>, �B�Ȥ᪬�A=���`(1)
  select count(*) into v_CountB from SO001 where 
     MduId=p_MduId and CustStatusCode=1;

  -- ����/�ݩ���[C]: ���ɤ��[A] - ���Ĥ��[B]
  v_CountC := v_CountA - v_CountB;

  -- ��s�Ӥj�Ӹ�Ƥ��U�ؤ�������
  update SO017 set AcceptCnt=v_CountA, InstCnt=v_CountB, 
     UnInstCnt=v_CountC where MduId=p_MduId;

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

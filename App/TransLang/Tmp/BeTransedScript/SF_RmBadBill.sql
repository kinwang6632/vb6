
create or replace function SF_RmBadBill(p_DueDate varchar2, 
  p_UpdTime varchar2, p_UpdName varchar2, p_RetMsg out varchar2) return number
  AS
  v_STCode number;
  v_STName varchar2(12);
  v_RcdCount number := 0;
  v_Today date;

  cursor cc(p_STCode number) is
    select * from SO033 where RealDate<=to_date(p_DueDate,'YYYYMMDD') and
      (STCode=p_STCode or UCCode is not null);

begin

  -- �Ѽ��ˬd
  if p_DueDate is null or p_UpdTime is null or p_UpdName is null then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  -- ���u����]�N�X�ɤ�(CD016)�ѦҸ���1(�b�b)���b�b�N�X/�W��
  begin
    select CodeNo, Description into v_STCode, v_STName from CD016 where RefNo=1;
  exception 
    when others then
      p_RetMsg := '�u����]�N�X�ɤ����]�ѦҸ���1�����';
      return -2;
  end;

  -- Lock���������
  begin
    lock table SO033 in exclusive mode nowait;
  exception
    when others then
      p_RetMsg := '�����������w����, �i��L�H���ϥΤ�';
      return -3;
  end;

  v_Today := to_date(to_char(sysdate, 'YYYYMMDD'),'YYYYMMDD');

  -- Loop �C�@��, �s�W���a�b����� (�u���̫�2�檺��줺�e���P)
  for cr in cc(v_STCode) loop
    insert into SO035 (CustId, BillNo, Item, CitemCode, CitemName, 
	OldAmt, OldPeriod, OldStartDate, OldStopDate, ShouldDate, 
	ShouldAmt, RealAmt, RealPeriod, RealStartDate, RealStopDate, 
	ClctEn, ClctName, CMCode, CMName, Note, 
	CreateTime, CreateEn, InvoiceTime, PrtClsTime, CompCode, 
	AddrNo, StrtCode, MduId, PTCode, PTName, 
	CancelFlag, ServCode, ClctAreaCode, PrtCount, OldClctEn, 
	OldClctName, PrintEn, PrintTime, EntryEn, ClassCode, 
	ManualNo, ClctYM, PrtSNo, GUINo, Quantity, 
	CustCount, AreaCode, 
	UCCode, UCName, STCode, STName, 
	RealDate, UpdTime, UpdEn)
      values (cr.CustId, cr.BillNo, cr.Item, cr.CitemCode, cr.CitemName, 
	cr.OldAmt, cr.OldPeriod, cr.OldStartDate, cr.OldStopDate, cr.ShouldDate, 
	cr.ShouldAmt, cr.RealAmt, cr.RealPeriod, cr.RealStartDate, cr.RealStopDate, 
	cr.ClctEn, cr.ClctName, cr.CMCode, cr.CMName, cr.Note, 
	cr.CreateTime, cr.CreateEn, cr.InvoiceTime, cr.PrtClsTime, cr.CompCode, 
	cr.AddrNo, cr.StrtCode, cr.MduId, cr.PTCode, cr.PTName, 
	cr.CancelFlag, cr.ServCode, cr.ClctAreaCode, cr.PrtCount, cr.OldClctEn, 
	cr.OldClctName, cr.PrintEn, cr.PrintTime, cr.EntryEn, cr.ClassCode, 
	cr.ManualNo, cr.ClctYM, cr.PrtSNo, cr.GUINo, cr.Quantity, 
	cr.CustCount, cr.AreaCode,
	NULL, NULL, v_STCode, v_STName, 
	v_Today, p_UpdTime, p_UpdName);

    v_RcdCount := v_RcdCount + 1;
  end loop;

  if v_RcdCount = 0 then
    p_RetMsg := '���槹��, ���C����=0';
    return 0;
  else
    -- �R����������ɤ��ŦX���󪺸��
    delete from SO033 where RealDate<=to_date(p_DueDate,'YYYYMMDD') and
      (STCode=v_STCode or UCCode is not null);

    p_RetMsg := '���槹��, ���C����='||to_char(v_RcdCount);
    commit;
    return v_RcdCount;
  end if;

exception
  when others then
    rollback;
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := '��L���~';
    return -99;
end;
/

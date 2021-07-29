/*
  Dynamic SQL �d�ҵ{��

  IN  ....
  OUT ....
  Return NUMBER�G���G�N�X
	...
        ??: p_RetMsg='SQL���~: '
	...
        -99�Gp_RetMsg='��L���~'
*/
create or replace function ABCXYZ
  (...) 
  return number as

  c39		char(1) := chr(39);
  v_SQL1 	varchar2(4000);		-- must long enough
  v_Column1	varchar2(200);
  v_Where1	varchar2(200);
  v_MoreTable   varchar2(100);
  v_Join	varchar2(100);

  type CurTyp is ref cursor;	--�ۭqcursor���A
  cDynamic CurTyp;          	--��dynamic SQL


begin

  --***** 1. �զXSQL�d�ߦr��: (���Ҹ�����, �NSQL���O��h�q�ӲզX, ��ڤW�i���²��)
  v_Column1 := 'select A.BillNo, A.Item, A.CustId, A.CitemCode, A.CitemName, A.RealDate, A.RealAmt, ' ||
		'A.RealPeriod, A.RealStartDate, A.RealStopDate, A.STCode, A.ClctEn, A.AddrNo, A.ServCode, A.ServiceType';
  v_Where1  :=  ' A.RealDate>=' || GiPackage.QryDTString0(to_date(p_RealDate1,'YYYYMMDD')) ||
		' and A.RealDate<=' || GiPackage.QryDTString0(to_date(p_RealDate2,'YYYYMMDD')) ||
		' and A.UCCode is null and A.CancelFlag=0 and nvl(A.RealAmt,0) !=0';

  -- �Y�����q�O����
  if p_CompCode is not null then
    v_SQL1 := v_SQL1 || ' and A.CompCode=' || to_char(p_CompCode);
  end if;

  -- �Y���A�����O����
  if p_ServiceType is not null then
    v_SQL1 := v_SQL1 || ' and A.ServiceType=' || c39 || p_ServiceType || c39;
  end if;

  -- �Y��������O����
  if p_BillTypeSQL is not null then
    v_SQL1 := v_SQL1 || ' and substr(A.BillNo,7,1) ' || p_BillTypeSQL;
  end if;

  -- ���SQL��_
  if v_Join is not null then
    v_Where1 := ' and' || v_Where1;
  end if;
  v_SQL1 := v_Column1 || ' from SO033 A' || v_MoreTable || ' where ' || v_Join  || v_Where1 || v_SQL1 || ' union ' ||
	    v_Column1 || ' from SO034 A' || v_MoreTable || ' where ' || v_Join || v_Where1 || v_SQL1 ;


  --*****  2. �}��Dynamic sql
  begin
    open cDynamic for v_SQL1;
  exception
    when others then
      rollback;
      p_RetMsg := 'SQL���~: ' || v_SQL1;
      return -4;
  end;


  --*****  3. loop���C�@����ƨӳB�z
  v_BadCnt := 0;
  loop       -- ���쪺�C�@�����
    fetch cDynamic INTO x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealAmt, x_RealPeriod, x_RealStartDate, x_RealStopDate, x_STCode, 
		x_ClctEn, x_AddrNo, x_ServCode, x_ServiceType;
    exit when cDynamic%NOTFOUND;
    .......
  end loop;


  --*****  4. Close dynamic cursor
  close cDynamic;


  commit;
  p_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||
		'��, ���ͩ��ӵ���='||to_char(v_Cnt);
  return 0;

exception
  when others then
    --*****  4. Close dynamic cursor
    close cDynamic;

    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    return -99;
end;
/
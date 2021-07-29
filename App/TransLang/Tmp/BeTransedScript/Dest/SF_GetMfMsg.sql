 
CREATE OR REPLACE FUNCTION SF_GetMfMsg
  (p_AddrNo IN NUMBER) RETURN VARCHAR2
  AS
  v_RetMsg   	varchar2(100) := 'Trouble area :  ';
  v_AddrSort 	varchar2(100);
  v_Noe1	number;
  v_Noe2	number;
  v_Noe3	number;
  v_Noe4	number;
 
  v_Flag 	number := 0;	-- �Y���, �h�]��1
 
  v_SNo		varchar2(8);
  v_Noe		number;
  v_Lane	varchar2(10);
  v_Alley	varchar2(10);
  v_Alley2	varchar2(10);
 
  v_ErrorTime	date;
  v_EndTime	date;
  v_MfName	varchar2(12);
 
  cursor cc1(p_AddrSort varchar2) is
    select a.SNo, a.Noe, a.Lane, a.Alley, a.Alley2
      from SO023 a,SO022 b where a.Sno=b.Sno and p_AddrSort >= a.AddrSortA and p_AddrSort <= a.AddrSortB
      and (sysdate<=b.FinTime or b.FinTime is Null)
      order by a.Ord desc;
 
begin
  -- �ˬd�Ѽ�
  if p_AddrNo is null then
    return '';
  end if;
 
  -- ���a�}�s���ҹ������a�}�ƧǦr��
  select AddrSort, Noe1, Noe2, Noe3, Noe4 
	into v_AddrSort, v_Noe1, v_Noe2, v_Noe3, v_Noe4 
	from SO014 where AddrNo=p_AddrNo;
 
  open cc1(v_AddrSort);
  loop
    fetch cc1 into v_SNo, v_Noe, v_Lane, v_Alley, v_Alley2;
    exit when cc1%notfound or v_Flag=1;
 
    -- �Y�L����������, �h���Ĥ@�����, �N�O�F
    if v_Noe = 0 then
      v_Flag := 1;
      exit;
    end if;
 
    -- �ˬd�O�_�ŦX����������
    if v_Alley2 is not null then	-- �Y������, �h�ˬd��O�_����������
      if v_Noe4 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    elsif v_Alley is not null then	-- �Y�˦���, �h�ˬd��O�_�ˤ�������
      if v_Noe3 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    elsif v_Lane is not null then	-- �Y�Ѧ���, �h�ˬd��O�_�Ѥ�������
      if v_Noe2 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    else				-- �Y�ѧ����ҵL��, �h�ˬd��O�_�󤺳�����
      if v_Noe1 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    end if;
  end loop;
  close cc1;
 
  if v_Flag = 1 then			-- �Y���, �h���ӵ������������
    select ErrorTime, EndTime, MfName into v_ErrorTime, v_EndTime, v_MfName
	from SO022 where SNo=v_SNo;
    v_RetMsg := v_RetMsg || GiPackage.GetDTString3(v_ErrorTime) || '~' ||
		GiPackage.GetDTString3(v_EndTime) || ',reason:' || v_MfName;
    return v_RetMsg;
 
  else
    return '';
  end if;
 
exception
  when others then
    if cc1%isopen then
      close cc1;
    end if;
    return '';
end;
/

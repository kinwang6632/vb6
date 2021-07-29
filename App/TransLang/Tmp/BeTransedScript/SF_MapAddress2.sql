
create or replace function SF_MapAddress2
  (p_UserId varchar2, p_StrtCode number, p_RetMsg out varchar2) return number
  as
  v_RetCode number;
  v_StartExecTime date;
  v_Flag	number;
  v_Ord 	number;
  v_Adjcomp	number;			-- �O�_�@�ֽվ�Ȥ����ɤ������q�O
  v_OldStrtCode number := -1;
  c39		char(1) := chr(39);

  v_Cmd 	varchar2(600);
  v_Cmd2 	varchar2(600);
  v_Cursor	NUMBER;			-- ��dynamic SQL

  v_ZipCode	varchar2(5);
  v_CircuitNo	varchar2(20);
  v_AreaCode	varchar2(3);
  v_AreaName	varchar2(12);
  v_ServCode	varchar2(3);
  v_ServName	varchar2(12);
  v_ClctAreaCode varchar2(3);
  v_ClctAreaName varchar2(12);
  v_ClctEn 	varchar2(10);
  v_ClctName 	varchar2(20);
  v_BTCode 	number;
  v_BTName 	varchar2(12);
  v_MduId 	varchar2(8);
  v_MduName 	varchar2(30);
  v_CompCode 	number;
  v_CompName 	varchar2(12);

  -- for SO081
  v_IAddrCnt1 	number;
  v_ICustCnt1 	number;
  v_IUCnt1 	number;
  v_CAddrCnt1 	number;
  v_CCustCnt1 	number;
  v_CUCnt1 	number;
  v_IAddrCnt2 	number;
  v_ICustCnt2 	number;
  v_IUCnt2 	number;
  v_CAddrCnt2 	number;
  v_CCustCnt2 	number;
  v_CUCnt2 	number;


  -- �Ҧ��a�}���
  cursor cc1 is
    select * from SO014 where StrtCode = p_StrtCode order by AddrSort;

  cursor cc2 is
    select * from Tmp002;

  -- �ݲ��ʦa�}���log��
  cursor cc3 is
    select UpdSQL from Tmp002;

 -- �ݲ��ʫȤ���log��
  cursor cc4 is
    select Content from SO080;

  -- �Y�˾��a�}�W���Ȥ᪬�A
  cursor ct1(v_ANo number) is 
    select CustId, CustStatusCode, CompCode, MduId, ServCode
	from SO001 where InstAddrNo=v_ANo;

  -- �Y���O�a�}�W���Ȥ᪬�A
  cursor ct2(v_ANo number) is 
    select CustId, CustStatusCode, ClctAreaCode from SO001 where ChargeAddrNo=v_ANo;

begin
  -- Check parameters
  if (p_UserId is null) or (p_StrtCode is null) then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  begin
    select nvl(AdjComp,0) into v_AdjComp from SO044;
  exception
    when others then
      p_RetMsg := '�w�]�Ѽ���(SO044)���L��ƩΦh�����';
      return -2;
  end;

  --�ˬd�O�_���إ߼Ȧs��Tmp002, �Y�L�h�إߤ�
  select count(*) into v_Flag from User_Tables where Table_Name = 'TMP002';
  if v_Flag = 0 then 
    v_Cmd := 'CREATE TABLE Tmp002 (CompCode NUMBER(3),AddrNo NUMBER(8),UpdSQL varchar2(600))';
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_Cmd);
  end if;

  v_OldStrtCode := -1;
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  delete from Tmp002;
  delete from SO081;
  delete from SO080;

  for cr1 in cc1 loop
    v_Cmd := null;
    v_Cmd2 := null;
    v_Flag := null;
    v_IAddrCnt1 := 0;
    v_ICustCnt1 := 0;
    v_IUCnt1 := 0;
    v_CAddrCnt1 := 0;
    v_CCustCnt1 := 0;
    v_CUCnt1 := 0;

    if cr1.StrtCode != v_OldStrtCode then	-- ���t�@�յ�D�F
      -- �C���������P��D��, �s�W1���ӵ�D��������έp��Ʀ�SO081
      if v_OldStrtCode != -1 then	-- �]�Ĥ@��v_OldStrtCode=-1
	insert into SO081 (StrtCode, Ord, IAddrCnt, ICustCnt, IUCnt, CAddrCnt, CCustCnt, CUCnt)
	  values (v_OldStrtCode, 9999, v_IAddrCnt2, v_ICustCnt2, v_IUCnt2, v_CAddrCnt2, 
	  v_CCustCnt2, v_CUCnt2);
      end if; 

      v_IAddrCnt2 := 0;
      v_ICustCnt2 := 0;
      v_IUCnt2 := 0;
      v_CAddrCnt2 := 0;
      v_CCustCnt2 := 0;
      v_CUCnt2 := 0;
      v_OldStrtCode := cr1.StrtCode;
    end if;

    v_Ord := SF_GetAddrMap3(cr1.StrtCode, cr1.AddrSort, cr1.Noe1, cr1.Noe2, cr1.Noe3, cr1.Noe4, 
		v_ZipCode, v_CircuitNo, v_AreaCode, v_AreaName, v_ServCode, 
		v_ServName, v_ClctAreaCode, v_ClctAreaName, v_ClctEn, v_ClctName, 
		v_BTCode, v_BTName, v_MduId, v_MduName, v_CompCode, v_CompName);

    if v_Ord < 0 then			-- �L�a�}����
      for t1 in ct1(cr1.AddrNo) loop
        if t1.CustStatusCode=1 then
	  v_ICustCnt2 := v_ICustCnt2 + 1;	-- �֭p�ӵ�D�W,���˾��a�}�p�⪺�����Ȥ��
        else
	  v_IUCnt2 := v_IUCnt2 + 1;		-- �֭p�ӵ�D�W,���˾��a�}�p�⪺�D�����Ȥ��
        end if;
	if ct1%rowcount=1 then
          v_IAddrCnt2 := v_IAddrCnt2 + 1;	-- �֭p�ӵ�D�W,���˾��a�}�p�⪺�a�}��
	end if;
      end loop;
      for t2 in ct2(cr1.AddrNo) loop
        if t2.CustStatusCode=1 then
	  v_CCustCnt2 := v_CCustCnt2 + 1;	-- �֭p�ӵ�D�W,�����O�a�}�p�⪺�����Ȥ��
        else
	  v_CUCnt2 := v_CUCnt2 + 1;		-- �֭p�ӵ�D�W,�����O�a�}�p�⪺�D�����Ȥ��
        end if;
	if ct2%rowcount=1 then
          v_CAddrCnt2 := v_CAddrCnt2 + 1;	-- �֭p�ӵ�D�W,�����O�a�}�p�⪺�a�}��
	end if;
      end loop;

      -- ���������, �H�U������a�}�ɤ��­�
      v_ZipCode		:= cr1.ZipCode;
      v_CircuitNo	:= cr1.CircuitNo;
      v_BTCode		:= cr1.BTCode;
      v_MduId 		:= cr1.MduId;
      v_CompCode 	:= cr1.CompCode;
      v_CompName	:= cr1.CompName;

    else				-- ���������
      for t1 in ct1(cr1.AddrNo) loop
	v_Cmd2 := null;

	-- �����n�ɤ~��s�Ȥ����ɤ������q�O
	if v_AdjComp=1 and t1.CompCode != v_CompCode then
	  v_Cmd2 := v_Cmd2 || ',CompCode='||v_CompCode||',CompName='||c39||v_CompName||c39;
	  --update SO001 set CompCode=v_CompCode, CompName=v_CompName where CustId=t1.CustId;
	end if;

        if t1.CustStatusCode=1 then
	  v_ICustCnt1 := v_ICustCnt1 + 1;	-- �֭p�ӵ�D�W,���˾��a�}�p�⪺�����Ȥ��
        else
	  v_IUCnt1 := v_IUCnt1 + 1;		-- �֭p�ӵ�D�W,���˾��a�}�p�⪺�D�����Ȥ��
        end if;
	if ct1%rowcount=1 then
	  v_IAddrCnt1 := v_IAddrCnt1 + 1;	-- �֭p�ӵ�D�W,���˾��a�}�p�⪺�a�}��
        end if;
	if t1.ServCode != v_ServCode then
	  v_Cmd2 := v_Cmd2 || ',ServCode='||c39||v_ServCode||c39||',ServArea='||c39||v_ServName||c39;
	end if;
	if t1.MduId != v_MduId then
	  v_Cmd2 := v_Cmd2 || ',MduId='||c39||v_MduId||c39;
	end if;
	if v_Cmd2 is not null then
	  insert into SO080 (CustId, Content) values 
		(t1.CustId, ltrim(v_Cmd2,',')||' where CustId='||t1.CustId);
	end if;
      end loop;

      for t2 in ct2(cr1.AddrNo) loop
	v_Cmd2 := null;

        if t2.CustStatusCode=1 then
	  v_CCustCnt1 := v_CCustCnt1 + 1;	-- �֭p�ӵ�D�W,�����O�a�}�p�⪺�����Ȥ��
        else
	  v_CUCnt1 := v_CUCnt1 + 1;		-- �֭p�ӵ�D�W,�����O�a�}�p�⪺�D�����Ȥ��
        end if;
	if ct2%rowcount=1 then
	  v_CAddrCnt1 := v_CAddrCnt1 + 1;	-- �֭p�ӵ�D�W,�����O�a�}�p�⪺�a�}��
	end if;
	if t2.ClctAreaCode != v_ClctAreaCode then
	  v_Cmd2 := v_Cmd2 || 'ClctAreaCode='||c39||v_ClctAreaCode||c39||',ClctAreaName='||c39||v_ClctAreaName||c39;
	  insert into SO080 (CustId, Content) 
		values (t2.CustId, v_Cmd2||' where CustId='||t2.CustId);
	end if;
      end loop;

      -- �֭p�Ӳ�(��D,����)������ƪ��U���έp��
      update SO081 set IAddrCnt=IAddrCnt+v_IAddrCnt1, ICustCnt=ICustCnt+v_ICustCnt1, 
	  IUCnt=IUCnt+v_IUCnt1, CAddrCnt=CAddrCnt+v_CAddrCnt1, CCustCnt=CCustCnt+v_CCustCnt1, 
	  CUCnt=CUCnt+v_CUCnt1  where StrtCode=v_OldStrtCode and Ord=v_Ord;

      -- �Y��L�Ӳ�(��D,����)������ƪ��U���έp��,�h�s�W�@��
      if SQL%rowcount=0 then 
	insert into SO081 (StrtCode, Ord, IAddrCnt, ICustCnt, IUCnt, CAddrCnt, CCustCnt, CUCnt)
	    values (v_OldStrtCode, v_Ord, v_IAddrCnt1, v_ICustCnt1, v_IUCnt1, v_CAddrCnt1, 
	    v_CCustCnt1, v_CUCnt1);

      end if;
    end if;

    if v_AreaCode!=cr1.AreaCode then
      v_Cmd := v_Cmd || ',AreaCode='||c39||v_AreaCode||c39
		||',AreaName='||c39||v_AreaName||c39;
      v_Flag := 1;
    end if;
    if v_ServCode!=cr1.ServCode then
      v_Cmd := v_Cmd || ',ServCode='||c39||v_ServCode||c39
		||',ServName='||c39||v_ServName||c39;
    end if;
    if v_ClctAreaCode!=cr1.ClctAreaCode then
      v_Cmd := v_Cmd || ',ClctAreaCode='||c39||v_ClctAreaCode||c39
		||',ClctAreaName='||c39||v_ClctAreaName||c39;
      v_Flag := 1;
    end if;
    if v_ZipCode!=cr1.ZipCode then
      v_Cmd := v_Cmd || ',ZipCode='||c39||v_ZipCode||c39;
      v_Flag := 1;
    end if;
    if v_ClctEn!=cr1.ClctEn then
      v_Cmd := v_Cmd || ',ClctEn='||c39||v_ClctEn||c39
		||',ClctName='||c39||v_ClctName||c39;
      v_Flag := 1;
    end if;
    if v_BTCode!=cr1.BTCode then
      v_Cmd := v_Cmd || ',BTCode='||c39||v_BTCode||c39
		||',BTName='||c39||v_BTName||c39;
      v_Flag := 1;
    end if;
    if v_MduId!=cr1.MduId then
      v_Cmd := v_Cmd || ',MduId='||c39||v_MduId||c39
		||',MduName='||c39||v_MduName||c39;
      v_Flag := 1;
    end if;
    if v_CircuitNo!=cr1.CircuitNo then
      v_Cmd := v_Cmd || ',CircuitNo='||c39||v_CircuitNo||c39;
      v_Flag := 1;
    end if;
    if v_CompCode!=cr1.CompCode then
      v_Cmd := v_Cmd || ',CompCode='||c39||v_CompCode||c39
		||',CompName='||c39||v_CompName||c39;
      v_Flag := 1;
    end if;
    
    -- Log �ܫݲ��ʦa�}�� Tmp002 (Must use dynamic SQL)
    if v_Flag = 1 then
      v_Cmd := ltrim(v_Cmd, ',')||' where CompCode='||cr1.CompCode||' and AddrNo='||cr1.AddrNo;
      insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	values (cr1.CompCode, cr1.AddrNo, v_Cmd);
    end if;
  end loop;

  -- �̫�@�յ�D���������έp��Ʃ|���g�J, �G�s�W��SO081
  if v_OldStrtCode!=-1 then
     insert into SO081 (StrtCode, Ord, IAddrCnt, ICustCnt, IUCnt, CAddrCnt, CCustCnt, CUCnt)
	values (v_OldStrtCode, 9999, v_IAddrCnt2, v_ICustCnt2, v_IUCnt2, v_CAddrCnt2, 
	v_CCustCnt2, v_CUCnt2);
  end if; 

  --dbms_output.put_line('<1> '||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'��');

  -- loop����s��ݲ��ʦa�}�ɤ������
  v_Cursor := DBMS_SQL.Open_cursor;
  for cr3 in cc3 loop 
    v_Cmd := 'update SO014 set '||cr3.UpdSQL;
    DBMS_SQL.Parse(v_Cursor, v_Cmd, DBMS_SQL.V7);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
  end loop;
  DBMS_SQL.Close_cursor(v_Cursor); 

  -- loop����s��ݲ��ʫȤ���(SO080)�������
  v_Cursor := DBMS_SQL.Open_cursor;
  for cr4 in cc4 loop 
    v_Cmd := 'update SO001 set '||cr4.Content;
    DBMS_SQL.Parse(v_Cursor, v_Cmd, DBMS_SQL.V7);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
  end loop;
  DBMS_SQL.Close_cursor(v_Cursor); 

  p_RetMsg := '���槹��, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'��';
  commit;
  return 0;

exception
  when others then
    if cc1%isopen then
      close cc1;
    end if;
    if cc3%isopen then
      close cc3;
    end if;
    if cc4%isopen then
      close cc4;
    end if;
    rollback;
    p_RetMsg := '��L���~'||v_Cmd;
    return -99;
end;
/

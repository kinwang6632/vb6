 
create or replace function SF_NewAddress
  (p_UserId varchar2, p_Option number, p_RetMsg out varchar2) return number
  as
 
  v_StartExecTime date;
  v_Addr1 	varchar2(100);
 
  v_Tmp		varchar2(600);
  v_Cmd 	varchar2(600);
  v_Cursor	NUMBER;			-- ��dynamic SQL
  v_RetCode number;
 
  -- �Ȥᤧ�a�}���
  cursor cc1 is 
    select CompCode, CustId, InstAddrNo, InstAddress, ChargeAddrNo, ChargeAddress, 
      MailAddrNo, MailAddress from SO001;
 
  -- �˾����u�椧�a�}���
  cursor cc2 is 
    select SNo, CompCode, AddrNo, Address from SO007;
 
  -- ���׬��u�椧�a�}���
  cursor cc3 is 
    select SNo, CompCode, AddrNo, Address from SO008;
 
  -- ��������u�椧�a�}���
  cursor cc4 is 
    select SNo, CompCode, OldAddrNo, OldAddress, ReInstAddrNo, ReInstAddress, 
	NewChargeAddrNo, NewChargeAddress, NewMailAddrNo, NewMailAddress from SO009;
 
  -- �a�}���v�ɤ��a�}���
  cursor cc5 is 
    select CustId, AddrNo, Address from SO015;
 
  -- �j�Ӥ��a�}���
  cursor cc6 is
    select MduId, MainAddrNo, MainAddress, ContAddrNo, ContAddress
	from SO017;
 
  -- �ݲ��ʸ�ƪ�SQL���O (�P�a�}���s�����\��@��Tmp002)
  cursor cc0 is
    select * from Tmp002 where CompCode is null and AddrNo is null;
 
begin
  -- Check parameters
  if (p_UserId is null) then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  delete from Tmp002;
 
  -- �B�z�Ȥᤧ�a�}���
  for cr1 in cc1 loop
    v_Cmd := null;
    v_Tmp := null;
 
    if cr1.InstAddrNo=cr1.ChargeAddrNo and cr1.InstAddrNo=cr1.MailAddrNo then  -- ���ۦP
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- ���a�}�r��
	-- ��X�ݲ��ʪ����
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
	if v_Tmp is not null then	-- �s�W�@���ݲ��ʸ��
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.1:�L���˾��a�}���,�Ƚs='||cr1.CustId);
      end;
 
    -- �l���a�}���P
    elsif cr1.InstAddrNo=cr1.ChargeAddrNo and cr1.InstAddrNo!=cr1.MailAddrNo then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- ���a�}�r��
	-- ��X�ݲ��ʪ����
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
        begin
	  select Address into v_Addr1 from SO014 
	    where CompCode=cr1.CompCode and AddrNo=cr1.MailAddrNo;  -- ���a�}�r��
	  if v_Addr1 != cr1.MailAddress then
	    v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	  end if;
        exception
	  when others then  -- Log�@�����~���
	    insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	      values (cr1.CompCode, cr1.MailAddrNo, '1.2:�L���l�H�a�}���,�Ƚs='||cr1.CustId);
        end;
 
	if v_Tmp is not  null then	-- �s�W�@���ݲ��ʸ��
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.3:�L���˾��a�}���,�Ƚs='||cr1.CustId);
      end;
 
    -- ���O�a�}���P
    elsif cr1.InstAddrNo=cr1.MailAddrNo and cr1.InstAddrNo!=cr1.ChargeAddrNo then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- ���a�}�r��
	-- ��X�ݲ��ʪ����
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
        begin
	  select Address into v_Addr1 from SO014 
	    where CompCode=cr1.CompCode and AddrNo=cr1.ChargeAddrNo;  -- ���a�}�r��
	  if v_Addr1 != cr1.ChargeAddress then
	    v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	  end if;
        exception
	  when others then  -- Log�@�����~���
	    insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	      values (cr1.CompCode, cr1.ChargeAddrNo, '1.4:�L�����O�a�}���,�Ƚs='||cr1.CustId);
        end;
 
	if v_Tmp is not null then	-- �s�W�@���ݲ��ʸ��
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.5:�L���˾��a�}���,�Ƚs='||cr1.CustId);
      end;
 
    -- �˾��a�}���P
    elsif cr1.InstAddrNo!=cr1.ChargeAddrNo and cr1.ChargeAddrNo=cr1.MailAddrNo then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.MailAddrNo;  -- ���a�}�r��
	-- ��X�ݲ��ʪ����
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
        begin
	  select Address into v_Addr1 from SO014 
	    where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- ���a�}�r��
	  if v_Addr1 != cr1.InstAddress then
	    v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	  end if;
        exception
	  when others then  -- Log�@�����~���
	    insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	      values (cr1.CompCode, cr1.InstAddrNo, '1.6:�L���˾��a�}���,�Ƚs='||cr1.CustId);
        end;
 
	if v_Tmp is not null then	-- �s�W�@���ݲ��ʸ��
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.MailAddrNo, '1.7:�L���l�H�a�}���,�Ƚs='||cr1.CustId);
      end;
 
    -- 3�Ӧa�}�����P
    else 
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.8:�L���˾��a�}���,�Ƚs='||cr1.CustId);
      end;
 
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.ChargeAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.ChargeAddrNo, '1.9:�L�����O�a�}���,�Ƚs='||cr1.CustId);
      end;
 
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.MailAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.MailAddrNo, '1.10:�L���l�H�a�}���,�Ƚs='||cr1.CustId);
      end;
 
      if v_Tmp is not null then	-- �s�W�@���ݲ��ʸ��
	v_Cmd := 'update SO001 set '||ltrim(v_Tmp,',')||' where CustId='||cr1.CustId;
	insert into Tmp002 (UpdSQL) values (v_Cmd);
      end if;
    end if;
  end loop;
 
  -- �B�z�˾����u�椧�a�}���
  for cr2 in cc2 loop
    begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr2.CompCode and AddrNo=cr2.AddrNo;  -- ���a�}�r��
	if v_Addr1 != cr2.Address then	-- �s�W�@���ݲ��ʸ��
	  v_Cmd := 'update SO007 set Address='||chr(39)||v_Addr1||chr(39)||' where SNo='||
			chr(39)||cr2.SNo||chr(39);
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
    exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr2.CompCode, cr2.AddrNo, '2.�L���a�}���,�u��='||cr2.SNo);
    end;
  end loop;
 
  -- �B�z���׬��u�椧�a�}���
  for cr3 in cc3 loop
    begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr3.CompCode and AddrNo=cr3.AddrNo;  -- ���a�}�r��
	if v_Addr1 != cr3.Address then	-- �s�W�@���ݲ��ʸ��
	  v_Cmd := 'update SO008 set Address='||chr(39)||v_Addr1||chr(39)||' where SNo='||
			chr(39)||cr3.SNo||chr(39);
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
    exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr3.CompCode, cr3.AddrNo, '3.�L���a�}���,�u��='||cr3.SNo);
    end;
  end loop;
 
  -- �B�z��������u�椧�a�}���
  for cr4 in cc4 loop
    v_Cmd := null;
    v_Tmp := null;
 
    begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.OldAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr4.OldAddress then
	  v_Tmp := ',OldAddress='||chr(39)||v_Addr1||chr(39);
	end if;
    exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.OldAddrNo, '4.1:�L���˾��a�}���,�u��='||cr4.SNo);
    end;
 
    if nvl(cr4.ReInstAddrNo,0) != 0 then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.ReInstAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr4.ReInstAddress then
	  v_Tmp := v_Tmp||',ReInstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.ReInstAddrNo, '4.2:�L���s�˾��a�}���,�u��='||cr4.SNo);
      end;
    end if;
 
    if nvl(cr4.NewChargeAddrNo,0) != 0 then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.NewChargeAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr4.NewChargeAddress then
	  v_Tmp := v_Tmp||',NewChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.NewChargeAddrNo, '4.3:�L���s���O�a�}���,�u��='||cr4.SNo);
      end;
    end if;
 
    if nvl(cr4.NewMailAddrNo,0) != 0 then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.NewMailAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr4.NewMailAddress then
	  v_Tmp := v_Tmp||',NewMailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.NewMailAddrNo, '4.4:�L���s�l�H�a�}���,�u��='||cr4.SNo);
      end;
    end if;
 
    if v_Tmp is not null then		-- �s�W�@���ݲ��ʸ��
      v_Cmd := 'update SO009 set '||ltrim(v_Tmp, ',')||' where SNo='||chr(39)||cr4.SNo||chr(39);
      insert into Tmp002 (UpdSQL) values (v_Cmd);
    end if;
  end loop;
 
  -- �B�z�a�}���v�ɤ��a�}���
  for cr5 in cc5 loop
    begin
	select Address into v_Addr1 from SO014 where AddrNo=cr5.AddrNo;  -- ���a�}�r��
	if v_Addr1 != cr5.Address then	-- �s�W�@���ݲ��ʸ��
	  v_Cmd := 'update SO015 set Address='||chr(39)||v_Addr1||chr(39)||
			' where AddrNo='||cr5.AddrNo;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
    exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr5.AddrNo, '5.�L�����v�a�}���,�Ƚs='||cr5.CustId);
    end;
  end loop;
 
  -- �B�z�j�Ӥ��a�}���
  for cr6 in cc6 loop
    v_Cmd := null;
    v_Tmp := null;
 
    -- �D�n�a�}�P�p���a�}'same'
    if nvl(cr6.MainAddrNo,0)!=0 and cr6.MainAddrNo=cr6.ContAddrNo then
      begin
        select Address into v_Addr1 from SO014 where AddrNo=cr6.MainAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr6.MainAddress or v_Addr1 != cr6.ContAddress then
          v_Tmp := ',MainAddress='||chr(39)||v_Addr1||chr(39)||
		',ContAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr6.MainAddrNo, '6.1:�L���D�n�a�}���,�j��='||cr6.MduId);
      end;
    end if;
 
    -- �D�n�a�}�P�p���a�}'Not same', ���D�n�a�}
    if nvl(cr6.MainAddrNo,0)!=0 and cr6.MainAddrNo!=cr6.ContAddrNo then
      begin
        select Address into v_Addr1 from SO014 where AddrNo=cr6.MainAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr6.MainAddress then
          v_Tmp := ',ainAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr6.MainAddrNo, '6.2:�L���D�n�a�}���,�j��='||cr6.MduId);
      end;
    end if;
 
    -- �D�n�a�}�P�p���a�}'Not same', ���p���a�}
    if nvl(cr6.ContAddrNo,0)!=0 and cr6.MainAddrNo!=cr6.ContAddrNo then
      begin
        select Address into v_Addr1 from SO014 where AddrNo=cr6.ContAddrNo;  -- ���a�}�r��
	if v_Addr1 != cr6.ContAddress then
          v_Tmp := ',ContAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log�@�����~���
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr6.ContAddrNo, '6.3:�L���p���a�}���,�j��='||cr6.MduId);
      end;
    end if;
 
    if v_Tmp is not null then		-- �s�W�@���ݲ��ʸ��
      v_Cmd := 'update SO017 set '||ltrim(v_Tmp, ',')||
		' where MduId='||chr(39)||cr6.MduId||chr(39);
      insert into Tmp002 (UpdSQL) values (v_Cmd);
    end if;
  end loop;
 
  dbms_output.put_line('<99> '||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'second');
 
  -- loop����s��ݲ��ʦa�}�ɤ������
  if p_Option = 1 then
    v_Cursor := DBMS_SQL.Open_cursor;
    for cr0 in cc0 loop 
      v_Cmd := cr0.UpdSQL;
      DBMS_SQL.Parse(v_Cursor, v_Cmd, DBMS_SQL.V7);
      v_RetCode := DBMS_SQL.Execute(v_Cursor);
    end loop;
    DBMS_SQL.Close_cursor(v_Cursor); 
  end if;
 
  p_RetMsg := 'Execution OK, it takes'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'second';
  commit;
  return 0;
 
exception
  when others then
    rollback;
    p_RetMsg := 'Other error.';
    return -99;
end;
/

 
create or replace function SF_NewAddress
  (p_UserId varchar2, p_Option number, p_RetMsg out varchar2) return number
  as
 
  v_StartExecTime date;
  v_Addr1 	varchar2(100);
 
  v_Tmp		varchar2(600);
  v_Cmd 	varchar2(600);
  v_Cursor	NUMBER;			-- 供dynamic SQL
  v_RetCode number;
 
  -- 客戶之地址欄位
  cursor cc1 is 
    select CompCode, CustId, InstAddrNo, InstAddress, ChargeAddrNo, ChargeAddress, 
      MailAddrNo, MailAddress from SO001;
 
  -- 裝機派工單之地址欄位
  cursor cc2 is 
    select SNo, CompCode, AddrNo, Address from SO007;
 
  -- 維修派工單之地址欄位
  cursor cc3 is 
    select SNo, CompCode, AddrNo, Address from SO008;
 
  -- 停拆移機派工單之地址欄位
  cursor cc4 is 
    select SNo, CompCode, OldAddrNo, OldAddress, ReInstAddrNo, ReInstAddress, 
	NewChargeAddrNo, NewChargeAddress, NewMailAddrNo, NewMailAddress from SO009;
 
  -- 地址歷史檔之地址欄位
  cursor cc5 is 
    select CustId, AddrNo, Address from SO015;
 
  -- 大樓之地址欄位
  cursor cc6 is
    select MduId, MainAddrNo, MainAddress, ContAddrNo, ContAddress
	from SO017;
 
  -- 待異動資料的SQL指令 (與地址重新對應功能共用Tmp002)
  cursor cc0 is
    select * from Tmp002 where CompCode is null and AddrNo is null;
 
begin
  -- Check parameters
  if (p_UserId is null) then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  v_StartExecTime := sysdate;		-- 開始執行時間
  delete from Tmp002;
 
  -- 處理客戶之地址欄位
  for cr1 in cc1 loop
    v_Cmd := null;
    v_Tmp := null;
 
    if cr1.InstAddrNo=cr1.ChargeAddrNo and cr1.InstAddrNo=cr1.MailAddrNo then  -- 全相同
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- 取地址字串
	-- 找出需異動的欄位
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
	if v_Tmp is not null then	-- 新增一筆待異動資料
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.1:無此裝機地址資料,客編='||cr1.CustId);
      end;
 
    -- 郵遞地址不同
    elsif cr1.InstAddrNo=cr1.ChargeAddrNo and cr1.InstAddrNo!=cr1.MailAddrNo then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- 取地址字串
	-- 找出需異動的欄位
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
        begin
	  select Address into v_Addr1 from SO014 
	    where CompCode=cr1.CompCode and AddrNo=cr1.MailAddrNo;  -- 取地址字串
	  if v_Addr1 != cr1.MailAddress then
	    v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	  end if;
        exception
	  when others then  -- Log一筆錯誤資料
	    insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	      values (cr1.CompCode, cr1.MailAddrNo, '1.2:無此郵寄地址資料,客編='||cr1.CustId);
        end;
 
	if v_Tmp is not  null then	-- 新增一筆待異動資料
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.3:無此裝機地址資料,客編='||cr1.CustId);
      end;
 
    -- 收費地址不同
    elsif cr1.InstAddrNo=cr1.MailAddrNo and cr1.InstAddrNo!=cr1.ChargeAddrNo then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- 取地址字串
	-- 找出需異動的欄位
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
        begin
	  select Address into v_Addr1 from SO014 
	    where CompCode=cr1.CompCode and AddrNo=cr1.ChargeAddrNo;  -- 取地址字串
	  if v_Addr1 != cr1.ChargeAddress then
	    v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	  end if;
        exception
	  when others then  -- Log一筆錯誤資料
	    insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	      values (cr1.CompCode, cr1.ChargeAddrNo, '1.4:無此收費地址資料,客編='||cr1.CustId);
        end;
 
	if v_Tmp is not null then	-- 新增一筆待異動資料
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.5:無此裝機地址資料,客編='||cr1.CustId);
      end;
 
    -- 裝機地址不同
    elsif cr1.InstAddrNo!=cr1.ChargeAddrNo and cr1.ChargeAddrNo=cr1.MailAddrNo then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.MailAddrNo;  -- 取地址字串
	-- 找出需異動的欄位
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := v_Tmp || ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := v_Tmp || ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
 
        begin
	  select Address into v_Addr1 from SO014 
	    where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- 取地址字串
	  if v_Addr1 != cr1.InstAddress then
	    v_Tmp := v_Tmp || ',InstAddress='||chr(39)||v_Addr1||chr(39);
	  end if;
        exception
	  when others then  -- Log一筆錯誤資料
	    insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	      values (cr1.CompCode, cr1.InstAddrNo, '1.6:無此裝機地址資料,客編='||cr1.CustId);
        end;
 
	if v_Tmp is not null then	-- 新增一筆待異動資料
	  v_Cmd := 'update SO001 set '||ltrim(v_Tmp, ',')||' where CustId='||cr1.CustId;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.MailAddrNo, '1.7:無此郵寄地址資料,客編='||cr1.CustId);
      end;
 
    -- 3個地址全不同
    else 
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.InstAddrNo;  -- 取地址字串
	if v_Addr1 != cr1.InstAddress then
	  v_Tmp := ',InstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.InstAddrNo, '1.8:無此裝機地址資料,客編='||cr1.CustId);
      end;
 
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.ChargeAddrNo;  -- 取地址字串
	if v_Addr1 != cr1.ChargeAddress then
	  v_Tmp := ',ChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.ChargeAddrNo, '1.9:無此收費地址資料,客編='||cr1.CustId);
      end;
 
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr1.CompCode and AddrNo=cr1.MailAddrNo;  -- 取地址字串
	if v_Addr1 != cr1.MailAddress then
	  v_Tmp := ',MailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr1.CompCode, cr1.MailAddrNo, '1.10:無此郵寄地址資料,客編='||cr1.CustId);
      end;
 
      if v_Tmp is not null then	-- 新增一筆待異動資料
	v_Cmd := 'update SO001 set '||ltrim(v_Tmp,',')||' where CustId='||cr1.CustId;
	insert into Tmp002 (UpdSQL) values (v_Cmd);
      end if;
    end if;
  end loop;
 
  -- 處理裝機派工單之地址欄位
  for cr2 in cc2 loop
    begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr2.CompCode and AddrNo=cr2.AddrNo;  -- 取地址字串
	if v_Addr1 != cr2.Address then	-- 新增一筆待異動資料
	  v_Cmd := 'update SO007 set Address='||chr(39)||v_Addr1||chr(39)||' where SNo='||
			chr(39)||cr2.SNo||chr(39);
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
    exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr2.CompCode, cr2.AddrNo, '2.無此地址資料,工單='||cr2.SNo);
    end;
  end loop;
 
  -- 處理維修派工單之地址欄位
  for cr3 in cc3 loop
    begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr3.CompCode and AddrNo=cr3.AddrNo;  -- 取地址字串
	if v_Addr1 != cr3.Address then	-- 新增一筆待異動資料
	  v_Cmd := 'update SO008 set Address='||chr(39)||v_Addr1||chr(39)||' where SNo='||
			chr(39)||cr3.SNo||chr(39);
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
    exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr3.CompCode, cr3.AddrNo, '3.無此地址資料,工單='||cr3.SNo);
    end;
  end loop;
 
  -- 處理停拆移機派工單之地址欄位
  for cr4 in cc4 loop
    v_Cmd := null;
    v_Tmp := null;
 
    begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.OldAddrNo;  -- 取地址字串
	if v_Addr1 != cr4.OldAddress then
	  v_Tmp := ',OldAddress='||chr(39)||v_Addr1||chr(39);
	end if;
    exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.OldAddrNo, '4.1:無此裝機地址資料,工單='||cr4.SNo);
    end;
 
    if nvl(cr4.ReInstAddrNo,0) != 0 then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.ReInstAddrNo;  -- 取地址字串
	if v_Addr1 != cr4.ReInstAddress then
	  v_Tmp := v_Tmp||',ReInstAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.ReInstAddrNo, '4.2:無此新裝機地址資料,工單='||cr4.SNo);
      end;
    end if;
 
    if nvl(cr4.NewChargeAddrNo,0) != 0 then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.NewChargeAddrNo;  -- 取地址字串
	if v_Addr1 != cr4.NewChargeAddress then
	  v_Tmp := v_Tmp||',NewChargeAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.NewChargeAddrNo, '4.3:無此新收費地址資料,工單='||cr4.SNo);
      end;
    end if;
 
    if nvl(cr4.NewMailAddrNo,0) != 0 then
      begin
	select Address into v_Addr1 from SO014 
	  where CompCode=cr4.CompCode and AddrNo=cr4.NewMailAddrNo;  -- 取地址字串
	if v_Addr1 != cr4.NewMailAddress then
	  v_Tmp := v_Tmp||',NewMailAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (cr4.CompCode, cr4.NewMailAddrNo, '4.4:無此新郵寄地址資料,工單='||cr4.SNo);
      end;
    end if;
 
    if v_Tmp is not null then		-- 新增一筆待異動資料
      v_Cmd := 'update SO009 set '||ltrim(v_Tmp, ',')||' where SNo='||chr(39)||cr4.SNo||chr(39);
      insert into Tmp002 (UpdSQL) values (v_Cmd);
    end if;
  end loop;
 
  -- 處理地址歷史檔之地址欄位
  for cr5 in cc5 loop
    begin
	select Address into v_Addr1 from SO014 where AddrNo=cr5.AddrNo;  -- 取地址字串
	if v_Addr1 != cr5.Address then	-- 新增一筆待異動資料
	  v_Cmd := 'update SO015 set Address='||chr(39)||v_Addr1||chr(39)||
			' where AddrNo='||cr5.AddrNo;
	  insert into Tmp002 (UpdSQL) values (v_Cmd);
	end if;
    exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr5.AddrNo, '5.無此歷史地址資料,客編='||cr5.CustId);
    end;
  end loop;
 
  -- 處理大樓之地址欄位
  for cr6 in cc6 loop
    v_Cmd := null;
    v_Tmp := null;
 
    -- 主要地址與聯絡地址'same'
    if nvl(cr6.MainAddrNo,0)!=0 and cr6.MainAddrNo=cr6.ContAddrNo then
      begin
        select Address into v_Addr1 from SO014 where AddrNo=cr6.MainAddrNo;  -- 取地址字串
	if v_Addr1 != cr6.MainAddress or v_Addr1 != cr6.ContAddress then
          v_Tmp := ',MainAddress='||chr(39)||v_Addr1||chr(39)||
		',ContAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr6.MainAddrNo, '6.1:無此主要地址資料,大樓='||cr6.MduId);
      end;
    end if;
 
    -- 主要地址與聯絡地址'Not same', 有主要地址
    if nvl(cr6.MainAddrNo,0)!=0 and cr6.MainAddrNo!=cr6.ContAddrNo then
      begin
        select Address into v_Addr1 from SO014 where AddrNo=cr6.MainAddrNo;  -- 取地址字串
	if v_Addr1 != cr6.MainAddress then
          v_Tmp := ',ainAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr6.MainAddrNo, '6.2:無此主要地址資料,大樓='||cr6.MduId);
      end;
    end if;
 
    -- 主要地址與聯絡地址'Not same', 有聯絡地址
    if nvl(cr6.ContAddrNo,0)!=0 and cr6.MainAddrNo!=cr6.ContAddrNo then
      begin
        select Address into v_Addr1 from SO014 where AddrNo=cr6.ContAddrNo;  -- 取地址字串
	if v_Addr1 != cr6.ContAddress then
          v_Tmp := ',ContAddress='||chr(39)||v_Addr1||chr(39);
	end if;
      exception
	when others then  -- Log一筆錯誤資料
	  insert into Tmp002 (CompCode, AddrNo, UpdSQL)
	    values (null, cr6.ContAddrNo, '6.3:無此聯絡地址資料,大樓='||cr6.MduId);
      end;
    end if;
 
    if v_Tmp is not null then		-- 新增一筆待異動資料
      v_Cmd := 'update SO017 set '||ltrim(v_Tmp, ',')||
		' where MduId='||chr(39)||cr6.MduId||chr(39);
      insert into Tmp002 (UpdSQL) values (v_Cmd);
    end if;
  end loop;
 
  dbms_output.put_line('<99> '||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'second');
 
  -- loop執行存於待異動地址檔中的資料
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

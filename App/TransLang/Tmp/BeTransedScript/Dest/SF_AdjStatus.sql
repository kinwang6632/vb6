create or replace function SF_AdjStatus(p_Option1 number, p_Option2 number, 
	p_RetMsg out varchar2) return number
  as
  v_Count	NUMBER := 0;
  v_Count2	number := 0;
  v_RcdNo	number := 0;
  v_Cmd 	varchar2(2000);
  v_Cursor	NUMBER;			-- 供dynamic SQL
  v_RetCode	NUMBER;
  v_ReConCode 	number;
  v_StartExecTime date;
  v_StopExecTime date;
 
  v_CustId	number;
  v_AcceptTime	date;
  v_WorkCode 	number;
  v_NewRefNo	number;
  v_NewWipCode  number;
  v_NewWipName  varchar2(12);
 
  v_RefNo1	number;
  v_RefNo2	number;
  v_RefNo3	number;
  v_AcceptTime1 date;
  v_AcceptTime2 date;
  v_AcceptTime3 date;
 
  v_NewTime	date;
  v_NewStatus	number;
  v_NewName	varchar2(12);
  v_NewBalance	number;
  v_TmpTime1 	date;
  v_TmpTime2 	date;
  v_TmpTime3 	date;
  v_TmpTime4 	date;
  v_TmpSNo	varchar2(9);
  v_ReasonCode	number;
  v_ReasonName	varchar2(12);
  v_NowStr	varchar2(19);
  v_TimeStr	varchar2(19);
 
  -- 所有客戶之相關狀態資料
  cursor cc1 is
    select CustId, CustStatusCode, WipCode1, WipName1, WipCode2, WipName2, WipCode3, WipName3,
	InstTime, InstSNo, StopTime, PR1SNo, PRTime, PR2SNo, DelDate, Balance, MduId, ChargeType
	from SO001 order by CustId; 
 
  -- *************************************************************
  -- 某客戶的裝機未完工log資料
  cursor wa0(v_Cid number) is
    select RefNo, AcceptTime from SO071 where CustId=v_Cid order by AcceptTime desc;
 
  -- 裝機派工檔中有未完工資料, 而裝機未完工log檔中卻無
  cursor wa1 is
    select CustId, SNo from SO007 where SNo not in (select SNo from SO071)
	and FinTime is null and SignDate is null;
 
  -- 裝機未完工log檔中有資料, 而裝機派工檔中卻無未完工資料
  cursor wa2 is
    select CustId, SNo from SO071 where SNo not in 
	(select SNo from SO007 where FinTime is null and SignDate is null);
 
  -- 或: select SNo from SO071 minus
  --       (select SNo from SO007 where FinTime is null and SignDate is null);
 
  -- *************************************************************
  -- 某客戶的維修未完工log資料
  --cursor wb0(v_Cid number) is
  --  select RefNo from SO072 where CustId=v_Cid order by AcceptTime desc;
 
  -- 維修派工檔中有未完工資料, 而維修未完工log檔中卻無
  cursor wb1 is
    select CustId, SNo from SO008 where SNo not in (select SNo from SO072)
	and FinTime is null and SignDate is null;
 
  -- 維修未完工log檔中有資料, 而維修派工檔中卻無未完工資料
  cursor wb2 is
    select CustId, SNo from SO072 where SNo not in 
	(select SNo from SO008 where FinTime is null and SignDate is null);
 
  -- *************************************************************
  -- 某客戶的停拆移機未完工log資料
  cursor wc0(v_Cid number) is
    select RefNo,AcceptTime from SO073 where CustId=v_Cid order by AcceptTime desc;
 
  -- 停拆移機派工檔中有未完工資料, 而停拆移機未完工log檔中卻無
  cursor wc1 is
    select CustId, SNo from SO009 where SNo not in (select SNo from SO073)
	and FinTime is null and SignDate is null;
 
  -- 停拆移機未完工log檔中有資料, 而停拆移機派工檔中卻無未完工資料
  cursor wc2 is
    select CustId, SNo, AcceptTime from SO073 where SNo not in 
	(select SNo from SO009 where FinTime is null and SignDate is null);
 
  -- 某客戶裝機完工資料的完工時間及單號
  cursor cc2(v_Cid number) is
    select A.FinTime, A.SNo from SO007 A, CD005 B where A.CustId=v_CId and 
	A.FinTime is not null and A.InstCode=B.CodeNo and (B.RefNo=1 or B.RefNo=5)
	order by A.FinTime desc;
 
  -- 某客戶停拆移機完工資料的拆機完工時間及單號
  cursor cc3(v_Cid number) is
    select A.FinTime, A.SNo, A.ReasonCode, A.ReasonName from SO009 A, CD007 B 
	where A.CustId=v_CId and 
	A.FinTime is not null and A.PRCode=B.CodeNo and (B.RefNo=2 or B.RefNo=5)
	order by A.FinTime desc;
 
  -- 某客戶停拆移機完工資料的停機完工時間及單號
  cursor cc4(v_Cid number) is
    select A.FinTime, A.SNo, A.ReasonCode, A.ReasonName from SO009 A, CD007 B 
	where A.CustId=v_CId and 
	A.FinTime is not null and A.PRCode=B.CodeNo and B.RefNo=1
	order by A.FinTime desc;
 
BEGIN
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_NowStr := GiPackage.GetDTString2(v_StartExecTime);
 
  -- check parameters
  if p_Option1 is null or p_Option2 is null then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  -- 取停機復機的代碼
  begin
    select CodeNo into v_ReConCode from CD005 where RefNo=7;
  exception
    when others then
      p_RetMsg := 'No RefNo=7 record in CD005    ';
      return -2;
  end;
 
  -- 刪除服務狀態重整log檔
  delete from SO080;
 
 
  -- *************************************************************
  -- (1)若裝機派工檔(SO007)中有未完工資料, 而裝機未完工log檔(SO071)中卻無, 則新增至SO071
  for rwa1 in wa1 loop
    select A.AcceptTime, A.InstCode, B.RefNo into v_AcceptTime1, v_WorkCode, v_RefNo1
	from SO007 A, CD005 B where SNo = rwa1.SNo and A.InstCode=B.CodeNo;
    if p_Option1=1 then
      insert into SO071 (CustId, SNo, AcceptTime, InstCode, RefNo)
	values (rwa1.CustId, rwa1.SNo, v_AcceptTime1, v_WorkCode, v_RefNo1);
    end if;
    v_RcdNo := v_RcdNo + 1;
    insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, rwa1.CustId, 11, rwa1.SNo, '裝機檔有未完工單子,裝機未完工log檔卻無資料. 派工類別='||v_WorkCode||',Reference no.='||v_RefNo1);
  end loop;
 
  -- (2)若裝機未完工log檔中有資料, 而裝機派工檔中卻無未完工資料, 則刪除SO071
  for rwa2 in wa2 loop
    if p_Option1=1 then
      delete from SO071 where SNo=rwa2.SNo;
    end if;
    v_RcdNo := v_RcdNo + 1;
    insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, rwa2.CustId, 12, rwa2.SNo, '裝機未完工log檔有資料, 裝機檔中卻無未完工單子');
  end loop;
 
  -- *************************************************************
  -- (3)若維修派工檔(SO008)中有未完工資料, 而維修未完工log檔(SO072)中卻無, 則新增至SO072
  for rwb1 in wb1 loop
    select A.AcceptTime, A.ServiceCode, B.RefNo into v_AcceptTime2, v_WorkCode, v_RefNo2
	from SO008 A, CD005 B where SNo = rwb1.SNo and A.ServiceCode=B.CodeNo;
    if p_Option1=1 then
      insert into SO072 (CustId, SNo, AcceptTime, ServiceCode, RefNo)
	values (rwb1.CustId, rwb1.SNo, v_AcceptTime2, v_WorkCode, v_RefNo2);
    end if;
    v_RcdNo := v_RcdNo + 1;
    insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, rwb1.CustId, 21, rwb1.SNo, '維修檔有未完工單子,維修未完工log檔卻無資料. 派工類別='||v_WorkCode||',Reference no.='||v_RefNo2);
  end loop;
 
  -- (4)若維修未完工log檔中有資料, 而派工檔中卻無未完工資料, 則刪除SO072
  for rwb2 in wb2 loop
    if p_Option1=1 then
      delete from SO072 where SNo=rwb2.SNo;
    end if;
    v_RcdNo := v_RcdNo + 1;
    insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, rwb2.CustId, 22, rwb2.SNo, '維修未完工log檔有資料, 維修檔中卻無未完工單子');
  end loop;
 
  -- *************************************************************
  -- (5)若停拆移機派工檔(SO009)中有未完工資料, 而停拆移機未完工log檔(SO073)中卻無, 則新增至SO073
  for rwc1 in wc1 loop
    select A.AcceptTime, A.PRCode, B.RefNo into v_AcceptTime3, v_WorkCode, v_RefNo3
	from SO009 A, CD005 B where SNo = rwc1.SNo and A.PRCode=B.CodeNo;
    if p_Option1=1 then
      insert into SO073 (CustId, SNo, AcceptTime, PRCode, RefNo)
	values (rwc1.CustId, rwc1.SNo, v_AcceptTime3, v_WorkCode, v_RefNo3);
    end if;
    v_RcdNo := v_RcdNo + 1;
    insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, rwc1.CustId, 13, rwc1.SNo, '停拆移機檔有未完工單子,停拆移機未完工log檔卻無資料. 派工類別='||v_WorkCode||',Reference no.='||v_RefNo3);
  end loop;
 
  -- (6)若停拆移機未完工log檔中有資料, 而停拆移機派工檔中卻無未完工資料, 則刪除SO073
  for rwc2 in wc2 loop
    if p_Option1=1 then
      delete from SO073 where SNo=rwc2.SNo;
    end if;
    v_RcdNo := v_RcdNo + 1;
    insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, rwc2.CustId, 32, rwc2.SNo, '停拆移機未完工log檔有資料, 停拆移機檔中卻無未完工單子');
  end loop;
 
  v_Cursor := DBMS_SQL.Open_cursor;
 
  -- (7)*************************************************************
  for cr1 in cc1 loop
    v_CustId := cr1.CustId;
    v_Cmd := null;
 
    -- *************************************************************
    -- (8_1)取最近一筆裝機未完工單的相關資料
    v_RefNo1 := null;
    open wa0(v_CustId);
      fetch wa0 into v_RefNo1, v_AcceptTime1;
    close wa0;
 
    -- (8_2)決定最新的裝機派工狀態
    if v_RefNo1 is null then
      v_NewWipCode := 0;		-- (無)
      v_NewWipName := '(none)';
    elsif v_RefNo1=0 then
      v_NewWipCode := 5;		-- 其他(室內移機,換線)
      v_NewWipName := 'other';
    elsif v_RefNo1=1 then
      v_NewWipCode := 1;		-- 裝機中
      v_NewWipName := 'in Inst.';
    elsif v_RefNo1=5 or v_RefNo1=7 then
      v_NewWipCode := 2;		-- 復機中
      v_NewWipName := 'in RC ';
    elsif v_RefNo1=6 then
      v_NewWipCode := 4;		-- 改裝
      v_NewWipName := '改裝';
    else
      v_NewWipCode := 3;		-- 設備加裝
      v_NewWipName := 'Add equip.';
    end if;
 
    -- (8_3)若與原派工狀態不符, 則更新為新的裝機派工狀態
    if v_NewWipCode != cr1.WipCode1 then
      v_RcdNo := v_RcdNo + 1;
      insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, v_CustId, 1, null, '裝機派工狀態: 由 '||cr1.WipCode1||' 改為 '||v_NewWipCode||'('||v_NewWipName||')');
      v_Cmd := v_Cmd || ',WipCode1='||to_char(v_NewWipCode)||',WipName1='||chr(39)||v_NewWipName||chr(39);
    end if;
 
    -- *************************************************************
    -- (9_1)是否有維修未完工log資料
    select count(*) into v_Count2 from SO072 where CustId=v_CustId;
    if v_Count2 > 0 then
      v_NewWipCode := 21;		-- 維修中
      v_NewWipName := 'in TC ';
    else
      v_NewWipCode := 0;		-- (無)
      v_NewWipName := '(none)';
    end if;
 
    -- (9_2)若與原派工狀態不符, 則更新為新的維修派工狀態
    if v_NewWipCode != cr1.WipCode2 then
      v_RcdNo := v_RcdNo + 1;
      insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, v_CustId, 2, null, '維修派工狀態: 由 '||cr1.WipCode2||' 改為 '||v_NewWipCode||'('||v_NewWipName||')');
      v_Cmd := v_Cmd || ',WipCode2='||to_char(v_NewWipCode)||',WipName2='||chr(39)||v_NewWipName||chr(39);
    end if;
 
    -- *************************************************************
    -- (10_1)取最近一筆停拆移機未完工單的相關資料
    v_RefNo3 := null;
    open wc0(v_CustId);
      fetch wc0 into v_RefNo3, v_AcceptTime3;
    close wc0;
 
    -- (10_2)決定最新的停拆移機派工狀態
    if v_RefNo3=3 then
      v_NewWipCode := 11;		-- 移機中
      v_NewWipName := 'in Move';
    elsif v_RefNo3=2 or v_RefNo3=5 then
      v_NewWipCode := 12;		-- 拆機中
      v_NewWipName := 'Pending disc.';
    elsif v_RefNo3=1 then
      v_NewWipCode := 13;		-- 停機中
      v_NewWipName := 'in Disc.';
    elsif v_RefNo3=4 then
      v_NewWipCode := 14;		-- 移拆中
      v_NewWipName := 'Dsic/Move';
    else
      v_NewWipCode := 0;		-- (無)
      v_NewWipName := '(none)';
    end if;
 
    -- (10_3)若與原派工狀態不符, 則更新為新的停拆移機派工狀態
    if v_NewWipCode != cr1.WipCode3 then
      v_RcdNo := v_RcdNo + 1;
      insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, v_CustId, 3, null, '停拆移機派工狀態: 由 '||cr1.WipCode3||' 改為 '||v_NewWipCode||'('||v_NewWipName||')');
      v_Cmd := v_Cmd || ',WipCode3='||to_char(v_NewWipCode)||',WipName3='||chr(39)||v_NewWipName||chr(39);
    end if;
 
    -- *************************************************************
    -- (11)若原裝機時間為空, 則取裝機檔中最近的完工時間; 若有值, 則更新裝機時間
    v_TmpTime1 := null;
    if cr1.InstTime is null then
      open cc2(v_CustId);
        fetch cc2 into v_Tmptime1, v_TmpSNo;
      close cc2;
      if v_TmpTime1 is not null then
	v_TimeStr := GiPackage.QryDTString1(v_TmpTime1);
	v_RcdNo := v_RcdNo + 1;
	insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	  values (v_RcdNo, v_CustId, 41, null, '裝機時間: 由無改為 '||v_TimeStr);
	v_Cmd := v_Cmd || ',InstTime='||v_TimeStr||',InstSNo='||chr(39)||v_TmpSNo||chr(39);
      end if;
    end if;
 
    -- *************************************************************
    -- (12)若原拆機時間為空, 則取拆機檔中最近的拆機完工時間; 若有值, 則更新拆機時間
    v_TmpTime3 := null;
    if cr1.PRTime is null then
      open cc3(v_CustId);
        fetch cc3 into v_TmpTime3, v_TmpSNo, v_ReasonCode, v_ReasonName;
      close cc3;
      if v_TmpTime3 is not null then
	v_TimeStr := GiPackage.QryDTString1(v_TmpTime3);
	v_RcdNo := v_RcdNo + 1;
	insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	  values (v_RcdNo, v_CustId, 43, null, '拆機時間: 由無改為 '||v_TimeStr);
	v_Cmd := v_Cmd || ',PRTime='||v_TimeStr||
		',PR2SNo='||chr(39)||v_TmpSNo||chr(39)||',PRCode='||
		to_char(v_ReasonCode)||',PRName='||chr(39)||v_ReasonName||chr(39);
      end if;
    end if;
 
    -- *************************************************************
    -- (13)若原停機時間為空, 則取拆機檔中最近的停機完工時間; 若有值, 則更新停機時間
    v_TmpTime2 := null;
    if cr1.StopTime is null then
      open cc4(v_CustId);
        fetch cc4 into v_TmpTime2, v_TmpSNo, v_ReasonCode, v_ReasonName;
      close cc4;
      if v_TmpTime2 is not null then  -- 尚需比較停機復機的完工時間
	select max(FinTime) into v_TmpTime4 from SO007 
	  where CustId=v_CustId and FinTime is not null and InstCode=v_ReConCode;
	if v_TmpTime4 is null or v_TmpTime2 > v_TmpTime4 then
	  v_TimeStr := GiPackage.QryDTString1(v_TmpTime2);
	  v_RcdNo := v_RcdNo + 1;
	  insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	    values (v_RcdNo, v_CustId, 42, null, '停機時間: 由無改為 '||v_TimeStr);
	  v_Cmd := v_Cmd || ',StopTime='||v_TimeStr||
		',PR1SNo='||chr(39)||v_TmpSNo||chr(39)||',StopCode='||
		to_char(v_ReasonCode)||',StopName='||chr(39)||v_ReasonName||chr(39);
	end if;
      end if;
    end if;
 
    -- *************************************************************
    -- (14)若有拆機中工單, 則調整客戶狀態為'Pending disc.', 否則: 若非註銷戶, 才決定新的客戶狀態
    -- 否則: no action
    v_NewTime := null;
    v_NewStatus := 5;
    if cr1.DelDate is null and (v_RefNo3=2 or v_RefNo3=5) and cr1.CustStatusCode!=6 then
      v_RcdNo := v_RcdNo + 1;
      insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, v_CustId, 0, null, '客戶狀態: 由 '||cr1.CustStatusCode||' 改為 6(拆機中)');
      v_Cmd := v_Cmd || ',CustStatusCode=6,CustStatusName='||chr(39)||'Pending disc.'||chr(39);
 
    elsif cr1.DelDate is null then
      -- DO NOT USE greatest()
      -- v_NewTime := greatest(cr1.InstTime, cr1.StopTime, cr1.PRTime, v_TmpTime1, v_TmpTime2, v_TmpTime3);
      if cr1.InstTime is not null then
	v_NewStatus := 1;
	v_NewTime := cr1.InstTime;
      end if;
      if cr1.StopTime is not null and (v_NewTime is null or cr1.StopTime > v_NewTime) then
	v_NewStatus := 2;
	v_NewTime := cr1.StopTime;
      end if;
      if cr1.PRTime is not null and (v_NewTime is null or cr1.PRTime > v_NewTime) then
	v_NewStatus := 3;
	v_NewTime := cr1.PRTime;
      end if;
      if v_TmpTime1 is not null and (v_NewTime is null or v_TmpTime1 > v_NewTime) then
	v_NewStatus := 1;
	v_NewTime := v_TmpTime1;
      end if;
      if v_TmpTime2 is not null and (v_NewTime is null or v_TmpTime2 > v_NewTime) then
	v_NewStatus := 2;
	v_NewTime := v_TmpTime2;
      end if;
      if v_TmpTime3 is not null and (v_NewTime is null or v_TmpTime3 > v_NewTime) then
	v_NewStatus := 3;
	v_NewTime := v_TmpTime3;
      end if;
 
      if v_NewStatus=5 then		-- 促銷
	v_NewName := 'On sales';
      elsif v_NewStatus=3 then		-- 拆機
	v_NewName := 'Disc.';
      elsif v_NewStatus=2 then 		-- 停機
	v_NewName := 'Tmp disc.';
      elsif v_NewStatus=1 then 		-- 正常
	v_NewName := 'Active';
      else
	v_NewStatus := 5;		-- 促銷
	v_NewName := 'On sales';
      end if;
 
      if cr1.CustStatusCode != v_NewStatus then
	v_RcdNo := v_RcdNo + 1;
	insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	  values (v_RcdNo, v_CustId, 0, null, '客戶狀態: 由 '||cr1.CustStatusCode||
	  ' 改為 '||v_NewStatus||'('||v_NewName||')');
	v_Cmd := v_Cmd || ',CustStatusCode='||to_char(v_NewStatus)||
		',CustStatusName='||chr(39)||v_NewName||chr(39);
      end if;
 
    end if;
 
    -- *************************************************************
    -- (15)若帳上欠費不同, 則調整之
    select sum(ShouldAmt) into v_NewBalance from SO033 
      where CustId=v_CustId and UCCode is not null;
    v_NewBalance := 0 - v_NewBalance;
    if cr1.Balance != v_NewBalance then
      v_RcdNo := v_RcdNo + 1;
      insert into SO080 (RcdNo, CustId, Type, SNo, Content)
	values (v_RcdNo, v_CustId, 100, null, '客戶欠費: 由 '||cr1.Balance||
	' 改為 '||v_NewBalance);
      v_Cmd := v_Cmd || ',Balance='||to_char(v_NewBalance);
    end if;
 
    -- *************************************************************
    -- (16)若無大樓編號,則設定收費屬性為1 (一般戶)
    if cr1.MduId is not null and cr1.ChargeType != 1 then
      v_Cmd := v_Cmd || ',ChargeType=1';
    end if;
 
    -- *************************************************************
    -- (99)執行異動指令
    if p_Option1=1 and v_Cmd is not null then
      if p_Option2 = 1 then
	v_Cmd := v_Cmd || ',UpdTime='||chr(39)||v_NowStr||chr(39)||
		',UpdEn='||chr(39)||'Adjust customer status'||chr(39);
      end if;
      v_Cmd := 'update SO001 set ' || ltrim(v_Cmd, ',') || ' where CustId=' || to_char(v_CustId);
      DBMS_OUTPUT.PUT_LINE(v_Cmd);
 
      begin
        DBMS_SQL.Parse(v_Cursor, v_Cmd, DBMS_SQL.V7);
        v_RetCode := DBMS_SQL.Execute(v_Cursor);
      exception
        when others then
          DBMS_SQL.Close_cursor(v_Cursor);
          DBMS_OUTPUT.PUT_LINE('SQL指令錯誤: ' || v_Cmd);
	  p_RetMsg := 'SQL指令錯誤: ' || v_Cmd;
          return -3;
      end;
    end if;
 
    v_Count := v_Count + 1;
  end loop; 
  DBMS_SQL.Close_cursor(v_Cursor);  
 
  -- *************************************************************
  v_StopExecTime := sysdate;
  p_RetMsg := 'Execution OK, it takes'||to_char(trunc(86400*(v_StopExecTime-v_StartExecTime)))||'second';
  commit;
  return v_RcdNo;
 
EXCEPTION
  WHEN OTHERS THEN
    DBMS_SQL.Close_cursor(v_Cursor);
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := 'Other error.';
    return -99;
END;
/

/*
@c:\gird\v350\script\SF_RptA1_2
@c:\gird\v300\script\SF_RptA1_2
variable nn number
variable msg varchar2(500)
exec :nn := SF_RptA1_2('20000101', '20000131', '20020531', 'C', 7, null, '=', null, null, 'IN (1, 20, 30, 36)', null, null, null, null, null, null, 2, 0, null, 'IN (2,3)', :msg);
print nn
print msg

prompt 測試版指令
exec :nn := SF_RptA1_2('19980601', '19980630', '19980731', 'C', 2, 455029, '=', null, null, null, null, null, null, null, null, null, 2, 5, null, null, :msg);

-- exec :nn := SF_RptA1_2('20000101', '20010215', '20010131', 'I', 1, 2, 5, :msg);
print nn
print msg

prompt 正式版指令
exec :nn := SF_RptA1_2('20000101', '20000131', '20010930', 'C', 1, 111526, '=',null ,null , null, null, null, null, null, null, null, 2, 0, null, null, :msg);
exec :nn := SF_RptA1_2('20010105', '20010630', '20010630', 'C', 3, null, null, null ,null , null, null, null, null, null, null, null, 2, 0, null, null, :msg);
print nn
print msg

  產生預收攤提報表(A1報表)明細資料: SF_RptA1()
  檔名: SF_RptA1.sql

  (1) 彙總或逐筆型式的報表, 可依據此明細檔產生
  (2) 排序方式於報表產生時, 再進行排序, 此明細檔資料不事先排序
  (3) 若報表尚需其他欄位, 可利用此明細檔去join其他檔, 以產生該欄位內容

  IN	p_RealDate1 varchar2: 	收費日起始, 'YYYYMMDD'
	p_RealDate2 varchar2:	收費日截止, 'YYYYMMDD'
	p_RefDate varchar2:	參考日期, 'YYYYMMDD'
	p_ServiceType char(1):	服務類別, 例: 'C', 'I' or null
	p_CompCode number:	公司別

	p_CustId number:	客戶編號
	p_CustIdOp varchar2:	客戶編號比較元, 例: '=', '>', '<=', ...
	p_ShouldDate1 varchar2:	應收日起始, 'YYYYMMDD'
	p_ShouldDate2 varchar2:	應收日截止, 'YYYYMMDD'
	p_CitemSQL varchar2:	收費項目SQL條件字串
	--p_CompSQL varchar2:	公司別SQL條件字串
	p_ClassSQL varchar2:	客戶類別SQL條件字串
	p_BillTypeSQL varchar2: 單據類別SQL條件字串
	p_ServSQL varchar2:	服務SQL條件字串
	p_AreaSQL varchar2:	行政區SQL條件字串
	p_StrtSQL varchar2:	街道SQL條件字串
	p_CircuitNo varchar2:	網路編號條件, null=全部, '-1'=空白者才統計, 其他字串=符合此字串才統計

	p_TaxMode number:	稅率計算方式, 1=未稅, 2=已稅
	p_ErrYear number:	計費起始日超過幾年視為錯誤期限
	p_MDUSQL varchar2:	大樓編號SQL條件字串
	p_Citem2SQL varchar2:	非週期收費項目SQL條件字串			<==== NEW added: 2001.11.15
  OUT 
      p_RetMsg VARCHAR2：結果訊息 (呼叫端變數至少需100 bytes)

  Return NUMBER：結果代碼
        0: 正常完畢
	-1: p_RetMsg='參數錯誤'
	-2: p_RetMsg='短收代碼檔中設定呆帳參考號有誤'
        -3: p_RetMsg='收費參數檔找不到該公司及該服務類別的資料'
        -4: p_RetMsg='SQL錯誤: '
	-5: p_RetMsg='所有週期性收費項目的稅率皆應設為一致'
        -99：p_RetMsg='其他錯誤'

  By: Pierre
  Date: 2001.02.23
	2001.05.14 加入大樓編號SQL條件
	2001.08.24 改成可處理多筆參考號設為呆帳的短收原因代碼
	2001.11.16 改: 加條件物件--非週期性項目
	2003.03.22 改: 程式回傳錯誤訊息時, 將Oracle的錯誤訊息加入, 以便區別是程式問題,
			資料問題,或者是Oracle問題

	2003.04.29 改:  1. 非週期性負項資料若有有效期限範圍, 則分攤至每月. 若無有效期限範圍, 則取其正項資料
			   的有效期限來計算該負項的有效期限, 之後再分攤至每月.
			2. 若負項資料可對應到正項, 則存於SO089時, 將其收費項目代碼與名稱改成同正項資料
	2003.05.02 改: 2處加upper(), 以解決前端傳來小寫'in'的問題
	2003.05.14 改: 若998負項之起始日比正項的起始日還早, 這是不合理的負項, 將負項之起始日改為正項之
			起始日來計算

  Modify: 2001.10.04 by Lawrence
        2001.10.04 so043加一參數(Taxrate)供本報表做含稅計算用,原"檢查收費項目代碼檔中的稅率設定情形, 若所有週期性收費項目的稅率皆應設為一致"取消此判斷
        2001.10.27 改: 加判斷起迄日同一日且SO043.Para3=1者
        2002.05.10 改: 加判斷實收期數=0且為週期性者insert into SO090
        2002.10.07 改: 每2000筆Commit一次


  By: Stanley
        2002.03.01 在SQL語法中加入So014的Join
        2003.05.07 增加公司別或服務類別條件 
*/
/*
-- 測試版 function header
create or replace function SF_RptA1_2
  (p_RealDate1 varchar2, p_RealDate2 varchar2, p_RefDate varchar2, 
   p_ServiceType char, p_CompCode number, p_TaxMode number, p_ErrYear number, 
   p_RetMsg out varchar2) 
  return number as
*/
-- 正式版 function header
create or replace function SF_RptA1_2
  (p_RealDate1 varchar2, p_RealDate2 varchar2, p_RefDate varchar2, 
   p_ServiceType char, p_CompCode number, p_CustId number, p_CustIdOp varchar2, 
   p_ShouldDate1 varchar2, p_ShouldDate2 varchar2, p_CitemSQL varchar2,
   p_ClassSQL varchar2, p_BillTypeSQL varchar2, p_ServSQL varchar2, 
   p_AreaSQL varchar2, p_StrtSQL varchar2, p_CircuitNo varchar2,
   p_TaxMode number, p_ErrYear number, p_MDUSQL varchar2, p_Citem2SQL varchar2, p_RetMsg out varchar2) 
  return number as

  v_Para3 number;
  v_Para17 number;
  v_OldP3 number;
  v_Cnt   number := 0;
  v_TaxCodeCnt number := 0;
  v_CurrTaxRate number;
  v_TmpErrCnt number := 0;
  v_BadCode  number;
  v_BadCnt   number := 0;
  v_StartExecTime date;

  c39		char(1) := chr(39);
  v_SQL1 	varchar2(4000);
  v_Column1	varchar2(1000);
  v_Where1	varchar2(500);
  v_MoreTable   varchar2(100);
  v_Join	varchar2(100);

  v_TotalDay number;
  v_PreDay number;
  v_PastAmt number;
  v_01Amt number;
  v_02Amt number;
  v_03Amt number;
  v_04Amt number;
  v_05Amt number;
  v_06Amt number;
  v_07Amt number;
  v_08Amt number;
  v_09Amt number; 
  v_10Amt number;
  v_11Amt number;
  v_12Amt number;
  v_ElseAmt number;
  v_Flag number;

  v_CurrMonth number := 0;
  v_RefMonth number := 0;
  v_CurrDate date;
  v_BeginDate date;
  v_LastDate date;

  v_CurrAmt number;
  v_LeftAmt number;
  v_LeftDay number;
  v_CurrDay number;
  v_Mndx number;
  v_ClctMonth number;
  v_TaxRate number;
  v_Count number:=0;

  type CurTyp is ref cursor;  --自訂cursor型態
  v_DyCursor CurTyp;          --供dynamic SQL

  x_BillNo varchar2(15);
  x_Item number;
  x_CustId number;
  x_CitemCode number;
  x_CitemName varchar2(20);
  x_RealDate date;
  x_RealAmt number;
  x_RealPeriod number;
  x_RealStartDate date;
  x_RealStopDate date;
  x_STCode number;
  x_ClctEn varchar2(10);
  x_AddrNo number;
  x_ServCode varchar2(3);
  x_ServiceType char(1);

  y_RealAmt number;
  y_PeriodFlag number;
  -- for 稅率檢查 
  cursor cTaxRate is
	select  B.CodeNo, B.Rate1 
	  from cd019 A, cd033 B where A.TaxCode=B.CodeNo(+) and A.PeriodFlag = 1 and (A.ServiceType=p_ServiceType or A.ServiceType is null)
	  group by b.CodeNo, B.Rate1;
  -- 呆帳
  cursor cBad is
	select CodeNo from CD016 where RefNo = 1 and (ServiceType=p_ServiceType or ServiceType is null);
  v_BadCodeString varchar2(300);

  -- 2003.04.29
  s_CitemCode number;		-- 退費項目所對應之原收費項目
  s_BillNo varchar2(15);	-- 退費項目所對應之原單號
  s_Item number;		-- 退費項目所對應之原項次
  x_ShouldDate date;		-- 應收日/應退日
  v_TmpCitem varchar2(500);	-- for週期收費項目處理用
  I number;
  v_Index  number;
  v_CitemCodeCnt number := 0;
  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER;
  CitemCode NumberAry;

  x_TmpStartDate date;

begin

  -- Check parameters
  if p_RealDate1 is null or p_RealDate2 is null or p_RefDate is null or
    p_CompCode is null or p_ServiceType is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

  -- 處理收費項目SQL條件字串, 將其轉成array, 以便後續判斷負項資料所對應之正項是否落於此條件內 2003.04.29
  DBMS_OUTPUT.PUT_LINE('1. 處理收費項目SQL條件字串');
  v_CitemCodeCnt := 0;
  if p_CitemSQL is not null then
    v_TmpCitem := rtrim(ltrim(p_CitemSQL));
    if upper(substr(v_TmpCitem, 1, 2)) = 'IN' then			-- 'IN (...,..,..)'
      v_TmpCitem := substr(v_TmpCitem, 5);
      v_TmpCitem := substr(v_TmpCitem, 1, length(v_TmpCitem)-1);
    elsif upper(substr(v_TmpCitem, 1, 6)) = 'NOT IN' then		-- 'NOT IN (...,..,..)'
      v_TmpCitem := substr(v_TmpCitem, 9);
      v_TmpCitem := substr(v_TmpCitem, 1, length(v_TmpCitem)-1);
    elsif substr(v_TmpCitem, 1, 1) = '=' then			-- '=...'
      v_TmpCitem := ltrim(substr(v_TmpCitem, 2));
    elsif substr(v_TmpCitem, 1, 2) = '!=' then			-- '!=...'
      v_TmpCitem := ltrim(substr(v_TmpCitem, 3));
    end if;

    v_Index := 1;
    I := 1;
    while v_TmpCitem is not null loop
      v_Index := instr(v_TmpCitem, ',');
      if v_Index > 0 then
	begin
	  CitemCode(I) := ltrim(rtrim(substr(v_TmpCitem, 1, v_Index-1)));
	  v_TmpCitem := substr(v_TmpCitem, v_Index+1);
	  I:=I+1;     
	exception
	  when others then
	    v_TmpCitem := null;
	end;   
      else
	CitemCode(I) := to_number(rtrim(ltrim(v_TmpCitem)));
	v_TmpCitem := null;     
      end if;         
    end loop;
    v_CitemCodeCnt := I;
  end if;

  -- 取收費參數檔相關參數: Para3: 次期起使日同本期截止日
  -- 目前用公司別=1的資料來做, 若非該公司, 則公司別應為必要欄位
  DBMS_OUTPUT.PUT_LINE('2. 取收費參數檔相關參數');
  if p_ServiceType is null then
    begin
      select Para3, Para17,TaxRate into v_Para3, v_Para17, v_TaxRate from SO043 where CompCode=p_CompCode;
    exception
      when others then
        p_RetMsg := '若多公司別且多服務種類的系統, 則服務類別為必要欄位';
        return -3;
    end;
  else
    begin
      select Para3, Para17, TaxRate into v_Para3, v_Para17, v_TaxRate from SO043 where CompCode=p_CompCode and ServiceType=p_ServiceType;
    exception
      when others then
        p_RetMsg := '收費參數檔找不到該公司別及該服務類別的資料';
        return -3;
    end;
  end if;

  v_OldP3 := v_Para3;
  if v_Para3=1 then
     v_Para3 := 0;
  else
     v_Para3 := 1; 
  end if; 

  -- 檢查短收代碼檔中設定呆帳參考號
/*
  begin
    select CodeNo into v_BadCode from CD016 where RefNo = 1 and (ServiceType=p_ServiceType or ServiceType is null);
  exception
    when others then
      p_RetMsg := '短收代碼檔中設定呆帳參考號有誤';
      return -2;
  end;
*/
  DBMS_OUTPUT.PUT_LINE('3. 取短收代碼');
  for cr1 in cBad loop
    v_BadCodeString := v_BadCodeString || '(' || ltrim(to_char(cr1.CodeNo,'999')) || ')';
    v_BadCnt := v_BadCnt + 1;
  end loop;
  -- DBMS_OUTPUT.PUT_LINE(v_BadCodeString);
  if v_BadCodeString is null then
    v_BadCodeString := '(-1)';
  end if;
/*  if v_BadCodeString is null then
    p_RetMsg := '短收代碼檔中設定呆帳參考號有誤(未設RefNo=1的資料)';
    return -2;
  end if; */

/*
  -- 檢查收費項目代碼檔中的稅率設定情形, 若所有週期性收費項目的稅率皆應設為一致, 則檢查之
  if p_TaxMode = 1 then
    for cr2 in cTaxRate loop
      v_TaxCodeCnt := v_TaxCodeCnt + 1;
      v_CurrTaxRate := nvl(cr2.Rate1, 0);
    end loop; 
    if v_TaxCodeCnt != 1 then
      p_RetMsg := '所有週期性收費項目的稅率皆應設為一致';
      return -5;
    end if;
  else		-- 視為已稅(稅內含)方式來計算
    v_CurrTaxRate := 0;
  end if;
*/

 -- 修改 2001.10.03 by Lawrence
  if p_TaxMode = 1 then
    v_CurrTaxRate := v_TaxRate;
  else		
    v_CurrTaxRate := 0;
  end if;

  -- 正式版SQL指令
  -- 組合SQL查詢字串: 目前畫面SO5A10上的所有查詢參數於SO033等資料檔中皆有, 故不用join其他檔
  DBMS_OUTPUT.PUT_LINE('4. 組合SQL查詢字串');
  v_Column1 := 'select A.BillNo, A.Item, A.CustId, A.CitemCode, A.CitemName, A.RealDate, A.RealAmt, ' ||
		'A.RealPeriod, A.RealStartDate, A.RealStopDate, A.STCode, A.ClctEn, A.AddrNo, A.ServCode, A.ServiceType, A.sCitemCode, A.sBillNo, A.sItem, A.ShouldDate';
  v_Where1  :=  ' A.RealDate>=' || GiPackage.QryDTString0(to_date(p_RealDate1,'YYYYMMDD')) ||
		' and A.RealDate<=' || GiPackage.QryDTString0(to_date(p_RealDate2,'YYYYMMDD')) ||
		' and A.UCCode is null and A.CancelFlag=0 and nvl(A.RealAmt,0) !=0';

  -- 若有客戶編號條件
  if p_CustId is not null then
    v_SQL1 := v_SQL1 || ' and A.CustId' || p_CustIdOp || to_char(p_CustId);
  end if;

  -- 若收費項目條件有值

  if p_CitemSQL is not null and p_Citem2SQL is not null then 
--    v_SQL1 := v_SQL1 || ' and (nvl(A.RealPeriod,0)>0 and A.CitemCode ' || p_CitemSQL || ' or A.CitemCode '|| p_Citem2SQL ||')';
    v_SQL1 := v_SQL1 || ' and (A.CitemCode ' || p_CitemSQL || ' or A.CitemCode '|| p_Citem2SQL ||')';
  end if;

  if p_CitemSQL is not null and p_Citem2SQL is null then 
--    v_SQL1 := v_SQL1 || ' and nvl(A.RealPeriod,0)>0 and A.CitemCode ' || p_CitemSQL;
    v_SQL1 := v_SQL1 || ' and A.CitemCode ' || p_CitemSQL;
  end if;
  if p_CitemSQL is null and p_Citem2SQL is not null then 
--    v_SQL1 := v_SQL1 || ' and (nvl(A.RealPeriod,0)>0 or A.CitemCode ' || p_Citem2SQL || ')';
    v_SQL1 := v_SQL1 || ' and  (A.CitemCode ' || p_Citem2SQL || ')';
  end if;
--  if p_CitemSQL is null and p_Citem2SQL is null then 
--    v_SQL1 := v_SQL1 || ' and nvl(A.RealPeriod,0)>0';
--  end if;
/*
  if p_CitemSQL is not null then 
     v_SQL1 := v_SQL1 || ' and A.CitemCode ' || p_CitemSQL;
  end if;
*/
  -- 若有應收日期條件
  if p_ShouldDate1 is not null then
    v_SQL1 := v_SQL1 || ' and A.ShouldDate>=' || GiPackage.QryDTString0(to_date(p_ShouldDate1,'YYYYMMDD')) ||
		 ' and A.ShouldDate<=' || GiPackage.QryDTString0(to_date(p_ShouldDate2,'YYYYMMDD'));
  end if;

  -- 若有公司別條件
  if p_CompCode is not null then
    v_SQL1 := v_SQL1 || ' and A.CompCode=' || to_char(p_CompCode);
  end if;

  -- 若有服務類別條件
  if p_ServiceType is not null then
    v_SQL1 := v_SQL1 || ' and A.ServiceType=' || c39 || p_ServiceType || c39;
  end if;

  -- 若有客戶類別條件
  if p_ClassSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.ClassCode ' || p_ClassSQL;
  end if;

  -- 若有服務區條件
  if p_ServSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.ServCode ' || p_ServSQL;
  end if;

  -- 若有行政區條件
  if p_AreaSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.AreaCode ' || p_AreaSQL;
  end if;

  -- 若有街道條件
  if p_StrtSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.StrtCode ' || p_StrtSQL;
  end if;

  -- 若有單據類別條件
  if p_BillTypeSQL is not null then
    v_SQL1 := v_SQL1 || ' and substr(A.BillNo,7,1) ' || p_BillTypeSQL;
  end if;

  -- 若有網路編號條件
  if p_CircuitNo is not null then
    v_Join := v_Join || ' A.AddrNo=B.AddrNo';
    v_MoreTable := v_MoreTable || ', SO014 B';
    if p_CircuitNo = '-1' then
      v_SQL1 := v_SQL1 || ' and B.CircuitNo is null';
    else
      v_SQL1 := v_SQL1 || ' and B.CircuitNo=' || c39 || p_CircuitNo || c39;
    end if;
  end if;

  -- 若有大樓條件		2001.05.14
  if p_MduSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.MduId ' || p_MduSQL;
  end if;

  -- 整組SQL串起
  if v_Join is not null then
    v_Where1 := ' and' || v_Where1;
  end if;
  v_SQL1 := v_Column1 || ' from SO033 A' || v_MoreTable || ' where ' || v_Join || v_Where1 || v_SQL1 || ' union ' ||
	    v_Column1 || ' from SO034 A' || v_MoreTable || ' where ' || v_Join || v_Where1 || v_SQL1 ;

/*************************************************
  -- Log SQL statement for checking
  delete from TmpSQL;
  insert into TmpSQL (SQLString) values (v_SQL1);
  commit;
  --return 0;
**************************************************/

/*
  -- 測試版SQL指令
  v_SQL1 := 'select BillNo, Item, CustId, CitemCode, CitemName, RealDate, RealAmt, RealPeriod,' ||
	    ' RealStartDate, RealStopDate, STCode, ClctEn, AddrNo, ServCode from SO033' ||
	    ' where RealDate >=' || GiPackage.QryDTString0(to_date(p_RealDate1,'YYYYMMDD')) || 
 	    ' and RealDate <=' || GiPackage.QryDTString0(to_date(p_RealDate2,'YYYYMMDD')) ||
	    ' and RealStartDate is not null and UCCode is null and CancelFlag = 0 and nvl(RealAmt,0) != 0';
	    --' and CustId = 8723';
*/

  v_StartExecTime := sysdate;		-- 開始執行時間
  p_RetMsg := '';
  v_RefMonth := to_number(substr(p_RefDate,1,6));

  DBMS_OUTPUT.PUT_LINE('5. 清除暫存檔內容');
  delete from SO089;			-- 清除暫存檔內容
  delete from SO090;
  commit;

  begin
    open v_DyCursor for v_SQL1;
  exception
    when others then
      rollback;
      p_RetMsg := 'SQL錯誤: ' || v_SQL1;
      return -4;
  end;

  DBMS_OUTPUT.PUT_LINE('6. loop開始細部動作');
  v_BadCnt := 0;
  loop       -- 取到的每一筆資料
    fetch v_DyCursor INTO x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealAmt, x_RealPeriod, x_RealStartDate, x_RealStopDate, x_STCode, 
		x_ClctEn, x_AddrNo, x_ServCode, x_ServiceType, s_CitemCode, s_BillNo, s_Item, x_ShouldDate;
    exit when v_DyCursor%NOTFOUND;

    v_PastAmt := 0;
    v_01Amt := 0;
    v_02Amt := 0;
    v_03Amt := 0;
    v_04Amt := 0;
    v_05Amt := 0;
    v_06Amt := 0;
    v_07Amt := 0;
    v_08Amt := 0;
    v_09Amt := 0;
    v_10Amt := 0;
    v_11Amt := 0;
    v_12Amt := 0;
    v_ElseAmt := 0;

    v_TotalDay := 0;
    v_PreDay := 0;
    v_Flag := 0;
    v_CurrAmt := 0;
    v_LeftAmt := 0;
    v_LeftDay := 0;
    v_CurrDay := 0;
    v_Mndx := 0;
    y_RealAmt := 0;

    -- 檢查此項退費的正項代碼是否落在查詢條件內, 若否, 則不計算此項退費資料 2003.04.29
    if nvl(s_CitemCode,0)!=0 then	-- 退費項目
      v_Index := 0;
      if v_CitemCodeCnt>0 then
	for I in 1..v_CitemCodeCnt loop
	  if CitemCode(I)=s_CitemCode then	-- 查到
	    v_Index := I;
	  end if;
	end loop;
	if v_Index = 0 then			-- 查不到, 不計算此項退費資料
	  goto GO_NEXT1;
	end if;
      end if;
    end if;

    /* 
	需再過濾呆帳與不合邏輯的資料:
	. 有問題的資料(填入SO090檔): 有效期限之一為空值, 有效期限超過年限
	. 呆帳資料: by pass
    */
    -- 2002.05.10 by Lawrence
    select PeriodFlag into y_PeriodFlag from CD019 where CodeNo=x_CitemCode and ServiceType=x_ServiceType;

    --if nvl(x_STCode, -1) = v_BadCode then
    if instr(v_BadCodeString, '('||ltrim(to_char(nvl(x_STCode, 0),'999'))||')') > 0 then
      v_BadCnt := v_BadCnt + 1;			-- by pass 呆帳資料
      goto GO_NEXT1;
    elsif x_RealPeriod>0 and (x_RealStopDate is null or x_RealStartDate is null) 
	or x_RealPeriod>0 and x_RealStopDate<x_RealStartDate then	-- 不合法的有效期限
      insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '不合法的有效期限');
      goto GO_NEXT1;
    elsif x_RealPeriod>0 and nvl(p_ErrYear,0)>0 and months_between(last_day(x_RealStopDate),last_day(x_RealStartDate))>p_ErrYear*12 then
      insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '有效期限超過年限'); 
      goto GO_NEXT1;
    elsif nvl(x_RealPeriod,0)=0 and nvl(y_PeriodFlag,0)=1 then
      insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '不合法的週期性項目'); 
      goto GO_NEXT1;
    end if;

    -- ******************************************************************************************
    if nvl(x_RealPeriod,0)=0 and s_CitemCode is null then	-- 合法資料, 非週期性收費, 且非退費
      if p_TaxMode = 1 then			-- 該月尚未分攤前, 所剩之總金額, 考慮已稅/未稅的影響
        y_RealAmt := round(x_RealAmt*100/(100+v_CurrTaxRate));
      else
	y_RealAmt := x_RealAmt;
      end if;
      --v_CurrMonth := TO_NUMBER(to_char(x_RealDate, 'YYYYMM'));
      v_ClctMonth := to_number(to_char(x_RealDate, 'YYYYMM'));	-- 收費日的月份
      v_Cnt := v_Cnt + 1;
      v_TotalDay := 0;
      v_PreDay := 0;
      v_CurrAmt := y_RealAmt;

      -- 若該月月份<收費日的月份, 則該月分攤金額應歸屬到收費日的月份去, 
      -- 然後才再看收費日月份的該月分攤金額, 應歸屬在第幾個月份
      if v_ClctMonth < v_RefMonth then 			-- 過去的月份: 累計至"過去實現收入"
	v_Mndx := 0;
      elsif v_ClctMonth = v_RefMonth then		-- 參考日該月: 累計至"本月實現收入"
	v_Mndx := 1;
      else						-- 未來的月份 (應不會發生)
	-- 決定該筆分攤金額, 應落在第幾個月份
	v_Mndx := months_between(to_date(to_char(v_ClctMonth)||'01', 'YYYYMMDD'),
		add_months(last_day(to_date(p_RefDate,'YYYYMMDD'))+1, -1)) + 1;
      end if;

  	  if v_Mndx = 0 then
	    v_PastAmt := v_CurrAmt;
	  elsif v_Mndx = 1 then
	    v_01Amt := v_CurrAmt;
	  elsif v_Mndx = 2 then
	    v_02Amt := v_CurrAmt;
	  elsif v_Mndx = 3 then
	    v_03Amt := v_CurrAmt;
	  elsif v_Mndx = 4 then
	    v_04Amt := v_CurrAmt;
	  elsif v_Mndx = 5 then
	    v_05Amt := v_CurrAmt;
	  elsif v_Mndx = 6 then
	    v_06Amt := v_CurrAmt;
	  elsif v_Mndx = 7 then
	    v_07Amt := v_CurrAmt;
	  elsif v_Mndx = 8 then
	    v_08Amt := v_CurrAmt;
	  elsif v_Mndx = 9 then
	    v_09Amt := v_CurrAmt;
	  elsif v_Mndx = 10 then
	    v_10Amt := v_CurrAmt;
	  elsif v_Mndx = 11 then
	    v_11Amt := v_CurrAmt;
	  elsif v_Mndx = 12 then
	    v_12Amt := v_CurrAmt;
	  else
	    v_ElseAmt := v_CurrAmt;
	  end if;

      -- 新增一筆明細至暫存檔: SO089
      begin
        insert into SO089 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
	  RealStartDate, RealStopDate, RealPeriod, RealAmt, TotalDay, PreDay, 
	  MonthPast, Month01, Month02, Month03, Month04, Month05, Month06, Month07, 
	  Month08, Month09, Month10, Month11, Month12, MonthElse,
	  ClctEn, AddrNo, ServCode) values 
	  (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate,
	  null, null, 0, y_RealAmt, v_TotalDay, v_PreDay,
	  v_PastAmt, v_01Amt, v_02Amt, v_03Amt, v_04Amt, v_05Amt, v_06Amt, v_07Amt,  
	  v_08Amt, v_09Amt, v_10Amt, v_11Amt, v_12Amt, v_ElseAmt,
	  x_ClctEn, x_AddrNo, x_ServCode);
      exception
	when others then
	  v_TmpErrCnt := v_TmpErrCnt + 1;
	  insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, '無法新增至SO089'); 
	  goto GO_NEXT1;
      end;

    -- ******************************************************************************************
    elsif nvl(x_RealPeriod,0)!=0 and s_CitemCode is null then	-- 合法資料, 週期性收費
      v_Cnt := v_Cnt + 1;
--      v_TotalDay := x_RealStopDate - x_RealStartDate + v_Para3;	-- 有效期限總日數
--90.07.13 Stanley 修改
--90.10.04 Lawrence 修改
      v_TotalDay := SF_Days360(x_RealStartDate,x_RealStopDate + v_Para3,v_Para17,p_RetMsg);   -- 有效期限總日數
      v_LastDate := x_RealStopDate;				-- 有效期限截止日
      v_ClctMonth := to_number(to_char(x_RealDate, 'YYYYMM'));	-- 收費日的月份

      if v_TotalDay = 0 then			-- 2001.10.27 added by Pierre
	v_TotalDay := 1;
      end if;

      -- 決定攤到每月所需之共用變數
      v_BeginDate := x_RealStartDate;		-- 該月首日
      v_LeftDay := v_TotalDay;			-- 該月尚未分攤前, 所剩之總日數

      if p_TaxMode = 1 then			-- 該月尚未分攤前, 所剩之總金額, 考慮已稅/未稅的影響
        v_LeftAmt := round(x_RealAmt*100/(100+v_CurrTaxRate));
      else
	v_LeftAmt := x_RealAmt;
      end if;
      y_RealAmt := v_LeftAmt;

      while v_Flag = 0 loop			-- Loop將有效期限拆成每一月份
        v_CurrMonth := to_number(to_char(v_BeginDate,'YYYYMM')); 	-- 該月月份
        if x_RealStopDate >= last_day(v_BeginDate) then		-- 該月首日至當月底之日數
          --v_CurrDay := last_day(v_BeginDate)-v_BeginDate+1;
-- 90.07.13 Stanley 修改
--90.10.04 Lawrence 修改
          v_CurrDay := SF_Days360(v_BeginDate,last_day(v_BeginDate)+1,v_Para17,p_RetMsg);

-- DBMS_OUTPUT.PUT_LINE(TO_CHAR(v_BeginDate,'YYYYMMDD') || ', ' || TO_CHAR(last_day(v_BeginDate),'YYYYMMDD') || ', ' || v_CurrDay);
	else
	  --v_CurrDay := to_char(x_RealStopDate,'DD') - v_OldP3;
-- 90.10.04 Lawrence 修改
          v_CurrDay := SF_Days360(to_date(to_char(x_RealStopDate,'YYYYMM')||'01','YYYYMMDD'),x_RealStopDate + v_Para3,v_Para17,p_RetMsg);

-- DBMS_OUTPUT.PUT_LINE(to_char(x_RealStopDate,'YYYYMM')||'01' || ', ' || TO_CHAR(x_RealStopDate,'YYYYMMDD') || ', ' || v_CurrDay);
	end if;
	if v_CurrDay > v_LeftDay then	-- 但: 若日數大於所剩之總日數, 則應修正成剩餘日數, 否則無法處理1日的問題
	  v_CurrDay := v_LeftDay;
	end if;

/*	-- test code: 2001.10.27
if v_LeftDay=0 or v_CurrDay=0 then
	DBMS_OUTPUT.PUT_LINE('客編=' || x_CustId || ', 單號=' || x_BillNo || ', 項目=' || x_CitemName);
	DBMS_OUTPUT.PUT_LINE('v_LeftAmt ='||v_LeftAmt||', v_LeftDay = '||v_LeftDay||', v_CurrDay = ' ||v_CurrDay);
end if;
*/
        v_CurrAmt := round(v_LeftAmt / v_LeftDay * v_CurrDay); 	 	-- 該月分攤金額
	
	-- 預收日數的定義: 該月月份大於參考日該月, 則屬之
        if v_CurrMonth > v_RefMonth then
	  v_PreDay := v_PreDay + v_CurrDay;
	end if;

	-- 若該月月份<收費日的月份, 則該月分攤金額應歸屬到收費日的月份去, 
	-- 然後才再看收費日月份的該月分攤金額, 應歸屬在第幾個月份
	if v_CurrMonth < v_ClctMonth then
	  v_CurrMonth := v_ClctMonth;
	end if;

          if v_CurrMonth < v_RefMonth then 		-- 過去的月份: 累計至"過去實現收入"
	    v_Mndx := 0;
	  elsif v_CurrMonth = v_RefMonth then		-- 參考日該月: 累計至"本月實現收入"
	    v_Mndx := 1;
	  else						-- 未來的月份
	    -- 決定該筆分攤金額, 應落在第幾個月份
	    v_Mndx := months_between(to_date(to_char(v_CurrMonth)||'01', 'YYYYMMDD'),
	  	add_months(last_day(to_date(p_RefDate,'YYYYMMDD'))+1, -1)) + 1;
	  end if;

  	  if v_Mndx = 0 then
	    v_PastAmt := v_PastAmt + v_CurrAmt;
	  elsif v_Mndx = 1 then
	    v_01Amt := v_01Amt + v_CurrAmt;
	  elsif v_Mndx = 2 then
	    v_02Amt := v_02Amt + v_CurrAmt;
	  elsif v_Mndx = 3 then
	    v_03Amt := v_03Amt + v_CurrAmt;
	  elsif v_Mndx = 4 then
	    v_04Amt := v_04Amt + v_CurrAmt;
	  elsif v_Mndx = 5 then
	    v_05Amt := v_05Amt + v_CurrAmt;
	  elsif v_Mndx = 6 then
	    v_06Amt := v_06Amt + v_CurrAmt;
	  elsif v_Mndx = 7 then
	    v_07Amt := v_07Amt + v_CurrAmt;
	  elsif v_Mndx = 8 then
	    v_08Amt := v_08Amt + v_CurrAmt;
	  elsif v_Mndx = 9 then
	    v_09Amt := v_09Amt + v_CurrAmt;
	  elsif v_Mndx = 10 then
	    v_10Amt := v_10Amt + v_CurrAmt;
	  elsif v_Mndx = 11 then
	    v_11Amt := v_11Amt + v_CurrAmt;
	  elsif v_Mndx = 12 then
	    v_12Amt := v_12Amt + v_CurrAmt;
	  else
	    v_ElseAmt := v_ElseAmt + v_CurrAmt;
	  end if;
/*
	DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_CurrMonth || ', ' || v_CurrDay || ', ' || v_CurrAmt);
        DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_PastAmt || ', ' || v_LeftDay || ', ' || v_LeftAmt);
*/
	-- 重新決定攤到每月所需之共用變數
        v_BeginDate := last_day(v_BeginDate) + 1;	-- 次月首日
        v_LeftDay := v_LeftDay - v_CurrDay;		-- 次月剩餘總日數
        v_LeftAmt := v_LeftAmt - v_CurrAmt;		-- 次月剩餘總額 

	-- 若已超出有效截止日,或所剩日數為0, 則表示應跳出迴圈
        if v_BeginDate > v_LastDate or v_LeftDay = 0 then	-- or v_LeftAmt = 0
          v_Flag := 1;
        end if;

        v_Count:=v_Count+1;
        if v_Count=2000 then
            Commit;
            v_Count:=0;
        end if;

      end loop;

/*
      DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_PastAmt || ', ' || v_01Amt || ', ' || v_02Amt || ', ' || 
	v_03Amt || ', ' || v_04Amt || ', ' || v_05Amt || ', ' || v_06Amt || ', ' || v_07Amt || ', ' || 
	v_08Amt || ', ' || v_09Amt || ', ' || v_10Amt || ', ' || v_11Amt || ', ' || v_12Amt || ', ' || v_ElseAmt);
*/
      -- 新增一筆明細至暫存檔: SO089
      -- 註: 實收金額應填 y_RealAmt, 不是 x_RealAmt
      begin
        insert into SO089 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
	  RealStartDate, RealStopDate, RealPeriod, RealAmt, TotalDay, PreDay, 
	  MonthPast, Month01, Month02, Month03, Month04, Month05, Month06, Month07, 
	  Month08, Month09, Month10, Month11, Month12, MonthElse,
	  ClctEn, AddrNo, ServCode) values 
	  (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate,
	  x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, v_TotalDay, v_PreDay,
	  v_PastAmt, v_01Amt, v_02Amt, v_03Amt, v_04Amt, v_05Amt, v_06Amt, v_07Amt,  
	  v_08Amt, v_09Amt, v_10Amt, v_11Amt, v_12Amt, v_ElseAmt,
	  x_ClctEn, x_AddrNo, x_ServCode);
      exception
	when others then
	  v_TmpErrCnt := v_TmpErrCnt + 1;
	  insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, '無法新增至SO089'); 
	  goto GO_NEXT1;
      end;

    -- ******************************************************************************************
    else		-- 退費項目	2003.04.29
      -- 取該筆退費資料的有效期限, 若有, 則以此範圍去分攤至每月, 
      -- 若無, 則依退費資料的應收日, 及取其對應正項資料的有效截止日當做有效期限來分攤至每月
      if x_RealStartDate is null or x_RealStopDate is null then
	-- x_RealStartDate := x_ShouldDate;
	begin
	  select RealStartDate, RealStopDate, CitemName into x_RealStartDate, x_RealStopDate, x_CitemName from v_Charge 
		where BillNo=s_BillNo and Item=s_Item and nvl(CancelFlag,0)=0;
	exception
	  when others then
	    insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '此退費資料無有效的正項資料'); 
	    goto GO_NEXT1;
	end;
	x_RealStartDate := greatest(x_ShouldDate, x_RealStartDate);	-- 2003.05.14
      else 	-- 已有退費資料的有效期限, 故只再取其正項收費資料之項目名稱
	begin
	  select Description into x_CitemName from CD019 where CodeNo=s_CitemCode;
	exception
	  when others then
	    insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '無正項資料的收費項目名稱, 代碼='||s_CitemCode); 
	    goto GO_NEXT1;
	end;

	-- 再取正項之起始日, 以最大的起始日為主 2003.05.14
	begin
	  select RealStartDate into x_TmpStartDate from v_Charge where BillNo=s_BillNo and Item=s_Item and nvl(CancelFlag,0)=0;
	exception
	  when others then
	    insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '此退費資料無有效的正項資料'); 
	    goto GO_NEXT1;
	end;
	x_RealStartDate := greatest(x_TmpStartDate, x_RealStartDate);
      end if;

      /**********************************************************************************************
	至此, 吾人已知此退費資料之對應正項資料的CitemCode, CitemName, 以及退費所含蓋之有效期限 !!!
	分攤退費金額至每月 
      ***********************************************************************************************/
      v_Cnt := v_Cnt + 1;
      v_TotalDay := SF_Days360(x_RealStartDate,x_RealStopDate + v_Para3,v_Para17,p_RetMsg);   -- 有效期限總日數
      v_LastDate := x_RealStopDate;				-- 有效期限截止日
      v_ClctMonth := to_number(to_char(x_RealDate, 'YYYYMM'));	-- 收費日的月份

      if v_TotalDay = 0 then			-- 2001.10.27 added by Pierre
	v_TotalDay := 1;
      end if;

      -- 決定攤到每月所需之共用變數
      v_BeginDate := x_RealStartDate;		-- 該月首日
      v_LeftDay := v_TotalDay;			-- 該月尚未分攤前, 所剩之總日數
      if p_TaxMode = 1 then			-- 該月尚未分攤前, 所剩之總金額, 考慮已稅/未稅的影響
        v_LeftAmt := round(x_RealAmt*100/(100+v_CurrTaxRate));
      else
	v_LeftAmt := x_RealAmt;
      end if;
      y_RealAmt := v_LeftAmt;

      while v_Flag = 0 loop			-- Loop將有效期限拆成每一月份
        v_CurrMonth := to_number(to_char(v_BeginDate,'YYYYMM')); 	-- 該月月份
        if x_RealStopDate >= last_day(v_BeginDate) then		-- 該月首日至當月底之日數
          v_CurrDay := SF_Days360(v_BeginDate,last_day(v_BeginDate)+1,v_Para17,p_RetMsg);
	else
          v_CurrDay := SF_Days360(to_date(to_char(x_RealStopDate,'YYYYMM')||'01','YYYYMMDD'),x_RealStopDate + v_Para3,v_Para17,p_RetMsg);
	end if;
	if v_CurrDay > v_LeftDay then	-- 但: 若日數大於所剩之總日數, 則應修正成剩餘日數, 否則無法處理1日的問題
	  v_CurrDay := v_LeftDay;
	end if;

/*	-- test code: 2001.10.27
if v_LeftDay=0 or v_CurrDay=0 then
	DBMS_OUTPUT.PUT_LINE('客編=' || x_CustId || ', 單號=' || x_BillNo || ', 項目=' || x_CitemName);
	DBMS_OUTPUT.PUT_LINE('v_LeftAmt ='||v_LeftAmt||', v_LeftDay = '||v_LeftDay||', v_CurrDay = ' ||v_CurrDay);
end if;
*/
        v_CurrAmt := round(v_LeftAmt / v_LeftDay * v_CurrDay); 	 	-- 該月分攤金額
	
	-- 預收日數的定義: 該月月份大於參考日該月, 則屬之
        if v_CurrMonth > v_RefMonth then
	  v_PreDay := v_PreDay + v_CurrDay;
	end if;

	-- 若該月月份<收費日的月份, 則該月分攤金額應歸屬到收費日的月份去, 
	-- 然後才再看收費日月份的該月分攤金額, 應歸屬在第幾個月份
	if v_CurrMonth < v_ClctMonth then
	  v_CurrMonth := v_ClctMonth;
	end if;

          if v_CurrMonth < v_RefMonth then 		-- 過去的月份: 累計至"過去實現收入"
	    v_Mndx := 0;
	  elsif v_CurrMonth = v_RefMonth then		-- 參考日該月: 累計至"本月實現收入"
	    v_Mndx := 1;
	  else						-- 未來的月份
	    -- 決定該筆分攤金額, 應落在第幾個月份
	    v_Mndx := months_between(to_date(to_char(v_CurrMonth)||'01', 'YYYYMMDD'),
	  	add_months(last_day(to_date(p_RefDate,'YYYYMMDD'))+1, -1)) + 1;
	  end if;

  	  if v_Mndx = 0 then
	    v_PastAmt := v_PastAmt + v_CurrAmt;
	  elsif v_Mndx = 1 then
	    v_01Amt := v_01Amt + v_CurrAmt;
	  elsif v_Mndx = 2 then
	    v_02Amt := v_02Amt + v_CurrAmt;
	  elsif v_Mndx = 3 then
	    v_03Amt := v_03Amt + v_CurrAmt;
	  elsif v_Mndx = 4 then
	    v_04Amt := v_04Amt + v_CurrAmt;
	  elsif v_Mndx = 5 then
	    v_05Amt := v_05Amt + v_CurrAmt;
	  elsif v_Mndx = 6 then
	    v_06Amt := v_06Amt + v_CurrAmt;
	  elsif v_Mndx = 7 then
	    v_07Amt := v_07Amt + v_CurrAmt;
	  elsif v_Mndx = 8 then
	    v_08Amt := v_08Amt + v_CurrAmt;
	  elsif v_Mndx = 9 then
	    v_09Amt := v_09Amt + v_CurrAmt;
	  elsif v_Mndx = 10 then
	    v_10Amt := v_10Amt + v_CurrAmt;
	  elsif v_Mndx = 11 then
	    v_11Amt := v_11Amt + v_CurrAmt;
	  elsif v_Mndx = 12 then
	    v_12Amt := v_12Amt + v_CurrAmt;
	  else
	    v_ElseAmt := v_ElseAmt + v_CurrAmt;
	  end if;
/*
	DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_CurrMonth || ', ' || v_CurrDay || ', ' || v_CurrAmt);
        DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_PastAmt || ', ' || v_LeftDay || ', ' || v_LeftAmt);
*/
	-- 重新決定攤到每月所需之共用變數
        v_BeginDate := last_day(v_BeginDate) + 1;	-- 次月首日
        v_LeftDay := v_LeftDay - v_CurrDay;		-- 次月剩餘總日數
        v_LeftAmt := v_LeftAmt - v_CurrAmt;		-- 次月剩餘總額 

	-- 若已超出有效截止日,或所剩日數為0, 則表示應跳出迴圈
        if v_BeginDate > v_LastDate or v_LeftDay = 0 then	-- or v_LeftAmt = 0
          v_Flag := 1;
        end if;

        v_Count:=v_Count+1;
        if v_Count=2000 then
            Commit;
            v_Count:=0;
        end if;
      end loop;

      -- 新增一筆明細至暫存檔: SO089
      -- 註: 實收金額應填 y_RealAmt, 不是 x_RealAmt
      -- 注意: "收費項目代碼/名稱"用正項資料的欄位取代,不用退費資料本身的收費項目代碼/名稱,以便報表展現與排序
      begin
        insert into SO089 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
	  RealStartDate, RealStopDate, RealPeriod, RealAmt, TotalDay, PreDay, 
	  MonthPast, Month01, Month02, Month03, Month04, Month05, Month06, Month07, 
	  Month08, Month09, Month10, Month11, Month12, MonthElse,
	  ClctEn, AddrNo, ServCode) values 
	  (x_BillNo, x_Item, x_CustId, s_CitemCode, x_CitemName, x_RealDate,
	  x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, v_TotalDay, v_PreDay,
	  v_PastAmt, v_01Amt, v_02Amt, v_03Amt, v_04Amt, v_05Amt, v_06Amt, v_07Amt,  
	  v_08Amt, v_09Amt, v_10Amt, v_11Amt, v_12Amt, v_ElseAmt,
	  x_ClctEn, x_AddrNo, x_ServCode);
      exception
	when others then
	  v_TmpErrCnt := v_TmpErrCnt + 1;
	  insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, s_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, '無法新增至SO089'); 
	  goto GO_NEXT1;
      end;

    end if;

    -- ******************************************************************************************
    <<GO_NEXT1>>
    NULL;
  end loop;
 
  commit;
  p_RetMsg := '產生完畢, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||
		'秒, 產生明細筆數='||to_char(v_Cnt);
  return 0;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    p_RetMsg := 'Oracle訊息: '||SQLERRM||', 錯誤資料: 客編=' || x_CustId || ', 單號=' || x_BillNo || ', 項目=' || x_CitemName || ', SBillNo='||s_BillNo||', SCitemCode='||s_CitemCode;
    --rollback;
    commit;
    return -99;
end;
/
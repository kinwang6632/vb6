/*
  Dynamic SQL 範例程式

  IN  ....
  OUT ....
  Return NUMBER：結果代碼
	...
        ??: p_RetMsg='SQL錯誤: '
	...
        -99：p_RetMsg='其他錯誤'
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

  type CurTyp is ref cursor;	--自訂cursor型態
  cDynamic CurTyp;          	--供dynamic SQL


begin

  --***** 1. 組合SQL查詢字串: (此例較複雜, 將SQL指令拆成多段來組合, 實際上可能教簡單)
  v_Column1 := 'select A.BillNo, A.Item, A.CustId, A.CitemCode, A.CitemName, A.RealDate, A.RealAmt, ' ||
		'A.RealPeriod, A.RealStartDate, A.RealStopDate, A.STCode, A.ClctEn, A.AddrNo, A.ServCode, A.ServiceType';
  v_Where1  :=  ' A.RealDate>=' || GiPackage.QryDTString0(to_date(p_RealDate1,'YYYYMMDD')) ||
		' and A.RealDate<=' || GiPackage.QryDTString0(to_date(p_RealDate2,'YYYYMMDD')) ||
		' and A.UCCode is null and A.CancelFlag=0 and nvl(A.RealAmt,0) !=0';

  -- 若有公司別條件
  if p_CompCode is not null then
    v_SQL1 := v_SQL1 || ' and A.CompCode=' || to_char(p_CompCode);
  end if;

  -- 若有服務類別條件
  if p_ServiceType is not null then
    v_SQL1 := v_SQL1 || ' and A.ServiceType=' || c39 || p_ServiceType || c39;
  end if;

  -- 若有單據類別條件
  if p_BillTypeSQL is not null then
    v_SQL1 := v_SQL1 || ' and substr(A.BillNo,7,1) ' || p_BillTypeSQL;
  end if;

  -- 整組SQL串起
  if v_Join is not null then
    v_Where1 := ' and' || v_Where1;
  end if;
  v_SQL1 := v_Column1 || ' from SO033 A' || v_MoreTable || ' where ' || v_Join  || v_Where1 || v_SQL1 || ' union ' ||
	    v_Column1 || ' from SO034 A' || v_MoreTable || ' where ' || v_Join || v_Where1 || v_SQL1 ;


  --*****  2. 開啟Dynamic sql
  begin
    open cDynamic for v_SQL1;
  exception
    when others then
      rollback;
      p_RetMsg := 'SQL錯誤: ' || v_SQL1;
      return -4;
  end;


  --*****  3. loop取每一筆資料來處理
  v_BadCnt := 0;
  loop       -- 取到的每一筆資料
    fetch cDynamic INTO x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealAmt, x_RealPeriod, x_RealStartDate, x_RealStopDate, x_STCode, 
		x_ClctEn, x_AddrNo, x_ServCode, x_ServiceType;
    exit when cDynamic%NOTFOUND;
    .......
  end loop;


  --*****  4. Close dynamic cursor
  close cDynamic;


  commit;
  p_RetMsg := '產生完畢, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||
		'秒, 產生明細筆數='||to_char(v_Cnt);
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
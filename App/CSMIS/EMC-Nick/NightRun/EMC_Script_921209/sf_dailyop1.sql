/*  
@C:\Gird\v400\Script\SF_DailyOp1
@c:\EMC_Script\SF_DailyOp1

variable nn number
variable TName varchar2(20)
variable msg varchar2(100)

exec :nn := SF_DailyOp1(7, 'C', '20021013', '200210141500', 'Pierre', 0, null, :TName, :msg);
print nn
print TName
print msg

  營運日報表: 每日統計客戶裝機/拆機數, MTD, YTD		***** V3.0 *****
  檔名: SF_DailyOp1.sql

  IN	p_CompCode number:	公司別
	p_ServiceType char(1):	服務類別, 例: 'C', 'I'
	p_StopDate  varchar2:	統計截止日, 'YYYYMMDD'
	p_CutOffTime varchar2:	Cut off時間, 'YYYYMMDDHHMI'
	p_Operator varchar2:	操作者名稱
	p_PRFlag number:	拆機中客戶是否列入有效戶, 0=否, 1=是
	p_ReturnSQL varchar2:	退單原因SQL條件

  OUT   p_TmpTableName varchar2: 統計結果暫存檔名 (呼叫端變數至少需20 bytes)
	p_RetMsg VARCHAR2：結果訊息 (呼叫端變數至少需100 bytes)

  Return NUMBER：結果代碼
        0: 正常完畢
	-1: p_RetMsg='參數錯誤'
	-2: p_RetMsg='暫存檔檔名建立錯誤, 請檢查是否有建立Sequence物件: S_DSSRPT_TableName';
	-3: p_RetMsg='SQL1錯誤: ' || v_SQL1
	-4: p_RetMsg='無法刪除暫存檔, 可能他人正使用本報表中'
	-5: p_RetMsg='無法建立暫存檔, 可能他人正使用本報表中, 請稍後再試'
	-6: p_RetMsg='存檔至營運日報log檔錯誤, SQLCODE='||SQLCODE||', '||'SQLERRM='||SQLERRM
        -99：p_RetMsg='其他錯誤'

  暫存檔架構:
	CompCode number(3) 			公司別
	ServiceType char(1) 			服務類別
	ItemCode1 number(3)			項目代碼1
	ItemName1 varchar2(40)			項目名稱1
	ItemCode2 number(3)			項目代碼2
	ItemName2 varchar2(40)			項目名稱2
	Value01 number(8)			數值1
	Value02 number(8)			數值2
	Value03 number(8)			數值3
	Value04 number(8)			數值4
	Value05 number(8)			數值5
	Value06 number(8)			數值6
	Value07 number(8)			數值7
	Value08 number(8)			數值8

  By: Pierre
  Date: 2001.03.20
	2001.03.26 調整: 使SO093成為正式的log檔
	2001.03.26 加入: 統計"其他"的取消裝機數
	2001.04.06 調整: 其他.拆機中/裝機中(MTD)的定義改為從頭至今日, 取消裝機中的YTD項目
	2001.09.07 調整: 總銷數 --> 總銷售
	2001.09.25 調整: 規格(1)~(27)項的完工時間 --> 受理時間
	2001.11.05 調整: 1. 小計.淨成長/有效客戶: 分成上下二行, 即增加"有效客戶(YTD)"這一行數據
			 2. 其他.拆機中, 其他.裝機中: 統計依據為受理日範圍 ==> 2001.09.25就改成如此
			 3. 其他.裝機中: 定義由參考號=1, 改為參考號=1或5
			 4. 其他.取消裝機數: 統計依據改為簽收日範圍 ==> 本來就如此
			 5. 其他.取消裝機數: 條件再加入退單原因的限制(3 or 6)
			 6. 增加輸入參數: 拆機中客戶是否列入有效戶, 退單原因的限制

	2002.01.09 加入: 定址解碼器相關21項統計資料, 欄位編號[57]~[77] (請參考公式文件)
	2002.01.16 修改: 公式[77]之bug
	2002.01.23 修改: 停機復機的bug
	2002.10.14 改: 改成V3.0的版本, CustStatusCode移至SO002中

  By: Lawrence
        2002.01.04 調整: 1. 裝機.復機: 定義加入參考號=7 (停機復機)
                         2. 裝機.合計: 定義加入參考號=7 (停機復機)
                         3. 拆機.拆機: 定義加入參考號=1 (停機)
                         4. 其他.拆機中: 定義加入參考號=1 (停機)
                         5. 其他.裝機中: 定義加入參考號=7 (停機復機)
                         6. 其他.取消裝機數: 定義加入參考號=7 (停機復機)
	2002.04.09 改: 當有勾選拆機中客戶是否列入有效戶, 計算有效客戶數bug
        2003.11.18 Mark delete用法,改為Truncate
	
*/
create or replace function SF_DailyOp1
  (p_CompCode number, p_ServiceType char, p_StopDate varchar2,
   p_CutoffTime varchar2, p_Operator varchar2,
   p_PRFlag number, p_ReturnSQL varchar2, 
   p_TmpTableName out varchar2, p_RetMsg out varchar2) 
  return number as

  -- p_StartDate varchar2

  v_StartExecTime date;
  v_StartDate	varchar2(8);
  c39		char(1) := chr(39);
  v_Cnt1	number;
  v_ErrCnt1	number;
  v_Cnt2	number;
  v_ErrCnt2	number;

  v_ThisYear	char(4);
  --v_ThisMonth	char(6);
  v_SQL1 	varchar2(2000);
  v_SQL2 	varchar2(1000);
  v_SQL3 	varchar2(500);
  v_Where1	varchar2(100);
  v_ReturnCode  varchar2(500);

  z_Type number;
  z_Amt number;

  type CurTyp is ref cursor;
  cDynamic CurTyp;

  -- ***********************
  v_0110_1 number := 0;
  v_0110_2 number := 0;
  v_0110_3 number := 0;
  v_0110_4 number := 0;
  v_0110_5 number := 0;

  v_0120_1 number := 0;
  v_0120_2 number := 0;
  v_0120_3 number := 0;
  v_0120_4 number := 0;
  v_0120_5 number := 0;

  v_0199_1 number := 0;
  v_0199_2 number := 0;
  v_0199_3 number := 0;
  v_0199_4 number := 0;
  v_0199_5 number := 0;


  -- ***********************
  v_0310_1 number := 0;
  v_0310_2 number := 0;
  v_0310_3 number := 0;
  v_0310_5 number := 0;

  v_0320_1 number := 0;
  v_0320_2 number := 0;
  v_0320_3 number := 0;
  v_0320_5 number := 0;

  v_0399_1 number := 0;
  v_0399_2 number := 0;
  v_0399_3 number := 0;
  v_0399_5 number := 0;

  -- ***********************
  v_1010_1 number := 0;
  v_1010_2 number := 0;
  v_1010_3 number := 0;
  v_1010_5 number := 0;

  v_1020_1 number := 0;
  v_1020_2 number := 0;
  v_1020_3 number := 0;
  v_1020_5 number := 0;

  v_1099_1 number := 0;
  v_1099_2 number := 0;
  v_1099_3 number := 0;
  v_1099_5 number := 0;

  -- ***********************
  v_2010_1 number := 0;
  v_2010_2 number := 0;
  v_2010_3 number := 0;
  v_2010_5 number := 0;

  v_3010_1 number := 0;
  v_3010_2 number := 0;
  v_3010_3 number := 0;
  v_3010_5 number := 0;

  v_3020_5 number := 0;				-- 2001.11.05

  -- ***********************
  --v_4010_1 number := 0;
  --v_4010_2 number := 0;
  v_4010_3 number := 0;
  --v_4010_5 number := 0;

  --v_4020_1 number := 0;
  v_4020_2 number := 0;
  v_4020_3 number := 0;
  v_4020_5 number := 0;

  v_4030_1 number := 0;
  v_4030_2 number := 0;
  v_4030_3 number := 0;
  v_4030_5 number := 0;

  -- ***********************	2002.01.09 added
  v_5010_1 number := 0;
  v_5010_2 number := 0;
  v_5010_3 number := 0;
  v_5010_5 number := 0;

  v_5020_1 number := 0;
  v_5020_2 number := 0;
  v_5020_3 number := 0;
  v_5020_5 number := 0;

  v_5030_1 number := 0;
  v_5030_2 number := 0;
  v_5030_3 number := 0;
  v_5030_5 number := 0;

  v_5040_1 number := 0;
  v_5040_2 number := 0;
  v_5040_3 number := 0;
  v_5040_5 number := 0;

  v_5099_1 number := 0;
  v_5099_2 number := 0;
  v_5099_3 number := 0;
  v_5099_5 number := 0;

  v_6010_3 number := 0;

begin
  p_TmpTableName := null;

  -- Check parameters
  if p_StopDate is null or
    p_CompCode is null or p_ServiceType is null or p_CutoffTime is null or
    p_PRFlag<0 or p_PRFlag>1 then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

  v_StartExecTime := sysdate;		-- 開始執行時間
  v_StartDate := substr(p_StopDate,1,6) || '01';
  -- DELETE FROM SO093;
  -- 2003.11.18 by Lawrence Mark delete用法,改為Truncate
  DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO093');  
  -- commit;

/* **********************************************************************************
  -- 組成暫存檔名
  begin
    select S_DSSRPT_TableName.NextVal into v_Cnt1 from dual;
  exception
    when others then
      p_RetMsg := '暫存檔檔名建立錯誤, 請檢查是否有建立Sequence物件: S_DSSRPT_TableName';
      return -2;
  end;
  p_TmpTableName := 'DSSRPT_' || ltrim(to_char(v_Cnt1, '09999999'));
  --DBMS_OUTPUT.PUT_LINE('暫存檔名: ' || p_TmpTableName);

  --檢查是否有建立暫存檔, 若有則先刪除, 再建立之
  select count(*) into v_Cnt1 from User_Tables where Table_Name = p_TmpTableName;
  if v_Cnt1 > 0 then 
    v_SQL3 := 'drop table ' || p_TmpTableName;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL3);
    exception
      when others then
        p_RetMsg := '無法刪除暫存檔, 可能他人正使用本報表中';
        return -4;
    end;
  end if;
  v_SQL3 := 'CREATE TABLE ' || p_TmpTableName || '(' ||
	'CompCode number(3), ServiceType char(1), ItemCode1 number(3), ItemName1 varchar2(40),
	ItemCode2 number(3), ItemName2 varchar2(40), Value01 number(8) default 0,
	Value02 number(8) default 0, Value03 number(8) default 0, Value04 number(8) default 0,
	Value05 number(8) default 0, Value06 number(8) default 0, Value07 number(8) default 0,
	Value08 number(8) default 0)';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL3);
  exception
    when others then
      p_RetMsg := '無法建立暫存檔, 可能他人正使用本報表中, 請稍後再試';
      return -5;
  end;
*********************************************************************************** */

  v_ThisYear := substr(p_StopDate,1,4);
  --v_ThisMonth := substr(p_StopDate,1,6);
  v_Where1 := 'A.CompCode='||p_CompCode||' and A.ServiceType='||c39||p_ServiceType||c39;

  -- ***************************************************************************
  -- 統計截止日當天: 銷售(1) -- 業務銷售 [新裝機(10), 復機(20), 總數(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.AcceptTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39||' group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- 新裝機
      v_0110_2 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- 拆機復機 , 停機復機 2002.01.04
      v_0120_2 := v_0120_2 + z_Amt;
    end if;
  end loop;
  v_0199_2 := v_0110_2 + v_0120_2;	-- 總數
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(1)-業務銷售--截止日當天: 新裝機='||v_0110_2||', 復機='||v_0120_2||', 總數='||v_0199_2);

  -- ***************************************************************************
  -- 統計截止日當天: 銷售(1) -- 總銷售 [新裝機(10), 復機(20), 總數(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.AcceptTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- 新裝機
      v_0110_3 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- 拆機復機 , 停機復機 2002.01.04
      v_0120_3 := v_0120_3 + z_Amt;
    end if;
  end loop;
  v_0199_3 := v_0110_3 + v_0120_3;	-- 總數
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(1)-總銷售--截止日當天: 新裝機='||v_0110_3||', 復機='||v_0120_3||', 總數='||v_0199_3);

  -- ***************************************************************************
  -- 統計截止日MTD: 銷售(1) [新裝機(10), 復機(20), 總數(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- 新裝機
      v_0110_4 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- 拆機復機 , 停機復機 2002.01.04
      v_0120_4 := v_0120_4 + z_Amt;
    end if;
  end loop;
  v_0199_4 := v_0110_4 + v_0120_4;	-- 客戶增加合計
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(1)-統計截止日MTD: 新裝機='||v_0110_4||', 復機='||v_0120_4||', 總數='||v_0199_4);


  -- ***************************************************************************
  -- 統計截止日YTD: 銷售(1) [新裝機(10), 復機(20), 總數(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- 新裝機
      v_0110_5 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- 拆機復機 , 停機復機 2002.01.04
      v_0120_5 := v_0120_5 + z_Amt;
    end if;
  end loop;
  v_0199_5 := v_0110_5 + v_0120_5;	-- 客戶增加合計
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(1)-統計截止日YTD: 新裝機='||v_0110_5||', 復機='||v_0120_5||', 總數='||v_0199_5);

  -- ***************************************************************************
  -- 統計截止日當天: 銷售(1) -- CSR [新裝機(10), 復機(20), 總數(99)]
  v_0110_1 := v_0110_3 - v_0110_2;
  v_0120_1 := v_0120_3 - v_0120_2;
  v_0199_1 := v_0199_3 - v_0199_2;


  -- ***************************************************************************
  -- 統計起始日至截止日前一天: 銷售(3) [業務銷售(20)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39 ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0320_1 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(3)-起始日至截止日前一天: 業務銷售='||v_0320_1);

  -- ***************************************************************************
  -- 統計起始日至截止日前一天: 銷售(3) [總銷售(99)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0399_1 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(3)-起始日至截止日前一天: 總銷售='||v_0399_1);

  -- ***************************************************************************
  -- 統計截止日YTD: 銷售(3) [業務銷售(20)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39 ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0320_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(3)-截止日YTD: 業務銷售='||v_0320_5);

  -- ***************************************************************************
  -- 統計截止日YTD: 銷售(3) [總銷售(99)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0399_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('銷售(3)-截止日YTD: 總銷售='||v_0399_5);

  -- ***************************************************************************
  v_0310_1 := v_0399_1 - v_0320_1;		-- 起始日至截止日前一天: 銷售(3) [CSR(10)]

  v_0310_2 := v_0199_1;				-- 統計截止日當天: 銷售(3) [CSR(10)]
  v_0320_2 := v_0199_2;				-- 統計截止日當天: 銷售(3) [業務銷售(20)]
  v_0399_2 := v_0199_3;				-- 統計截止日當天: 銷售(3) [總銷售(99)]

  v_0310_3 := v_0310_1 + v_0310_2;		-- 起始日至截止當日MTD: 銷售(3) [CSR(10)]
  v_0320_3 := v_0320_1 + v_0320_2;		-- 起始日至截止當日MTD: 銷售(3) [業務銷售(20)]
  v_0399_3 := v_0399_1 + v_0399_2;		-- 起始日至截止當日MTD: 銷售(3) [總銷售(99)]

  v_0310_5 := v_0399_5 - v_0320_5;		-- 截止日YTD: 銷售(3) [CSR(10)]


  -- ***************************************************************************
  -- 統計起始日至截止日前一天: 裝機(10) [新裝機(10), 復機(20), 客戶增加合計(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- 新裝機
      v_1010_1 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- 拆機復機 , 停機復機 2002.01.04
      v_1020_1 := v_1020_1 + z_Amt;
    end if;
  end loop;
  v_1099_1 := v_1010_1 + v_1020_1;	-- 客戶增加合計
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('起始日至截止日前一天: 新裝機='||v_1010_1||', 復機='||v_1020_1||', 客戶增加合計='||v_1099_1);

  -- ***************************************************************************
  -- 統計截止日當天: 裝機(10) [新裝機(10), 復機(20), 客戶增加合計(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.FinTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- 新裝機
      v_1010_2 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- 拆機復機 , 停機復機 2002.01.04
      v_1020_2 := v_1020_2 + z_Amt;
    end if;
  end loop;
  v_1099_2 := v_1010_2 + v_1020_2;	-- 客戶增加合計
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('截止日當天: 新裝機='||v_1010_2||', 復機='||v_1020_2||', 客戶增加合計='||v_1099_2);

  -- ***************************************************************************
  -- 統計本月至截止日當天(MTD): 裝機(10) [新裝機(10), 復機(20), 客戶增加合計(99)]
  v_1010_3 := v_1010_1 + v_1010_2;
  v_1020_3 := v_1020_1 + v_1020_2;
  v_1099_3 := v_1099_1 + v_1099_2;

  -- ***************************************************************************
  -- 統計YTD: 裝機(10) [新裝機(10), 復機(20), 客戶增加合計(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- 新裝機
      v_1010_5 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- 拆機復機 , 停機復機 2002.01.04
      v_1020_5 := v_1020_5 + z_Amt;
    end if;
  end loop;
  v_1099_5 := v_1010_5 + v_1020_5;	-- 客戶增加合計
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('YTD: 新裝機='||v_1010_5||', 復機='||v_1020_5||', 客戶增加合計='||v_1099_5);

  -- ***************************************************************************
  -- 統計起始日至截止日前一天: 拆機(20) [拆機(10)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_2010_1 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('起始日至截止日前一天: 拆機='||v_2010_1);

  -- ***************************************************************************
  -- 統計截止日當天: 拆機(20) [拆機(10)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.FinTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;

  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_2010_2 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('截止日當天: 拆機='||v_2010_2);

  -- ***************************************************************************
  -- 統計MTD: 拆機(20) [拆機(10)]
  v_2010_3 := v_2010_1 + v_2010_2;

  -- ***************************************************************************
  -- 統計YTD: 拆機(20) [拆機(10)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_2010_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('YTD: 拆機='||v_2010_5);

  -- ***************************************************************************
  -- 統計: 小計(30) [淨成長(10)]
  v_3010_1 := v_1099_1 - v_2010_1;
  v_3010_2 := v_1099_2 - v_2010_2;
  v_3010_3 := v_1099_3 - v_2010_3;
  v_3010_5 := v_1099_5 - v_2010_5;

  -- ***************************************************************************
  -- 統計: 小計(30) [有效客戶(20) YTD]				2001.11.05
  /* v_3020_5: 
	拆機中客戶有可能不列入有效客戶, 視參數決定
  */

  --** v_SQL1 := 'select count(*) Amt from SO001 where ';
  v_SQL1 := 'select count(*) Amt from SO002 where ';
  if p_PRFlag = 0 then			-- 拆機中客戶不算
    v_SQL1 := v_SQL1 || 'CustStatusCode=1';
  else					-- 拆機中客戶要算
    v_SQL1 := v_SQL1 || '(CustStatusCode=1 or CustStatusCode=6)';
  end if;
  v_SQL1 := v_SQL1 ||' and ServiceType='||c39||p_ServiceType||c39;

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_3020_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('小計: 有效客戶(YTD)='||v_3020_5);


  -- ***************************************************************************
/* -- 統計起始日至截止日MTD: 其他(40) [拆機中(10)] 
  v_SQL1 := 'select count(*) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=2 or B.RefNo=5)'; */
  -- 統計從頭至截止日: 其他(40) [拆機中(10)] 
  -- 2001.04.06
  v_SQL1 := 'select count(*) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)';

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4010_3 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('起始日至截止日MTD: 拆機中='||v_4010_3);

  -- ***************************************************************************
  -- 統計截止日當日: 其他(40) [裝機中(20)]
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and trunc(A.AcceptTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)'; 
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4020_2 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('截止日當日: 裝機中='||v_4020_2);

  -- ***************************************************************************
/*  -- 統計起始日至截止日MTD: 其他(40) [裝機中(20)] 
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and B.RefNo=1'; */
  -- 統計從頭至截止日: 其他(40) [裝機中(20)] 
  -- 2001.04.06
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)'; 
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4020_3 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('起始日至截止日MTD: 裝機中='||v_4020_3);
/*
  -- ***************************************************************************
  -- 統計起始日至截止日YTD: 其他(40) [裝機中(20)] 
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=5)';

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4020_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('起始日至截止日YTD: 裝機中='||v_4020_5);
*/

  -- ***************************************************************************
  v_ReturnCode := ' and A.ReturnCode is not null';	-- not null是必要條件
  if p_ReturnSQL is not null then
    v_ReturnCode := v_ReturnCode || ' and A.ReturnCode ' || p_ReturnSQL;
  end if;

  -- 統計起始日至截止日前一天: 其他(40) [取消裝機數(30)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.SignDate>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.SignDate<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' || v_ReturnCode;
  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_4030_1 := z_Amt;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('起始日至截止日前一天: 取消裝機數='||v_4030_1);

  -- ***************************************************************************
  -- 統計截止日當天: 其他(40) [取消裝機數(30)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.SignDate=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' || v_ReturnCode;
  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_4030_2 := z_Amt;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('截止日當天: 取消裝機數='||v_4030_2);

  -- ***************************************************************************
  -- 統計本月至截止日當天MTD: 其他(40) [取消裝機數(30)]
  v_4030_3 := v_4030_1 + v_4030_2;

  -- ***************************************************************************
  -- 統計YTD: 其他(40) [取消裝機數(30)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.SignDate>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.SignDate<=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' || v_ReturnCode;
  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_4030_5 := z_Amt;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('YTD: 取消裝機數='||v_4030_5);

  -- ***************************************************************************
  -- 統計: 定址解碼器.銷售/出租/借出/客戶自備.當日
  v_SQL1 := 'select A.BuyCode Type,nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1 group by A.BuyCode' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then				-- 銷售
      v_5010_2 := z_Amt;
    elsif z_Type=3 then				-- 出租
      v_5020_2 := z_Amt;
    elsif z_Type=4 then				-- 借出
      v_5030_2 := z_Amt;
    elsif z_Type=5 then				-- 客戶自備
      v_5040_2 := z_Amt;
    end if;
  end loop;
  v_5099_2 := v_5010_2 + v_5020_2 + v_5030_2 + v_5040_2;	-- 總數
  close cDynamic;

  -- ***************************************************************************
  -- 統計: 定址解碼器.銷售/出租/借出/客戶自備.當月MTD
  v_SQL1 := 'select A.BuyCode Type,nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.InstDate<=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1 group by A.BuyCode' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then				-- 銷售
      v_5010_3 := z_Amt;
    elsif z_Type=3 then				-- 出租
      v_5020_3 := z_Amt;
    elsif z_Type=4 then				-- 借出
      v_5030_3 := z_Amt;
    elsif z_Type=5 then				-- 客戶自備
      v_5040_3 := z_Amt;
    end if;
  end loop;
  v_5099_3 := v_5010_3 + v_5020_3 + v_5030_3 + v_5040_3;	-- 總數
  close cDynamic;

  -- ***************************************************************************
  -- 統計: 定址解碼器.銷售/出租/借出/客戶自備.當月至前一日
  v_5010_1 := v_5010_3 - v_5010_2;
  v_5020_1 := v_5020_3 - v_5020_2;
  v_5030_1 := v_5030_3 - v_5030_2;
  v_5040_1 := v_5040_3 - v_5040_2;
  v_5099_1 := v_5099_3 - v_5099_2;

  -- ***************************************************************************
  -- 統計: 定址解碼器.銷售/出租/借出/客戶自備.YTD
  v_SQL1 := 'select A.BuyCode Type,nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.InstDate<=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1 group by A.BuyCode' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then				-- 銷售
      v_5010_5 := z_Amt;
    elsif z_Type=3 then				-- 出租
      v_5020_5 := z_Amt;
    elsif z_Type=4 then				-- 借出
      v_5030_5 := z_Amt;
    elsif z_Type=5 then				-- 客戶自備
      v_5040_5 := z_Amt;
    end if;
  end loop;
  v_5099_5 := v_5010_5 + v_5020_5 + v_5030_5 + v_5040_5;	-- 總數
  close cDynamic;

  -- ***************************************************************************
  -- 統計: 定址解碼器.待裝解碼器.MTD/YTD (MTD, YTD值皆一樣)
  v_SQL1 := 'select nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate is null and A.PrDate is null ' ||
		' and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_6010_3 := z_Amt;
  close cDynamic;


  -- ***************************************************************************
  -- 存至log檔: SO093
  begin
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value04, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 1, '銷售', 10, '新裝機', 
	v_0110_1, v_0110_2, v_0110_3, v_0110_4, v_0110_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value04, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 1, '銷售', 20, '復機', 
	v_0120_1, v_0120_2, v_0120_3, v_0120_4, v_0120_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value04, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 1, '銷售', 99, '總數', 
	v_0199_1, v_0199_2, v_0199_3, v_0199_4, v_0199_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 3, '銷售', 10, '客服銷售', 
	v_0310_1, v_0310_2, v_0310_3, v_0310_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 3, '銷售', 20, '業務銷售', 
	v_0320_1, v_0320_2, v_0320_3, v_0320_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 3, '銷售', 99, '總銷售', 
	v_0399_1, v_0399_2, v_0399_3, v_0399_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 10, '裝機', 10, '新裝機', 
	v_1010_1, v_1010_2, v_1010_3, v_1010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 10, '裝機', 20, '復機', 
	v_1020_1, v_1020_2, v_1020_3, v_1020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 10, '裝機', 99, '客戶增加合計', 
	v_1099_1, v_1099_2, v_1099_3, v_1099_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 20, '拆機', 10, '拆機', 
	v_2010_1, v_2010_2, v_2010_3, v_2010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 30, '小計', 10, '淨成長', 
	v_3010_1, v_3010_2, v_3010_3, v_3010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 30, '小計', 20, '有效客戶', 
	v_3020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value03, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '其他', 10, '拆機中', 
	v_4010_3, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
/*    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '其他', 20, '裝機中', 
	v_4020_2, v_4020_3, v_4020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator); */
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value02, Value03, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '其他', 20, '裝機中', 
	v_4020_2, v_4020_3, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '其他', 30, '取消裝機數', 
	v_4030_1, v_4030_2, v_4030_3, v_4030_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    -- 2002.01.09
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '定址解碼器', 10, '銷售', 
	v_5010_1, v_5010_2, v_5010_3, v_5010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '定址解碼器', 20, '出租', 
	v_5020_1, v_5020_2, v_5020_3, v_5020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '定址解碼器', 30, '借出', 
	v_5030_1, v_5030_2, v_5030_3, v_5030_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '定址解碼器', 40, '客戶自備', 
	v_5040_1, v_5040_2, v_5040_3, v_5040_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '定址解碼器', 99, '總數', 
	v_5099_1, v_5099_2, v_5099_3, v_5099_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 60, '定址解碼器', 10, '待裝解碼器', 
	null, null, v_6010_3, v_6010_3, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

  exception
    when others then
      p_RetMsg := '存檔至營運日報log檔錯誤, SQLCODE='||SQLCODE||', '||'SQLERRM='||SQLERRM;
      rollback;
      return -6;
  end;


  commit;
  p_RetMsg := '產生完畢, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'秒';
  return 0;

exception
  when others then
    close cDynamic;
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    return -99;
end;
/
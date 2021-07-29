/*
@c:\EMC_script\SP_Gen5DRptData1 
set serveroutput on
exec SP_Gen5DRptData1;  

  批次作業: EMC營運管理報表資料(SO5D01/02/03/04/07)產生作業
  檔名: SP_Gen5DRptData1.sql

  說明: 
	1. 此程式可用Night-run來自動每月5日凌晨執行一次
	2. 此程式的執行順序應該放在[EMC每個月備份重要資料程式(SP_BackupCustData)]之後
	3. 此程式的統計資料來源檔為EMCxxx_SO001ALL_yyyymm(每月每SO一個), 及裝機派工資
	   料檔/停拆移機派工資料檔
	4. 統計結果存放於"營運管理報表資料檔"(SO511A/B/C), 用公司別及資料月份來區別, 並透
	   過備份機制(Snapshot), 將各SO每月統計資料傳送到總公司的備份主機中, 供Web report
	   或前端報表程式取用
	5. 因為"有效戶數"為統計當時的值, 並不是上月底的值, 故此程式會統計上月底之後
	   的裝機戶數與報拆戶數, 並反應到統計值去
	6. 如何驗證已掛上Oracle night-run?
		執行 select * from User_Jobs指令, 看看是否有該SP_Gen5DRptData1的job
	7. 如何檢查每次night-run的結果?
		執行 select * from SO030A where JobId=999987

  結果檔案結構: 
  1. 5D01營運管理報表資料檔: SO511A
	. 公司別		CompCode	N(3)
	. 資料月份		RptYM		N(6), 西曆年月
	. 客戶大類別代碼	ClassCode	N(3), 1~7
	. 繳費期別代碼		PeriodCode	N(3), 1~5
	. 下收月份		NextYM		N(6), 西曆年月
	. 客戶數		CustCount	N(8)
	. 資料種類		DataType	N(1), 0=A項, 1=D項
     索引:
	. RptYm
	. CompCode, RptYM, ClassCode, PeriodCode, NextYM

  2. 5D02收費單數預估統計表: SO511B
	. 公司別		CompCode	N(3)
	. 資料月份		RptYM		N(6), 西曆年月
	. 客戶大類別代碼	ClassCode	N(3), 1~7
	. 繳費期別代碼		PeriodCode	N(3), 1~5
	. 下收月份		NextYM		N(6), 西曆年月
	. 客戶數		CustCount	N(8)
     索引:
	. RptYm
	. CompCode, RptYM, ClassCode, PeriodCode, NextYM

  3. 5D03收費金額預估統計表: SO511C
	. 公司別		CompCode	N(3)
	. 資料月份		RptYM		N(6), 西曆年月
	. 客戶大類別代碼	ClassCode	N(3), 1~7
	. 繳費期別代碼		PeriodCode	N(3), 1~5
	. 下收月份		NextYM		N(6), 西曆年月
	. 金額			Amount		N(10)
     索引:
	. RptYm
	. CompCode, RptYM, ClassCode, PeriodCode, NextYM

  4. 各繳費期別之牌價表: SO511Z
	. 公司別		CompCode	N(3)
	. 客戶大類別代碼	ClassCode	N(3), 1~7
	. 繳費期別代碼		PeriodCode	N(3), 1~5
	. 牌價金額		UnitPrice	N(8)
     索引:
	. CompCode, ClassCode, PeriodCode

  By: Pierre
  Date: 2003.09.04
	2003.09.05 改:	1. 客戶類別702改成701
			2. A項客戶數要計算筆數, 不是CustCount加起來

	2003.09.26 改: 上月底之後的裝機戶數報拆戶數改成由A項的數字減去, 原先由D項中扣除是不對的
*/

CREATE OR REPLACE PROCEDURE SP_Gen5DRptData1  IS
  v_RetMsg	varchar2(100);
  v_RetCode	number	:= -99;
  v_CompCode	number;
  v_OwnerName	varchar2(20);
  v_JobID	number;
  v_StartExecTime date;
  v_StopExecTime  date;
  v_ServiceType     char(1);
  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_SQL		varchar2(4000); -- for SQL statement
  v_Time	date;
  v_RptYM	number;		-- 資料月份
  v_StartDate	date;		-- 資料截止日之上月首日
  v_StopDate	date;		-- 資料截止日
  v_TableName	varchar2(100);	-- 統計資料來源檔檔名
  v_cnt1	number;		-- counter 1
  c39		char(1) := chr(39);
  I		number;
  J		number;
  K		number;
  v_StartMonth	number;		-- 資料截止日上月之總月數, 例: 200307==>2003*12+7
  v_CntA	number;
  v_CntD	number;
  v_Delta	number;
  v_UnitPrice	number;

  type CurTyp is ref cursor;	-- 自訂cursor型態
  v_DyCursor1 CurTyp;		-- 供dynamic SQL
  v_DyCursor2 CurTyp;		-- 供dynamic SQL
  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER; --宣告一個Table的型態(用法和陣列同)
  ACustCount NumberAry;		-- A項結果陣列
  DCustCount NumberAry;		-- D項結果陣列
  MonthAry NumberAry;		-- 下收月份陣列, 內存v_StartDate起的4個月份

  TYPE VC2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER; --宣告一個Table的型態(用法和陣列同)
  TNameAry VC2Ary;		-- array for table names

  x_ClassCode	number;		-- 從[統計資料來源檔]取出的客戶類別代碼
  x_Period	number;		-- 從[統計資料來源檔]取出的期數
  x_CustCount	number;		-- 從[統計資料來源檔]取出的客戶數
  x_ClctDate	date;		-- 從[統計資料來源檔]取出的下收日
  z_ClassCode	number;		-- 客戶大類別代碼
  z_PeriodCode	number;		-- 繳費期別代碼
  z_MonthIndex	number;		-- 下收日所落的月份index, 可能值為1~14

begin
  -- ********************************************************
  -- 設定參數初值
  -- ********************************************************
  v_JobID	:= 999987;		-- 開博內定Job編輯
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_ServiceType := 'C';
  v_FuncName	:= 'EMC營運管理報表資料(SO5D01/02/03/04/07)產生作業';	-- 功能名稱
  v_PrgName	:= 'SP_GEN5DRPTDATA1';	-- 程式名稱

  -- ********************************************************
  -- 決定[資料截止日], [資料月份]及各變數初值
  -- ********************************************************
  v_StopDate	:= last_day(add_months(v_StartExecTime,-1));
  --v_StopDate	:= to_date('20030731', 'YYYYMMDD');	-- for testing
  v_RptYM	:= to_number(to_char(v_StopDate, 'YYYYMM'));
  v_StartDate	:= last_day(add_months(v_StartExecTime, -3))+1;
  v_StartMonth	:= to_number(to_char(v_StartDate,'YYYY'))*12 + to_number(to_char(v_StartDate,'MM'));
  for I in 1..14 loop
    MonthAry(I) := to_number(to_char(add_months(v_StartDate,I-1),'YYYYMM'));
  end loop;

  -- ********************************************************
  -- 取公司別代碼
  -- ********************************************************
  begin 
    select CompCode into v_CompCode from SO501A where RowNum<=1;
  exception
    when others then
      v_RetMsg := '取公司代碼時發生錯誤, '||SQLERRM;
      v_RetCode := -1;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  -- ********************************************************
  -- 取該資料區主要Owner name
  -- ********************************************************
  begin
    select upper(Description) into v_OwnerName from SO507 where CodeNo=v_CompCode;
    ---- DBMS_OUTPUT.PUT_LINE('Owner name = '||v_OwnerName);
  exception
    when others then
      v_RetMsg := '取該資料區主要Owner name時發生錯誤, '||SQLERRM;
      v_RetCode := -2;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  -- ********************************************************
  -- 組成[統計資料來源檔]檔名
  -- ********************************************************
  v_TableName := v_OwnerName||'_SO001ALL_'||ltrim(to_char(v_RptYM,'999999'));
  ---- DBMS_OUTPUT.PUT_LINE('Table name = '||v_TableName);

  -- ********************************************************
  -- 檢查資料區中是否已有此[統計資料來源檔], 若無, 則log錯誤並結束程式
  -- ********************************************************
  select count(*) into v_cnt1 from user_tables where table_name=v_TableName;
  if v_cnt1 = 0 then 
    v_RetMsg := '此資料區無統計資料來源檔, 檔名: '||v_TableName;
    v_RetCode := -3;
    DBMS_OUTPUT.PUT_LINE(v_RetMsg);
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
    return;
  end if; 

  -- ********************************************************
  -- 檢查資料區中是否已有[營運管理報表資料檔], 若無, 則log錯誤並結束程式
  -- ********************************************************
  TNameAry(1) := 'SO511A';
  TNameAry(2) := 'SO511B';
  TNameAry(3) := 'SO511C';
  TNameAry(4) := 'SO511Z';
  for I in 1..4 loop
    select count(*) into v_cnt1 from user_tables where table_name=TNameAry(I);
    if v_cnt1 = 0 then 
      v_RetMsg := '此資料區無營運管理報表資料檔, 檔名: '||TNameAry(I);
      v_RetCode := -4;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
    end if;
  end loop;

  -- ********************************************************
  -- 清除[營運管理報表資料檔]中同一[資料月份]的統計資料, 以便可重複計算此程式
  -- ********************************************************
  delete from SO511A where CompCode=v_CompCode and RptYM=v_RptYM;
  delete from SO511B where CompCode=v_CompCode and RptYM=v_RptYM;
  delete from SO511C where CompCode=v_CompCode and RptYM=v_RptYM;
  -- commit;	-- 不要commit, 以免rollback segment改變

  -- ********************************************************
  -- 設定[A項結果陣列], [D項結果陣列]初值
  -- ********************************************************
  for I in 1..7*5*14 loop
    ACustCount(I) := 0;
    DCustCount(I) := 0;
  end loop;

  -----------------------------------------------------------
  -- 產生5D01營運管理報表資料檔: SO511A
  -----------------------------------------------------------
  -- ********************************************************
  -- 由[統計資料來源檔]中計算A項客戶數
  -- ********************************************************
  begin
    v_SQL := 'select ClassCode1, nvl(Period,0), nvl(CustCount,0), ClctDate from '||v_TableName||
	' where CompCode='||v_CompCode||' and ClctDate>='||GiPackage.QryDTString0(v_StartDate)||
	' and CustStatusCode=1 and ClassCode1>=100 and ClassCode1<=799 and ClassCode1!=701';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor1 for v_SQL;
  exception
    when others then
      v_RetMsg := '由[統計資料來源檔]中計算A項客戶數時發生錯誤, '||SQLERRM;
      v_RetCode := -5;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop每一筆符合條件的客戶資料
    fetch v_DyCursor1 INTO x_ClassCode, x_Period, x_CustCount, x_ClctDate; 
    exit when v_DyCursor1%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- 客戶大類別代碼
    if x_Period<=1 then				-- 繳費期別代碼
      z_PeriodCode := 1;
    elsif x_Period=2 then
      z_PeriodCode := 2;
    elsif x_Period>=3 and x_Period<=5 then
      z_PeriodCode := 3;
    elsif x_Period>=6 and x_Period<=11 then
      z_PeriodCode := 4;
    else
      z_PeriodCode := 5;
    end if;

    if x_CustCount = 0 then			-- 若客戶數為0或null, 則轉成1
      x_CustCount := 1;
    end if;

    -- 計算[下收月份]所歸屬的月份, 可能值為1~14
    z_MonthIndex := to_number(to_char(x_ClctDate,'YYYY'))*12 + to_number(to_char(x_ClctDate,'MM')) - v_StartMonth + 1;
    if z_MonthIndex > 14 then
      z_MonthIndex := 14;
    end if;

    -- 將[客戶數]累計至[A項結果陣列]對應的元素中 ==> 2003.09.05 是筆數, 不是客戶數
    ACustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) := 
	ACustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) + 1;

  end loop;		-- Loop每一筆符合條件的客戶資料
  close v_DyCursor1;

  -- ********************************************************	2003.09.26
  -- 統計[資料截止日]次日凌晨之後裝機的戶數(各種客戶類別的戶數), 
  -- 並反應到[A項結果陣列]對應月份中(扣除)
  -- 條件: SO007中派工類別代碼之參考號是1或5, 且完工時間晚於[資料截止日]的各類別客戶數
  -- ********************************************************
  begin
    v_SQL := 'select B.ClassCode1, count(A.CustId) from SO007 A, SO001 B, CD005 C'||
	' where A.CustId=B.CustId and A.InstCode=C.CodeNo and A.FinTime>'||GiPackage.QryDTString0(v_StopDate)||
	' and (C.RefNo=1 or C.RefNo=5) and B.ClassCode1>=100 and B.ClassCode1<=799 and B.ClassCode1!=701'||
	' group by B.ClassCode1';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor1 for v_SQL;
  exception
    when others then
      v_RetMsg := '統計[資料截止日]次日凌晨之後裝機的戶數時發生錯誤, '||SQLERRM;
      v_RetCode := -7;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop每一筆符合條件的客戶資料
    fetch v_DyCursor1 INTO x_ClassCode, x_CustCount; 
    exit when v_DyCursor1%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- 客戶大類別代碼

    -- 若客戶類別為400~499(大樓統收), 則將戶數由[客戶大類別代碼]為4的第1個月之月繳戶數中扣除
    -- 否則, 將戶數由[客戶大類別代碼]該類的第14個月之年繳戶數中扣除
    if z_ClassCode = 4 then
      ACustCount((4-1)*70+1) := ACustCount((4-1)*70+1) - x_CustCount;
    else
      ACustCount((z_ClassCode-1)*70+(5-1)*14+14) := ACustCount((z_ClassCode-1)*70+(5-1)*14+14) - x_CustCount;
    end if;
  end loop;
  close v_DyCursor1;

  -- ********************************************************	2003.09.26
  -- 統計[資料截止日]次日凌晨之後報拆機的戶數(各種客戶類別的戶數), 
  -- 並反應到[A項結果陣列]對應月份中(加回)
  -- 條件: SO009中派工類別代碼之參考號是2或5, 且受理時間晚於[資料截止日]的各類別客戶數
  -- ********************************************************
  begin
    v_SQL := 'select B.ClassCode1, count(A.CustId) from SO009 A, SO001 B, CD007 C'||
	' where A.CustId=B.CustId and A.PrCode=C.CodeNo and A.AcceptTime>'||GiPackage.QryDTString0(v_StopDate)||
	' and (C.RefNo=2 or C.RefNo=5) and B.ClassCode1>=100 and B.ClassCode1<=799 and B.ClassCode1!=701'||
	' group by B.ClassCode1';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor1 for v_SQL;
  exception
    when others then
      v_RetMsg := '統計[資料截止日]次日凌晨之後報拆機的戶數時發生錯誤, '||SQLERRM;
      v_RetCode := -8;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop每一筆符合條件的客戶資料
    fetch v_DyCursor1 INTO x_ClassCode, x_CustCount; 
    exit when v_DyCursor1%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- 客戶大類別代碼

    -- 若客戶類別為400~499(大樓統收), 則將戶數累計至[客戶大類別代碼]為4的第1個月之月繳戶數
    -- 否則, 將戶數累計至[客戶大類別代碼]該類的第14個月之年繳戶數
    if z_ClassCode = 4 then
      ACustCount((4-1)*70+1) := ACustCount((4-1)*70+1) + x_CustCount;
    else
      ACustCount((z_ClassCode-1)*70+(5-1)*14+14) := ACustCount((z_ClassCode-1)*70+(5-1)*14+14) + x_CustCount;
    end if;
  end loop;
  close v_DyCursor1;

  -- ********************************************************
  -- 由[統計資料來源檔]中計算D項客戶數
  -- ********************************************************
  begin
    v_SQL := 'select ClassCode1, nvl(Period,0), nvl(CustCount,0), ClctDate from '||v_TableName||
	' where CompCode='||v_CompCode||' and (RealStopDate+1)>='||GiPackage.QryDTString0(v_StartDate)||
	' and CustStatusCode=1 and ClassCode1>=100 and ClassCode1<=799 and ClassCode1!=701';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor2 for v_SQL;
  exception
    when others then
      v_RetMsg := '由[統計資料來源檔]中計算D項客戶數時發生錯誤, '||SQLERRM;
      v_RetCode := -6;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop每一筆符合條件的客戶資料
    fetch v_DyCursor2 INTO x_ClassCode, x_Period, x_CustCount, x_ClctDate; 
    exit when v_DyCursor2%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- 客戶大類別代碼
    if x_Period<=1 then				-- 繳費期別代碼
      z_PeriodCode := 1;
    elsif x_Period=2 then
      z_PeriodCode := 2;
    elsif x_Period>=3 and x_Period<=5 then
      z_PeriodCode := 3;
    elsif x_Period>=6 and x_Period<=11 then
      z_PeriodCode := 4;
    else
      z_PeriodCode := 5;
    end if;

    if x_CustCount = 0 then			-- 若客戶數為0或null, 則轉成1
      x_CustCount := 1;
    end if;

    -- 計算[下收月份]所歸屬的月份, 可能值為1~14
    z_MonthIndex := to_number(to_char(x_ClctDate,'YYYY'))*12 + to_number(to_char(x_ClctDate,'MM')) - v_StartMonth + 1;
    if z_MonthIndex > 14 then
      z_MonthIndex := 14;
    end if;

    -- 將[客戶數]累計至[D項結果陣列]對應的元素中
    DCustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) := 
	DCustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) + x_CustCount;
  end loop;		-- Loop每一筆符合條件的客戶資料
  close v_DyCursor2;

  -- ********************************************************
  /* 計算D項與A項的依[客戶大類別代碼]小計之客戶數差異, 並反應到[D項結果陣列]對應月份中, 
     使兩者若依[客戶大類別代碼]小計之總戶數完全一致
	.. Loop 7種[客戶大類別]
		... 從[A項結果陣列]中小計14個月的A項客戶數 ==> A
		... 從[D項結果陣列]中小計14個月的D項客戶數 ==> B
		... 計算上述兩值的差異(A-B) ==> C
		... 若客戶類別為400~499(大樓統收), 則將C值累計至該[客戶大類別代碼]的第1個月之月繳戶數
		... 否則, 將C值累計至該[客戶大類別代碼]的第14個月之年繳戶數
  */
  -- ********************************************************
  v_Delta := 0;
  for I in 1..7 loop
    v_CntA := 0;
    v_CntD := 0;
    for J in 1..5 loop
      for K in 1..14 loop
	v_CntA := v_CntA + ACustCount((I-1)*70+(J-1)*14+K);	-- 該客戶大類別的A項客戶數小計
	v_CntD := v_CntD + DCustCount((I-1)*70+(J-1)*14+K);	-- 該客戶大類別的D項客戶數小計
      end loop;
    end loop;

    v_Delta := v_Delta + v_CntA-v_CntD;
    if I=4 then
      DCustCount((4-1)*70+1) := DCustCount((4-1)*70+1) + (v_CntA-v_CntD);
    else
      DCustCount((I-1)*70+(5-1)*14+14) := DCustCount((I-1)*70+(5-1)*14+14) + (v_CntA-v_CntD);
    end if;
  end loop;

  -- ********************************************************
  -- 將[A項結果陣列]儲存至[營運管理報表資料檔SO511A]	2003.09.26
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511A (CompCode, RptYM, ClassCode, PeriodCode, NextYM, CustCount, DataType)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), ACustCount((I-1)*70+(J-1)*14+K), 0);
      end loop;
    end loop;
  end loop;

  -- ********************************************************
  -- 將[D項結果陣列]儲存至[營運管理報表資料檔SO511A]
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511A (CompCode, RptYM, ClassCode, PeriodCode, NextYM, CustCount, DataType)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), DCustCount((I-1)*70+(J-1)*14+K), 1);
      end loop;
    end loop;
  end loop;

  -----------------------------------------------------------
  -- 產生5D02收費單數預估統計表: SO511B
  ------------------------------------------------------------
  /* 
  做法: 從[D項結果陣列]推算出每個月份累計的收費單數
  . Loop I.每種客戶大類別
	.. Loop J.每種繳費期別
		... Loop K.每個下收月份的客戶數
			* 若當月份客戶數不是0, 則計算[新下收月份]
				[新下收月份] = 該月份加上繳費期別月數
			* 若[新下收月份]未超過14個月, 則將該客戶數累加至[新下收月份]的客戶數
		... end of Loop K
	.. end of Loop J
  . end of Loop I
  */
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	if DCustCount((I-1)*70+(J-1)*14+K) != 0 then
	  if J <=3 then
	    z_MonthIndex := k + j;	-- 1=月繳, 2=雙月繳, 3=季繳
	  elsif J = 4 then
	    z_MonthIndex := k + 6;	-- 4=半年繳
	  else
	    z_MonthIndex := k + 12;	-- 5=年繳
	  end if;
	  if z_MonthIndex <= 14 then
	    DCustCount((I-1)*70+(J-1)*14+z_MonthIndex) := DCustCount((I-1)*70+(J-1)*14+z_MonthIndex) + 
		DCustCount((I-1)*70+(J-1)*14+K);
	  end if;
	end if;
      end loop;
    end loop;
  end loop;

  -- ********************************************************
  -- 將[D項結果陣列]儲存至[收費單數預估統計表SO511B]
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511B (CompCode, RptYM, ClassCode, PeriodCode, NextYM, CustCount)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), DCustCount((I-1)*70+(J-1)*14+K));
      end loop;
    end loop;
  end loop;


  ------------------------------------------------------------
  -- 產生5D03收費金額預估統計表: SO511C
  ------------------------------------------------------------
  /*
  做法: 從5D02所計算之[D項結果陣列], 乘上依[客戶大類別]及[繳費期別]
	  所對應的牌價, 可推算出每個月份累計的收費金額
  . Loop I.每種客戶大類別
	.. Loop J.每種繳費期別
		... 取[客戶大類別]及[繳費期別]所對應的[牌價]
		... 若取不到[牌價], 則log一筆錯誤資料, 並將[牌價]設為0
		... Loop K.每個下收月份
			* 將該月份的累計收費單數乘上[牌價]
			* 再存入對應的[D項結果陣列]月份中
		... end of Loop K
	.. end of Loop J
  . end of Loop I
  */
  for I in 1..7 loop
    for J in 1..5 loop
      -- 取[客戶大類別]及[繳費期別]所對應的[牌價]
      begin
	select UnitPrice into v_UnitPrice from SO511Z where ClassCode=I and PeriodCode=J;
      exception
	when others then
	  v_UnitPrice := 0;
	  v_RetMsg := '牌價檔(SO511Z)中無客戶大類別'||I||'與繳費期別'||J||'所對應到的牌價';
	  v_RetCode := -9;
	  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	    TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	    values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	    trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      end;

      for K in 1..14 loop
	if DCustCount((I-1)*70+(J-1)*14+K) != 0 then
	  -- 將該月份的累計收費單數乘上[牌價], 再存入對應的[D項結果陣列]月份中
	  DCustCount((I-1)*70+(J-1)*14+K) := DCustCount((I-1)*70+(J-1)*14+K) * v_UnitPrice;
	end if;
      end loop;
    end loop;
  end loop;

  -- ********************************************************
  -- 將[D項結果陣列]儲存至[收費金額預估統計表SO511C]
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511C (CompCode, RptYM, ClassCode, PeriodCode, NextYM, Amount)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), DCustCount((I-1)*70+(J-1)*14+K));
      end loop;
    end loop;
  end loop;


  -- ********************************************************
  -- log執行成功記錄至[Night-run結果log檔](SO030A)
  -- ********************************************************
  v_StopExecTime := sysdate;		-- 結束執行時間
  v_RetMsg := '產生完畢, 共花費'||to_char(trunc(86400*(v_StopExecTime-v_StartExecTime)))||'秒';
  v_RetCode := 0;
  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);

  commit;
  DBMS_OUTPUT.PUT_LINE(v_RetMsg);

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code='||SQLCODE||' '|| 'Error Message='||SQLERRM);
    v_RetMsg := '其他錯誤 :'||SQLERRM;
    v_RetCode := -99;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
    return;
END;
/
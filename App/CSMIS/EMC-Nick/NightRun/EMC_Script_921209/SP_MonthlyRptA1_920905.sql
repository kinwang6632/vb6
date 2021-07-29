/*
@C:\EMC_Script\SP_MonthlyRptA1

exec SP_MonthlyRptA1

  批次作業__每月自動產生加值頻道單項產品分析資料 (呼叫SF_RptA1或SF_RptA1_2)

  檔名: SP_MonthlyRptA1.sql 
  流程: 
	. 取執行參數
	. 呼叫A1報表後端程式
	. 刪除同一月份的過去統計資料
	. 將結果彙集成統計用資料, 並複製到A1每月資料檔中(SO508)
	. log執行結果

  A1報表執行參數:
	v_RealDate1 varchar2: 	收費日起始, 'YYYYMMDD'		==> '20030101'
	v_RealDate2 varchar2:	收費日截止, 'YYYYMMDD'		==> 執行當時之上月月底
	v_RefDate varchar2:	參考日期, 'YYYYMMDD'		==> 執行當時之上月月底
	v_ServiceType char(1):	服務類別			==> SO501A.ServiceType
	v_CompCode number:	公司別				==> SO501A.CompCode
	v_CustId number:	客戶編號			==> null
	v_CustIdOp varchar2:	客戶編號比較元			==> '='
	v_ShouldDate1 varchar2:	應收日起始, 'YYYYMMDD'		==> null
	v_ShouldDate2 varchar2:	應收日截止, 'YYYYMMDD'		==> null
	v_CitemSQL varchar2:	收費項目SQL條件字串		==> Refo=2 (含會員費)
	v_ClassSQL varchar2:	客戶類別SQL條件字串		==> null
	v_BillTypeSQL varchar2: 單據類別SQL條件字串		==> null
	v_ServSQL varchar2:	服務SQL條件字串			==> null
	v_AreaSQL varchar2:	行政區SQL條件字串		==> null
	v_StrtSQL varchar2:	街道SQL條件字串			==> null
	v_CircuitNo varchar2:	網路編號條件			==> null
	v_TaxMode number:	稅率計算方式, 1=未稅, 2=已稅	==> 2 
	v_ErrYear number:	計費起始日超過幾年視為錯誤期限	==> 5
	v_MDUSQL varchar2:	大樓編號SQL條件字串		==> null
	v_Citem2SQL varchar2:	非週期收費項目SQL條件字串	==> '=998'

  By: Pierre
  Date: 2003.05.05
	2003.05.16 改: 客戶數定義由統計客戶數改為當月之客戶數
	2003.05.21 改: 加上一筆當月總收入，收費項目代碼為0，收費項目名稱為"總收入"，
			當月實收金額(Amt)，客戶數(Cnt)，因此其平均單價就可以由Excel來計算為Amt/Cnt

	2003.07.23 改: 加上一筆(有效STB數)資料，收費項目代碼為-1，收費項目名稱為"有效STB數"，
			當月實收金額(Amt)為0，客戶數(Cnt)為有效STB數, 如此可算出每個STB的平均單價
			有效STB數之定義: 
			. [參考日]該日前裝機, 且尚未拆機的盒子 (目前此時仍使用中的STB)
			. [參考日]該日前裝機, 且拆機日落在[參考日]後的盒子 (該日仍使用中的STB)

	2003.08.04 改: 加上一筆(有效STB戶數)資料，收費項目代碼為-2，收費項目名稱為"有效STB戶數"，
			當月實收金額(Amt)為0，客戶數(Cnt)為有效STB數, 如此可算出每個STB戶的平均單價
			有效STB戶數之定義: 符合以下條件之STB之戶數(不同客編)
			. [參考日]該日前裝機, 且尚未拆機的盒子 (目前此時仍使用中的STB)
			. [參考日]該日前裝機, 且拆機日落在[參考日]後的盒子 (該日仍使用中的STB)

	2003.09.05 改: 有效戶數與有效台數的定義再加上
			. CATV屬於有效戶(正常戶), 且
			. EMC尚需加上客戶類別<700, 其他MSO則無此項
		Q: 若CATV客戶狀態需屬於有效戶, 則此程式應該於每月1日凌晨執行, 否則會有誤差.
*/

create or replace procedure SP_MonthlyRptA1
as
  v_RealDate1 	varchar2(8) := '20030101';
  v_RealDate2 	varchar2(8);
  v_RefDate 	varchar2(8);
  v_ServiceType char(1);
  v_CompCode 	number;
  v_CustId 	number;
  v_CustIdOp 	varchar2(10);
  v_ShouldDate1 varchar2(8);
  v_ShouldDate2 varchar2(8);
  v_CitemSQL 	varchar2(1000);
  v_ClassSQL 	varchar2(500);
  v_BillTypeSQL varchar2(100);
  v_ServSQL 	varchar2(500);
  v_AreaSQL 	varchar2(500);
  v_StrtSQL 	varchar2(500);
  v_CircuitNo 	varchar2(500);
  v_TaxMode 	number := 2;
  v_ErrYear 	number := 5;
  v_MDUSQL 	varchar2(500);
  v_Citem2SQL 	varchar2(500) := '=998';

  v_Para16	number;
  v_RptYM	number;		-- 報表資料歸屬月份, 西曆YYYYMM
  v_Operator 	varchar2(10);
  v_JobId 	number;
  c39		char(1) := chr(39);
  v_RetCode 	number := 0;
  v_RetMsg 	varchar2(1000) := '無結果';
  v_StartExecTime date;
  v_StopExecTime date;
  v_Second	number;
  v_Selection 	varchar2(3000);
  v_Now 	date;
  v_PrgName	varchar2(100);
  v_FuncName	varchar2(100);
  v_CitemCnt	number := 0;
  v_TotalCustCnt number := 0;
  v_PrDate1	date;
  v_PrDate2	date;
  v_FlowId	number := 0;		-- MSO代號

  cursor cc1 is
    select CodeNo from CD019 where CodeNo>=800 and RefNO=2 order by CodeNo;

  cursor cc2 is
    select CitemCode, count(distinct CustId) CustCnt from SO089 where nvl(Month01,0)!=0 group by CitemCode;

begin
  -- ********************************************************
  -- (1) 取night-run所需參數
  -- ********************************************************
  v_JobId		:= 999990;
  v_StartExecTime 	:= sysdate;		-- 開始執行時間
  v_Operator		:= 'Night-run';
  v_PrgName		:= 'SP_MonthlyRptA1';
  v_FuncName		:= '每月產生加值頻道單項產品分析資料';

  v_RealDate2 := to_char(last_day(add_months(v_StartExecTime,-1)), 'YYYYMMDD');
  -- ***************************************************************************************
  -- 若要手動計算各月資料, 則啟用下行, 修改成各月日期, compile且執行之, 
  -- 但記得執行完後, 要改回原狀(註解該行), 且再compile一次, 例如
  --   v_RealDate2 := '20030131'; -- 可計算92.01月的相關資料
  -- ***************************************************************************************
  v_RefDate := v_RealDate2;
  v_RptYM := to_number(substr(v_RefDate,1,6));

  -- 取v_CompCode, v_ServiceType
  begin
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501A where Rownum<=1;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, 'SO501A: 執行參數未設定ok');
      commit;
      return;
  end;

  -- 取MSO代號: 0=EMC, 1=TBC, ...
  begin
    select FlowId into v_FlowId from SO041 where RowNum<=1;
  exception
    when others then
      v_FlowId := 0;
  end;

  -- 取v_CitemSQL 
  v_CitemCnt := 0;
  v_CitemSQL := '';
  for cr1 in cc1 loop		-- 將收費項目代碼用','方式串起來
    v_CitemSQL := v_CitemSQL || ',' || ltrim(to_char(cr1.CodeNo,'999'));
    v_CitemCnt := v_CitemCnt + 1;
  end loop;
  if v_CitemCnt > 0 then	-- 除去最左邊的','
    v_CitemSQL := substr(v_CitemSQL, 2);
  end if;
  if v_CitemCnt=1 then
    v_CitemSQL := '= ' || v_CitemSQL;
  elsif v_CitemCnt > 1 then
    v_CitemSQL := 'IN ('||v_CitemSQL||')';
  end if;
  DBMS_OUTPUT.PUT_LINE('v_CitemSQL: ' || v_CitemSQL);

  -- 取Para16, 以便決定呼叫, 0=SF_RptA1, 1=SF_RptA1_2
  begin
    select Para16 into v_Para16 from SO043 where CompCode=v_CompCode and ServiceType=v_ServiceType;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '無法取SO043.Para16');
      commit;
      return;
  end;

  -- ********************************************************
  -- (2) 呼叫A1報表後端程式
  -- ********************************************************
  if v_Para16 = 0 then
    v_RetCode := SF_RptA1(v_RealDate1, v_RealDate2, v_RefDate, v_ServiceType, v_CompCode, 
		v_CustId, v_CustIdOp, v_ShouldDate1, v_ShouldDate2, v_CitemSQL, v_ClassSQL, 
		v_BillTypeSQL, v_ServSQL, v_AreaSQL, v_StrtSQL, v_CircuitNo, v_TaxMode,	v_ErrYear, 
		v_MDUSQL, v_Citem2SQL, v_RetMsg);
  else
    v_RetCode := SF_RptA1_2(v_RealDate1, v_RealDate2, v_RefDate, v_ServiceType, v_CompCode, 
		v_CustId, v_CustIdOp, v_ShouldDate1, v_ShouldDate2, v_CitemSQL, v_ClassSQL, 
		v_BillTypeSQL, v_ServSQL, v_AreaSQL, v_StrtSQL, v_CircuitNo, v_TaxMode,	v_ErrYear, 
		v_MDUSQL, v_Citem2SQL, v_RetMsg);
  end if;

  -- ********************************************************
  -- (3) 將結果彙集成統計用資料, 並複製到A1每月資料檔中(SO508)
  -- ********************************************************
  -- 先刪除同一月份的統計資料
  delete SO508 where RptYM=v_RptYM;

  -- 將SO089的暫存資料統計後, 複製到SO508
  begin
    insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, CitemCode, b.Description, 
	sum(nvl(RealAmt,0)), sum(nvl(MonthPast,0)), sum(nvl(Month01,0)), 0 
	from SO089 a, Cd019 b where a.CitemCode=b.CodeNo group by a.CitemCode, b.Description);
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '無法將統計結果複製到A1每月資料檔(SO0508)');
      commit;
      return;
  end;

  -- 再統計當月之客戶數, 並回填SO508中的CustCnt		2003.05.16
  for cr2 in cc2 loop
    update SO508 set CustCnt=cr2.CustCnt where RptYM=v_RptYM and CitemCode=cr2.CitemCode;

    -- 因為目前Cd019中的CodeNo為PK, 故為效率考量, 以下指令效果同上, 故使用上列指令
    --update SO508 set CustCnt=cr2.CustCnt where RptYM=v_RptYM and CompCode=v_CompCode and ServiceType=v_ServiceType
    --  and CitemCode=cr2.CitemCode;
  end loop;

  -- 加上一筆當月總收入，收費項目代碼為0，收費項目名稱為'總收入'
  -- 當月實收金額(Amt)，客戶數(Cnt)，因此其平均單價就可以由Excel來計算為Amt/Cnt 	2003.05.21
  begin
    insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, 0, '總收入', 
	sum(nvl(RealAmt,0)), sum(nvl(MonthPast,0)), sum(nvl(Month01,0)), count(distinct CustId) 
	from SO089 where nvl(Month01,0)!=0 );
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '無法統計當月總收入資料');
      commit;
      return;
  end;

  -- 統計: 當月有效STB數
  -- 加上一筆(有效STB數)資料，收費項目代碼為-1，收費項目名稱為"有效STB數"，
  -- 當月實收金額(Amt)為0，客戶數(Cnt)為有效STB數, 如此可算出每個STB的平均單價		2003.07.23
  -- 有效戶數與有效台數的定義再加強			2003.09.01
  v_PrDate1 := to_date(substr(v_RefDate,1,6)||'01', 'YYYYMMDD');	-- 該月首日
  v_PrDate2 := to_date(v_RefDate,'YYYYMMDD')+1;		-- 參考日次日
  begin
    if v_FlowId=0 then
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -1, '有效STB數', 
	0, 0, 0, count(*) from SO004 A, CD022 B, SO002 C, SO001 D where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.CustId=D.CustId and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1 and D.ClassCode1<700);
    else
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -1, '有效STB數', 
	0, 0, 0, count(*) from SO004 A, CD022 B, SO002 C where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1);
    end if;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '無法統計當月有效STB數資料');
      commit;
      return;
  end;

  -- 統計: 當月有效STB戶數
  -- 加上一筆(有效STB數)資料，收費項目代碼為-2，收費項目名稱為"有效STB戶數"，
  -- 當月實收金額(Amt)為0，客戶數(Cnt)為有效STB數, 如此可算出每個STB戶的平均單價	2003.08.04
  -- 有效戶數與有效台數的定義再加強			2003.09.01
  begin
    if v_FlowId=0 then
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -2, '有效STB戶數', 
	0, 0, 0, count(distinct A.CustId) from SO004 A, CD022 B, SO002 C, SO001 D where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.CustId=D.CustId and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1 and D.ClassCode1<700);
    else
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -2, '有效STB戶數', 
	0, 0, 0, count(distinct A.CustId) from SO004 A, CD022 B, SO002 C where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1);
    end if;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '無法統計當月有效STB戶數資料');
      commit;
      return;
  end;

  -- 將SO090的異常資料複製到SO508A
  begin
    insert into SO508A (BillNo, Item, CustId, CitemCode, CitemName, RealDate, RealStartDate, RealStopDate, 
	RealPeriod, RealAmt, ErrorType) 
	(select BillNo, Item, CustId, CitemCode, CitemName, RealDate, RealStartDate, RealStopDate, 
	RealPeriod, RealAmt, ErrorType from SO090);
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '無法將SO090的異常資料複製到SO508A');
      commit;
      return;
  end;

  -- ********************************************************
  -- (4) Log into operation log file: so030a, so054
  -- ********************************************************
  v_StopExecTime := sysdate;
  v_Second := trunc(86400*(sysdate-v_StartExecTime));

  if v_RetCode < 0 then
    DBMS_OUTPUT.PUT_LINE(v_FuncName||'--執行A1後端程式錯誤, Return code: ' || v_RetCode);
  else
    DBMS_OUTPUT.PUT_LINE(v_FuncName||'--執行A1後端程式完畢, Return code: ' || v_RetCode);
    v_RetMsg := '執行完畢';
    v_Selection := '收費日起始:'||v_RealDate1||';收費日截止:'||v_RealDate2||';參考日期:'||v_RefDate||
		';收費項目SQL條件字串:'||v_CitemSQL||';稅率計算方式:'||v_TaxMode||
		';非週期收費項目SQL條件字串:'||v_Citem2SQL;
    insert into SO054 (PRGNAME, EMPNO, RUNDATETIME, SECOND, SELECTION, STOPDATETIME)
	values (v_PrgName, v_Operator, v_StartExecTime, v_Second, v_Selection, v_StopExecTime);
  end if;

  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, v_StopExecTime, 
	v_Second, v_PrgName, v_FuncName, v_RetCode, v_RetMsg);


  commit;

  <<GO_NEXT1>>
  NULL;


exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;

    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '其他錯誤');
    commit;

end;
/
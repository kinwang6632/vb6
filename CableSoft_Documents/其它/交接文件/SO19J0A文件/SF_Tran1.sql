/*
@C:\gird\v400\script\SF_Tran1
variable nn number
variable msg varchar2(500)
exec :nn := sf_tran1(0, '20030903', '1110010', 'aaaaa', '2129050', 'bbbb', 1, 1, 'Hammer', 'C', 200111, '1101010', 'cccc', :msg);
print nn
print msg

--exec :nn := SF_Tran1( :retmsg);

  日結帳:  SF_Tran1()
  檔名:  SF_Tran1.sql

  IN  p_TranOption number: 1=正式日結, 0=只預視
      p_StopDate varchar2: 過帳截止日日期字串, 西曆YYYYMMDD
      p_A1Code varchar2: 庫存現金會計科目代號
      p_A1Name varchar2: 庫存現金會計科目名稱
      p_A2Code varchar2: 銷項稅額會計科目代號
      p_A2Name varchar2: 銷項稅額會計科目名稱
      p_CompCode varchar2: 公司別代碼
      p_PrtOption number: 列印種類, 1=各項明細, 2=各項加種
      p_UserId varchar2: 執行人員代號
      p_ServiceType varchar2: 服務類別
      p_ActYM number: 帳款歸屬年月
      p_A3Code varchar2: STB應收會計科目代號
      p_A3Name varchar2: STB應收會計科目名稱
      (p_UpdName varchar2: 執行人員姓名 not use now)
  OUT  p_RetMsg varchar2: 結果訊息 (至少200 bytes)
  Return number: 處理結果代碼, 對應訊息存於p_RetMsg中
         >=0: p_RetMsg='執行完畢, 提列筆數=<筆數>'
         -1: p_RetMsg='參數錯誤'
         -2: p_RetMsg='短收原因代碼檔中未設參考號為1的資料'
         -3: p_RetMsg='收費參數檔中未設定完整資料'
         -4: p_RetMsg='收費項目代碼檔中未設參考號為1(基本台費用)的資料'
         -5: p_RetMsg='收費項目代碼檔中有對應不到稅率資料 or 您忘了設定稅率!!!'
         -6: p_RetMsg='找不到SO043參數設定內容'
         -7: p_RetMsg='SO043參數檔為空檔設定內容'
         -8: p_RetMsg := '呆帳的代碼太多'
         ?: p_RetMsg='應收資料檔鎖定失敗, 可能他人正使用中'
        -95: p_RetMsg='TranA2錯誤'
        -96: p_RetMsg='TranA1錯誤'
        -97: p_RetMsg='過帳時發票資料有誤!!! 執行完畢, 問題筆數 = ???'
        -98: p_RetMsg='過帳時部分資料有誤'
        -99: p_RetMsg='其他錯誤'

  By: Pierre
  Date: 2000.07.27

  Edit: Lawrence by 2001.09.20 Initizine v_OldAmt Value = 0 每迴圈中
                    2001.10.08 Add 新增欄位過入
                    2001.12.06 搬移至已日結資料檔SO034中修改Detete from入Begin...exception
                    2001.12.18 回填實收日期至發票檔中
                    2002.01.03 調整回填實收日期至發票檔的BUGS
                    2002.06.17 for New Schema
                    2002.06.18 調整muti-service
                    2002.07.26 修改存so079的bug, 少存CompCode, ServiceType值
                    2002.11.26 加呼叫sf_TranA1, sf_TranA2後端程式
                    2002.12.25 調整sf_TranA1, sf_TranA2後端回傳值判斷機制
                    2002.12.31 增加參數p_A3Code varchar2, p_A3Name varchar2
                    2003.02.21 加填SO034與SO035之新增欄位
                    2003.04.15 加填InvDate
                    2003.07.01 加判斷以分開非週期性有有效期限的狀況
                    2003.08.26 改Return -5的訊息內容
                    2003.09.04 包裝select語法begin.....end
                    2003.09.26 for Lydia.發票拋檔程式修正(920925).doc
                    2003.10.07 原存取SO080改為SO080B已避免程式相衝突
                    2003.10.09 增加規格(for Jacy)CSMISV4.0_DMIST_新版發票拋檔_修正_JACY_921008.doc
                    2003.10.23 改: delete語法改為TRUNCATE TABLE.....以加快本程式效率 and so091.ErrType長度加至100 
                                             調整log訊息SO002A=>帳戶[SO1100F]
                    2003.11.17 調整where條件加CustId=cr1.CustId
                    2004.11.11 讀取SO033.FaciSNo存入SO034 and SO035中(設備資料by機by戶機制)
                    2005.08.08 1731調整insert into so034 and so035語法(ChEvenCode, ChEvenName, AuthenticId, CardBillNo, CardCode, CardName)
                    2005.08.31 依CM/CP專案增加欄位
                    2005.09.03 調整invtitle長度
                    2005.09.16 應PM需求改欄位名稱BUNDLEPRODCODENO,BUNDLEPRODNAME -> BPCODENO,BPNAME
                    2005.09.19 變數v_ErrName加長至VC2(100), 儲存SQLERRM值
                    2005.10.28 add Trace return value to VB
                    2005.11.18 add CP通信費銷帳(調改)資料SO166
                    2005.11.21 add CP月租費資料SO167; add CP通信費退費明細SO170

                    2002.01.07 Stanley 修改無法Update Inv007的BUG 
                    2002.12.25 Stanley 增加處理SO033Debt Table內之資料拋至會計
*/

create or replace function SF_Tran1(p_TranOption number, p_StopDate varchar2, p_A1Code varchar2,
  p_A1Name varchar2, p_A2Code varchar2, p_A2Name varchar2, p_CompCode varchar2,
  p_PrtOption number, p_UserId varchar2, p_ServiceType varchar2, p_ActYM number, p_A3Code varchar2, p_A3Name varchar2, p_RetMsg out varchar2) return number
  AS

  v_Tmp number;
  v_Para3 number;
  v_Para5 number;
  v_Para16 number;
  v_Para17 number;
  v_Str1 varchar2(16);
  v_BCCode number;
  v_NextBCDate date;

  v_CitemCode number := 0;
  v_TaxCode number;
  v_TaxRate number := 0;
  v_Sign char(1);
  v_CurrSign char(1);
  v_Note  varchar2(100);
  v_Note2 varchar2(100);
  v_AccIdP1 varchar2(15);
  v_AccIdM1 varchar2(15);
  v_AccIdP2 varchar2(15);
  v_AccIdM2 varchar2(15);

  v_TaxAmt number := 0;
  v_PureAmt number :=0;
  v_CurrAmt number :=0;
  v_NextAmt number :=0;

  v_InstTime date;
  v_AccNo varchar2(15);
  v_TotalTax number := 0;
  v_TotalAmt number := 0;
  v_SumAmt number := 0;
  v_RcdCount number := 0;
  v_OldAmt number := 0;
  v_Now date;
  v_NextMoney number :=0;
  v_Cnt number :=0;
  v_Cnt1 number :=0;
  v_AmtChk number :=0;
  v_RefNo  number; 
  v_CodeNo varchar2(500);
  v_ErrName varchar2(100);
  v_Invoice number :=0;
  v_SQL varchar2(100);
  v_ErrSQL varchar2(1000);  
  
  -- 2002.11.26 by Lawrence
  v_SeqNo number;  
  v_msg varchar2(500);
  v_Test number;
  v_PeriodFlag number;
  -- 2003.09.26 by Lawrence
  v_Cnt2 number :=0;
  v_Flag number :=0;
  v_InvTitle varchar2(50);
  v_InvoiceType number;
  v_ChargeTitle varchar2(50);
  v_Count number;		-- 2005.11.18 by Lawrence
  v_Count2 number;		-- 2005.11.18 by Lawrence
  v_CPAdjCitemCode varchar2(10);-- 2005.11.18 by Lawrence
  v_STCode number;		-- 2005.11.18 by Lawrence
  v_STName varchar2(20);	-- 2005.11.18 by Lawrence
  v_CPAdjAmount number:=0;	-- 2005.11.18 by Lawrence
  v_CPCitemCode varchar2(10);	-- 2005.11.18 by Lawrence
  v_New char(1);		-- 2005.11.18 by Lawrence
  v_MonthAmt number;		-- 2005.11.18 by Lawrence
  v_StartDate date;		-- 2005.11.21 by Lawrence
  v_stopDate date;		-- 2005.11.21 by Lawrence
  v_Realamt number;		-- 2005.11.21 by Lawrence
  v_shouldamt number;		-- 2005.11.21 by Lawrence
  v_stamt number;		-- 2005.11.21 by Lawrence
  v_TRealamt number;		-- 2005.11.21 by Lawrence
  v_TShouldamt number;		-- 2005.11.21 by Lawrence
  v_TSTamt number;		-- 2005.11.21 by Lawrence

  i number :=0;
  h number;			-- 2005.11.21 by Lawrence
  cursor cc1 is
    select * from SO033 where RealDate<=to_date(p_StopDate,'YYYYMMDD') 
      and UCCode is null 
      and CompCode=p_CompCode and ServiceType=p_ServiceType
      order by CitemCode;

  cursor cc2 is
    select CodeNo from CD016 where RefNo=1 and (ServiceType=p_ServiceType or ServiceType is null);

  cursor cc3 is
    select * from SO033Debt 
     where CompCode=p_CompCode 
       and ServiceType=p_ServiceType
       and ShouldDate<=to_date(p_StopDate,'YYYYMMDD') 
       and ClsTime is null
     order by CitemCode;
begin

  -- 參數檢查
  if p_TranOption is null or (p_TranOption not in (0, 1)) or
    p_StopDate is null or p_A1Code is null or p_A1Name is null or
    p_A2Code is null or p_A2Name is null or p_CompCode is null or
    p_PrtOption is null or p_UserId is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;
  v_Test:=0;
  -- 取次期起始日同本期截止日, 延後幾日為次收日
  if p_ServiceType is not null then
    begin
      select Para3, Para5, Para16, Para17 into v_Para3, v_Para5, v_Para16, v_Para17 from SO043 where CompCode=p_CompCode and ServiceType=p_ServiceType;
    exception
      when others then
        begin
          select Para3, Para5, Para16, Para17 into v_Para3, v_Para5, v_Para16, v_Para17 from SO043 where rownum=1;
        exception
          when others then
            p_RetMsg := 'SO043參數檔為空檔設定內容';
            return -7;
        end;
    end;
  else
    begin
      select Para3, Para5, Para16, Para17 into v_Para3, v_Para5, v_Para16, v_Para17 from SO043 where rownum=1;
    exception
      when others then
        p_RetMsg := '找不到SO043參數設定內容';
        return -6;
    end;
  end if;
  --Stanley 90/08/24 修改  
  -- 取短收原因代碼檔中設為呆帳的代碼
  begin
    for c2 in cc2 loop
      v_CodeNo := v_CodeNo||'('||c2.CodeNo||')';
    end loop;  
  exception
    when others then
        p_RetMsg := '呆帳的代碼太多';
       return -8;
  end;
  -- 檢查是否有發票管理 2001.12.18
  begin
    SELECT count(*) into v_Invoice FROM USER_TABLES WHERE TABLE_NAME='INV007';
  exception
    when others then
      p_RetMsg := 'Inv007 table not found';
      return -99;
  end;
  -- 取基本台訊號費代碼
  begin
    select CodeNo into v_BCCode from CD019 where RefNo=1 and ServiceType=p_ServiceType;
  exception
    when others then
      p_RetMsg := '收費項目代碼檔中未設參考號為1(基本台費用)的資料. ServiceType='||p_ServiceType;
      return -4;
  end;
  -- 檢查稅率是否設定正確
  begin
     select count(*) into v_RcdCount from CD019 where TaxCode not in (select CodeNo from CD033);
  exception
    when others then
      p_RetMsg := '有稅率設定不正確資料[select count(*) into v_RcdCount from CD019 where TaxCode not in (select CodeNo from CD033)]';
      return -99;
  end;
  if v_RcdCount > 0 then
      p_RetMsg := '收費項目代碼檔中有對應不到稅率資料 or 您忘了設定稅率!!!';
      return -5;
  end if;

  begin
-- 2003.10.22 by Lawrence    
--    delete from SO079;
--    delete from SO091;                    -- 清除暫存檔內容
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO079');
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO091');
  exception
    when others then
      p_RetMsg := 'Can not delete so079 or so091';
      return -99;      
  end;
  commit;
  v_Str1 := '截止日:'||ltrim(to_char(to_number(substr(p_StopDate,1,4))-1911, '09'))||substr(p_StopDate,5);
  v_Now := sysdate;
  v_RcdCount := 0;

  -- Loop SO033 符合條件之每一筆, 進行以下動作
  for cr1 in cc1 loop
    v_OldAmt := 0;
    --90/08/24 Stanley 修改 短收原因代碼的參考號為1時表示'呆帳'不做日結動作
    If CR1.STCode IS NOT NULL Then
      if instr(v_CodeNo,'('||cr1.stcode||')')>0 then goto GO_NEXT; end if;
    End If;
    v_Test:=-99;
    -- 2003.07.01 Lawrence Get PeriodFlag Value分開非週期性有有效期限的狀況
    begin
      select PeriodFlag into v_PeriodFlag from CD019 where CodeNo=cr1.CitemCode and ServiceType=p_ServiceType;
    exception
      when others then
        p_RetMsg := 'CD019.PeriodFlag not found! CitemCode :'||cr1.CitemCode;
        return -99;        
    end;
    v_Test:=-98;
    if (cr1.RealStopDate is null or cr1.RealStartDate is null) and nvl(cr1.RealPeriod,0)<>0 and v_PeriodFlag=1 and cr1.CancelFlag=0 then     -- 不合法的有效期限
        insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '不合法的有效期限');
        v_Cnt:=v_Cnt+1;
        v_Test:=-981;
    elsif cr1.RealStopDate is not null and cr1.RealStartDate is not null and cr1.RealPeriod<=0 and v_PeriodFlag=1 and cr1.CancelFlag=0 then
         insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '不合法的期數');
        v_Cnt:=v_Cnt+1;
        v_Test:=-982;
    elsif cr1.RealStopDate < cr1.RealStartDate and cr1.CancelFlag=0 then
         insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '有效截止日不可小於起始日期');
        v_Cnt:=v_Cnt+1;
        v_Test:=-983;
    else  -- 合法資料
       v_Test:=-9831;
       -- 2003.09.26 by Lawrence (for Lydia)加Log
       if cr1.AccountNo is not null and cr1.CancelFlag=0 then
          select count(*) into v_cnt2 from so002a where CustId=cr1.CustId and AccountNo=cr1.AccountNo and CompCode=cr1.CompCode and BankCode=cr1.BankCode;
          v_Test:=-9832;
          if v_cnt2=0 then
             v_Test:=-984;
             insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '在帳戶[SO1100F]中找不到帳號:'||cr1.AccountNo||'銀行代碼:'||to_char(cr1.BankCode));
             v_Cnt:=v_Cnt+1;
             v_Flag:=0;
          else
            v_Test:=-984;
            v_Flag:=1;
            -- 2003.10.09 by Lawrence (for Jacy)加Log
            -- 2003.11.17 by Lawrence where條件加CustId=cr1.CustId
            select InvTitle, InvoiceType, ChargeTitle into v_InvTitle, v_InvoiceType, v_ChargeTitle from so002a where AccountNo=cr1.AccountNo and CompCode=cr1.CompCode and BankCode=cr1.BankCode and CustId=cr1.CustId;
            if v_InvoiceType=3 and v_InvTitle is null then
               v_Test:=-985;
               insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '無發票抬頭內容[SO1100F]帳號:'||cr1.AccountNo||'銀行代碼:'||to_char(cr1.BankCode));
               v_Cnt:=v_Cnt+1;
               v_Flag:=0;
            end if;
            if v_ChargeTitle is null then
               v_Test:=-986;
               insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '無收件人內容[SO1100F]帳號:'||cr1.AccountNo||'銀行代碼:'||to_char(cr1.BankCode));
               v_Cnt:=v_Cnt+1;
               v_Flag:=0;
            end if;
            ---------------------------------------
          end if;
       else
          v_Flag:=1;
          v_Test:=-987;
       end if;
       -------------------------------------
       if cr1.CancelFlag = 1 and v_Flag = 1 then  -- 作廢資料搬移至壞帳資料檔SO035
         if p_TranOption=1 then
           insert into SO035 (AddrNo, AreaCode, BillNo, CancelFlag, CitemCode,
             CitemName, ClassCode, ClctAreaCode, ClctEn, ClctName, ClctYM,
             ClsEn, ClsTime, CMCode, CMName, CompCode, CreateEn,
             CreateTime, CustCount, CustId, EntryEn, GUINo, InvoiceTime,
             Item, ManualNo, MduId, Note, OldAmt, OldClctEn, OldClctName,
             OldPeriod, OldStartDate, OldStopDate, PrintEn,
             PrintTime, PrtClsTime, PrtCount, PrtSNo, PTCode, PTName,
             Quantity, RealAmt, RealDate, RealPeriod, RealStartDate,
             RealStopDate, ServCode, ShouldAmt, ShouldDate, STCode, STName,
             StrtCode, UCCode, UCName, UpdEn, UpdTime, FirstTime,ServiceType, CancelCode, CancelName,
             DiscountCode, DiscountName, DiscountDate1, DiscountDate2, DiscountAmt, CitibankATM, NodeNo, PreInvoice,
             ProcessMark, AccountNo, CombineNo, Amt, AuthorizeNo, OrderNo, SeqNo, OldBillNo, PackageNo, PackageName,
             FirstFlag, BudgetPeriod, ChEven, SBillNo, SItem, SCitemCode, Budget, FaciSeqNo, SNO, BankCode, BankName, 
             MediaBillNo, InvDate, FaciSNo, ChEvenCode, ChEvenName, AuthenticId, CardBillNo, CardCode, CardName,
             PromCode, PromName, BPCode, BPName, ContNo, Punish, MergePrint)
             values (cr1.AddrNo, cr1.AreaCode, cr1.BillNo, cr1.CancelFlag, cr1.CitemCode,
             cr1.CitemName, cr1.ClassCode, cr1.ClctAreaCode, cr1.ClctEn, cr1.ClctName, cr1.ClctYM,
             p_UserId, v_Now, cr1.CMCode, cr1.CMName, cr1.CompCode, cr1.CreateEn,
             cr1.CreateTime, cr1.CustCount, cr1.CustId, cr1.EntryEn, cr1.GUINo, cr1.InvoiceTime,
             cr1.Item, cr1.ManualNo, cr1.MduId, cr1.Note, cr1.OldAmt, cr1.OldClctEn,
             cr1.OldClctName, cr1.OldPeriod, cr1.OldStartDate, cr1.OldStopDate, cr1.PrintEn,
             cr1.PrintTime, cr1.PrtClsTime, cr1.PrtCount, cr1.PrtSNo, cr1.PTCode, cr1.PTName,
             cr1.Quantity, cr1.RealAmt, cr1.RealDate, cr1.RealPeriod, cr1.RealStartDate,
             cr1.RealStopDate, cr1.ServCode, cr1.ShouldAmt, cr1.ShouldDate, cr1.STCode, cr1.STName,
             cr1.StrtCode, cr1.UCCode, cr1.UCName, cr1.UpdEn, cr1.UpdTime, cr1.UpdTime,cr1.ServiceType, cr1.CancelCode, cr1.CancelName,
             cr1.DiscountCode, cr1.DiscountName, cr1.DiscountDate1, cr1.DiscountDate2, cr1.DiscountAmt, cr1.CitibankATM, cr1.NodeNo, cr1.PreInvoice,
             cr1.ProcessMark,  cr1.AccountNo,  cr1.CombineNo,  cr1.Amt,  cr1.AuthorizeNo,  cr1.OrderNo,  cr1.SeqNo,  cr1.OldBillNo,  cr1.PackageNo,  cr1.PackageName,
             cr1.FirstFlag,  cr1.BudgetPeriod,  cr1.ChEven,  cr1.SBillNo,  cr1.SItem,  cr1.SCitemCode,  cr1.Budget,  cr1.FaciSeqNo,  cr1.SNO,  cr1.BankCode,  cr1.BankName,
             cr1.MediaBillNo, cr1.InvDate, cr1.FaciSNo, cr1.ChEvenCode, cr1.ChEvenName, cr1.AuthenticId, cr1.CardBillNo, cr1.CardCode, cr1.CardName,
             cr1.PromCode, cr1.PromName, cr1.BPCode, cr1.BPName, cr1.ContNo, cr1.Punish, cr1.MergePrint);
           delete from SO033 where BillNo=cr1.BillNo and Item=cr1.Item;
         end if;
         v_Test:=-988;
       elsif cr1.CancelFlag = 0 and v_Flag = 1 then   -- 一般資料
         if cr1.CitemCode != v_CitemCode then
           -- 取該收費項目之當期收入科目,當期預收科目,次期收入科目,次期預收科目
           v_CitemCode := cr1.CitemCode;
           v_Test:=v_CitemCode;
           begin
              select AccIdP1, AccIdM1, AccIdP2, AccIdM2, TaxCode, Sign
                into v_AccIdP1, v_AccIdM1, v_AccIdP2, v_AccIdM2, v_TaxCode, v_Sign
                from CD019 where CodeNo=v_CitemCode and ServiceType=p_ServiceType;
           exception
               when others then
                  begin
                    select AccIdP1, AccIdM1, AccIdP2, AccIdM2, TaxCode, Sign
                      into v_AccIdP1, v_AccIdM1, v_AccIdP2, v_AccIdM2, v_TaxCode, v_Sign
                      from CD019 where CodeNo=v_CitemCode and rownum=1;
                  exception
                    when others then
                      p_RetMsg:='CD019 not found! CodeNo : '||v_CitemCode;
                      return -99;                    
                  end;
           end;
           if v_Sign='+' then
              v_Sign:='-';
           else
              v_Sign:='+';
           end if;
           if v_TaxCode is not null then        -- 取目前稅率
              begin
                select nvl(Rate1,0) into v_TaxRate from CD033 where CodeNo=v_TaxCode;
              exception
                when others then
                  v_TaxRate:=0;
                  p_RetMsg:='CD033 not found Tax rate. CodeNo='||v_TaxCode;
                  return -99;
              end;
           else
              v_TaxRate := 0;
           end if;
         end if;
         v_CurrSign := '-';
         v_Note2 := cr1.BillNo || '-' || cr1.CitemName || ',' || GiPackage.ED2CDString(cr1.RealDate);
         v_TaxAmt := 0;
         v_PureAmt := 0;
         v_CurrAmt := 0;
         v_NextAmt := 0;
         if Nvl(cr1.RealPeriod,0) > 0 then        -- 週期性費用
           -- 計算期初攤提
           if v_Para16=0 then --日平均法       
             v_Tmp := SF_NGet1(cr1.RealAmt, cr1.RealDate, cr1.RealStartDate, cr1.RealStopDate,
                             v_TaxRate, v_Para3, v_TaxAmt, v_PureAmt, v_CurrAmt , v_NextAmt);
           else
             v_Tmp := SF_NGet1_2(cr1.RealAmt, cr1.RealDate, cr1.RealStartDate, cr1.RealStopDate,
                             v_TaxRate, v_Para3, v_Para17, v_TaxAmt, v_PureAmt, v_CurrAmt , v_NextAmt);
           end if;
           if v_Sign = '-' then
             v_PureAmt := 0 + v_PureAmt;
             v_CurrAmt := 0 + v_CurrAmt;
             v_NextAmt := 0 + v_NextAmt;
           else
             v_PureAmt := 0 - v_PureAmt;
             v_CurrAmt := 0 - v_CurrAmt;
             v_NextAmt := 0 - v_NextAmt;
           end if;
           -- 取該客戶之安裝時間, 當本筆為基本台費時, 再取基本台次收日
           if v_CitemCode = v_BCCode then
              begin
                select a.InstTime, b.NextBCDate into v_InstTime, v_NextBCDate
                    from SO002 a, SO001 b where a.CustId=cr1.CustId and a.ServiceType=p_ServiceType and b.CustId=cr1.CustId;
              exception
                when NO_DATA_FOUND then
                   v_InstTime := Null;
                   v_NextBCDate := Null;
              end;   -- 更新該客戶之基本台次收日
              if (cr1.RealStopDate+v_Para5 > v_NextBCDate) and p_TranOption=1 then
                 select sum(RealAmt) into v_NextMoney from so033 where BillNo=cr1.BillNo;
                 update SO001 set NextBCDate=cr1.RealStopDate+v_Para5, NextMoney=v_NextMoney where CustId=cr1.CustId;
              elsif v_NextBCDate is null and p_TranOption=1 then
                 select sum(RealAmt) into v_NextMoney from so033 where BillNo=cr1.BillNo;
                 update SO001 set NextBCDate=cr1.RealStopDate+v_Para5, NextMoney=v_NextMoney where CustId=cr1.CustId;
              end if;
           else
             begin
                select InstTime into v_InstTime from SO002 where CustId=cr1.CustId and ServiceType=p_ServiceType;
             exception
              when others then
                 p_RetMsg:='SO002 not Found. CustId='||cr1.CustId||', ServiceType='||p_ServiceType;
                 return -99;
             end;
           end if;
           if v_InstTime>=cr1.RealStartDate and v_InstTime<cr1.RealStopDate+1 then -- 首次費
             v_AccNo := v_AccIdP1;         -- 當期收入科目
             v_Note := v_Str1;
             if p_PrtOption = 1 then
               v_Note := v_Note || ',' || v_Note2;
             end if;
             if v_CurrAmt != 0 then        -- 新增一筆至暫存檔
                if v_OldAmt != cr1.RealAmt then
                      insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
                                         RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
                                 values (v_AccNo, cr1.CustId, abs(v_CurrAmt), cr1.RealDate, cr1.RealAmt,
                                         cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note, p_CompCode, p_ServiceType);
                      v_OldAmt := cr1.RealAmt;
                else
                      insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
                                         RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
                                 values (v_AccNo, cr1.CustId, abs(v_CurrAmt), cr1.RealDate, 0,
                                         cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note, p_CompCode, p_ServiceType);
                end if;
             end if;
             v_AccNo := v_AccIdM1;         -- 當期預收科目
           else                    -- 非首次費
             v_AccNo := v_AccIdP2;         -- 次期收入科目
             v_Note := v_Str1;
             if p_PrtOption = 1 then
               v_Note := v_Note || ',' || v_Note2;
             end if;
             if v_CurrAmt != 0 then        -- 新增一筆至暫存檔
                if v_OldAmt != cr1.RealAmt then
                    insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
                                       RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
                               values (v_AccNo, cr1.CustId, abs(v_CurrAmt), cr1.RealDate, cr1.RealAmt,
                                       cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note, p_CompCode, p_ServiceType);
                    v_OldAmt := cr1.RealAmt;
                else
                    insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
                                       RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
                               values (v_AccNo, cr1.CustId, abs(v_CurrAmt), cr1.RealDate, 0,
                                       cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note, p_CompCode, p_ServiceType);
                end if;
             end if;
             v_AccNo := v_AccIdM2;         -- 次期預收科目
           end if;
         else                              -- 非週期性費用
           if v_TaxRate > 0 then           -- 計算銷售額
             v_PureAmt := round(cr1.RealAmt/(1+v_TaxRate/100));
           else
             v_PureAmt := cr1.RealAmt;
           end if;
           v_TaxAmt := cr1.RealAmt - v_PureAmt;    -- 計算稅額
           v_AccNo := v_AccIdP1;                   -- 當期收入科目
           v_Note := v_Str1;
           if p_PrtOption = 1 then
             v_Note := v_Note || ',' || v_Note2;
           end if;
           v_NextAmt := v_PureAmt;
         end if;
         v_Test:=-989;
       end if;
       if cr1.CancelFlag != 1 and v_Flag=1 then
          if v_NextAmt != 0 and v_AccNo is not null then   -- 新增一筆至暫存檔
                if v_OldAmt != cr1.RealAmt then
                    insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
                           RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
                           values (v_AccNo, cr1.CustId, abs(v_NextAmt), cr1.RealDate, cr1.RealAmt,
                           0, cr1.CitemCode, v_Sign, v_Note, p_CompCode, p_ServiceType);
                    v_OldAmt := cr1.RealAmt;
                else
                    insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
                           RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
                           values (v_AccNo, cr1.CustId, abs(v_NextAmt), cr1.RealDate, 0,
                           0, cr1.CitemCode, v_Sign, v_Note, p_CompCode, p_ServiceType);
                end if;
          end if;
          if v_Sign = '-' then
               v_TotalTax := v_TotalTax + nvl(v_TaxAmt,0);
               v_TotalAmt := v_TotalAmt + nvl(v_PureAmt,0);
          else
               v_TotalTax := v_TotalTax - abs(nvl(v_TaxAmt,0));
               v_TotalAmt := v_TotalAmt - abs(nvl(v_PureAmt,0));
          end if;
          insert into SO080B (Rcdno, Custid, Type) values (abs(nvl(v_TaxAmt,0)), abs(nvl(v_PureAmt,0)), v_AmtChk);
      end if;
      -- 搬移至已日結資料檔SO034
      if p_TranOption=1 and cr1.CancelFlag != 1 and v_Flag=1 then
         begin
           insert into SO034 (AddrNo, AreaCode, BillNo, CancelFlag, CitemCode,
              CitemName, ClassCode, ClctAreaCode, ClctEn, ClctName, ClctYM,
              ClsEn, ClsTime, CMCode, CMName, CompCode, CreateEn,
              CreateTime, CustCount, CustId, EntryEn, GUINo, InvoiceTime,
              Item, ManualNo, MduId, Note, OldAmt, OldClctEn, OldClctName,
              OldPeriod, OldStartDate, OldStopDate, PrintEn,
              PrintTime, PrtClsTime, PrtCount, PrtSNo, PTCode, PTName,
              Quantity, RealAmt, RealDate, RealPeriod, RealStartDate,
              RealStopDate, ServCode, ShouldAmt, ShouldDate, STCode, STName,
              StrtCode, UCCode, UCName, UpdEn, UpdTime, ActYM, FirstTime, ServiceType, CitibankATM, NodeNo, PreInvoice,
              DiscountCode, DiscountName, DiscountDate1, DiscountDate2, DiscountAmt, ProcessMark, AccountNo, CombineNo,
              Amt, AuthorizeNo, OrderNo, SeqNo, OldBillNo, PackageNo, PackageName, FirstFlag, BudgetPeriod, ChEven, SBillNo, SItem, 
              SCitemCode, Budget, FaciSeqNo, SNO, BankCode, BankName, MediaBillNo, InvDate, FaciSNo, ChEvenCode, ChEvenName, AuthenticId, CardBillNo, CardCode, CardName,
              PromCode, PromName, BPCode, BPName, ContNo, Punish, MergePrint)
           values (cr1.AddrNo, cr1.AreaCode, cr1.BillNo, cr1.CancelFlag, cr1.CitemCode,
              cr1.CitemName, cr1.ClassCode, cr1.ClctAreaCode, cr1.ClctEn, cr1.ClctName, cr1.ClctYM,
              p_UserId, v_Now, cr1.CMCode, cr1.CMName, cr1.CompCode, cr1.CreateEn,
              cr1.CreateTime, cr1.CustCount, cr1.CustId, cr1.EntryEn, cr1.GUINo, cr1.InvoiceTime,
              cr1.Item, cr1.ManualNo, cr1.MduId, cr1.Note, cr1.OldAmt, cr1.OldClctEn,
              cr1.OldClctName, cr1.OldPeriod, cr1.OldStartDate, cr1.OldStopDate, cr1.PrintEn,
              cr1.PrintTime, cr1.PrtClsTime, cr1.PrtCount, cr1.PrtSNo, cr1.PTCode, cr1.PTName,
              cr1.Quantity, cr1.RealAmt, cr1.RealDate, cr1.RealPeriod, cr1.RealStartDate,
              cr1.RealStopDate, cr1.ServCode, cr1.ShouldAmt, cr1.ShouldDate, cr1.STCode, cr1.STName,
              cr1.StrtCode, cr1.UCCode, cr1.UCName, cr1.UpdEn, cr1.UpdTime, p_ActYM, cr1.UpdTime,cr1.ServiceType, cr1.CitibankATM, cr1.NodeNo, cr1.PreInvoice,
              cr1.DiscountCode, cr1.DiscountName, cr1.DiscountDate1, cr1.DiscountDate2, cr1.DiscountAmt, cr1.ProcessMark, cr1.AccountNo, cr1.CombineNo,
              cr1.Amt, cr1.AuthorizeNo, cr1.OrderNo, cr1.SeqNo, cr1.OldBillNo, cr1.PackageNo, cr1.PackageName, cr1.FirstFlag, cr1.BudgetPeriod, cr1.ChEven, cr1.SBillNo, cr1.SItem, 
              cr1.SCitemCode, cr1.Budget, cr1.FaciSeqNo, cr1.SNO, cr1.BankCode, cr1.BankName, cr1.MediaBillNo, cr1.InvDate, cr1.FaciSNo, cr1.ChEvenCode, cr1.ChEvenName, cr1.AuthenticId,
              cr1.CardBillNo, cr1.CardCode, cr1.CardName, cr1.PromCode, cr1.PromName, cr1.BPCode, cr1.BPName, cr1.ContNo, cr1.Punish, cr1.MergePrint);
           -- 2001.12.18 回填實收日期至發票檔中
           -- 2002.1.7  stanely 修改
           begin
              if v_Invoice>0 and cr1.GUINO is not null and cr1.RealDate is not null then
                 v_SQL := 'update INV007 set CHARGEDATE=To_Date('||chr(39)||cr1.RealDate||chr(39)||') where INVID='||chr(39)||cr1.GUINO||chr(39)||' and CHARGEDATE is null';
                 EXECUTE IMMEDIATE v_SQL;
              end if;
           exception
             when others then
               v_ErrName:=substrb(SQLERRM,1,100);
               insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                      RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                      values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                     cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, v_ErrName);
               v_Cnt1:=v_Cnt1+1;
           end;
           -- 2005.11.18 by Lawrence add CP通信費銷帳(調改)資料(CNS_CP需求文件_與台固之資料交換介面_Lawrence_941110.doc)
           if cr1.ServiceType='P' then
              v_CPAdjCitemCode:='';
              if cr1.RealAmt>=0 then
                 select count(*) into v_Count from so165 where CitemCode1=cr1.CitemCode and CPProperty=1;
                 select CPAdjCitemCode into v_CPAdjCitemCode from so165 where CitemCode1=cr1.CitemCode and CPProperty=1;
              else
                 select count(*) into v_Count from so165 where CitemCode2=cr1.CitemCode and CPProperty=1;
                 select CPAdjCitemCode into v_CPAdjCitemCode from so165 where CitemCode2=cr1.CitemCode and CPProperty=1;
              end if;
              if cr1.STCode is not null then
                 select count(*) into v_Count2 from cd016 where CodeNo=cr1.STCode and RefNo=2;
              else
                 v_Count2:=1;
              end if;
              if v_Count=1 and v_count2=1 then
                 v_STCode:=cr1.STCode;
                 v_STName:=cr1.STName;
                 insert into SO166 
                    (CustId, BillNo, Item, CitemCode, CitemName, ShouldDate, ShouldAmt, RealDate, 
                    RealAmt, STCode, STName, CompCode, ServiceType, FaciSNO, TFNBillNo, TFNServiceID, CPShouldYM, 
                    CPAdjCitemCode, CPAdjCode, CPAdjName, CPRate, CPExport)
                 values
                    (cr1.custid, cr1.BillNo, cr1.Item, cr1.CitemCode, cr1.CitemName, cr1.ShouldDate, cr1.ShouldAmt, cr1.RealDate,
                    cr1.RealAmt, cr1.STCode, cr1.STName, cr1.CompCode, cr1.ServiceType, cr1.FaciSNo, cr1.TFNBillNo, cr1.TFNServiceID, cr1.CPShouldYM,
                    v_CPAdjCitemCode, v_STCode, v_STName, null, 0);
              end if;
           end if;
           -- 2005.11.18 by Lawrence add CP月租費資料(CNS_CP需求文件_與台固之資料交換介面_Lawrence_941110.doc)
           if cr1.ServiceType='P' then
              v_CPAdjCitemCode:='';
              v_CPCitemCode:='';
              if cr1.RealAmt>=0 then
                 select count(*) into v_Count from so165 where CitemCode1=cr1.CitemCode and CPProperty=2;
                 select CPCitemCode, CPAdjCitemCode into v_CPCitemCode, v_CPAdjCitemCode from so165 where CitemCode1=cr1.CitemCode and CPProperty=2;
              else
                 select count(*) into v_Count from so165 where CitemCode2=cr1.CitemCode and CPProperty=2;
                 select CPCitemCode, CPAdjCitemCode into v_CPCitemCode, v_CPAdjCitemCode from so165 where CitemCode2=cr1.CitemCode and CPProperty=2;
              end if;
              if cr1.STCode is not null then
                 select count(*) into v_Count2 from cd016 where CodeNo=cr1.STCode and RefNo=2;
              else
                 v_Count2:=1;
              end if;
              if v_Count=1 and v_count2=1 then
                 v_STCode:=cr1.STCode;
                 v_STName:=cr1.STName;
                 v_CPAdjAmount:=cr1.ShouldAmt-cr1.RealAmt;
                 if substr(cr1.BillNo,7,1)='I' then
                  v_New:='N';
                 else
                  v_New:='R';
                 end if;
                 select MonthAmt into v_MonthAmt from cd078a where BPCode=cr1.BPCode and CitemCode=cr1.CitemCode and ServiceType='P';
                 -- 2005.11.21 by Lawrence
                 if cr1.RealPeriod>1 then
                    v_startDate:=cr1.RealStartDate;
                    if add_Months(cr1.RealStartDate,1)-1>cr1.RealStopDate then
                       v_stopDate:=cr1.RealStopDate;
                    else
                       v_stopDate:=add_Months(cr1.RealStartDate,1)-1;
                    end if;
                    v_shouldamt:=ROUND(cr1.shouldamt/cr1.RealPeriod,0);
                    v_realamt:=ROUND(cr1.RealAmt/cr1.RealPeriod,0);
                    v_stamt:=ROUND(v_CPAdjAmount/cr1.RealPeriod,0);
                    v_TShouldAmt:=0;
                    v_TRealAmt:=0;
                    v_TSTAmt:=0;
                    -- 帳款日結資料,以每筆資料的有效期間依月份切割後轉入SO167
                    for h in 1..cr1.RealPeriod loop
                      if h=cr1.RealPeriod then
                         v_ShouldAmt:=cr1.ShouldAmt-v_TShouldAmt;
                         v_RealAmt:=cr1.RealAmt-v_TRealAmt;
                         v_STAmt:=v_CPAdjAmount-v_TSTAmt;
                      end if;
                      insert into SO167 
                        (CustId, BillNo, Item, CitemCode, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod, 
                        STCode, STName, CompCode, ClctYM, ServiceType, FaciSNO, BPCode, BPName, TFNServiceID, CPAdjCode, 
                        CPAdjName, CPAdjAmount, CPCitemCode, CPContractMode, CPStartDate, CPStopDate, MonthAmt, BelongYM, 
                        CPExport, CPReturnRent)
                      values
                        (cr1.custid, cr1.BillNo, cr1.Item, cr1.CitemCode, cr1.CitemName, cr1.ShouldDate, v_ShouldAmt, cr1.RealDate, v_RealAmt, cr1.RealPeriod,
                        cr1.STCode, cr1.STName, cr1.CompCode, cr1.ClctYM, cr1.ServiceType, cr1.FaciSNo, cr1.BPCode, cr1.BPName, cr1.TFNServiceID, v_STCode,
                        v_STName, v_stamt, v_CPCitemCode, v_New, to_char(v_StartDate,'YYYYMMDD'), to_char(v_StopDate,'YYYYMMDD'), v_MonthAmt, to_char(v_startDate,'YYYYMM'),
                        0, 0);
                      v_startDate:=v_stopDate+1;
                      if add_Months(v_startDate,1)-1>cr1.RealStopDate then
                         v_stopDate:=cr1.RealStopDate;
                      else
                         v_stopDate:=add_Months(v_startDate,1)-1;
                      end if; 
                      v_TShouldAmt:=v_TShouldAmt+v_shouldamt;
                      v_TRealAmt:=v_TRealAmt+v_realamt;
                      v_TSTAmt:=v_TSTAmt+v_STAmt;
                    end loop;
                 else
                    -- 2005.11.18 by Lawrence
                    insert into SO167 
                      (CustId, BillNo, Item, CitemCode, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod, 
                      STCode, STName, CompCode, ClctYM, ServiceType, FaciSNO, BPCode, BPName, TFNServiceID, CPAdjCode, 
                      CPAdjName, CPAdjAmount, CPCitemCode, CPContractMode, CPStartDate, CPStopDate, MonthAmt, BelongYM, 
                      CPExport, CPReturnRent)
                    values
                      (cr1.custid, cr1.BillNo, cr1.Item, cr1.CitemCode, cr1.CitemName, cr1.ShouldDate, cr1.ShouldAmt, cr1.RealDate, cr1.RealAmt, cr1.RealPeriod,
                      cr1.STCode, cr1.STName, cr1.CompCode, cr1.ClctYM, cr1.ServiceType, cr1.FaciSNo, cr1.BPCode, cr1.BPName, cr1.TFNServiceID, v_STCode,
                      v_STName, v_CPAdjAmount, v_CPCitemCode, v_New, to_char(cr1.RealStartDate,'YYYYMMDD'), to_char(cr1.RealStopDate,'YYYYMMDD'), v_MonthAmt, to_char(cr1.ShouldDate,'YYYYMM'),
                      0, 0);
                 end if; 
              end if;
           end if;
           -- 2005.11.21 by Lawrence add CP通信費退費明細(CNS_CP需求文件_與台固之資料交換介面_Lawrence_941110.doc)
           if cr1.ServiceType='P' then
              v_CPAdjCitemCode:='';
              if cr1.RealAmt>=0 then
                 select count(*) into v_Count from so165 where CitemCode1=cr1.CitemCode and CPProperty=4;
                 select CPAdjCitemCode into v_CPAdjCitemCode from so165 where CitemCode1=cr1.CitemCode and CPProperty=4;
              else
                 select count(*) into v_Count from so165 where CitemCode2=cr1.CitemCode and CPProperty=4;
                 select CPAdjCitemCode into v_CPAdjCitemCode from so165 where CitemCode2=cr1.CitemCode and CPProperty=4;
              end if;
              if v_Count=1 then
                 v_STCode:=cr1.STCode;
                 v_STName:=cr1.STName;
                 insert into SO170 
                    (CustId, BillNo, Item, CitemCode, CitemName, ShouldDate, ShouldAmt, RealDate, 
                    RealAmt, STCode, STName, CompCode, ServiceType, FaciSNO, TFNBillNo, TFNServiceID, CPShouldYM, 
                    CPAdjCitemCode, CPAdjCode, CPAdjName, CPRate, CPExport)
                 values
                    (cr1.custid, cr1.BillNo, cr1.Item, cr1.CitemCode, cr1.CitemName, cr1.ShouldDate, cr1.ShouldAmt, cr1.RealDate,
                    cr1.RealAmt, cr1.STCode, cr1.STName, cr1.CompCode, cr1.ServiceType, cr1.FaciSNo, cr1.TFNBillNo, cr1.TFNServiceID, cr1.CPShouldYM,
                    v_CPAdjCitemCode, v_STCode, v_STName, null, 0);
              end if;
           end if;
           delete from SO033 where BillNo=cr1.BillNo and Item=cr1.Item;
         exception
           when others then
             v_ErrName:=substrb(SQLERRM,1,100);
             insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, v_ErrName);
             v_Cnt:=v_Cnt+1;
         end;
      end if;
      if cr1.CancelFlag != 1 and v_Flag=1 then
         v_RcdCount := v_RcdCount + 1;
      end if;
    end if;
    v_Test:=-97;
   --90/04/04 Stanley 修改
   <<GO_NEXT>>
   NULL;
  end loop;
  v_CitemCode:=0;
  --處理SO033Debt機制
  for cr3 in cc3 loop
     v_Sign:='';
     v_Test:=-96;
     if cr3.CitemCode != v_CitemCode then
        -- 取該收費項目之當期收入科目,借貸別
        v_CitemCode := cr3.CitemCode;
        begin
          select AccIdP1, Sign
            into v_AccIdP1, v_Sign
            from CD019 where CodeNo=cr3.CitemCode and ServiceType=p_ServiceType;
        exception
          when others then
            begin
              select AccIdP1, Sign
                into v_AccIdP1, v_Sign
                from CD019 where CodeNo=cr3.CitemCode and rownum=1;
            exception
              when others then
                p_RetMsg:='CD019 not found! CodeNo : '||cr3.CitemCode;
                return -99;
            end;
        end;

        if v_Sign='+' then
           v_Sign:='-';
        else
           v_Sign:='+';
        end if;
     end if;
     v_Test:=-95;
     v_Note := v_Str1 || ',' || v_Note2;
     insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, CitemCode, Sign, 
                        Note, CompCode, ServiceType, OrderNo, SNO, Type)
            values (v_AccIdP1, cr3.CustId, abs(cr3.ShouldAmt), cr3.ShouldDate, cr3.CitemCode, v_Sign, 
                        v_Note, p_CompCode, p_ServiceType, cr3.OrderNo, cr3.SNO, 1);
     v_Test:=-94;
     insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, CitemCode, Sign, 
                        Note, CompCode, ServiceType, OrderNo, SNO, Type)
            values (p_A3Code, cr3.CustId, abs(cr3.ShouldAmt), cr3.ShouldDate, cr3.CitemCode, '+', 
                        v_Note, p_CompCode, p_ServiceType, cr3.OrderNo, cr3.SNO, 2);
     v_Test:=-93;
     if p_TranOption=1 then
        update SO033Debt set ClsTime=Trunc(sysdate) where BillNo=cr3.BillNo and Item=cr3.Item;
     end if;
     v_Test:=-92;
     if v_Sign = '-' then
        v_TotalAmt := v_TotalAmt + nvl(cr3.ShouldAmt,0);
     else
        v_TotalAmt := v_TotalAmt - abs(nvl(cr3.ShouldAmt,0));
     end if;
     v_Test:=-91;
     v_RcdCount:=v_RcdCount+1;
  end loop;
  v_SumAmt:=(v_TotalAmt+v_TotalTax);
  v_Test:=-90;
  if v_SumAmt != 0 then         -- 新增一筆至暫存檔
    if v_SumAmt < 0 then
       insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
           RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
           values (p_A1Code, 0, ABS(v_SumAmt), to_date(p_StopDate,'YYYYMMDD'), ABS(v_SumAmt),
           0, null, '-', v_Str1, p_CompCode, p_ServiceType);
    else
       insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
           RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
           values (p_A1Code, 0, v_SumAmt, to_date(p_StopDate,'YYYYMMDD'), v_SumAmt,
           0, null, '+', v_Str1, p_CompCode, p_ServiceType);
    end if;
  end if;
  v_Test:=-89;
  if v_TotalTax != 0 then               -- 新增一筆至暫存檔
    if v_TotalTax < 0 then
       insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
           RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
           values (p_A2Code, 0,ABS(v_TotalTax), to_date(p_StopDate,'YYYYMMDD'), 0,
           0, null, '+', v_Str1, p_CompCode, p_ServiceType);
    else
       insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt,
           RealPeriod, CitemCode, Sign, Note, CompCode, ServiceType)
           values (p_A2Code, 0,v_TotalTax, to_date(p_StopDate,'YYYYMMDD'), 0,
           0, null, '-', v_Str1, p_CompCode, p_ServiceType);
    end if;
  end if;
  v_Test:=-88;
/*
  -- 2002.11.26 by Lawrence 
  -- 2002.12.25 by Lawrence update
  -- 2003.04.28 by Lawrence Mark (已不用)
  v_Tmp:=SF_TRANA1(p_StopDate, p_CompCode, p_ServiceType, p_UserId, 1, v_SeqNo, v_msg);
  if v_SeqNo>0 and v_Tmp>=0 then
     v_Tmp:=SF_TRANA2(v_SeqNo, p_CompCode, v_msg);
     if v_Tmp>=0 then
        if v_Cnt>0 then
           p_RetMsg := '過帳時部分資料有誤!!!'||chr(13)||chr(10)||'執行完畢, 提列筆數 = ' ||to_char(v_RcdCount,'999999')||' 稅前金額 = '||v_TotalAmt||' 稅額 = '||v_TotalTax||' 稅後金額 = '||v_SumAmt;
           return -98;
        else
          if v_Cnt1>0 then
              p_RetMsg := '過帳時發票資料有誤!!!'||chr(13)||chr(10)||'執行完畢, 問題筆數 = ' ||to_char(v_Cnt1,'999999');
              return -97;
          else
              p_RetMsg := '執行完畢, 提列筆數 = '||to_char(v_RcdCount,'999999')||' 稅前金額 = '||v_TotalAmt||' 稅額 = '||v_TotalTax||' 稅後金額 = '||v_SumAmt;
              return v_RcdCount;
          end if;
        end if;
     else
        p_RetMsg := 'SF_TranA2 : Code='||to_char(v_Tmp)||'.'||v_Msg;
        return -95;
     end if;
  else
     p_RetMsg := 'SF_TranA1 : Code='||to_char(v_Tmp)||'.'||v_Msg;
     return -96;
  end if;
*/
  -- 2002.12.25 by Lawrence Mark
  -- 2003.04.28 by Lawrence 解除Mark
  if v_Cnt>0 then
     p_RetMsg := '過帳時部分資料有誤!!!'||chr(13)||chr(10)||'執行完畢, 提列筆數 = ' ||to_char(v_RcdCount,'999999')||' 稅前金額 = '||v_TotalAmt||' 稅額 = '||v_TotalTax||' 稅後金額 = '||v_SumAmt;
     return -98;
  else
    if v_Cnt1>0 then
        p_RetMsg := '過帳時發票資料有誤!!!'||chr(13)||chr(10)||'執行完畢, 問題筆數 = ' ||to_char(v_Cnt1,'999999');
        return -97;
    else
        p_RetMsg := '執行完畢, 提列筆數 = '||to_char(v_RcdCount,'999999')||' 稅前金額 = '||v_TotalAmt||' 稅額 = '||v_TotalTax||' 稅後金額 = '||v_SumAmt;
        return v_RcdCount;
    end if;
  end if;

exception
  when others then
    rollback;
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := SQLERRM||'/'||to_char(v_Test);
    return -99;
end;
/
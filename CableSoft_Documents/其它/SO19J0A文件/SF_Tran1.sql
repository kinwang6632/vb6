/*
@C:\gird\v400\script\SF_Tran1
variable nn number
variable msg varchar2(500)
exec :nn := sf_tran1(0, '20030903', '1110010', 'aaaaa', '2129050', 'bbbb', 1, 1, 'Hammer', 'C', 200111, '1101010', 'cccc', :msg);
print nn
print msg

--exec :nn := SF_Tran1( :retmsg);

  �鵲�b:  SF_Tran1()
  �ɦW:  SF_Tran1.sql

  IN  p_TranOption number: 1=�����鵲, 0=�u�w��
      p_StopDate varchar2: �L�b�I������r��, ���YYYYMMDD
      p_A1Code varchar2: �w�s�{���|�p��إN��
      p_A1Name varchar2: �w�s�{���|�p��ئW��
      p_A2Code varchar2: �P���|�B�|�p��إN��
      p_A2Name varchar2: �P���|�B�|�p��ئW��
      p_CompCode varchar2: ���q�O�N�X
      p_PrtOption number: �C�L����, 1=�U������, 2=�U���[��
      p_UserId varchar2: ����H���N��
      p_ServiceType varchar2: �A�����O
      p_ActYM number: �b���k�ݦ~��
      p_A3Code varchar2: STB�����|�p��إN��
      p_A3Name varchar2: STB�����|�p��ئW��
      (p_UpdName varchar2: ����H���m�W not use now)
  OUT  p_RetMsg varchar2: ���G�T�� (�ܤ�200 bytes)
  Return number: �B�z���G�N�X, �����T���s��p_RetMsg��
         >=0: p_RetMsg='���槹��, ���C����=<����>'
         -1: p_RetMsg='�Ѽƿ��~'
         -2: p_RetMsg='�u����]�N�X�ɤ����]�ѦҸ���1�����'
         -3: p_RetMsg='���O�Ѽ��ɤ����]�w������'
         -4: p_RetMsg='���O���إN�X�ɤ����]�ѦҸ���1(�򥻥x�O��)�����'
         -5: p_RetMsg='���O���إN�X�ɤ�����������|�v��� or �z�ѤF�]�w�|�v!!!'
         -6: p_RetMsg='�䤣��SO043�ѼƳ]�w���e'
         -7: p_RetMsg='SO043�Ѽ��ɬ����ɳ]�w���e'
         -8: p_RetMsg := '�b�b���N�X�Ӧh'
         ?: p_RetMsg='�����������w����, �i��L�H���ϥΤ�'
        -95: p_RetMsg='TranA2���~'
        -96: p_RetMsg='TranA1���~'
        -97: p_RetMsg='�L�b�ɵo����Ʀ��~!!! ���槹��, ���D���� = ???'
        -98: p_RetMsg='�L�b�ɳ�����Ʀ��~'
        -99: p_RetMsg='��L���~'

  By: Pierre
  Date: 2000.07.27

  Edit: Lawrence by 2001.09.20 Initizine v_OldAmt Value = 0 �C�j�餤
                    2001.10.08 Add �s�W���L�J
                    2001.12.06 �h���ܤw�鵲�����SO034���ק�Detete from�JBegin...exception
                    2001.12.18 �^��ꦬ����ܵo���ɤ�
                    2002.01.03 �վ�^��ꦬ����ܵo���ɪ�BUGS
                    2002.06.17 for New Schema
                    2002.06.18 �վ�muti-service
                    2002.07.26 �ק�sso079��bug, �֦sCompCode, ServiceType��
                    2002.11.26 �[�I�ssf_TranA1, sf_TranA2��ݵ{��
                    2002.12.25 �վ�sf_TranA1, sf_TranA2��ݦ^�ǭȧP�_����
                    2002.12.31 �W�[�Ѽ�p_A3Code varchar2, p_A3Name varchar2
                    2003.02.21 �[��SO034�PSO035���s�W���
                    2003.04.15 �[��InvDate
                    2003.07.01 �[�P�_�H���}�D�g���ʦ����Ĵ��������p
                    2003.08.26 ��Return -5���T�����e
                    2003.09.04 �]��select�y�kbegin.....end
                    2003.09.26 for Lydia.�o�����ɵ{���ץ�(920925).doc
                    2003.10.07 ��s��SO080�אּSO080B�w�קK�{���۽Ĭ�
                    2003.10.09 �W�[�W��(for Jacy)CSMISV4.0_DMIST_�s���o������_�ץ�_JACY_921008.doc
                    2003.10.23 ��: delete�y�k�אּTRUNCATE TABLE.....�H�[�֥��{���Ĳv and so091.ErrType���ץ[��100 
                                             �վ�log�T��SO002A=>�b��[SO1100F]
                    2003.11.17 �վ�where����[CustId=cr1.CustId
                    2004.11.11 Ū��SO033.FaciSNo�s�JSO034 and SO035��(�]�Ƹ��by��by�����)
                    2005.08.08 1731�վ�insert into so034 and so035�y�k(ChEvenCode, ChEvenName, AuthenticId, CardBillNo, CardCode, CardName)
                    2005.08.31 ��CM/CP�M�׼W�[���
                    2005.09.03 �վ�invtitle����
                    2005.09.16 ��PM�ݨD�����W��BUNDLEPRODCODENO,BUNDLEPRODNAME -> BPCODENO,BPNAME
                    2005.09.19 �ܼ�v_ErrName�[����VC2(100), �x�sSQLERRM��
                    2005.10.28 add Trace return value to VB
                    2005.11.18 add CP�q�H�O�P�b(�է�)���SO166
                    2005.11.21 add CP�믲�O���SO167; add CP�q�H�O�h�O����SO170

                    2002.01.07 Stanley �ק�L�kUpdate Inv007��BUG 
                    2002.12.25 Stanley �W�[�B�zSO033Debt Table������Ʃߦܷ|�p
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

  -- �Ѽ��ˬd
  if p_TranOption is null or (p_TranOption not in (0, 1)) or
    p_StopDate is null or p_A1Code is null or p_A1Name is null or
    p_A2Code is null or p_A2Name is null or p_CompCode is null or
    p_PrtOption is null or p_UserId is null then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;
  v_Test:=0;
  -- �������_�l��P�����I���, ����X�鬰������
  if p_ServiceType is not null then
    begin
      select Para3, Para5, Para16, Para17 into v_Para3, v_Para5, v_Para16, v_Para17 from SO043 where CompCode=p_CompCode and ServiceType=p_ServiceType;
    exception
      when others then
        begin
          select Para3, Para5, Para16, Para17 into v_Para3, v_Para5, v_Para16, v_Para17 from SO043 where rownum=1;
        exception
          when others then
            p_RetMsg := 'SO043�Ѽ��ɬ����ɳ]�w���e';
            return -7;
        end;
    end;
  else
    begin
      select Para3, Para5, Para16, Para17 into v_Para3, v_Para5, v_Para16, v_Para17 from SO043 where rownum=1;
    exception
      when others then
        p_RetMsg := '�䤣��SO043�ѼƳ]�w���e';
        return -6;
    end;
  end if;
  --Stanley 90/08/24 �ק�  
  -- ���u����]�N�X�ɤ��]���b�b���N�X
  begin
    for c2 in cc2 loop
      v_CodeNo := v_CodeNo||'('||c2.CodeNo||')';
    end loop;  
  exception
    when others then
        p_RetMsg := '�b�b���N�X�Ӧh';
       return -8;
  end;
  -- �ˬd�O�_���o���޲z 2001.12.18
  begin
    SELECT count(*) into v_Invoice FROM USER_TABLES WHERE TABLE_NAME='INV007';
  exception
    when others then
      p_RetMsg := 'Inv007 table not found';
      return -99;
  end;
  -- ���򥻥x�T���O�N�X
  begin
    select CodeNo into v_BCCode from CD019 where RefNo=1 and ServiceType=p_ServiceType;
  exception
    when others then
      p_RetMsg := '���O���إN�X�ɤ����]�ѦҸ���1(�򥻥x�O��)�����. ServiceType='||p_ServiceType;
      return -4;
  end;
  -- �ˬd�|�v�O�_�]�w���T
  begin
     select count(*) into v_RcdCount from CD019 where TaxCode not in (select CodeNo from CD033);
  exception
    when others then
      p_RetMsg := '���|�v�]�w�����T���[select count(*) into v_RcdCount from CD019 where TaxCode not in (select CodeNo from CD033)]';
      return -99;
  end;
  if v_RcdCount > 0 then
      p_RetMsg := '���O���إN�X�ɤ�����������|�v��� or �z�ѤF�]�w�|�v!!!';
      return -5;
  end if;

  begin
-- 2003.10.22 by Lawrence    
--    delete from SO079;
--    delete from SO091;                    -- �M���Ȧs�ɤ��e
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO079');
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO091');
  exception
    when others then
      p_RetMsg := 'Can not delete so079 or so091';
      return -99;      
  end;
  commit;
  v_Str1 := '�I���:'||ltrim(to_char(to_number(substr(p_StopDate,1,4))-1911, '09'))||substr(p_StopDate,5);
  v_Now := sysdate;
  v_RcdCount := 0;

  -- Loop SO033 �ŦX���󤧨C�@��, �i��H�U�ʧ@
  for cr1 in cc1 loop
    v_OldAmt := 0;
    --90/08/24 Stanley �ק� �u����]�N�X���ѦҸ���1�ɪ��'�b�b'�����鵲�ʧ@
    If CR1.STCode IS NOT NULL Then
      if instr(v_CodeNo,'('||cr1.stcode||')')>0 then goto GO_NEXT; end if;
    End If;
    v_Test:=-99;
    -- 2003.07.01 Lawrence Get PeriodFlag Value���}�D�g���ʦ����Ĵ��������p
    begin
      select PeriodFlag into v_PeriodFlag from CD019 where CodeNo=cr1.CitemCode and ServiceType=p_ServiceType;
    exception
      when others then
        p_RetMsg := 'CD019.PeriodFlag not found! CitemCode :'||cr1.CitemCode;
        return -99;        
    end;
    v_Test:=-98;
    if (cr1.RealStopDate is null or cr1.RealStartDate is null) and nvl(cr1.RealPeriod,0)<>0 and v_PeriodFlag=1 and cr1.CancelFlag=0 then     -- ���X�k�����Ĵ���
        insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '���X�k�����Ĵ���');
        v_Cnt:=v_Cnt+1;
        v_Test:=-981;
    elsif cr1.RealStopDate is not null and cr1.RealStartDate is not null and cr1.RealPeriod<=0 and v_PeriodFlag=1 and cr1.CancelFlag=0 then
         insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '���X�k������');
        v_Cnt:=v_Cnt+1;
        v_Test:=-982;
    elsif cr1.RealStopDate < cr1.RealStartDate and cr1.CancelFlag=0 then
         insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '���ĺI��餣�i�p��_�l���');
        v_Cnt:=v_Cnt+1;
        v_Test:=-983;
    else  -- �X�k���
       v_Test:=-9831;
       -- 2003.09.26 by Lawrence (for Lydia)�[Log
       if cr1.AccountNo is not null and cr1.CancelFlag=0 then
          select count(*) into v_cnt2 from so002a where CustId=cr1.CustId and AccountNo=cr1.AccountNo and CompCode=cr1.CompCode and BankCode=cr1.BankCode;
          v_Test:=-9832;
          if v_cnt2=0 then
             v_Test:=-984;
             insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '�b�b��[SO1100F]���䤣��b��:'||cr1.AccountNo||'�Ȧ�N�X:'||to_char(cr1.BankCode));
             v_Cnt:=v_Cnt+1;
             v_Flag:=0;
          else
            v_Test:=-984;
            v_Flag:=1;
            -- 2003.10.09 by Lawrence (for Jacy)�[Log
            -- 2003.11.17 by Lawrence where����[CustId=cr1.CustId
            select InvTitle, InvoiceType, ChargeTitle into v_InvTitle, v_InvoiceType, v_ChargeTitle from so002a where AccountNo=cr1.AccountNo and CompCode=cr1.CompCode and BankCode=cr1.BankCode and CustId=cr1.CustId;
            if v_InvoiceType=3 and v_InvTitle is null then
               v_Test:=-985;
               insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '�L�o�����Y���e[SO1100F]�b��:'||cr1.AccountNo||'�Ȧ�N�X:'||to_char(cr1.BankCode));
               v_Cnt:=v_Cnt+1;
               v_Flag:=0;
            end if;
            if v_ChargeTitle is null then
               v_Test:=-986;
               insert into SO091 (BillNo, Item, CustId, CitemCode, CitemName, RealDate,
                   RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrType)
                   values (cr1.BillNo, cr1.Item, cr1.CustId, cr1.CitemCode, cr1.CitemName, cr1.RealDate,
                   cr1.RealStartDate, cr1.RealStopDate, cr1.RealPeriod, cr1.RealAmt, '�L����H���e[SO1100F]�b��:'||cr1.AccountNo||'�Ȧ�N�X:'||to_char(cr1.BankCode));
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
       if cr1.CancelFlag = 1 and v_Flag = 1 then  -- �@�o��Ʒh�����a�b�����SO035
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
       elsif cr1.CancelFlag = 0 and v_Flag = 1 then   -- �@����
         if cr1.CitemCode != v_CitemCode then
           -- ���Ӧ��O���ؤ�������J���,����w�����,�������J���,�����w�����
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
           if v_TaxCode is not null then        -- ���ثe�|�v
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
         if Nvl(cr1.RealPeriod,0) > 0 then        -- �g���ʶO��
           -- �p������u��
           if v_Para16=0 then --�饭���k       
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
           -- ���ӫȤᤧ�w�ˮɶ�, �������򥻥x�O��, �A���򥻥x������
           if v_CitemCode = v_BCCode then
              begin
                select a.InstTime, b.NextBCDate into v_InstTime, v_NextBCDate
                    from SO002 a, SO001 b where a.CustId=cr1.CustId and a.ServiceType=p_ServiceType and b.CustId=cr1.CustId;
              exception
                when NO_DATA_FOUND then
                   v_InstTime := Null;
                   v_NextBCDate := Null;
              end;   -- ��s�ӫȤᤧ�򥻥x������
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
           if v_InstTime>=cr1.RealStartDate and v_InstTime<cr1.RealStopDate+1 then -- �����O
             v_AccNo := v_AccIdP1;         -- ������J���
             v_Note := v_Str1;
             if p_PrtOption = 1 then
               v_Note := v_Note || ',' || v_Note2;
             end if;
             if v_CurrAmt != 0 then        -- �s�W�@���ܼȦs��
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
             v_AccNo := v_AccIdM1;         -- ����w�����
           else                    -- �D�����O
             v_AccNo := v_AccIdP2;         -- �������J���
             v_Note := v_Str1;
             if p_PrtOption = 1 then
               v_Note := v_Note || ',' || v_Note2;
             end if;
             if v_CurrAmt != 0 then        -- �s�W�@���ܼȦs��
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
             v_AccNo := v_AccIdM2;         -- �����w�����
           end if;
         else                              -- �D�g���ʶO��
           if v_TaxRate > 0 then           -- �p��P���B
             v_PureAmt := round(cr1.RealAmt/(1+v_TaxRate/100));
           else
             v_PureAmt := cr1.RealAmt;
           end if;
           v_TaxAmt := cr1.RealAmt - v_PureAmt;    -- �p��|�B
           v_AccNo := v_AccIdP1;                   -- ������J���
           v_Note := v_Str1;
           if p_PrtOption = 1 then
             v_Note := v_Note || ',' || v_Note2;
           end if;
           v_NextAmt := v_PureAmt;
         end if;
         v_Test:=-989;
       end if;
       if cr1.CancelFlag != 1 and v_Flag=1 then
          if v_NextAmt != 0 and v_AccNo is not null then   -- �s�W�@���ܼȦs��
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
      -- �h���ܤw�鵲�����SO034
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
           -- 2001.12.18 �^��ꦬ����ܵo���ɤ�
           -- 2002.1.7  stanely �ק�
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
           -- 2005.11.18 by Lawrence add CP�q�H�O�P�b(�է�)���(CNS_CP�ݨD���_�P�x�T����ƥ洫����_Lawrence_941110.doc)
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
           -- 2005.11.18 by Lawrence add CP�믲�O���(CNS_CP�ݨD���_�P�x�T����ƥ洫����_Lawrence_941110.doc)
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
                    -- �b�ڤ鵲���,�H�C����ƪ����Ĵ����̤�����Ϋ���JSO167
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
           -- 2005.11.21 by Lawrence add CP�q�H�O�h�O����(CNS_CP�ݨD���_�P�x�T����ƥ洫����_Lawrence_941110.doc)
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
   --90/04/04 Stanley �ק�
   <<GO_NEXT>>
   NULL;
  end loop;
  v_CitemCode:=0;
  --�B�zSO033Debt����
  for cr3 in cc3 loop
     v_Sign:='';
     v_Test:=-96;
     if cr3.CitemCode != v_CitemCode then
        -- ���Ӧ��O���ؤ�������J���,�ɶU�O
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
  if v_SumAmt != 0 then         -- �s�W�@���ܼȦs��
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
  if v_TotalTax != 0 then               -- �s�W�@���ܼȦs��
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
  -- 2003.04.28 by Lawrence Mark (�w����)
  v_Tmp:=SF_TRANA1(p_StopDate, p_CompCode, p_ServiceType, p_UserId, 1, v_SeqNo, v_msg);
  if v_SeqNo>0 and v_Tmp>=0 then
     v_Tmp:=SF_TRANA2(v_SeqNo, p_CompCode, v_msg);
     if v_Tmp>=0 then
        if v_Cnt>0 then
           p_RetMsg := '�L�b�ɳ�����Ʀ��~!!!'||chr(13)||chr(10)||'���槹��, ���C���� = ' ||to_char(v_RcdCount,'999999')||' �|�e���B = '||v_TotalAmt||' �|�B = '||v_TotalTax||' �|����B = '||v_SumAmt;
           return -98;
        else
          if v_Cnt1>0 then
              p_RetMsg := '�L�b�ɵo����Ʀ��~!!!'||chr(13)||chr(10)||'���槹��, ���D���� = ' ||to_char(v_Cnt1,'999999');
              return -97;
          else
              p_RetMsg := '���槹��, ���C���� = '||to_char(v_RcdCount,'999999')||' �|�e���B = '||v_TotalAmt||' �|�B = '||v_TotalTax||' �|����B = '||v_SumAmt;
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
  -- 2003.04.28 by Lawrence �Ѱ�Mark
  if v_Cnt>0 then
     p_RetMsg := '�L�b�ɳ�����Ʀ��~!!!'||chr(13)||chr(10)||'���槹��, ���C���� = ' ||to_char(v_RcdCount,'999999')||' �|�e���B = '||v_TotalAmt||' �|�B = '||v_TotalTax||' �|����B = '||v_SumAmt;
     return -98;
  else
    if v_Cnt1>0 then
        p_RetMsg := '�L�b�ɵo����Ʀ��~!!!'||chr(13)||chr(10)||'���槹��, ���D���� = ' ||to_char(v_Cnt1,'999999');
        return -97;
    else
        p_RetMsg := '���槹��, ���C���� = '||to_char(v_RcdCount,'999999')||' �|�e���B = '||v_TotalAmt||' �|�B = '||v_TotalTax||' �|����B = '||v_SumAmt;
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
/*
@C:\Gird\v400\Script\SF_TranInv3
@C:\Gird\v400\Script\待測程式區2\SF_TranInv2

variable nn number
variable msg varchar2(30000)
--exec :nn := SF_TranInv3( 7, 'IN (''C'')', null, null, null, null, null, null, '20080501', '20080501', null, null, null, null, null, 1, 1, 1, 1, null, null, 0, 3, :Msg);
--exec :nn := SF_TranInv2( 3, 'IN (''D'')', NULL, NULL, null, null, null, null, '20100228', '20100228', null, null, null, null, null, 1, 1, 1, 1, null, null, 0, 3, :Msg);
print NN
PRINT MSG

  預開發票拋檔作業(針對開博新發票系統):  SF_TranInv3()
  檔名:  SF_TranInv3.sql
  IN  
      p_CompCode number: 公司別代碼
      p_ServiceTypeSQL varchar2: 服務類別
      p_CustId1 number: 起始客戶編號
      p_CustId2 number: 截止客戶編號
      p_ServCodeSQL varchar2: 服務區SQL條件字串
      p_ClassSQL varchar2: 客戶類別SQL條件字串 
      p_BillTypeSQL varchar2: 單據種類SQL條件字串 ????
      p_CitemSQL varchar2: 收費項目SQL條件字串  
      p_ShouldDate1 varchar2: 起始應收日期, 'YYYYMMDD'
      p_ShouldDate2 varchar2: 截止應收日期, 'YYYYMMDD'
      p_ClctEnSQL varchar2: 收費人員代號
      p_UpdDate1 varchar2: 起始異動日期, 'YYYYMMDD'
      p_UpdDate2 varchar2: 截止異動日期, 'YYYYMMDD'
      p_UpdEn varchar2: 異動人員姓名 (不是代碼呦)
      p_CMCodeSQL varchar2: 收費方式代碼
      p_InvType number: 發票種類(1=所有, 2=二聯, 3=三聯)
      p_AddrType number: 寄件地址種類(1=裝機, 2=收費, 3=郵寄)
      p_InvoiceType number: 開發票型態(1=預開, 2=全部)
      p_Period number: 是否含分期資料總額(0=否, 1=是) 
      p_RealDate1 varchar2: 起始實收日期, 'YYYYMMDD'
      p_RealDate2 varchar2: 截止實收日期, 'YYYYMMDD'    
      p_AreaCity number: 加縣市行政區字串(0=否, 1=是) 
      p_Mode number: 1=由SO4131, 2=由SO4132, 3=由SO4141, 4=由SO4142
  OUT  
      p_RetMsg varchar2: 結果訊息 (至少1000 bytes)
  Return number: 處理結果代碼, 對應訊息存於p_RetMsg中
      >=0: p_RetMsg='產生完畢, 共花費<秒數>秒, 處理筆數=<筆數>' 
         -1: p_RetMsg='參數錯誤'
         -2: p_RetMsg='SQL錯誤: <SQL指令>'
         -3: p_RetMsg='找不到收費項目代碼'
         -4: p_RetMsg='服務類別沒設定'
         -5: p_RetMsg='SQL3錯誤: <SQL3指令>'
       -99: p_RetMsg='其他錯誤'

  By: Lawrence
  Date: 2003.08.25 重寫for Jacy規格(多帳戶處理+無帳戶處理)
        2003.09.08 包裝select SQL.begin...end
        2003.09.17 調整get so002a語法
        2003.09.22 加v_ContName and v_ChkCustId判斷式
        2003.09.26 for Jacy規格##增加 9.收件人 chargetitle,--增加 16.服務類別 servicetype(CSMISV4.0_DMIST_新版發票拋檔_修正_JACY_920926.doc)
        2003.10.09 for Jacy規格(CSMISV4.0_DMIST_新版發票拋檔_修正_JACY_921008.doc)
        2003.11.18 Mark delete用法,改為Truncate
        2004.05.24 調整變數z_VouString,z_VouString1為varchar2(305); z_Note為varchar2(750)
        2004.12.20 for 需求編號：1342(應收日期相同, 且在拋檔時只要收費資料的帳號相同, 則需拋在同一筆資料下)
        2005.09.03 調整CitemName, invtitle長度
        2006.06.06 Adj Address相關欄位加長 for 2408
        2006.11.08 for 2841
        2006.12.28 for 2930 有關地址加串縣市行政區字串
        2007.01.03 PM又說地址欄位要加大至142, 已請Casper update問題集文件說明
        2007.03.08 調整處理無帳戶的部分,地址加串縣市行政區字串
        2007.12.06 for 3398
        2007.12.06 for 3436 原讀SO002A改讀SO138
        2008.03.11 for 3785
        2008.03.25 改Bug.無帳號少傳z_InvSeqNo參數
        2008.04.17 add AccountNo is null and InvSeqNo>0之讀取SO138
        2008.04.24 調整一下程式無帳號處理的z_strInvNo, z_ContName變數名稱
        2008.08.07 for Minchen 調整nvl(z_ContName,'xya')改為nvl(z_ContName,'xyz'); 查無當時為何要預設'xya'
        2008.08.25 3883 如invseqno = 0 or invseqno is null 則依據原有發票拋檔的順序來處理
        2008.08.29 3986 不管有無帳戶皆增加依InvSeqNo值處理
        2009.12.22 5439
        2010.02.10 for 5542 v_SQL相關變數加大至10000
        2010.03.23 for 5534經討論結果(與Jacy)先做第4點(SF_TranInv1,SF_TranInv4不再使用)
        2010.05.11 FOR 5438 Jacky說CombCitemCode and CombStartDate有值, 則改讀取CombAmount, CombStartDate, CombStopDate欄位
        2010.05.13 5656 調整先判斷SO138有值則取SO138.PreInvoice，如無值則仍取SO002.PreInvoice {利用判斷式分開[nvl(A.InvSeqNo,0)=0 及 nvl(A.InvSeqNo,0)>0]並做Union後處理}

  以下為Pierre的修改註解:
	2010.07.15 for 5668, 拋檔中的字串增加"發票開立種類、常用E-MAIL、行動電話、受贈單位代碼、受贈單位名稱" 5個欄位
*/

CREATE OR REPLACE function SF_TranInv3(p_CompCode number, p_ServiceTypeSQL varchar2,
      p_CustId1 number, p_CustId2 number, p_ServCodeSQL varchar2, p_ClassSQL varchar2, 
      p_BillTypeSQL varchar2, p_CitemSQL varchar2, 
      p_ShouldDate1 varchar2, p_ShouldDate2 varchar2, 
      p_ClctEnSQL varchar2, p_UpdDate1 varchar2, p_UpdDate2 varchar2, 
      p_UpdEn varchar2, p_CMCodeSQL varchar2, p_InvType number, p_AddrType number, 
      p_InvoiceType number, p_Period number, p_RealDate1 varchar2, p_RealDate2 varchar2, p_AreaCity number, p_Mode number, p_RetMsg out varchar2) return number
  AS
  c39 		char(1):=chr(39);
  v_StartExecTime 	date;
  v_Now 	  	date;
  v_ChkShouldDate 	date;
  v_Char  		char(1);
  v_Char1       	char(1);
  z_ServiceType	char(1);		-- 2010.07.15 by Pierre, v_ServiceType改名為z_ServiceType
  v_ChkBillNo   	varchar2(15);
  v_AccountNo 	varchar2(16);
  v_ContName	varchar2(50);
  v_SQL  		varchar2(30000);
  v_SQL1 		varchar2(30000);
  v_SQL2 		varchar2(30000);
  v_SQL3   	varchar2(30000);
  v_SQL4   	varchar2(30000);
  v_SQL5   	varchar2(30000);	-- 2010.05.13 by Lawrence
  v_SQL6   	varchar2(30000);	-- 2010.05.13 by Lawrence
  v_ChkMediaBillNo 	varchar2(11);
  v_ChkAccountNo   	varchar2(16);
  v_TaxCode 	number;
  v_ChkCustId   	number;
  v_TaxFlag     	number;
  v_InvCompCode 	number;
  v_CustId		number;
  v_ShouldAmt   	number:=0;
  v_RcdCount 	number:=0;
  v_Sno     		number:=1;
  v_CurrCitemCode 	number;
  v_chkInvSeqNo    	number:=0;  	--2008.08.15 by Corey
  z_ShouldDate	date;
  z_UpdA		date;
  z_RealStartDate 	date;
  z_RealStopDate  	date;
  z_TaxType	char(1);
  z_CustId 		number;
  z_ShouldAmt	number;
  z_CitemCode	number;
  z_AddrNo	number;
  z_InvoiceType 	number;
  z_TaxRate     	number;
  z_TaxMoney    	number;
  z_Qty         	number;
  z_TaxCode     	number;
  z_Item        	number;
  z_UpdB		number;
  z_ZipCode	varchar2(5);
  z_strInvNo    	varchar2(8);
  z_InvNo		varchar2(8);
  z_BillNo		varchar2(15);
  z_CitemName	varchar2(40);
  z_UpdTime	varchar2(19);
  z_Address	varchar2(142);
  z_AccountNo	varchar2(16);
  z_MediaBillNo	varchar2(11);
  z_ContName	varchar2(50);
  z_InvTitle	varchar2(50);
  z_InvAddress	varchar2(142);
  z_Tel1        	varchar2(20);
  z_ClctName    	varchar2(20);
  z_ChargeTitle	varchar2(50);
  z_Note		varchar2(750);
  z_VouString   	varchar2(500);	-- 2008.03.11 by Lawrence 改至380, 2010.07.15 by Pierre 改至500
  z_VouString1  	varchar2(500);	-- 2008.03.11 by Lawrence 改至380,, 2010.07.15 by Pierre 改至500
  z_AreaName 	varchar2(20);	-- 2006.12.28
  z_CityName 	varchar2(20);	-- 2006.12.28
  z_InvSeqNo 	number;		-- 2007.12.06 by Lawrence
  z_InvPurposeCode 	number;		-- 2008.03.11 by Lawrence
  z_InvPurposeName 	varchar2(20);	-- 2008.03.11 by Lawrence
  z_CombCitemCode	number;		-- 2009.12.22 by Lawrence
  z_CombCitemName	varchar2(40);	-- 2009.12.22 by Lawrence
  z_CombAmount	number;		-- 2010.05.11 by Lawrence
  z_CombStartDate	date;		-- 2010.05.11 by Lawrence
  z_CombStopDate	date;		-- 2010.05.11 by Lawrence
  x_ContName 	varchar2(50);

  TYPE CurTyp IS REF CURSOR;  
  v_DyCursor CurTyp;          

  -- 2010.07.15 by Pierre
  z_FaciSeqNo	varchar2(15);		-- SO033/SO034的設備流水號
  z_InvoiceKind	number;			-- 發票開立種類
  z_Email	varchar2(50);		-- 常用E-MAIL
  z_ContMobile	varchar2(20);		-- 行動電話
  z_DenRecCode	number;			-- 受贈單位代碼
  z_DenRecName	varchar2(20);		-- 受贈單位名稱
  x_RcdCnt1	number:=0;		-- 有帳戶的筆數
  x_RcdCnt2	number:=0;		-- 無帳戶的筆數
BEGIN
  v_Char := chr(01);   -- 欄位間格符號
  -- 參數檢查
  if p_InvType is null or (p_InvType not in (1,2,3)) or p_AddrType is null or (p_AddrType not in (1,2,3)) or p_Period is null or (p_Period not in (0,1)) or (p_Mode not in (3,4)) then
     p_RetMsg := '參數錯誤';
     return -1;
  end if;
  -- 多服務時, InvCompCode需設相同否則會拋錯帳
  v_SQL:='select InvCompCode from CD046 where CodeNo '||p_ServiceTypeSQL||' and InvCompCode is not null and RowNum=1';
  begin
     EXECUTE IMMEDIATE v_SQL INTO v_InvCompCode;     
  exception
    when others then
       p_RetMsg:='服務類別or發票系統公司別沒設定';
       return -4;         
  end;
  -- 刪除各暫存檔
  -- 2007.12.06 by Lawrence Mark Truncate改為delete...where...
  delete from Tmp004 where MFlag=p_Mode;
  delete from Tmp005 where MFlag=p_Mode;
  -- 設定各變數初值
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_SQL1 := null;
  -- 若有客戶編號條件
  if nvl(p_CustId1,0)>0 then
     v_SQL1 := v_SQL1 || ' and A.CustId>=' || p_CustId1 || ' and A.CustId<=' || p_CustId2;
  end if;  
  -- 若有收費項目條件
  if p_CitemSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.CitemCode ' || p_CitemSQL;
  end if;
  if p_ServCodeSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.ServCode ' || p_ServCodeSQL;
  end if;
  -- 若有客戶類別條件
  if p_ClassSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.ClassCode ' || p_ClassSQL;
  end if;
  -- 若有公司別條件
  if p_CompCode is not null then
     v_SQL1 := v_SQL1 || ' and A.CompCode=' || p_CompCode;
  end if;
  -- 若有應收日期條件
  if p_ShouldDate1 is not null then
     v_SQL1 := v_SQL1 || ' and A.ShouldDate>='||'to_date('||chr(39)||p_ShouldDate1||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')'||
                       ' and A.ShouldDate<=to_date('||chr(39)||p_ShouldDate2||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')';
  end if;
  -- 若有收費人員條件
  if p_ClctEnSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.ClctEn ' || p_ClctEnSQL;
  end if;
  -- 若有異動日期條件
  if p_UpdDate1 is not null then
     v_SQL1 := v_SQL1 || ' And A.UpdTime>='||c39||ltrim(to_char(to_number(substr(p_UpdDate1,1,4))-1911,'99'))||
                       '/'||substr(p_UpdDate1,5,2)||'/'||substr(p_UpdDate1,7,2)||c39||
                       ' and A.UpdTime<='||c39||ltrim(to_char(to_number(substr(p_UpdDate2,1,4))-1911,'99'))||
                       '/'||substr(p_UpdDate2,5,2)||'/'||substr(p_UpdDate2,7,2)||' 23:59:59'||c39;
  end if;
  -- 若有異動人員條件
  if p_UpdEn is not null then
     v_SQL1 := ' and A.UpdEn=' || c39 || p_UpdEn || c39;
  end if;
  -- 若有收費方式條件
  -- 2006.11.08 Lawrence p_CMCode改成p_CMCodeSQL複選
  if p_CMCodeSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.CMCode ' || p_CMCodeSQL;
  end if;
  v_SQL1 := v_SQL1 || ' and a.InvoiceTime is null and a.CancelFlag=0 and a.GUINo is null';
  -- 若有單據種類條件
  if p_BillTypeSQL is not null then
     v_Char1 := substr(p_BillTypeSQL, 1, 1);
     if v_Char1 = '=' or v_Char1 = '!' then
        if p_BillTypeSQL like '%B%' or p_BillTypeSQL like '%T%'  then
           v_SQL1 := v_SQL1 || ' and substr(A.BillNo,7,1) '||p_BillTypeSQL;
        else
           v_SQL1 := v_SQL1 || ' and substr(A.BillNo,1,1) '||p_BillTypeSQL;
        end if;
     else
        if p_BillTypeSQL like '%B%' and p_BillTypeSQL like '%T%'  then
           v_SQL1 := v_SQL1 || ' and (substr(A.BillNo,7,1) '||p_BillTypeSQL||')';
        else
           v_SQL1 := v_SQL1 || ' and (substr(A.BillNo,1,1) '||p_BillTypeSQL||' or substr(A.BillNo,7,1) '||p_BillTypeSQL||')';
        end if;
     end if;
  end if;
  v_SQL2 := 'select A.CustId,A.BillNo,A.ShouldDate,A.ShouldAmt,A.CitemCode,A.CitemName,'||
                    'A.UpdTime,A.RealStartDate,A.RealStopDate,A.Note,A.Quantity,A.ClctName,A.Item,C.InvNo,C.InvoiceType,';
  -- 取地址資料
  if p_AddrType=1 then
     v_SQL2 := v_SQL2 || 'B.InstAddress Address, B.InstAddrNo AddrNo';
  elsif p_AddrType=2 then
     v_SQL2 := v_SQL2 || 'B.ChargeAddress Address, B.ChargeAddrNo AddrNo';
  else
     v_SQL2 := v_SQL2 || 'B.MailAddress Address, B.MailAddrNo AddrNo';
  end if;
  -- 2010.03.23 by Lawrence #5534若有實收日期條件(由p_Period=1的區域移至此處)
  v_SQL4:='';
  if p_RealDate1 is not null then
     v_SQL4 := ' and A.RealDate>='||'to_date('||chr(39)||p_RealDate1||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')'||
                       ' and A.RealDate<=to_date('||chr(39)||p_RealDate2||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')';
  end if;
  --加判斷是否預開制(PreInvQice=1)
  -- 2010.07.15 by Pierre, 加取FaciSeqNo
  if p_InvoiceType=1 then
     --處理無帳戶的部分(依客戶編號,發票統一編號,發票抬頭,聯絡人(新加欄位) GROUP BY 資料,將相同的產生在一起)
     -- 2010.05.13 by Lawrence Add 條件and nvl(A.InvSeqNo,0)=0
     v_SQL3 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and C.PreInvoice=1 and A.Budget<>1 and nvl(A.InvSeqNo,0)=0' ||  v_SQL4 || 
     v_SQL1 || ' and A.AccountNo is null';
     --處理有帳戶的部分(依AccountNo,MediaBillNo GROUP BY資料,將相同的產生在一起)
     v_SQL := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and C.PreInvoice=1 and A.Budget<>1 and nvl(A.InvSeqNo,0)=0' ||  v_SQL4 || 
                     v_SQL1 || ' and A.AccountNo is not null';
     -- 2010.05.13 by Lawrence Add v_SQL5 , v_SQL6
     v_SQL6 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and A.Budget<>1 and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo and D.PreInvoice=1' ||  v_SQL4 || 
     v_SQL1 || ' and A.AccountNo is null';
     --處理有帳戶的部分(依AccountNo,MediaBillNo GROUP BY資料,將相同的產生在一起)
     v_SQL5 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and A.Budget<>1 and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo and D.PreInvoice=1' ||  v_SQL4 || 
                     v_SQL1 || ' and A.AccountNo is not null';
  else
     --處理無帳戶的部分(依客戶編號,發票統一編號,發票抬頭,聯絡人(新加欄位) GROUP BY 資料,將相同的產生在一起)
     -- 2010.05.13 by Lawrence Add 條件and nvl(A.InvSeqNo,0)=0
     v_SQL3 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and A.Budget<>1 and nvl(A.InvSeqNo,0)=0' ||  v_SQL4 || 
                       v_SQL1 || ' and A.AccountNo is null';
     --處理有帳戶的部分(依AccountNo,MediaBillNo GROUP BY資料,將相同的產生在一起)
     v_SQL := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and A.Budget<>1 and nvl(A.InvSeqNo,0)=0' ||  v_SQL4 || 
                     v_SQL1 || ' and A.AccountNo is not null';
     -- 2010.05.13 by Lawrence Add v_SQL5 , v_SQL6
     v_SQL6 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and A.Budget<>1 and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo' ||  v_SQL4 || 
                       v_SQL1 || ' and A.AccountNo is null';
     --處理有帳戶的部分(依AccountNo,MediaBillNo GROUP BY資料,將相同的產生在一起)
     v_SQL5 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033 A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and A.Budget<>1 and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo' ||  v_SQL4 || 
                     v_SQL1 || ' and A.AccountNo is not null';
  end if;
  --判斷是否該取另一Table SO033Debt的資料
  if p_Period<>1 then
     -- 2010.05.13 by Lawrence Add v_SQL5 , v_SQL6
     v_SQL3 := v_SQL3||' Union '||v_SQL6||' order by CustId,InvNo,InvTitle,ContName';
     -- 2004.12.20 by Lawrence order by A.ShouldDate,A.MediaBillNo對調
     v_SQL := v_SQL||' Union '||v_SQL5||' order by AccountNo,ShouldDate,MediaBillNo';
  else
     -- 2010.07.15 by Pierre, 加取FaciSeqNo
     if p_InvoiceType=1 then
        -- 2010.05.13 by Lawrence Add 條件and nvl(A.InvSeqNo,0)=0
        v_SQL3 := v_SQL3||' Union '||v_SQL2||',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and C.PreInvoice=1 and nvl(A.InvSeqNo,0)=0'|| v_SQL4 || 
                          v_SQL1 || ' and A.AccountNo is null';
        v_SQL := v_SQL||' Union '||v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and C.PreInvoice=1 and nvl(A.InvSeqNo,0)=0'|| v_SQL4 || 
                        v_SQL1; 
        -- 2010.05.13 by Lawrence Add v_SQL5 , v_SQL6
        v_SQL6 := v_SQL6||' Union '||v_SQL2||',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo and D.PreInvoice=1'|| v_SQL4 || 
                          v_SQL1 || ' and A.AccountNo is null';
        v_SQL5 := v_SQL5||' Union '||v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo and D.PreInvoice=1'|| v_SQL4 || 
                        v_SQL1; 
     else
        -- 2010.05.13 by Lawrence Add 條件and nvl(A.InvSeqNo,0)=0
        v_SQL3 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and nvl(A.InvSeqNo,0)=0'||v_SQL4|| 
                          v_SQL1 || ' and A.AccountNo is null';
        -- 2004.12.20 by Lawrence order by ShouldDate,MediaBillNo對調
        v_SQL := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and nvl(A.InvSeqNo,0)=0'||v_SQL4|| 
                        v_SQL1 || ' and A.AccountNo is not null';
        -- 2010.05.13 by Lawrence Add v_SQL5 , v_SQL6
        v_SQL6 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvTitle,C.ContName,C.InvAddress,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo'||v_SQL4|| 
                          v_SQL1 || ' and A.AccountNo is null';
        -- 2004.12.20 by Lawrence order by ShouldDate,MediaBillNo對調
        v_SQL5 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO033Debt A, SO001 B, SO002 C, SO138 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and nvl(A.InvSeqNo,0)>0 and D.InvSeqNo=A.InvSeqNo'||v_SQL4|| 
                        v_SQL1 || ' and A.AccountNo is not null';
     end if;
     -- 2010.05.13 by Lawrence Add Union to Parameter
     v_SQL3 := v_SQL3||' Union '||v_SQL6||' order by CustId,InvNo,InvTitle,ContName';
     v_SQL := v_SQL||' Union '||v_SQL5||' order by AccountNo,ShouldDate,MediaBillNo';
  end if;
  
  --1.處理有帳戶的部分
  begin
     OPEN v_DyCursor FOR v_SQL;
  exception
    when others then
       rollback;
       p_RetMsg := 'SQL錯誤: ' || v_SQL;
       return -2;
  end;
  v_AccountNo :='xyz';
  v_CurrCitemCode := 0;
  v_ChkAccountNo := 'abc';
  v_ChkShouldDate := to_Date('19000101','YYYYMMDD');
  v_ChkMediaBillNo := '19000100001';
  v_ChkCustId := 0;
  loop
    -- 2010.07.15 by Pierre, 加取FaciSeqNo
    FETCH v_DyCursor INTO z_CustId,z_BillNo,z_ShouldDate,z_ShouldAmt,z_CitemCode,z_CitemName,z_UpdTime,z_RealStartDate,z_RealStopDate,z_Note,z_Qty,z_ClctName,z_Item,z_strInvNo,z_InvoiceType,z_Address,z_AddrNo,z_AccountNo,z_MediaBillNo,z_ServiceType,z_InvSeqNo, z_InvPurposeCode, z_InvPurposeName, z_CombCitemCode, z_CombCitemName, z_CombAmount, z_CombStartDate, z_CombStopDate, z_FaciSeqNo;
    EXIT WHEN v_DyCursor%NOTFOUND;
    DBMS_OUTPUT.PUT_LINE('v_AccountNo:'||v_AccountNo||' / z_AccountNo:'||z_AccountNo||' / z_MediaBillNo:'||z_MediaBillNo);
    -- 2010.05.11 by Lawrence
    if z_CombCitemCode is not null and z_CombStartDate is not null then
       z_ShouldAmt := z_CombAmount;
       z_RealStartDate := z_CombStartDate;
       z_RealStopDate := z_CombStopDate;
    end if;
    -- 2006.12.28 by Lawrence
    if p_AreaCity=1 then
       select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
       z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
    end if;
    -- 2008.08.29 by Lawrence Add or (v_chkInvSeqNo != nvl(z_InvSeqNo,0))
    if (v_AccountNo != z_AccountNo) or (v_ChkCustId != z_CustId) or (v_chkInvSeqNo != nvl(z_InvSeqNo,0)) then
       v_AccountNo := z_AccountNo;
       v_chkInvSeqNo := nvl(z_InvSeqNo,0);	-- 2008.08.29 by Lawrence
       -- 取客戶的發票資料
       begin
         DBMS_OUTPUT.PUT_LINE('a.'||z_InvTitle);
         begin
           -- 取SO002A資料
           if p_AddrType=2 and  nvl(z_InvSeqNo,0)>0 then
             -- 2007.12.06 by Lawrence
             -- 2010.07.15 by Pierre, 加取3個欄位
             select InvNo, InvTitle, InvAddress, InvoiceType,ChargeAddress, ChargeAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_InvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ChargeTitle, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                from SO138 where InvSeqNo=z_InvSeqNo; 
             -- 2006.12.28 by Lawrence
             if p_AreaCity=1 then
                select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
                z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
             end if;
           elsif p_AddrType=3 and  nvl(z_InvSeqNo,0)>0 then
             -- 2007.12.06 by Lawrence
             -- 2010.07.15 by Pierre, 加取3個欄位
             select InvNo, InvTitle, InvAddress, InvoiceType,MailAddress, MailAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_InvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ChargeTitle, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                from SO138 where InvSeqNo=z_InvSeqNo;
             -- 2006.12.28 by Lawrence
             if p_AreaCity=1 then
                select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
                z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
             end if;
           elsif  nvl(z_InvSeqNo,0)>0 then
             -- 2006.12.28 by Lawrence
             -- 2010.07.15 by Pierre, 加取3個欄位
             select InvNo, InvTitle, InvAddress, InvoiceType, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_InvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_ChargeTitle, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                from SO138 where InvSeqNo=z_InvSeqNo;
           elsif nvl(z_InvSeqNo,0)=0 then
             -- 2010.07.15 by Pierre, 若[發票流水號]無值, 則至客戶各服務基本資料檔(SO002)中, 取CustId+CompCode+ServiceType對應資料的發票開立種類(InvoiceKind)、受贈單位代碼(DenRecCode)、受贈單位名稱(DenRecName)
             begin
               select InvoiceKind, DenRecCode, DenRecName into z_InvoiceKind, z_DenRecCode, z_DenRecName from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
             exception
               when others then
                 z_InvoiceKind := 0;
                 z_DenRecCode := null;
                 z_DenRecName  := null;
             end;
           end if;
         exception
           when others then
             z_InvNo:=null;
             z_InvTitle:=null;
             z_InvAddress:=null;
             z_ChargeTitle:=null;
             -- 2010.07.15 by Pierre, 加3個欄位
             z_InvoiceKind := 0;
             z_DenRecCode := null;
             z_DenRecName  := null;
             DBMS_OUTPUT.PUT_LINE('b1.'||z_InvTitle||'//'||SQLERRM);         
         end;
         -- 2003.09.26 by Lawrence
         x_ContName:=z_ChargeTitle;
         if z_InvTitle is null then  -- 第一順位
            z_InvTitle:=z_ChargeTitle;
         end if;
         DBMS_OUTPUT.PUT_LINE('c.'||z_InvNo||'-- '||z_InvTitle||'/z_InvoiceType='||z_InvoiceType||'/p_InvType='||p_InvType||' -- '||z_InvAddress);
         -- check發票種類, 符合的產生資料
         if p_InvType!=1 then
            if z_InvoiceType<>p_InvType then
               GOTO go_next;
            end if;
         else
            if z_InvoiceType=0 then
               GOTO go_next;
            end if;
         end if;
         if z_InvAddress is null then
            z_InvAddress := z_Address;
         end if;
       exception
         when others then
            if z_InvAddress is null then
               z_InvAddress := z_Address;
            end if;
            z_InvNo := null;
       end;
       -- 取寄件地址的郵遞區號
       begin
          select ZipCode into z_ZipCode from SO014 where AddrNo=z_AddrNo and CompCode=p_CompCode;
       exception
         when others then
            z_ZipCode := null;
       end;
       -- 取客戶電話一
       begin
          select tel1 into z_Tel1 from SO001 where CustId=z_CustId and CompCode=p_CompCode and instr(ServiceType,z_ServiceType)>0;
       exception
         when others then
            z_Tel1 := null;
       end;
       -- *****************************************************
       -- 2010.07.15 by Pierre, 取客戶的常用E-MAIL、行動電話
       begin
         if z_FaciSeqNo is not null then
           -- 因z_FaciSeqNo內容可能為客編, 造成SO004無該筆設備資料, 此狀況就要去SO001,SO002取Email, ContMobile
           begin
             select Email, ContMobile into z_Email, z_ContMobile from SO004 where SeqNo=z_FaciSeqNo;
           exception
             when others then
               select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
               select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
           end;
         else
           select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
           select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
         end if;
       exception
         when others then
           z_Email := null;
           z_ContMobile := null;
       end;
       -- *****************************************************
    else
       if z_InvAddress is null then
          z_InvAddress := z_Address;
       end if;
    end if;
    -- 決定課稅別
    if v_CurrCitemCode != z_CitemCode then
       v_CurrCitemCode := z_CitemCode;
       begin
          select TaxCode into v_TaxCode from CD019 where CodeNo=v_CurrCitemCode and (ServiceType=z_ServiceType or ServiceType is null);
          if v_TaxCode = 1 then
             z_TaxType := '1';
             -- 取稅率
             begin
                select Rate1 into z_TaxRate from CD033 where CodeNo=v_TaxCode;
             exception
               when others then
                  z_TaxRate := Nvl(z_TaxRate,0);
             end;
          elsif v_TaxCode = 3 then
             z_TaxType := '3';
             z_TaxRate:=0;
          elsif v_TaxCode is null then
             z_TaxType := null;
             z_TaxRate:=0;
          else
             z_TaxType := '2';
             z_TaxRate:=0;
          end if;
       exception
         when others then
            p_RetMsg := '找不到收費項目代碼';
            return -3;        
       end;
       --取稅率代碼檔中的"是否課說"
       if v_TaxCode is not null then
          begin
             select TaxFlag into v_TaxFlag from CD033 where CodeNo=v_TaxCode;
          exception
            when others then
               p_RetMsg := 'CD033.TaxFlag Not Found ?';
               return -99;
          end;
       end if;
       if v_TaxFlag=0 then 
          z_TaxType := null;
          z_TaxRate :=0;
       end if;
    end if;
    -- 取異動日期與時間
    begin
      if substr(z_UpdTime,3,1)='/' then
         z_UpdA := to_date(to_char(to_number(substr(z_UpdTime,1,2))+1911)||substr(z_UpdTime,4,2)||substr(z_UpdTime,7,2),'YYYYMMDD');
         z_UpdB := to_number(substr(z_UpdTime, 10, 2)||substr(z_UpdTime, 13, 2));
      else
         if z_UpdTime is not null then
            z_UpdA := to_date(substr(z_UpdTime,1,4)||substr(z_UpdTime,6,2)||substr(z_UpdTime,9,2),'YYYYMMDD');
         else
            z_UpdA :=Null;
         end if;
         z_UpdB := to_number(substr(z_UpdTime, 12, 2)||substr(z_UpdTime, 15, 2));
      end if;
    exception
      when others then
         p_RetMsg := '異動日期錯誤! '||z_UpdTime||' '||z_UpdA;
         return -6;              
    end;
    -- 新增至發票拋檔暫存檔
    -- 2004.12.20 by Lawrence 保留原行, 再依需求調整程式
    -- 2008.08.15 by Corey 增加INVSEQNO相同的條件 from IMS3986
    if (v_ChkAccountNo = v_AccountNo) and (v_ChkShouldDate = z_ShouldDate) and (v_ChkCustId=z_CustId) and (v_chkInvSeqNo= nvl(z_InvSeqNo,0)) then
       z_VouString1 := '';
       if z_TaxRate>0 then
          z_TaxMoney := z_ShouldAmt - round((z_ShouldAmt/(1+(z_TaxRate/100))),0);
       else
          z_TaxMoney := 0;
       end if;
       if z_ShouldAmt>0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_ShouldDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_ShouldAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_ShouldAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       DBMS_OUTPUT.PUT_LINE('4.1-VouString '||z_VouString||' /ShouldAmt:'||z_ShouldAmt||' /TaxType:'||z_TaxType);
    else
       -- 2010.07.15 by Pierre, 加串"發票開立種類、常用E-MAIL、行動電話、受贈單位代碼、受贈單位名稱" 5個欄位
       z_VouString1 := '##'||to_char(v_InvCompCode)||v_char||to_char(z_CustId)||v_char||rtrim(z_Tel1)||v_char||z_InvNo||v_char||z_InvTitle||v_char||z_Address||v_char||z_ZipCode||v_char||z_InvAddress||v_char||x_ContName||v_char||to_char(z_InvPurposeCode)||v_char||z_InvPurposeName||
                       v_char||to_char(z_InvoiceKind)||v_char||rtrim(z_Email)||v_char||rtrim(z_ContMobile)||v_char||to_char(z_DenRecCode)||v_char||rtrim(z_DenRecName) ;
       DBMS_OUTPUT.PUT_LINE('3-VouString '||z_VouString1||' /TaxRate:'||z_TaxRate);
       if z_TaxRate>0 then
          z_TaxMoney := z_ShouldAmt - round((z_ShouldAmt/(1+(z_TaxRate/100))),0);
       else
          z_TaxMoney := 0;
       end if;
       if z_ShouldAmt>0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_ShouldDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_ShouldAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_ShouldAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
    end if;
    if z_TaxType is not null and z_TaxType<>' ' and z_ShouldAmt>0 then 
       v_ChkAccountNo := v_AccountNo;
       v_ChkMediaBillNo := z_MediaBillNo;
       v_ChkShouldDate := z_ShouldDate;
       v_ChkCustId := z_CustId;
       v_chkInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString1,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       end if;
       -- Write Detail 
       insert into Tmp004
          (VouString,Sno,CustId,MFlag) 
       values
          (z_VouString,v_Sno,z_CustId,p_Mode);
       v_Sno := v_Sno+1;
       insert into Tmp005 
          (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
       values 
          (z_CustId, z_InvNo, z_InvTitle, z_InvAddress,  z_BillNo, z_ShouldDate, z_ShouldAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       v_ShouldAmt := v_ShouldAmt + z_ShouldAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt1 := x_RcdCnt1 + 1;		-- 2010.07.15 by Pierre
     end if;
     <<go_next>>
         Null;
  end loop;
  CLOSE v_DyCursor;             --close cursor
  
  --2.處理無帳戶的部分
  begin
     OPEN v_DyCursor FOR v_SQL3;
  exception
    when others then
       rollback;
       p_RetMsg := 'SQL3錯誤: ' || v_SQL3;
       return -5;
  end;
  v_CustId:=0;
  v_ContName:='xyz';
  v_CurrCitemCode := 0;
  v_ChkCustId := -2;
  v_ChkShouldDate := to_Date('19000101','YYYYMMDD');
  v_chkInvSeqNo:=0;
  loop
     -- 2010.07.15 by Pierre, 加取FaciSeqNo
     FETCH v_DyCursor INTO z_CustId,z_BillNo,z_ShouldDate,z_ShouldAmt,z_CitemCode,z_CitemName,z_UpdTime,z_RealStartDate,z_RealStopDate,z_Note,z_Qty,z_ClctName,z_Item,z_strInvNo,z_InvoiceType,z_Address,z_AddrNo,z_AccountNo,z_MediaBillNo,z_ServiceType,z_InvTitle,z_ContName,z_InvAddress, z_InvSeqNo, z_InvPurposeCode, z_InvPurposeName, z_CombCitemCode, z_CombCitemName, z_CombAmount, z_RealStartDate, z_RealStopDate, z_FaciSeqNo;
     EXIT WHEN v_DyCursor%NOTFOUND;
     -- 2010.05.11 by Lawrence
     if z_CombCitemCode is not null and z_CombStartDate is not null then
        z_ShouldAmt := z_CombAmount;
        z_RealStartDate := z_CombStartDate;
        z_RealStopDate := z_CombStopDate;
     end if;
     -- 2008.04.17 by Lawrence Add讀取SO138
     if nvl(z_InvSeqNo,0)>0 then
        -- 2008.04.24 by Lawrence 調整一下程式無帳號處理的z_strInvNo, z_ContName變數名稱
        -- 2008.08.25 by Lawrence Add p_AddrType=2 and p_AddrType=3 處理
        if p_AddrType=2 then
           begin
               -- 2010.07.15 by Pierre, 加取3個欄位
               select InvNo, InvTitle, InvAddress, InvoiceType,ChargeAddress, ChargeAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_strInvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ContName, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                  from SO138 where InvSeqNo=z_InvSeqNo; 
              DBMS_OUTPUT.PUT_LINE('-- z_ContName=SO138.ContName['||z_ContName||']');
           exception
              when others then
                 z_strInvNo:=null;
                 z_InvTitle:=null;
                 z_InvAddress:=null;
                 z_ContName:=null;
                 DBMS_OUTPUT.PUT_LINE('-- z_ContName:=null');              
           end;
        elsif p_AddrType=3 then
           begin
              -- 2010.07.15 by Pierre, 加取3個欄位
              select InvNo, InvTitle, InvAddress, InvoiceType,MailAddress, MailAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_strInvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ContName, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                 from SO138 where InvSeqNo=z_InvSeqNo; 
              DBMS_OUTPUT.PUT_LINE('-- z_ContName=SO138.ContName['||z_ContName||']');
           exception
              when others then
                 z_strInvNo:=null;
                 z_InvTitle:=null;
                 z_InvAddress:=null;
                 z_ContName:=null;
                 DBMS_OUTPUT.PUT_LINE('-- z_ContName:=null');              
           end;
        else
           begin
              -- 2010.07.15 by Pierre, 加取3個欄位
              select InvNo, InvTitle, InvAddress, InvoiceType, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_strInvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_ContName, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                 from SO138 where InvSeqNo=z_InvSeqNo; 
              DBMS_OUTPUT.PUT_LINE('-- z_ContName=SO138.ContName['||z_ContName||']');
           exception
              when others then
                 z_strInvNo:=null;
                 z_InvTitle:=null;
                 z_InvAddress:=null;
                 z_ContName:=null;
                 DBMS_OUTPUT.PUT_LINE('-- z_ContName:=null');              
           end;
        end if; 
      else
        -- 2010.07.15 by Pierre, 若[發票流水號]無值, 則至客戶各服務基本資料檔(SO002)中, 取CustId+CompCode+ServiceType對應資料的發票開立種類(InvoiceKind)、受贈單位代碼(DenRecCode)、受贈單位名稱(DenRecName)
        begin
          select InvoiceKind, DenRecCode, DenRecName into z_InvoiceKind, z_DenRecCode, z_DenRecName from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
        exception
          when others then
            z_InvoiceKind := 0;
            z_DenRecCode := null;
            z_DenRecName  := null;
        end;
     end if;
     -- 2007.03.08 by Lawrence
     if p_AreaCity=1 then
        select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
        z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
     end if;
     -- 2003.10.23 by Lawrence
     x_ContName:='';
     DBMS_OUTPUT.PUT_LINE('v_AccountNo:'||v_AccountNo||' / z_AccountNo:'||z_AccountNo||' / z_MediaBillNo:'||z_MediaBillNo);
     if (v_CustId != z_CustId) or (nvl(v_ContName,'xyz') != nvl(z_ContName,'xyz')) then
        DBMS_OUTPUT.PUT_LINE('v_CustId:'||v_CustId||' / z_CustId:'||z_CustId||' --- v_ContName:'||v_ContName||' / z_ContName:'||z_ContName);
        v_CustId := z_CustId;
        v_ContName := z_ContName;
        -- 取客戶的發票資料
        begin
           -- check發票種類, 符合的產生資料
           if p_InvType!=1 then
              if z_InvoiceType<>p_InvType then
                 GOTO go_next;
              end if;
           else
              if z_InvoiceType=0 then
                 GOTO go_next;
              end if;
           end if;
           ---------------------------
           if z_InvTitle is null then  --第二順位
              begin
                select custname into z_InvTitle from so001 where CustId=v_CustId and CompCode=p_CompCode and instr(ServiceType,z_ServiceType)>0; 
              exception
                when others then
                   z_InvTitle := null;   
              end;
           end if;
           -- 2003.09.26 by Lawrence第一順位
           if x_ContName is null then
              begin
                select custname into x_ContName from so001 where CustId=z_CustId and CompCode=p_CompCode;
                DBMS_OUTPUT.PUT_LINE('x_ContName is null => Get SO001.custname');
              exception
                when others then
                   x_ContName:=null; 
              end;
           end if;
           if z_InvAddress is null then
              z_InvAddress := z_Address;
           end if;
           z_InvNo := z_StrInvNo;
        exception
          when others then
             if z_InvTitle is null then  --第二順位
                begin
                   select custname into z_InvTitle from so001 where CustId=v_CustId and CompCode=p_CompCode and instr(ServiceType,z_ServiceType)>0; 
                exception
                  when others then
                     z_InvTitle := null;   
                end;
             end if;
             -- 2003.09.26 by Lawrence第一順位
             if x_ContName is null then
                begin
                   select custname into x_ContName from so001 where CustId=z_CustId and CompCode=p_CompCode;
                exception
                  when others then
                     x_ContName:=null; 
                end;
             end if;
             if z_InvAddress is null then
                z_InvAddress := z_Address;
             end if;
             z_InvNo := null;
        end;
        -- 取寄件地址的郵遞區號
        begin
           select ZipCode into z_ZipCode from SO014 where AddrNo=z_AddrNo and CompCode=p_CompCode;
        exception
          when others then
             z_ZipCode := null;
        end;
        -- 取客戶電話一
        begin
           select tel1 into z_Tel1 from SO001 where CustId=z_CustId and CompCode=p_CompCode and instr(ServiceType,z_ServiceType)>0;
        exception
          when others then
             z_Tel1 := null;
        end;
        -- *****************************************************
        -- 2010.07.15 by Pierre, 取客戶的常用E-MAIL、行動電話
        begin
          if z_FaciSeqNo is not null then
            -- 因z_FaciSeqNo內容可能為客編, 造成SO004無該筆設備資料, 此狀況就要去SO001,SO002取Email, ContMobile
            begin
              select Email, ContMobile into z_Email, z_ContMobile from SO004 where SeqNo=z_FaciSeqNo;
            exception
              when others then
                select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
                select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
            end;
          else
            select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
            select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
          end if;
        exception
          when others then
            z_Email := null;
            z_ContMobile := null;
        end;
        -- *****************************************************
     else
        if z_InvTitle is null then  --第二順位
           begin
              select custname into z_InvTitle from so001 where CustId=z_CustId and CompCode=p_CompCode; 
           exception
             when others then
                z_InvTitle := null;   
           end;
        end if;
        -- 2003.09.26 by Lawrence第一順位
        if x_ContName is null then
           begin
              select custname into x_ContName from so001 where CustId=z_CustId and CompCode=p_CompCode;
           exception
             when others then
                x_ContName:=null; 
           end;
        end if;
        if z_InvAddress is null then
           z_InvAddress := z_Address;
        end if;
     end if;
     -- 決定課稅別
     if v_CurrCitemCode != z_CitemCode then
       v_CurrCitemCode := z_CitemCode;
       begin
         select TaxCode into v_TaxCode from CD019 where CodeNo=v_CurrCitemCode and (ServiceType=z_ServiceType or ServiceType is null);
         if v_TaxCode = 1 then
            z_TaxType := '1';
            -- 取稅率
            begin
               select Rate1 into z_TaxRate from CD033 where CodeNo=v_TaxCode;
            exception
              when others then
                 z_TaxRate := Nvl(z_TaxRate,0);
            end;
         elsif v_TaxCode = 3 then
            z_TaxType := '3';
            z_TaxRate:=0;
         elsif v_TaxCode is null then
            z_TaxType := null;
            z_TaxRate:=0;
         else
            z_TaxType := '2';
            z_TaxRate:=0;
         end if;
       exception
         when others then
            p_RetMsg := '找不到收費項目代碼';
            return -3;        
       end;
       --取稅率代碼檔中的"是否課說"
       if v_TaxCode is not null then
          begin
             select TaxFlag into v_TaxFlag from CD033 where CodeNo=v_TaxCode;
          exception
            when others then
               p_RetMsg := 'CD033.TaxFlag not found ?';
               return -99;
          end;
       end if;
       if v_TaxFlag=0 then 
          z_TaxType := null;
          z_TaxRate :=0;
       end if;
    end if;
    -- 取異動日期與時間
    begin
       if substr(z_UpdTime,3,1)='/' then
          z_UpdA := to_date(to_char(to_number(substr(z_UpdTime,1,2))+1911)||substr(z_UpdTime,4,2)||substr(z_UpdTime,7,2),'YYYYMMDD');
          z_UpdB := to_number(substr(z_UpdTime, 10, 2)||substr(z_UpdTime, 13, 2));
       else
          if z_UpdTime is not null then
             z_UpdA := to_date(substr(z_UpdTime,1,4)||substr(z_UpdTime,6,2)||substr(z_UpdTime,9,2),'YYYYMMDD');
          else
             z_UpdA :=Null;
          end if;
          z_UpdB := to_number(substr(z_UpdTime, 12, 2)||substr(z_UpdTime, 15, 2));
       end if;
    exception
      when others then
         p_RetMsg := '異動日期錯誤! '||z_UpdTime||' '||z_UpdA;
         return -6;              
    end;
    -- 新增至發票拋檔暫存檔
    if (v_ChkCustId = z_CustId) and (v_ChkShouldDate = z_ShouldDate) and (nvl(v_ContName,'xyz') = nvl(z_ContName,'xyz')) and (v_chkInvSeqNo=nvl(z_InvSeqNo,0)) then
       z_VouString1 := '';
       if z_TaxRate>0 then
          z_TaxMoney := z_ShouldAmt - round((z_ShouldAmt/(1+(z_TaxRate/100))),0);
       else
          z_TaxMoney := 0;
       end if;
       if z_ShouldAmt>0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_ShouldDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_ShouldAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_ShouldAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       DBMS_OUTPUT.PUT_LINE('4.1-VouString '||z_VouString||' /ShouldAmt:'||z_ShouldAmt||' /TaxType:'||z_TaxType);
    else
       -- 2010.07.15 by Pierre, 加串"發票開立種類、常用E-MAIL、行動電話、受贈單位代碼、受贈單位名稱" 5個欄位
       z_VouString1 := '##'||to_char(v_InvCompCode)||v_char||to_char(z_CustId)||v_char||rtrim(z_Tel1)||v_char||z_InvNo||v_char||z_InvTitle||v_char||z_Address||v_char||z_ZipCode||v_char||z_InvAddress||v_char||x_ContName||v_char||to_char(z_InvPurposeCode)||v_char||z_InvPurposeName||
                       v_char||to_char(z_InvoiceKind)||v_char||rtrim(z_Email)||v_char||rtrim(z_ContMobile)||v_char||to_char(z_DenRecCode)||v_char||rtrim(z_DenRecName) ;
       DBMS_OUTPUT.PUT_LINE('3-VouString '||z_VouString1||' /TaxRate:'||z_TaxRate);
       if z_TaxRate>0 then
          z_TaxMoney := z_ShouldAmt - round((z_ShouldAmt/(1+(z_TaxRate/100))),0);
       else
          z_TaxMoney := 0;
       end if;
       if z_ShouldAmt>0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_ShouldDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_ShouldAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_ShouldAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
    end if;
    if z_TaxType is not null and z_TaxType<>' ' and z_ShouldAmt>0 then 
       v_ChkCustId := z_CustId;
       v_ChkShouldDate := z_ShouldDate;
       v_ContName := nvl(z_ContName,'xyz');
       v_chkInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString1,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       end if;
       -- Write Detail 
       insert into Tmp004
          (VouString,Sno,CustId,MFlag) 
       values
          (z_VouString,v_Sno,z_CustId,p_Mode);
       v_Sno := v_Sno+1;
       insert into Tmp005 
          (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
       values 
          (z_CustId, z_InvNo, z_InvTitle, z_InvAddress, z_BillNo, z_ShouldDate, z_ShouldAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       v_ShouldAmt := v_ShouldAmt + z_ShouldAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt2 := x_RcdCnt2 + 1;		-- 2010.07.15 by Pierre
     end if;
     <<go_next>>
         Null;
  end loop;
  CLOSE v_DyCursor;             --close cursor
  p_RetMsg := '產生完畢, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'秒, 處理筆數='||to_char(v_RcdCount)||', 處理金額='||to_char(v_ShouldAmt);
  DBMS_OUTPUT.PUT_LINE('有帳戶筆數='||x_RcdCnt1||', 無帳戶筆數='||x_RcdCnt2);		-- 2010.07.15 by Pierre
  commit;
  return v_RcdCount;
EXCEPTION
  WHEN OTHERS THEN 
     rollback;
     CLOSE v_DyCursor;             --close cursor
     DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
     p_RetMsg := SQLERRM;
     return -99;
END;
/
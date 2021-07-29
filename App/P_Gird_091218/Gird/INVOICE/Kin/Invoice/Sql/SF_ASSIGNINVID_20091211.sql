create or replace function sf_assigninvid
  ( p_User                     in varchar2,
    p_CompId                   in integer,
    p_LinkToMis                in varchar2,
    p_DbLink                   in varchar2,
    p_HowToCreate              in integer,
    p_InvDateEqualToChargeDate in integer,
    p_InvDate                  in varchar2,
    p_InvYearMonth             in varchar2,
    p_ChargeStartdate          in varchar2,
    p_ChargeStopDate           in varchar2,
    p_IdentifyID1              in varchar2,
    p_IdentifyID2              in integer,
    p_SystemID                 in varchar2,
    p_PrefixString             in varchar2,
    p_OrderBy                  in integer,
    p_MisDbOwner               in varchar2,
    p_Test                 in integer,
    p_RetCode                 out integer,
    p_RetMsg                  out varchar2,
    p_LogDateTime             out varchar2
  )
return number as

/*

  注意事項: 發票 Schema 的 owner 若是與 客服系統 Schema 的 owner 不同, 務必 grant 權限給發票系統的 owner
            客服系統所須 table         : grant select, update on SO033 to <發票系統的owner>
                                         grant select, update on SO034 to <發票系統的owner>
                                         grant select on so002a to <發票系統的owner>


  發票開立
  檔名: SF_AssignInvID.sql

  說明: 處理發票系統之發票開立(給號)

  參數 IN:

	p_User          		        varchar2 :  程式操作者
	p_CompCode      		        number   :  公司別代碼
	p_LinkToMis                 varchar2 :  是否更新客服系統   Y=>更新; N=>不更新
	p_HowToCreate   		        number   :  1=>預開;   2=>後開
	p_DbLink                    varchar2 :  跨資料庫存取時, 透過 DBLink 方式下SQL
  p_InvDateEqualToChargeDate	number   :  1=>發票日期等於收費日期;   2=>發票日期不等於收費日期
	p_InvDate			              varchar2 :  發票日期
	p_InvYearMonth 			        varchar2 :  發票年月
	p_ChargeStartdate 	        varchar2 :  收費起始日
	p_ChargeStopDate 		        varchar2 :  收費截止日
	p_IdentifyID1 			        number   :  識別碼(保留), 應該等於1
	p_IdentifyID2 			        number   :  識別碼(保留), 應該等於0
	p_SystemID 			            varchar2 :  System ID
  p_PrefixString 		          varchar2 :  發票字軌
	p_OrderBy 			            number   :  排序方式 : 0=>收費日期
	                                                   1=>郵遞區號
	                                                   2=>收費日期+郵遞區號
	                                                   3=>收費日期+郵遞區號+客編
  p_MisDbOwner                narchar2 :  回填至MIS的Table Space 的 Owner



  參數 OUT:

	p_RetMsg        		       varchar2  :     結果訊息
	p_RetCode       		       number    :     結果代碼


  Return
	number: 處理結果代碼, 對應訊息存於p_RetMsg中

         0:  p_RetMsg = '發票開立執行完畢'
        -1:  p_RetMsg = '參數錯誤'
       -99:  p_RetMsg = '其他錯誤'

*/

/*
    版本異動記錄
    First release    : 2002.12.23
    Modify By Jackal : 2003.03.11  原一次僅能針對一捲發票開立,改成可以同時多捲發票開立
    Modify By Jackal : 2003.04.19  1.篩選若發票金額 <0 的不用開發票並 Log 資料至 INV033
                                   2.檢查資料是否已開過發票,若開過不用開發票並 Log 資料至 INV033
    Modify By Jackal : 2003.04.24  INV016 多一欄位 ShouldBeAssigned , 用於判斷此收費項目是否要開立發票
    Modify By Jackal : 2003.07.22  多傳一個參數p_MisDbOwner
    Modify By Jackal : 2004.01.16  將變數s_CustSName由Varchar2(26)改為Varchar2(50)
    Modify By Jackal : 2004.05.13  1.若Inv016主檔與Inv017明細檔的合計金額不合,則不開立並Log 至 inv033
                                   2.檢核是否有相同seq,但收費明細的稅率別不同應不能開立並Log 至 inv033
                                   3.開立回填時若該筆資料不存在 so034或so033 中,則不開立並Log 至 inv033
                                   4.開立回填時若該筆資料其guino 已有值時(代表已開立過),則不開立並Log
    Modify By Jackal : 2004.05.14  INV008及INV032各加一個銷售額欄位(SaleAmount),銷售額=單價*數量(並四捨五入取整數)
    Modify By Nick   : 2005.01.17  發票檢查碼計算不對, 發票本檔 INV099 STARTNUM, CURNUM, ENDNUM 欄位改成 VARCHAR2(8)
    Modify By Nick   : 2005.04.22  抓取開立發票資料多加一項判斷, 稅別 TAXTYPE = 0 , 不予開立
    Modify By Nick   : 2005.09.30  針對 EMC 使用, 抓客服系統資料增加 DBLink 字串
    Modify By Nick   : 2005.10.04  增加入參數, p_Dblink, p_linkToMis,
                                   若 p_linkToMis = Y , 才更新客服系統資料,
                                   若 p_Dblink 有值, 則 sql 多加入 dblink 例如: select * from So033@catvc
    Modify By Nick   : 2005.10.20  修正因為多公司別檢查收費項目重覆, 而導致不能開立發票(多加判斷 CompId )
    Modify By Nick   : 2006.01.16  INV033 增加 CompId 公司別欄位, 多公司開立發票完後, 異常報表才會抓對資料
    Modify By Nick   : 2006.01.19  1.重新改寫
                                   2.讀取公司別設定的參數 INV001.AutoCreateNum, 增加開立判斷,
                                     若有設定 ( p_AutoCreateNum > 0 ) 代表有自動換開發票功能, 當發票明細筆數大
                                     於此數時必須自動換開第2張發票
    Modify By Nick   : 2006.02.09  1.使用 SQL Plus 編譯此 store function 時, 會無法編譯, 修改此 Bug.
                                   2.更新客服系統 SO033, SO034 欄位 ITEMNO 欄位打錯, 應為 ITEM
                                   3.發票本 INV099 的 LastInvDate 沒有更新到, 更新規則錯誤, 應為每次開立就必須更新
    Modify By Nick   : 2006.02.14  1.INV008, INV017, INV032 增加 LinkToMis 欄位, 是否更新或檢核客服系統控制到該筆
                                     明細, INV001 的 LinkToMis 是總開關
                                     Ex: INV001.LINKTOMIS = 'Y' AND INV008.LINKTOMIS = 'Y' 這樣才會真的去檢核及更新
                                     客服系統, 若 INV001.LINKTOMIS <> 'Y' or INV008.LINKTOMIS <> 'Y' 則不檢核及更新客服系統
    Modify By Nick   : 2006.04.10  1.增加發票明細收費項目合併功能
                                     新增 Table INV032A, INV008A
                                     新增欄位 INV017.ACCOUNTNO 文字檔匯入時會多此資料(只限CNS)
                                     開立時, 先判斷明細資料是否有須要合併成同一筆明細資料,
                                     Ex: 代收-亞太CM上網費1M, 數量:1, 金額:300 --|
                                         代收-亞太CM上網費1M, 數量:2, 金額:600 --|--> 合併成一筆 【代收-亞太CM上網費1M】, 數量:1, 金額: 1200
                                         代收-亞太CM上網費1M, 數量:1, 金額:300 --|
                                     若此明細為P服務, 則還要多判斷為同一CP門號才可以合併, 若不同門號則要併到另一個明細
                                   2.若此明細為信用卡繳費, 則多取信號卡號, 列印發票時會用到

    Modify By Nick   : 2006.06.26  1.INV001增加參數, 控制不同設備門號時, 發票收費明細是否須要合併成同一收費項目
                                     INV001.FACICOMBINE = 1 時, 若此明細為P服務, 則不同門號一樣合併到同一明細
                                   2.INV008 增加2個欄位, FACISNO, ACCOUNTNO 當不為合併項目時, 此2個欄位儲存 CP 門號及信用卡號
                                   3.開立發票排序方式, 新增 收費日期+郵遞區號

    Modify By Nick   : 2006.08.28  1.開立時, 重新計算銷售額及稅額, 銷售額=(總金額/1.05),  稅額=(總金額-銷售額)
                                   2.所以使用重新計算過後的銷售額及稅額, 回頭更新主檔的發票金額, 這可能會更原本客服拋過來的金額不一致
                                     尤其是合併過後的

    Modify By Nick   : 2006.11.10  1. 2006.08.28 版本有bug,當自動換開發票時,主發票總金額,主發票銷售額,主發票稅額欄位皆為該張的金額
    Modify By Nick   : 2007.05.24  開立發票排序方式, 新增 收費日期+郵遞區號+客編
    Modify By Nick   : 2008.04.07  發票開立增加 1.發票用途代碼 2.發票用途代碼說明 (INV016, INV031, INV007)
    Modify By Nick   : 2008.08.12  發票開立時, 增加回寫客服系統 SO033 及 SO034 欄位 INVPURPOSECODE, INVPURPOSENAME
	Modify By Kin     :  2009.12.11  問題集5358 有合併項目則單據起迄日期取INV008A的Min(StartDate)與Max(EndDate)
*/



  /* 存放前端傳進來的多組發票字軌 */

  type TInv099 is record (
       CompId       inv099.compid%type,
       YearMonth    inv099.yearmonth%type,
       Prefix       inv099.prefix%type,
       StartNum     inv099.startnum%type,
       EndNum       inv099.endnum%type,
       CurNum       inv099.curnum%type,
       Useful       inv099.useful%type,
       LastInvDate  inv099.lastinvdate%type,
       HasChange    boolean );

  type TInv099Table is table of TInv099 index by binary_integer;

  aInv099             TInv099Table;
  aInv099MaxBounds    binary_integer;
  aInv099Index        binary_integer;
  aTempStr            varchar2(2000);
  aPos                integer;
  aDetailCount        integer;
  aMasterCanCreate    boolean;
  aCheckNo            varchar2(2);
  aNotes              inv033.notes%type;
  aInv017TotalAmount  inv017.totalamount%type;
  aAutoCreateNum      inv001.autocreatenum%type;
  aFaciCombine        inv001.facicombine%type;
  aInvId              inv007.invid%type;
  aRealInvDate        inv007.invdate%type;
  aInvTaxType         inv007.taxtype%type;
  aInvTaxRate         inv007.taxrate%type;
  aDetailSeq          integer;
  aOldSeq             integer;
  aItemDesc           inv005.description%type;
  aMergeUnitPrice     inv008.unitprice%type;
  aMergeTaxAmount     inv008.taxamount%type;
  aMergeSaleAmount    inv008.saleamount%type;
  aMergeTotalAmount   inv008.totalamount%type;
  aMergeQuantity      inv008.quantity%type;
  aMergeCount         integer;
  aAccountNo          inv008a.accountno%type;
  aFacisNo            inv008a.facisno%type;
  aOldItemRefId       inv008a.facisno%type;

  type TAutoCreate is record (
       InvId          inv007.invid%type,
       SaleAmount     inv007.saleamount%type,
       TaxAmount      inv007.taxamount%type,
       InvAmount      inv007.invamount%type,
       MainSaleAmount     inv007.saleamount%type,
       MainTaxAmount      inv007.taxamount%type,
       MainInvAmount      inv007.invamount%type);

  type TAutoCreateTable is table of TAutoCreate index by binary_integer;

  aAutoCreateList       TAutoCreateTable;
  aAutoCreateMaxBounds  binary_integer;
  aCount                integer;

              /* ------------------------------------------------------------------ */

              /*  計算發票檢查碼 */

              function CalcInvCheckSum(aInvId in varchar2, aInvParam in varchar2)
                 return varchar2
              as
                 type IntegerArray is table of varchar2(2) index by binary_integer;
                 aCalcResult IntegerArray;
                 aLength binary_integer;
                 aResult integer := 0;
              begin
                 for aIndex in 1..Length( aInvId )
                 loop
                     aLength := aIndex;
                     aCalcResult( aLength ) := Lpad( to_char(
                       substr( aInvId, aIndex, 1 ) * substr( aInvParam, aIndex, 1 ) ), 2, '0' );
                 end loop;
                 for aIndex in 1..aLength
                 loop
                    if ( aIndex = 1 ) then
                      aCalcResult( aIndex ) := substr( aCalcResult( aIndex ),
                        length( aCalcResult( aIndex ) ), 1 );
                    else
                      aCalcResult( aIndex ) :=
                        ( substr( aCalcResult( aIndex ),
                           length( aCalcResult( aIndex ) ), 1 ) ) +
                        ( substr( aCalcResult( aIndex ),
                           length( aCalcResult( aIndex - 1 ) ) - 1, 1 ) );
                    end if;
                 end loop;
                 for aIndex in 1..aLength
                 loop
                    aResult := ( aResult + to_number( aCalcResult( aIndex ) ) );
                 end loop;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 取出可用的發票號碼, 順便更新發票本可用狀態 */

              function GetInvNumber(aIndex in out binary_integer, aId out varchar2, aDate in date) return boolean
              as
              begin
                aId := null;
                 << RE_GET >>
                   null;
                 if ( aIndex between 1 and aInv099MaxBounds ) then
                    /* 現在這本發票本是否可用 */
                    if ( aInv099( aIndex ).CurNum <= ( aInv099( aIndex ).EndNum ) ) then
                       /* 取發票號碼 */
                       aId := ( aInv099( aIndex ).prefix || Lpad( to_char( to_number( aInv099( aIndex ).CurNum ) ), 8, '0' ) );
                       aInv099( aIndex ).CurNum := lpad( to_char( to_number( aInv099( aIndex ).CurNum ) + 1 ), 8, '0' );
                       aInv099( aIndex ).LastInvDate := aDate;
                       aInv099( aIndex ).HasChange := true;
                    end if;
                    /* 判斷是否用完 */
                    if ( aInv099( aIndex ).CurNum > ( aInv099( aIndex ).EndNum ) ) then
                      aInv099( aIndex ).Useful := 'N';
                      aInv099( aIndex ).LastInvDate := aDate;
                      aInv099( aIndex ).HasChange := true;
                      aIndex := ( aIndex + 1 );
                      /* 沒有發票號碼, 再取一次 */
                      if ( aId is null ) then goto RE_GET; end if;
                    end if;
                 end if;
                 if ( aId is null ) then
                   p_RetMsg := '取出可用發票號碼失敗, 取出號碼為空值或是發票本號碼已用完';
                   p_RetCode := -99;
                 end if;
                 return ( aId is not null );
              end;

              /* ------------------------------------------------------------------- */


              /*  取 DBLink 字串 */

              function GetDBLink return varchar2
              as
                 aResult varchar2(30);
              begin
                 aResult := null;
                 if ( p_DbLink is not null ) then
                   if substr( p_DbLink, 1, 1 ) <> '@' then
                     aResult := '@' || p_DbLink;
                   else
                     aResult := p_DbLink;
                   end if;
                 end if;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* HowToCreate Mapping To CCB TableName */

              function GetTableNameByHowToCreate return varchar2
              as
                 aResult varchar2(100);
              begin
                 if ( p_HowToCreate = '1' ) then
                   aResult := 'SO033';
                 else
                   aResult := 'SO034';
                 end if;
                 return ( aResult || GetDBLink );
              end;

              /* ------------------------------------------------------------------- */

              function GetInv001Param(aNum in out integer, aFaciCombine in out integer) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    select nvl( autocreatenum, 0 ), nvl( facicombine, 0 )
                      into aNum, aFaciCombine
                      from inv001
                     where identifyid1 = p_IdentifyID1
                       and identifyid2 = p_IdentifyID2
                       and compid = p_CompId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '取公司別參數(自動換開發票及設備門號是否合併)時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 檢核客服系統此筆發票明細資料 */

              function CheckMisSystem(aId in varchar2, aBillNo in varchar2, aItem in number, aMsg out varchar2 ) return boolean
              as
                 aSql varchar2(255);
                 aTableName varchar2(100);
                 aGuiNo varchar2(10);
              begin
                 aMsg := null;
                 aTableName := GetTableNameByHowToCreate;
                 aSql := ' SELECT GUINO FROM ' || p_MisDbOwner || '.' || aTableName ||
                         '  WHERE CUSTID = ''' || aId || ''' AND BILLNO = ''' || aBillNo || ''' AND ITEM = ''' || aItem || '''';
                 begin
                   execute immediate aSql into aGuiNo;
                   if ( aGuiNo is not null ) then
                      aMsg := '客服系統內此筆收費單號已有發票號碼, 客服系統內發票號碼 : ' || aGuiNo;
                   end if;
                 exception
                   when others then aMsg := '客服系統無此筆對應資料。';
                 end;
                 return ( aMsg is null );
              end;


              /* ------------------------------------------------------------------- */

              /* 檢核發票系統內是否已有此收費單號跟項次 ( 已做廢不算 ), 若有, 則把發票號碼記下來 */

              function CheckAlreadyCreate(aBillId in inv008.billid%type, aItemNo in inv008.billiditemno%type,
                aMsg out varchar2) return boolean
              as
              begin
                 aMsg := null;
                 for cCheck in ( select a.invid from inv007 a, inv008 b where a.invid = b.invid
                    and a.compid = p_CompId  and a.isobsolete = 'N' and b.billid = aBillId and b.billiditemno = aItemNo
                    order by a.invid )
                 loop
                    if ( aMsg is null ) then
                      aMsg := '此收費單號已開立發票, 發票號碼:' || cCheck.invid;
                    else
                      aMsg := ( aMsg || ',' || cCheck.invid );
                    end if;
                 end loop;
                 return ( aMsg is not null );
              end;


              /* ------------------------------------------------------------------- */

              /* 發票開立檢核為不須開立時, 必須 Log 至 INV033, 同時註明原因 */

              function LogToInv033(aSeq in inv033.seq%type, aBillId in inv033.billid%type, aItemNo in inv033.billiditemno%type,
                 aCustId in varchar2, aCustName in varchar2, aNotes in inv033.notes%type ) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv033 ( seq, billid, billiditemno, taxtype, chargedate, itemid,
                       description, quantity, unitprice, taxrate, taxamount, totalamount, startdate,
                       enddate, chargeen, notes, custid, custname, logtime, upten, compid )
                    select a.seq, a.billid, a.billiditemno, a.taxtype, a.chargedate, a.itemid,
                       a.description, a.quantity, a.unitprice, a.taxrate, a.taxamount, a.totalamount, a.startdate,
                       a.enddate, a.chargeen, substr( aNotes, 1, 200 ), aCustId, aCustName,
                       to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User, p_CompId
                      from inv017 a
                     where a.seq = aseq and a.billid = aBillId and a.billiditemno = aItemNo;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV033 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 檢核客戶主檔是否有此客戶資料 */

              function CheckInv002Exists(aCustId in varchar2) return boolean
              as
                 aCount integer;
              begin
                 select count(1) into aCount from inv002
                  where custid = aCustId and compid = p_CompId
                    and identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2;
                 return ( aCount > 0 );
              end;

              /* ------------------------------------------------------------------- */

              /* 檢核客戶明細檔是否有此客戶資料 */

              function CheckInv019Exists(aCustId in varchar2) return boolean
              as
                 aCount integer;
              begin
                 select count(1) into aCount from inv019
                  where custid = aCustId and compid = p_CompId
                    and identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2
                    and titleid = 1;
                 return ( aCount > 0 );
              end;


              /* ------------------------------------------------------------------- */

              /* 寫發票系統的客戶主檔 */

              function WriteToInv002(aSeq in varchar2) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv002 ( identifyid1, identifyid2, compid, custid, custsname, custname, mzipcode,
                       mailaddr, tel1, isselfcreated, upttime, upten )
                    select p_IdentifyID1, p_IdentifyID2, p_CompId, a.custid, a.chargetitle, a.title, a.zipcode,
                       a.mailaddr, a.tel, 'N', to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User
                     from inv016 a
                    where a.seq = aSeq
                      and a.compid = p_CompId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV002 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 寫發票系統的客戶明細檔 */

              function WriteToInv019(aSeq in varchar2) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv019 ( identifyid1, identifyid2, compid, custid, titleid, titlesname,
                       titlename, businessid, mzipcode, mailaddr, invaddr, upttime, upten )
                    select p_IdentifyID1, p_IdentifyID2, p_CompId, a.custid, 1, a.chargetitle, a.title,
                       a.businessid, a.zipcode, a.mailaddr, a.invaddr, to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User
                      from inv016 a
                     where a.seq = aSeq
                       and a.compid = p_CompId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV019 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 寫發票暫存主檔 */

              function WriteToInv031(aSeq in varchar2, aId in varchar2, aNo in varchar2, aInvDate in date,
                aTaxType out varchar2, aTaxRate out number)
                 return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv031 ( identifyid1, identifyid2, checkno, invid, canmodify, invdate, chargedate,
                        compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat, taxtype,
                        taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete, howtocreate,
                        upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount, maininvamount, invuseid, invusedesc )
                    select p_IdentifyID1, p_IdentifyID2, aNo, aId, 'Y', aInvDate, a.chargedate,
                        p_CompId, a.custid, a.chargetitle, a.title, a.zipcode, a.invaddr, a.mailaddr, a.businessid, '1', a.taxtype,
                        a.taxrate, 0, 0, 0, 0, 'N', 'N', p_HowToCreate,
                        to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User, '0', aId, 0, 0, 0, a.invuseid, a.invusedesc
                    from inv016 a
                   where a.seq = aSeq
                     and a.compid = p_CompId;
                    /* 抓 TaxType, TaxRate 回傳 */
                    aTaxType := null;
                    aTaxRate := 0;
                    select taxtype, taxrate into aTaxType, aTaxRate from inv031
                     where identifyid1 = p_IdentifyID1
                       and identifyid2 = p_IdentifyID2
                       and invid = aId;
                    if ( aTaxType = '1' ) and ( aTaxRate = 0 ) then aTaxRate := 5; end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV031 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 寫合併發票明細檔 */

              function WriteToInv032A(aSeq in varchar2, aId in varchar2, aBillId in varchar2, aBillidItemno in number,
                aAcctNo in varchar2, aFacisNo in varchar2)
                 return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv032A ( itemidref, invid, billid, billiditemno, seq, startdate, enddate,
                       itemid, description, quantity, unitprice, taxamount, saleamount, totalamount,
                       servicetype, chargeen, linktomis, facisno, accountno )
                    select b.itemidref, aId, aBillId, aBillidItemno, null, a.startdate, a.enddate,
                       a.itemid, a.description, a.quantity, a.unitprice, a.taxamount, ( a.quantity * a.unitprice ), a.totalamount,
                       a.servicetype, a.chargeen, nvl( a.linktomis, 'Y' ), aFacisNo, aAcctNo
                     from inv017 a, inv005 b
                    where b.IDENTIFYID1(+) = p_IdentifyID1
                      and b.IDENTIFYID2(+) = p_IdentifyID2
                      and b.compid(+) = p_CompId
                      and a.itemid = b.itemid(+)
                      and a.seq = aSeq  and a.billid = aBillId and a.billiditemno = aBillidItemno;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV032A 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 寫發票暫存明細檔 */

              function WriteToInv032(aId in varchar2, aSeq in number, aBillId in varchar2, aBillidItemno in number,
                aItemId in varchar2, aDesc in varchar2, aUnit in number, aTax in number, aSale in number,
                aTotal in number, aQuantity in integer) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv032 ( identifyid1, identifyid2, invid, billid, billiditemno,
                       seq, startdate, enddate, itemid, description, quantity, unitprice,
                       taxamount, totalamount, chargeen, upttime, upten, servicetype, saleamount, linktomis,
                       facisno, accountno )
                    select p_IdentifyID1, p_IdentifyID2, aId, a.billid, a.billiditemno,
                       aSeq, a.startdate, a.enddate, aItemId, aDesc, aQuantity, aUnit,
                       aTax, aTotal, a.chargeen, to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ),
                       p_User, a.servicetype, aSale, a.linktomis,
                       decode( a.itemidref, null, a.facisno, null ),
                       decode( a.itemidref, null, a.accountno, null )
                    from inv032A a
                   where a.invid = aId and a.billid = aBillId and a.billiditemno = aBillidItemno;
                   aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV032 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* ------------------------------------------------------------------- */

              /* 更新發票暫存主檔 */

              function UpdateInv031(aMainInvId in varchar2, aRecord in TAutoCreate ) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    update inv031
                       set maininvid = aMainInvId,
                           saleamount = aRecord.SaleAmount,
                           taxamount = aRecord.TaxAmount,
                           invamount = aRecord.InvAmount,
                           mainsaleamount = aRecord.MainSaleAmount,
                           maintaxamount = aRecord.MainTaxAmount,
                           maininvamount = aRecord.MainInvAmount
                     where invid = aRecord.InvId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE INV031 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* -------------------------------------------------------------------- */


              function UpdateInv032A(aId in varchar2, aOldSeq in number, aNewSeq in number) return boolean
              as
                aResult boolean;
              begin
                 aResult := false;
                 begin
                    update inv032a set invid = aId, seq = aNewSeq
                     where invid = 'X' and seq = aOldSeq;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE INV032A 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* -------------------------------------------------------------------- */

              function UpdateInv016(aSeq in varchar2) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    update inv016 set beassignedinvid = 'Y' where seq = aSeq;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE INV016 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* ------------------------------------------------------------------- */

              /* 更新客服系統, 回填發票資料 */

              function UpdateMisSystem(aId in varchar2, aDate in date, aCId in varchar2, billid in varchar2,
                  billiditemno in number, aInvPCode in varchar2, aInvPName in varchar2) return boolean
              as
                 aResult boolean;
                 aSql varchar2(1024);
                 aTableName varchar2(100);
                 aPreInvoice varchar2(1);
              begin
                 aResult := false;
                 aTableName := GetTableNameByHowToCreate;
                 aPreInvoice := '0';
                 if ( p_HowToCreate = 1 ) then aPreInvoice := '1'; end if;
                 aSql := ' UPDATE ' || p_MisDbOwner || '.' || aTableName ||
                         '    SET GUINO = ' || '''' || aId || '''' || ',' ||
                         '        PREINVOICE = ' || '''' || aPreInvoice || '''' || ',' ||
                         '        INVDATE = TO_DATE( ' || '''' || to_char( aDate, 'YYYYMMDD' ) || '''' || ', ''YYYYMMDD'' ), ' ||
                         '        INVOICETIME = SYSDATE, ' ||
                         '        INVPURPOSECODE = ' || '''' || aInvPCode || '''' || ',' ||
                         '        INVPURPOSENAME = ' || '''' || aInvPName || '''' ||
                         '  WHERE CUSTID = ' || '''' || aCId || '''' ||
                         '    AND BILLNO = ' || '''' || billid || '''' ||
                         '    AND ITEM = ' || '''' || to_char( billiditemno ) || '''';
                 begin
                    execute immediate aSql;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE ' || aTableName ||' 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 取發票合併項目 */

              function GetMergeItem(aId in varchar2, aDesc in out varchar2 ) return boolean
              as
                aResult  boolean;
              begin
                aResult := true;
                begin
                   select a.description into aDesc from inv005 a
                   where a.Identifyid1 = p_IdentifyID1 and a.Identifyid2 = p_IdentifyID2
                     and a.compid = p_CompId and a.itemid = aId;
                exception
                   when others then aDesc := null;
                end;
                return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 取設備序號跟扣帳帳號 */

              function GetMisInfo(aCId in varchar2, billid in varchar2, billiditemno in number,
                aAcctNo in out varchar2, aFacisNo in out varchar2) return boolean
              as
                aResult boolean;
                aSql varchar2(1024);
                aTableName varchar2(100);
                aTableName2 varchar2(100);
                aId number(1);
                aServiceType char(1);
              begin
                 aResult := false;
                 aTableName := GetTableNameByHowToCreate;
                 aTableName2 := 'SO002A' || GetDbLink;
                 aSql := ' SELECT A.ACCOUNTNO, A.FACISNO, A.SERVICETYPE, B.ID ' ||
                         '   FROM ' || p_MisDbOwner || '.' || aTableName || ' A, ' || p_MisDbOwner || '.' || aTableName2 || ' B ' ||
                         '  WHERE A.CUSTID = B.CUSTID(+)  ' ||
                         '    AND A.ACCOUNTNO = B.ACCOUNTNO(+) ' ||
                         '    AND A.CUSTID = ' || '''' || aCId || '''' ||
                         '    AND A.BILLNO = ' || '''' || billid || '''' ||
                         '    AND A.ITEM = ' || '''' || to_char( billiditemno ) || '''';
                 begin
                    execute immediate aSql into aAcctNo, aFacisNo, aServiceType, aId;
                    /* 不是 P 服務別, 不取網路電話門號 */
                    if ( nvl( aServiceType, 'X' ) <> 'P' ) then aFacisNo := null; end if;
                    /* 帳戶別不為1, 非信用卡帳戶, 不取 AccountNo */
                    if ( nvl( aId, 0 ) <> 1  ) then aAcctNo := null; end if;
                    /* 信用卡號碼, 只取後4碼 */
                    if ( Length( Nvl( aAcctNo, 'X' ) ) > 4 ) then
                      aAcctNo := substr( aAcctNo, length( aAcctNo ) - 3, 4 );
                    end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '取信用卡號及設備序號 ( ' || aTableName || ' ) 時發生錯誤, 客編  = ' || aCId || ', SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


							/* ------------------------------------------------------------------- */


							function GetMergeAmount(
							  aId in varchar2, aSeq in number, aTaxType in varchar2, aTaxRate in number,
							  aUnit in out number, aTax in out number,
							  aSale in out number, aTotal in out number,
							  aQuan in out integer, aCount in out integer) return boolean
							as
							  aResult boolean;
							begin
                 aResult := false;
                 begin
                    select sum( unitprice ), sum( taxamount ), sum( saleamount ), sum( totalamount ), sum( quantity ), count(1)
                      into aUnit, aTax, aSale, aTotal, aQuan, aCount
                      from inv032a
                     where invid = aId and seq = aSeq;
                    /* 若為應稅則重新計算 銷售額, 稅額, 因為若是合併過後, 金額會有些誤差 */
                    if ( aTaxType = 1 ) then
                       aSale := round( aTotal / ( 1 + ( aTaxRate / 100 ) ) );
                       aTax := ( aTotal - aSale );
                       /* 有合併過的, 重新將單價設定成跟 銷售額一致 */
                       if ( aCount > 1 ) then
                          aUnit := aSale;
                       end if;
                    end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '計算合併發票明細金額/數量時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
							end;
				/*-----------------------------------------------------------------*/
				/*5358 找出合併項目的最小的單據起始日與最大的單據迄日 回填至發票明細 */
				function UpdInv032Date(aId in varchar2, aSeq in number) return boolean
					   as
						 aResult boolean;
						 aSqlMin varchar2(1024);
						 aSqlMax varchar2(1024);
						 aSQL  varchar2(2048);
					   begin
							aResult := false;
							begin
								aSqlMin := '(Select Min(A.StartDate) From ' || p_MisDbOwner || '.inv032A A,' || p_MisDbOwner || '.Inv005 B ' ||
										' Where A.invid=''' || aId || ''' And A.Seq=' || aSeq || ' And A.ItemID=B.ItemId And B.Sign=''' || '+'''||
										' And b.identifyid1 =''' || p_IdentifyID1 || ''' And b.identifyid2 = ' || p_IdentifyID2 || 
										' And B.CompId=' || p_CompId || ' And A.ItemIdRef is Not Null)' ;
								
								aSqlMax := '(Select Max(A.EndDate) From ' || p_MisDbOwner || '.inv032A A,' || p_MisDbOwner || '.Inv005 B ' ||
										' Where A.invid=''' || aId || ''' And A.Seq=' || aSeq || ' And A.ItemID=B.ItemId And B.Sign=''' || '+'''||
										' And b.identifyid1 =''' || p_IdentifyID1 || ''' And b.identifyid2 = ' || p_IdentifyID2 || 
										' And B.CompId=' || p_CompId || ' And A.ItemIdRef is Not Null)';
										
								aSQL := 'Update ' || p_MisDBOwner || '.inv032' ||
									 	' set StartDate='|| aSqlMin ||',' ||
										' EndDate='|| aSqlMax ||
										' Where seq ='||aSeq|| ' And invid='''||aid||''''||
										' And identifyid1='''||p_IdentifyID1||''''||
										' And identifyid2='||p_IdentifyID2;
							  execute immediate aSQL;
							  
												 
							  aResult := true;
							exception
							  when others then
							  begin
								 p_RetMsg := '填入發票明起迄日期發生錯誤, SQLCODE = ' || SQLCODE || ',SQLERRM = ' || SQLERRM;
							  end;
							end;
							return aResult;
					   end;	 
							/* ------------------------------------------------------------------- */


begin


    if ( p_CompId is null ) then
       p_RetMsg := '參數錯誤: 公司別不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_PrefixString is null ) then
       p_RetMsg := '參數錯誤: 指定開立發票本不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_ChargeStartdate is null ) or ( p_ChargeStopDate is null ) then
       p_RetMsg := '參數錯誤: 開立發票指定收費日(起)(迄)不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_SystemID is null ) then
       p_RetMsg := '參數錯誤: 系統代碼(SystemID)不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_InvDateEqualToChargeDate <> 1 ) and ( p_InvDate is null ) then
       p_RetMsg := '參數錯誤: 發票日期不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_LinkToMis = 'Y' ) and ( p_MisDbOwner is null ) then
       p_RetMsg := '參數錯誤: 已指定更新客服系統, 資料表擁有者不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;

    if ( p_IdentifyID1 is null ) or ( p_IdentifyID2 is null ) then
       p_RetMsg := '參數錯誤: 識別碼(IdentifyID)不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_OrderBy is null ) then
       p_RetMsg := '參數錯誤: 發票開立排序方式不可為空值。';
       p_RetCode := -1;
       return p_RetCode;
    end if;



    /* 先將傳進來的多組發票字軌拆解 */

    aTempStr := p_PrefixString;
    aInv099MaxBounds := 0;

    begin
       while ( aTempStr is not null )
       loop
          aInv099MaxBounds := ( aInv099MaxBounds + 1 );
	        aInv099( aInv099MaxBounds ).CompId := p_CompId;
	        aInv099( aInv099MaxBounds ).YearMonth := p_InvYearMonth;
	        aInv099( aInv099MaxBounds ).Prefix := substr( aTempStr, 1, 2 );
	        aInv099( aInv099MaxBounds ).StartNum := substr( aTempStr, 3, 8 );
	        aInv099( aInv099MaxBounds ).EndNum := null;
	        aInv099( aInv099MaxBounds ).CurNum := null;
	        aInv099( aInv099MaxBounds ).Useful := 'Y';
	        aInv099( aInv099MaxBounds ).LastInvDate := null;
	        aInv099( aInv099MaxBounds ).HasChange := false;
          aPos := instr( aTempStr, ',' );
          if ( aPos <= 0 ) then
            aTempStr := null;
          else
            aTempStr := substr( aTempStr, aPos + 1, length( aTempStr ) - aPos );
          end if;
       end loop;
    exception
      when others then
      begin
        p_RetCode := -99;
        p_RetMsg := '拆解多組發票字軌發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
        return p_RetCode;
      end;
    end;


    /* 取每一本發票本的資料 */

    for aIndex in 1..aInv099MaxBounds
    loop
       begin
          select a.curnum, a.endnum, a.lastinvdate
            into aInv099( aIndex ).CurNum, aInv099( aIndex ).EndNum, aInv099( aIndex ).LastInvDate
            from inv099 a
           where a.compid = aInv099( aIndex ).CompId
             and a.yearmonth = aInv099( aIndex ).YearMonth
             and a.prefix = aInv099( aIndex ).Prefix
             and a.startnum = aInv099( aIndex ).StartNum
             and a.useful = 'Y';
       exception
          when others then
          begin
             p_RetCode := -99;
             p_RetMsg := '讀取發票本時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
             return p_RetCode;
          end;
       end;
    end loop;


    /* 清空 Temp Table */

    begin
       dbms_utility.exec_ddl_statement( 'truncate table inv031' );
       dbms_utility.exec_ddl_statement( 'truncate table inv032' );
       dbms_utility.exec_ddl_statement( 'truncate table inv032a' );
    exception
       when others then
       begin
          p_RetMsg := '刪除 INV031/INV032/INV032A失敗, SQLCODE = ' || SQLCODE ||', SQLERRM = ' || SQLERRM;
          p_RetCode := -99;
          return p_RetCode;
       end;
    end;


    /* 取設定自動換開發票明細筆數設定 */
    if not GetInv001Param( aAutoCreateNum, aFaciCombine ) then return p_RetCode; end if;


    /* 本次處理時間, 前端程式抓 Log 檔用 */
    select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') into p_LogDateTime from dual;


    /* 開始使用的第一本發票本 */
    aInv099Index := 1;


    /* 抓取指定的發票主檔匯入資料, 先全部抓出來, 如果判斷後不須開立, 則直接寫 Log 進 INV033 */
    for c1 in (
              select * from inv016 a where a.compid = p_CompId
                   and a.chargedate between to_date( p_ChargeStartdate, 'YYYY/MM/DD' ) and to_date( p_ChargeStopDate, 'YYYY/MM/DD' )
                   and a.beassignedinvid = 'N' and a.isvalid = 'Y'
		               and a.howtocreate = p_howtocreate
		               ---and a.invamount > 0 and a.taxtype <> '0' and a.shouldbeassigned = 'Y'
                   and a.stopflag = 0  --- PS: stopflag = 1 的話, 表示被 user 指定刪除, 這裏不抓出來
                 order by
                   decode( p_OrderBy, 1, a.zipcode, to_char( a.chargedate, 'YYYYMMDD' ) ),
                   decode( p_OrderBy, 0, to_char( a.chargedate, 'YYYYMMDD' ), a.zipcode  ),
                   decode( p_OrderBy, 0, to_char( a.chargedate, 'YYYYMMDD' ), 1, a.zipcode, 2, a.zipcode, lpad( a.custid, 8, '0' ) )
               )
    loop

         aMasterCanCreate := true;
         aDetailCount := 0;
         aInv017TotalAmount := 0;

         /* 明細全抓除了被標示為刪除的之外 ( 因為對 user 來說, 此筆資料已被刪除 ), 逐一判斷, 不開立的話, 寫 Log  */

         for c2 in ( select * from inv017 where seq = c1.seq )
         loop

                aNotes := null;

                /* 判斷此筆是否須要開立( 主副檔 ) */

                /* 主檔就標示為不開立 */
                if ( c1.shouldbeassigned = 'N' ) then
                   aNotes := '指定為不須開立';
                /* 主檔的金額為 0, 或明細標示為開立, 但是明細該筆的總金額為零 ( 注意: 有可能為負項資料, 小於零為正常 ) */
                elsif ( c1.invamount <= 0 ) or ( ( c2.shouldbeassigned = 'Y' ) and ( c2.totalamount = 0 ) ) then
                   aNotes := '發票金額小於或等於零';
                /* 主檔的稅別為不須開立, 或明細標示為開立, 但是稅別為不開立 */
                elsif ( c1.taxtype = '0' ) or ( ( c2.shouldbeassigned = 'Y' ) and ( c2.taxtype = '0' ) ) then
                   aNotes := '發票稅別為不開立';
                end if;


                /* 寫 Log 後, 跳下筆 */
                if ( aNotes is not null ) then
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                     return p_RetCode;
                   end if;
                   aMasterCanCreate := false;
                   goto DoNothing;
                end if;


                /* 判斷此明細是否須要開立, ( 此筆明細不開立不代表全部都不開, 因為其它明細可能須要開立 ) */
                if ( c2.shouldbeassigned = 'N' ) then
                   aNotes := '指定為不須開立';
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                     return p_RetCode;
                   end if;
                   goto DoNothing;
                end if;



                /* 此部驟先檢查待開立明細是否可以開立發票, 只要其中一筆不符合條件, 整張就不予以開立 */
                /* 檢核時每一筆檢核完, 不可以因為其中一筆不過就直接跳過其它筆檢核, 因為每一筆都要 Log */

                /* 1. 是否須要檢核客服系統 */
                if ( p_LinkToMis = 'Y' ) and ( c2.linktomis = 'Y' ) then
                    /* 檢核客服系統資料 */
                    if not CheckMisSystem( c1.custid, c2.billid, c2.billiditemno, aNotes ) then
                       if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                         return p_RetCode;
                       end if;
                       aMasterCanCreate := false;
                       goto DoNothing;
                    end if;
                end if;


                /* 2. 檢核發票系統內是否已有此收費單號跟項次 */
                if ( CheckAlreadyCreate( c2.billid, c2.billiditemno, aNotes ) ) then
                    if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                      return p_RetCode;
                    end if;
                    aMasterCanCreate := false;
                    goto DoNothing;
                end if;


                /* 3. 判斷明細稅別是否與主檔不符 */
                if ( c1.taxtype <> c2.taxtype ) then
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, '收費明細的稅率別與主檔不同' ) then
                     return p_RetCode;
                   end if;
                   aMasterCanCreate := false;
                   goto DoNothing;
                end if;


                /* 4. 加總明細金額 */
                aInv017TotalAmount := nvl( aInv017TotalAmount, 0 ) + c2.totalamount;

                /* 明細確定可以開立的筆數 */

                aDetailCount := ( aDetailCount + 1 );

                << DoNothing >>

                   null;

         end loop;



         /* 明細結束, 如果判斷是可以開立的發票, 則再判斷主副檔的金額是否相符 */
         if ( aMasterCanCreate ) and ( aDetailCount > 0 ) and  ( c1.invamount <> aInv017TotalAmount ) then
            for c2 in ( select billid, billiditemno from inv017 where seq = c1.seq and shouldbeassigned = 'Y' )
            loop
              if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, '主副檔發票金額不符' ) then
                return p_RetCode;
              end if;
            end loop;
            aMasterCanCreate := false;
         end if;


         /* 確定可以開立此張發票 */
         if ( aMasterCanCreate ) and ( aDetailCount > 0 ) then

               /* 寫發票系統的客戶主檔 */
               if not CheckInv002Exists( c1.custid ) then
                  if not WriteToInv002( c1.seq ) then  return p_RetCode; end if;
               end if;


               /* 寫發票系統的客戶明細檔 */
               if not CheckInv019Exists( c1.custid  ) then
                  if not WriteToInv019( c1.seq ) then return p_RetCode; end if;
               end if;


               aDetailSeq := 0;

               /* 抓可以開立的明細, 先寫入 INV032A 準備做分類 */
               for c2 in ( select * from inv017 where seq = c1.seq and shouldbeassigned = 'Y'
                  and totalamount <> 0 and taxtype <> '0' order by billid, billiditemno )
               loop

                   /* 先指定為文字檔匯入的帳號/卡號  */
                   aAccountNo := c2.accountno;
                   /* 設備號碼並沒有從文字檔匯入 */
                   aFacisNo := null;

                   /* 若是與客服系統介接, 則取信用卡號及CP網路電話號碼(服務別為P) */
                   if ( p_LinkToMis = 'Y' ) and ( nvl( c2.linktomis, 'Y' ) = 'Y' ) then
                      if not GetMisInfo( c1.custid, c2.billid, c2.billiditemno, aAccountNo, aFacisNo ) then
                        return p_RetCode;
                      end if;

                   end if;

                   /* 寫明細檔 INV032A */
                   /*
                      注意: 此時寫到 INV032A 不可先取發票號碼, 後面須要做分類, 合併, 處理完後確定才可取發票號碼
                      流程: write to inv032a --> 分類,合併 --> inv032
                   */

                   if not WriteToInv032A( c2.seq, 'X', c2.billid, c2.billiditemno, aAccountNo, aFacisNo ) then
                     return p_RetCode;
                   end if;


               end loop;



               /* 合併發票明細  1.分類  2.編序號 (INV032A) */


               aDetailSeq := 0;
               aOldItemRefId := null;

               for c2 in (  select itemidref, facisno, count(1) from inv032a
                             where invid = 'X'
                               and itemidref is not null
                             group by itemidref, facisno
                          union all
                            select itemidref, facisno, 1 from inv032a
                             where invid = 'X'
                               and itemidref is null
                             order by itemidref, facisno   )
               loop


                    /* 1.有合併項目且有設備序號 */
                    if ( c2.itemidref is not null ) and ( c2.facisno is not null ) then


                        /* 參數設定為不同設備序號, 不須要合併 -> FaciCombine = 0 */
                        if ( aFaciCombine = 0 ) then
                           aDetailSeq := ( aDetailSeq + 1 );
                        /* 參數設定為不同設備序號, 須要合併  -> FaciCombine = 1 */
                        elsif ( nvl( aOldItemRefId, -1 ) <> c2.itemidref ) then
                           aDetailSeq := ( aDetailSeq + 1 );
                           aOldItemRefId := c2.itemidref;
                        end if;


                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref = c2.itemidref and facisno = c2.facisno
                           and seq is null;


                    /* 2. 有合併項目, 沒有設備序號 */
                    elsif ( c2.itemidref is not null ) and ( c2.facisno is null ) then

                        aDetailSeq := ( aDetailSeq + 1 );
                        aOldItemRefId := null;

                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref = c2.itemidref and facisno is null
                           and seq is null;

                    /* 3. 沒有合併項目 */
                    else

                        aDetailSeq := ( aDetailSeq + 1 );
                        aOldItemRefId := null;

                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref is null and seq is null and rownum <= 1;

                    end if;


               end loop;


               /* 分類完, 取發票號碼, 並寫入 INV032 */

               aDetailSeq := 0;              --- 真正寫入到發票明細的序號
               aAutoCreateMaxBounds := 0;    --- 自動換開記錄 List 的數量
               aOldSeq := null;              --- 比較基準序號

               for c2 in ( select * from inv032A where invid = 'X' order by seq )
               loop

                   if ( c2.seq <> nvl( aOldSeq, 0 ) ) then

                       /* 先判斷明細筆數是否已大於設定自動換開的筆數, 若是已大於換開筆數, 則重置明細序號 */
                       if ( aAutoCreateNum > 0 ) then
                         if ( aDetailSeq >= aAutoCreateNum ) then
                            aDetailSeq := 0;
                         end if;
                       end if;

                       aDetailSeq := ( aDetailSeq + 1 );

                       /* 明細第一筆, 取發票號碼 */
                       if ( aDetailSeq = 1 ) then

                           /* 取真正的發票日 */
                           if ( p_InvDateEqualToChargeDate = 1 ) then
                             aRealInvDate := c1.chargedate; ---發票日等於收費日
                           else
                             aRealInvDate := to_date( p_InvDate, 'YYYY/MM/DD' );  ---發票日等於傳入的發票日
                           end if;

                           /* 取發票號碼 */
                           if not GetInvNumber( aInv099Index, aInvId, aRealInvDate ) then return p_RetCode;  end if;

                           /* 計算發票檢查碼 */
                           aCheckNo := CalcInvCheckSum( substr( aInvId, 3, length( aInvId ) - 2 ),
                             p_SystemID || to_char( to_number( to_char( aRealInvDate, 'YYYYMMDD' ) ) - 19110000 ) );

                           /* 寫主檔, 順便取回該張發票的稅別及稅率, 合併明細重新計算稅額及銷售額時,會用到 */
                           if not WriteToInv031( c1.seq, aInvId, aCheckNo, aRealInvDate, aInvTaxType, aInvTaxRate ) then
                              return p_RetCode;
                           end if;

                           /* 將本次開立的發票號碼記錄下來, 金額重置 */
                           aAutoCreateMaxBounds := ( aAutoCreateMaxBounds + 1 );
                           aAutoCreateList( aAutoCreateMaxBounds ).InvId := aInvId;
                           aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).InvAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainSaleAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainTaxAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainInvAmount := 0;

                       end if;

                       /* 先更新發票號碼及正確的序號回 INV032A */
                       if not UpdateInv032A( aInvId, c2.seq, aDetailSeq ) then return p_RetCode; end if;

                       aItemDesc := null;

                       /* 合併參考號有值, 取參考號的名稱 */
                       /* 沒取到, 或是沒有設定參考號, 用匯檔進來的名稱 */
                       if ( c2.itemidref is not null ) then
                         if not GetMergeItem( c2.itemidref, aItemDesc ) then return p_RetCode; end if;
                       end if;

                       /* 取合併後該筆明細的金額, 必須用發票號碼 + 新的序號 */

                       aMergeUnitPrice := 0;
                       aMergeTaxAmount := 0;
                       aMergeSaleAmount := 0;
                       aMergeTotalAmount := 0;
                       aMergeQuantity := 0;
                       aMergeCount := 0;

                       if not GetMergeAmount( aInvId, aDetailSeq, aInvTaxType, aInvTaxRate,
                         aMergeUnitPrice, aMergeTaxAmount, aMergeSaleAmount,
                         aMergeTotalAmount, aMergeQuantity, aMergeCount ) then
                          return p_RetCode;
                       end if;


                       /* 如果合併的筆數超過 1 筆, 則寫入到 INV032 的 數量則為 1, 否則依照原本的數量 */
                       if ( aMergeCount > 1 ) then
                         aMergeQuantity := 1;
                       end if;


                       /* 將 INV032A 的合併項目回寫到 INV032 明細檔 */
                       if not WriteToInv032( aInvId, aDetailSeq, c2.billid, c2.billiditemno,
                           nvl( c2.itemidref, c2.itemid ), nvl( aItemDesc, c2.description ),
                           aMergeUnitPrice, aMergeTaxAmount, aMergeSaleAmount, aMergeTotalAmount, aMergeQuantity ) then
                         return p_RetCode;
                       end if;

                       aOldSeq := c2.seq;


                       /* 累計該張發票主檔下的明細金額 */
                       aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount :=
                         ( aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount + aMergeSaleAmount );
                       aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount :=
                         ( aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount + aMergeTaxAmount );
                       aAutoCreateList( aAutoCreateMaxBounds ).InvAmount :=
                        ( aAutoCreateList( aAutoCreateMaxBounds ).InvAmount + aMergeTotalAmount );


                       /* 先將主發票金額記錄在第一張開立的發票中*/
                       aAutoCreateList( 1 ).MainSaleAmount :=
                         ( aAutoCreateList( 1 ).MainSaleAmount + aMergeSaleAmount );
                       aAutoCreateList( 1 ).MainTaxAmount :=
                         ( aAutoCreateList( 1 ).MainTaxAmount + aMergeTaxAmount );
                       aAutoCreateList( 1 ).MainInvAmount :=
                        ( aAutoCreateList( 1 ).MainInvAmount + aMergeTotalAmount );


                   end if;


               end loop;
               for c2 in ( select A.seq from inv032A  A,inv005 B where A.invid = ainvId
                                and A.ItemIdRef is not null And A.ItemId=B.ItemId And B.Sign='+'
								and b.identifyid1=p_IdentifyID1 and b.identifyid2=p_IdentifyID2
								and B.CompId=p_CompId
								group by seq
                                order by seq )
                loop
					if not UpdInv032Date(ainvId,C2.Seq) then
						return p_RetCode;
					end if;
                   
                end loop;

               /* 每處理完一張主檔下的明細, 檢核是否有漏掉的 */

               select count(1) into aCount from inv032a where invid = 'X';

               if ( aCount > 0 ) then
                  p_RetMsg := '合併發票明細時發生錯誤, 尚有明細未完成合併。';
                  p_RetCode := -99;
                  return p_RetCode;
               end if;


               /* 將所有發票的主發票金額填入 */
               for aIndex in 1..aAutoCreateMaxBounds loop
                 aAutoCreateList( aIndex ).MainSaleAmount := aAutoCreateList( 1 ).MainSaleAmount;
                 aAutoCreateList( aIndex ).MainTaxAmount := aAutoCreateList( 1 ).MainTaxAmount;
                 aAutoCreateList( aIndex ).MainInvAmount := aAutoCreateList( 1 ).MainInvAmount;
               end loop;


               /* 回頭更新主檔的 maininvid, saleamount, taxamount, invamount */
               /* PS: MainInvId 回填最後一張發票號碼為主發票號碼, 若是只有一張發票等同於 invid = maininvid */
               for aIndex in 1..aAutoCreateMaxBounds loop
                      if not UpdateInv031( aInvId, aAutoCreateList( aIndex ) ) then
                      return p_RetCode;
                   end if;
               end loop;


               /* 更新發票來源檔 */
               if not UpdateInv016( c1.Seq ) then return p_RetCode; end if;

         end if;

    end loop;



    /* 寫入發票主檔 */
    begin
       insert into inv007 ( identifyid1, identifyid2, checkno, invid, canmodify, invdate, chargedate,
           compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat,
           taxtype, taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete,
           howtocreate, upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount,
           maininvamount, invuseid, invusedesc )
       select p_IdentifyID1, p_IdentifyID2, checkno, invid, canmodify, invdate, chargedate,
           compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat,
           taxtype, taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete,
           howtocreate, upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount,
           maininvamount, invuseid, invusedesc
         from inv031
        where identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2
          and compid = p_CompId;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '寫入發票主檔時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;


    /* 寫入發票明細檔 */
    begin
       insert into inv008 ( identifyid1, identifyid2, invid, billid, billiditemno, seq,
          startdate, enddate, itemid, description, quantity, unitprice,
          taxamount, totalamount, chargeen, upttime, upten, servicetype, saleamount, linktomis )
       select p_IdentifyID1, p_IdentifyID2, b.invid, b.billid, b.billiditemno, b.seq,
          b.startdate, b.enddate, b.itemid, b.description, b.quantity, b.unitprice,
          b.taxamount, b.totalamount, b.chargeen, b.upttime, b.upten, b.servicetype, b.saleamount, nvl( b.LinkToMis, 'Y' )
         from inv031 a, inv032 b
        where a.identifyid1 = b.identifyid1 and a.identifyid2 = b.identifyid2
          and a.invid = b.invid
          and a.identifyid1 = p_IdentifyID1 and a.identifyid2 = p_IdentifyID2
          and a.compid = p_CompId;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '寫入發票明細檔時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;



    /* 寫入發票合併檔, 只有合併的才寫過去 */
    begin
       insert into inv008A ( itemidref, invid, billid, billiditemno, seq,
          startdate, enddate, itemid, description, quantity, unitprice,
          taxamount, totalamount, servicetype, saleamount, facisno, accountno )
       select b.itemidref, b.invid, b.billid, b.billiditemno, b.seq,
          b.startdate, b.enddate, b.itemid, b.description, b.quantity, b.unitprice,
          b.taxamount, b.totalamount, b.servicetype, b.saleamount, b.facisno, b.accountno
         from inv031 a, inv032a b
        where a.invid = b.invid
          and a.identifyid1 = p_IdentifyID1 and a.identifyid2 = p_IdentifyID2
          and a.compid = p_CompId
          and b.itemidref is not null;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '寫入發票合併檔時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;



    /* 發票處理完畢, 更新客服系統, 抓 INV032A 的合併明細資料來更新 */
    if ( p_LinkToMis = 'Y' ) then
       for c1 in ( select a.invdate, a.invid, a.custid, a.invuseid, a.invusedesc, b.billid, b.billiditemno from inv031 a, inv032a b
          where a.invid = b.invid and nvl( b.linktomis, 'Y' ) = 'Y' and a.compid = p_CompId  order by b.invid, b.seq )
       loop
          if not UpdateMisSystem( c1.invid, c1.invdate, c1.custid, c1.billid, c1.billiditemno, c1.invuseid, c1.invusedesc ) then
             return p_RetCode;
          end if;
       end loop;
    end if;



    /* 更新發票本 */
    for aIndex in 1..aInv099MaxBounds
    loop
       if ( aInv099( aIndex ).HasChange ) then
          begin
             update inv099 set curnum = aInv099( aIndex ).CurNum, useful = aInv099( aIndex ).Useful,
                lastinvdate = aInv099( aIndex ).LastInvDate, upttime = sysdate, upten = p_User
              where identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2 and compid = p_CompId
                and yearmonth = aInv099( aIndex ).YearMonth and prefix = aInv099( aIndex ).Prefix
                and startnum = aInv099( aIndex ).startnum;
          exception
             when others then
             begin
                p_RetMsg := 'UPDTATE INV099 時發生錯誤, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                p_RetCode := -99;
                return p_RetCode;
             end;
          end;
       end if;
    end loop;

    p_RetCode := 0;
    p_RetMsg := '發票開立執行完畢';
    return p_RetCode;

end;
/

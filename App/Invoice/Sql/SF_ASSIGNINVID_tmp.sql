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
    p_RetCode                 out integer,
    p_RetMsg                  out varchar2,
    p_LogDateTime             out varchar2,
    p_ShowFaci                 in integer
  )
return number as

/*

  猔種ㄆ兜: 祇布 Schema  owner 璝琌籔 狝╰参 Schema  owner ぃ, 叭ゲ grant 舦倒祇布╰参 owner
            狝╰参┮斗 table         : grant select, update on SO033 to <祇布╰参owner>
                                         grant select, update on SO034 to <祇布╰参owner>
                                         grant select on so002a to <祇布╰参owner>


  祇布秨ミ
  郎: SF_AssignInvID.sql

  弧: 矪瞶祇布╰参ぇ祇布秨ミ(倒腹)

  把计 IN:

	p_User          		        varchar2 :  祘Α巨
	p_CompCode      		        number   :  そ絏
	p_LinkToMis                 varchar2 :  琌穝狝╰参   Y=>穝; N=>ぃ穝
	p_HowToCreate   		        number   :  1=>箇秨;   2=>秨
	p_DbLink                    varchar2 :  阁戈畐, 硓筁 DBLink よΑSQL
  p_InvDateEqualToChargeDate	number   :  1=>祇布ら戳单Μ禣ら戳;   2=>祇布ら戳ぃ单Μ禣ら戳
	p_InvDate			              varchar2 :  祇布ら戳
	p_InvYearMonth 			        varchar2 :  祇布る
	p_ChargeStartdate 	        varchar2 :  Μ禣癬﹍ら
	p_ChargeStopDate 		        varchar2 :  Μ禣篒ゎら
	p_IdentifyID1 			        number   :  醚絏(玂痙), 莱赣单1
	p_IdentifyID2 			        number   :  醚絏(玂痙), 莱赣单0
	p_SystemID 			            varchar2 :  System ID
  p_PrefixString 		          varchar2 :  祇布瓂
	p_OrderBy 			            number   :  逼よΑ : 0=>Μ禣ら戳
	                                                   1=>秎患跋腹
	                                                   2=>Μ禣ら戳+秎患跋腹
	                                                   3=>Μ禣ら戳+秎患跋腹+絪
  p_MisDbOwner                narchar2 :  恶MISTable Space  Owner



  把计 OUT:

	p_RetMsg        		       varchar2  :     挡狦癟
	p_RetCode       		       number    :     挡狦絏


  Return
	number: 矪瞶挡狦絏, 癸莱癟p_RetMsgい

         0:  p_RetMsg = '祇布秨ミ磅︽Ч拨'
        -1:  p_RetMsg = '把计岿粇'
       -99:  p_RetMsg = 'ㄤ岿粇'

*/

/*
    セ钵笆癘魁
    First release    : 2002.12.23
    Modify By Jackal : 2003.03.11  Ω度皐癸辈祇布秨ミ,эΘ辈祇布秨ミ
    Modify By Jackal : 2003.04.19  1.縵匡璝祇布肂 <0 ぃノ秨祇布 Log 戈 INV033
                                   2.浪琩戈琌秨筁祇布,璝秨筁ぃノ秨祇布 Log 戈 INV033
    Modify By Jackal : 2003.04.24  INV016 逆 ShouldBeAssigned , ノ耞Μ禣兜ヘ琌璶秨ミ祇布
    Modify By Jackal : 2003.07.22  肚把计p_MisDbOwner
    Modify By Jackal : 2004.01.16  盢跑计s_CustSNameパVarchar2(26)эVarchar2(50)
    Modify By Jackal : 2004.05.13  1.璝Inv016郎籔Inv017灿郎璸肂ぃ,玥ぃ秨ミLog  inv033
                                   2.浪琌Τseq,Μ禣灿祙瞯ぃ莱ぃ秨ミLog  inv033
                                   3.秨ミ恶璝赣掸戈ぃ so034┪so033 い,玥ぃ秨ミLog  inv033
                                   4.秨ミ恶璝赣掸戈ㄤguino Τ(秨ミ筁),玥ぃ秨ミLog
    Modify By Jackal : 2004.05.14  INV008のINV032綪扳肂逆(SaleAmount),綪扳肂=虫基*计秖(彼き俱计)
    Modify By Nick   : 2005.01.17  祇布浪琩絏璸衡ぃ癸, 祇布セ郎 INV099 STARTNUM, CURNUM, ENDNUM 逆эΘ VARCHAR2(8)
    Modify By Nick   : 2005.04.22  ъ秨ミ祇布戈兜耞, 祙 TAXTYPE = 0 , ぃぉ秨ミ
    Modify By Nick   : 2005.09.30  皐癸 EMC ㄏノ, ъ狝╰参戈糤 DBLink ﹃
    Modify By Nick   : 2005.10.04  糤把计, p_Dblink, p_linkToMis,
                                   璝 p_linkToMis = Y , 穝狝╰参戈,
                                   璝 p_Dblink Τ, 玥 sql  dblink ㄒ: select * from So033@catvc
    Modify By Nick   : 2005.10.20  タそ浪琩Μ禣兜ヘ滦, τ旧璓ぃ秨ミ祇布(耞 CompId )
    Modify By Nick   : 2006.01.16  INV033 糤 CompId そ逆, そ秨ミ祇布Ч, 钵盽厨穦ъ癸戈
    Modify By Nick   : 2006.01.19  1.穝э糶
                                   2.弄そ砞﹚把计 INV001.AutoCreateNum, 糤秨ミ耞,
                                     璝Τ砞﹚ ( p_AutoCreateNum > 0 ) Τ笆传秨祇布, 讽祇布灿掸计
                                     计ゲ斗笆传秨材2眎祇布
    Modify By Nick   : 2006.02.09  1.ㄏノ SQL Plus 絪亩 store function , 穦礚猭絪亩, э Bug.
                                   2.穝狝╰参 SO033, SO034 逆 ITEMNO 逆ゴ岿, 莱 ITEM
                                   3.祇布セ INV099  LastInvDate ⊿Τ穝, 穝砏玥岿粇, 莱–Ω秨ミ碞ゲ斗穝
    Modify By Nick   : 2006.02.14  1.INV008, INV017, INV032 糤 LinkToMis 逆, 琌穝┪浪狝╰参北赣掸
                                     灿, INV001  LinkToMis 琌羆秨闽
                                     Ex: INV001.LINKTOMIS = 'Y' AND INV008.LINKTOMIS = 'Y' 硂妓穦痷浪の穝
                                     狝╰参, 璝 INV001.LINKTOMIS <> 'Y' or INV008.LINKTOMIS <> 'Y' 玥ぃ浪の穝狝╰参
    Modify By Nick   : 2006.04.10  1.糤祇布灿Μ禣兜ヘㄖ
                                     穝糤 Table INV032A, INV008A
                                     穝糤逆 INV017.ACCOUNTNO ゅ郎蹲穦戈(CNS)
                                     秨ミ, 耞灿戈琌Τ斗璶ㄖΘ掸灿戈,
                                     Ex: Μ-ㄈびCM呼禣1M, 计秖:1, 肂:300 --|
                                         Μ-ㄈびCM呼禣1M, 计秖:2, 肂:600 --|--> ㄖΘ掸 Μ-ㄈびCM呼禣1M, 计秖:1, 肂: 1200
                                         Μ-ㄈびCM呼禣1M, 计秖:1, 肂:300 --|
                                     璝灿P狝叭, 玥临璶耞CP腹ㄖ, 璝ぃ腹玥璶ㄖ灿
                                   2.璝灿獺ノ煤禣, 玥獺腹腹, 祇布穦ノ

    Modify By Nick   : 2006.06.26  1.INV001糤把计, 北ぃ砞称腹, 祇布Μ禣灿琌斗璶ㄖΘΜ禣兜ヘ
                                     INV001.FACICOMBINE = 1 , 璝灿P狝叭, 玥ぃ腹妓ㄖ灿
                                   2.INV008 糤2逆, FACISNO, ACCOUNTNO 讽ぃㄖ兜ヘ, 2逆纗 CP 腹の獺ノ腹
                                   3.秨ミ祇布逼よΑ, 穝糤 Μ禣ら戳+秎患跋腹

    Modify By Nick   : 2006.08.28  1.秨ミ, 穝璸衡綪扳肂の祙肂, 綪扳肂=(羆肂/1.05),  祙肂=(羆肂-綪扳肂)
                                   2.┮ㄏノ穝璸衡筁綪扳肂の祙肂, 繷穝郎祇布肂, 硂穦セ狝┻筁ㄓ肂ぃ璓
                                     ぷㄤ琌ㄖ筁

    Modify By Nick   : 2006.11.10  1. 2006.08.28 セΤbug,讽笆传秨祇布,祇布羆肂,祇布綪扳肂,祇布祙肂逆赣眎肂
    Modify By Nick   : 2007.05.24  秨ミ祇布逼よΑ, 穝糤 Μ禣ら戳+秎患跋腹+絪
    Modify By Nick   : 2008.04.07  祇布秨ミ糤 1.祇布ノ硚絏 2.祇布ノ硚絏弧 (INV016, INV031, INV007)
    Modify By Nick   : 2008.08.12  祇布秨ミ, 糤糶狝╰参 SO033 の SO034 逆 INVPURPOSECODE, INVPURPOSENAME
	Modify By Kin     : 2009.12.11  拜肈栋5358 Τㄖ兜ヘ玥虫沮癬ùら戳INV008AMin(StartDate)籔Max(EndDate)
	Modify By Kin     : 2009.12.19  拜肈栋5358 狦砞﹚陪ボ砞称,璶盢灿戈砞称腹Keep癬ㄓ糶INV007.Memo1
	Modify By Kin     : 2009.12.25  拜肈栋5439 ㄖΜ禣戈ㄖ兜ヘ,ㄖЧㄖ祇布セōㄖ兜ヘ
	Modify By Kin     : 2010.04.09  拜肈栋5600 浪祇布戈ぃ璶﹃絪
	Modify By Kin     : 2010.07.12   拜肈栋5668 Head糤5逆 Detail糤SmartCardNo
*/



  /* 玡狠肚秈ㄓ舱祇布瓂 */

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
  aTempInv031     varchar2(2000);
  aTempInv032     varchar2(2000);
  aTempInv032A  varchar2(2000);
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
  aMemo1              inv007.memo1%type;
  aSmartCardNo    inv008a.SmartCardNo%type;
  aTmpSql varchar2(2048);
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

              /*  璸衡祇布浪琩絏 */

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

              /* ノ祇布腹絏, 抖獽穝祇布セノ篈 */

              function GetInvNumber(aIndex in out binary_integer, aId out varchar2, aDate in date) return boolean
              as
              begin
                aId := null;
                 << RE_GET >>
                   null;
                 if ( aIndex between 1 and aInv099MaxBounds ) then
                    /* 瞷硂セ祇布セ琌ノ */
                    if ( aInv099( aIndex ).CurNum <= ( aInv099( aIndex ).EndNum ) ) then
                       /* 祇布腹絏 */
                       aId := ( aInv099( aIndex ).prefix || Lpad( to_char( to_number( aInv099( aIndex ).CurNum ) ), 8, '0' ) );
                       aInv099( aIndex ).CurNum := lpad( to_char( to_number( aInv099( aIndex ).CurNum ) + 1 ), 8, '0' );
                       aInv099( aIndex ).LastInvDate := aDate;
                       aInv099( aIndex ).HasChange := true;
                    end if;
                    /* 耞琌ノЧ */
                    if ( aInv099( aIndex ).CurNum > ( aInv099( aIndex ).EndNum ) ) then
                      aInv099( aIndex ).Useful := 'N';
                      aInv099( aIndex ).LastInvDate := aDate;
                      aInv099( aIndex ).HasChange := true;
                      aIndex := ( aIndex + 1 );
                      /* ⊿Τ祇布腹絏, Ω */
                      if ( aId is null ) then goto RE_GET; end if;
                    end if;
                 end if;
                 if ( aId is null ) then
                   p_RetMsg := 'ノ祇布腹絏ア毖, 腹絏┪琌祇布セ腹絏ノЧ';
                   p_RetCode := -99;
                 end if;
                 return ( aId is not null );
              end;

              /* ------------------------------------------------------------------- */


              /*   DBLink ﹃ */

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
                       p_RetMsg := 'そ把计(笆传秨祇布の砞称腹琌ㄖ)祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 浪狝╰参掸祇布灿戈 */

              function CheckMisSystem(aId in varchar2, aBillNo in varchar2, aItem in number, aMsg out varchar2 ) return boolean
              as
                 aSql varchar2(255);
                 aTableName varchar2(100);
                 aGuiNo varchar2(10);
              begin
                 aMsg := null;
                 aTableName := GetTableNameByHowToCreate;
                 /*aSql := ' SELECT GUINO FROM ' || p_MisDbOwner || '.' || aTableName ||
                         '  WHERE CUSTID = ''' || aId || ''' AND BILLNO = ''' || aBillNo || ''' AND ITEM = ''' || aItem || '''';*/
				/* #5600 ぃ璶﹃絪 By Kin 2010/04/09 */
				aSql := ' SELECT GUINO FROM ' || p_MisDbOwner || '.' || aTableName ||
                         '  WHERE  BILLNO = ''' || aBillNo || ''' AND ITEM = ''' || aItem || '''';
                 begin
                   execute immediate aSql into aGuiNo;
                   if ( aGuiNo is not null ) then
                      aMsg := '狝╰参ず掸Μ禣虫腹Τ祇布腹絏, 狝╰参ず祇布腹絏 : ' || aGuiNo;
                   end if;
                 exception
                   when others then aMsg := '狝╰参礚掸癸莱戈';
                 end;
                 return ( aMsg is null );
              end;


              /* ------------------------------------------------------------------- */

              /* 浪祇布╰参ず琌ΤΜ禣虫腹蛤兜Ω ( 暗紀ぃ衡 ), 璝Τ, 玥р祇布腹絏癘ㄓ */

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
                      aMsg := 'Μ禣虫腹秨ミ祇布, 祇布腹絏:' || cCheck.invid;
                    else
                      aMsg := ( aMsg || ',' || cCheck.invid );
                    end if;
                 end loop;
                 return ( aMsg is not null );
              end;


              /* ------------------------------------------------------------------- */

              /* 祇布秨ミ浪ぃ斗秨ミ, ゲ斗 Log  INV033, 爹 */

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
                       p_RetMsg := 'INSERT INV033 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 浪め郎琌Τめ戈 */

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

              /* 浪め灿郎琌Τめ戈 */

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

              /* 糶祇布╰参め郎 */

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
                       p_RetMsg := 'INSERT INV002 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 糶祇布╰参め灿郎 */

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
                       p_RetMsg := 'INSERT INV019 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 糶祇布既郎 */

              function WriteToInv031(aSeq in varchar2, aId in varchar2, aNo in varchar2, aInvDate in date,
                aTaxType out varchar2, aTaxRate out number)
                 return boolean
              as
                 aResult boolean;
				 asql varchar2(2048);
              begin
                 aResult := false;
                 begin
					asql := 'INSERT INTO ' || aTempInv031 || ' (identifyid1, identifyid2, checkno, invid, canmodify, invdate, chargedate,' ||
						 ' compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat, taxtype,' ||
						 ' taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete, howtocreate,' ||
						 ' upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount,' ||
						 ' maininvamount, invuseid, invusedesc,InvoiceKind,Email,Contmobile,GiveUnitId,GiveUnitDesc )' ||
						 'SELECT ''' || p_IdentifyID1 ||''',' || p_IdentifyID2 ||',''' || aNo ||''','''|| aId || ''','''|| 'Y' ||''',' ||
						 'TO_CHAR(''' || aInvDate || '''),' ||
						 ' a.chargedate ,' || p_CompId ||',a.custid, a.chargetitle, ' ||
						 'a.title, a.zipcode, a.invaddr, a.mailaddr, a.businessid,'''|| '1' ||''', a.taxtype,' ||
						 'a.taxrate, 0, 0, 0, 0,''' || 'N' ||''', '''|| 'N' ||''',' || p_HowToCreate ||',' ||
						 'to_date(''' || p_LogDateTime || ''', ''' ||'YYYY/MM/DD HH24:MI:SS' ||''' ),'''||'' || p_User ||''',''' ||
						 '0' ||''' ,''' || aId ||''', 0, 0, 0,' ||
						 'a.invuseid, a.invusedesc,InvoiceKind,a.Email,a.Contmobile,a.GiveUnitId,a.GiveUnitDesc ' ||
						 ' FROM INV016 a ' ||
						 ' WHERE  a.seq = ' || aSeq ||
						 ' AND a.compid = ''' || p_CompId ||'''';
				 	aTaxType := null;
                    aTaxRate := 0;	  
				    execute immediate asql;
					asql := 'select taxtype,taxrate from ' || aTempInv031 ||
							' where identifyid1 = ''' || p_IdentifyID1 || '''' ||
							' and identifyid2 =  ' || p_IdentifyID2 ||
							' and invid = ''' || aId || '''';
					execute immediate asql into aTaxType,aTaxRate;
					if ( aTaxType = '1' ) and ( aTaxRate = 0 ) then aTaxRate := 5; end if;
					/*
                    insert into  inv031 ( identifyid1, identifyid2, checkno, invid, canmodify, invdate, chargedate,
                        compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat, taxtype,
                        taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete, howtocreate,
                        upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount, 
						maininvamount, invuseid, invusedesc,InvoiceKind,Email,Contmobile,GiveUnitId,GiveUnitDesc )
                    select p_IdentifyID1, p_IdentifyID2, aNo, aId, 'Y', aInvDate, a.chargedate,
                        p_CompId, a.custid, a.chargetitle, a.title, a.zipcode, a.invaddr, a.mailaddr, a.businessid, '1', a.taxtype,
                        a.taxrate, 0, 0, 0, 0, 'N', 'N', p_HowToCreate,
                        to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User, '0', aId, 0, 0, 0,
						a.invuseid, a.invusedesc,InvoiceKind,Email,Contmobile,GiveUnitId,GiveUnitDesc
                    from inv016 a
                   where a.seq = aSeq
                     and a.compid = p_CompId;
					 */
					 
                    /* ъ TaxType, TaxRate 肚 */
					/*
                    aTaxType := null;
                    aTaxRate := 0;
					
					
                    select taxtype, taxrate into aTaxType, aTaxRate from inv031
                     where identifyid1 = p_IdentifyID1
                       and identifyid2 = p_IdentifyID2
                       and invid = aId;
                    if ( aTaxType = '1' ) and ( aTaxRate = 0 ) then aTaxRate := 5; end if;
					*/
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV031 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 糶ㄖ祇布灿郎 */

              function WriteToInv032A(aSeq in varchar2, aId in varchar2, aBillId in varchar2, aBillidItemno in number,
                aAcctNo in varchar2, aFacisNo in varchar2,aSmartCardNo in varchar2)
                 return boolean
              as
                 aResult boolean;
				 aFind boolean;
				 asql varchar(2048);
              begin
                 aResult := false;
                 begin
					/*---------------------------------------------------------------------*/
					/* #5439盢INV017.ItemId穝ΘΜ禣戈ㄖ挡狦*/
					
					
					Update inv017 set itemid = decode(combcitemcode,null,itemid,combcitemcode),
					    	description=decode(combcitemcode,null,description,combcitemname),
							olditemid=itemid,olddescription=description
					where seq=aSeq and billid=aBillId and billiditemno=aBillidItemno;
					/*---------------------------------------------------------------------*/
				   
				   asql := 'insert into ' || aTempInv032a || '( itemidref, invid, billid, billiditemno, seq, startdate, enddate, ' ||
						'itemid, description, quantity, unitprice, taxamount, saleamount, totalamount,' ||
						'servicetype, chargeen, linktomis,olditemid,olddescription,combcitemcode,' ||
						'combcitemname,facisno, accountno,SmartCardNo )' ||
						'select b.itemidref,''' || aId || ''',' || '''' || aBillId || ''',' ||  aBillidItemno || ''', null, a.startdate, a.enddate,' ||
						'a.itemid, a.description, a.quantity, a.unitprice, a.taxamount, ( a.quantity * a.unitprice ), a.totalamount,' ||
						'a.servicetype, a.chargeen, nvl( a.linktomis, ''' || 'Y' ||''' ), a.olditemid,olddescription,combcitemcode, ' ||
						'combcitemname,''' || aFacisNo || ''',''' || aAcctNo || ''',''' || aSmartCardNo || '''' ||
						' from inv017 a,inv005 b ' ||
						' where b.IDENTIFYID1(+) = ''' || p_IdentifyID1 || '''' ||
						'  and b.IDENTIFYID2(+) = ' ||  p_IdentifyID2 ||
						' and b.compid(+) = ''' || p_CompId || '''' ||
						' and a.itemid =  b.itemid(+) ' ||
                        ' and a.seq = ' || aSeq || '  and a.billid = ''' || aBillId || ''' and a.billiditemno = ' || aBillidItemno;
					execute immediate asql;	
					asql := 'update ' || aTempInv032A || ' itemidref=(select combcitemcode from inv017 ' ||
										'where seq=' || aSeq || ' and billid=''' || aBillId || ''' and billiditemno= ' || aBillidItemno || ') ' ||
							' where seq is Null and billid=''' || aBillId || ''' and billiditemno=' || aBillIdItemno || ' and itemidref is null' ;
					execute immediate asql;
					asql :=	' update inv032A set itemid= olditemid ,description =  olddescription ' ||
					    'where seq is Null and billid='''|| aBillId || ' and billiditemno=' || aBillIdItemno;
					execute immediate asql;
				   
					/*
                    insert into inv032A ( itemidref, invid, billid, billiditemno, seq, startdate, enddate,
                       itemid, description, quantity, unitprice, taxamount, saleamount, totalamount,
                       servicetype, chargeen, linktomis,olditemid,olddescription,combcitemcode,
					   combcitemname,facisno, accountno,SmartCardNo )
                    select b.itemidref, aId, aBillId, aBillidItemno, null, a.startdate, a.enddate,
                       a.itemid, a.description, a.quantity, a.unitprice, a.taxamount, ( a.quantity * a.unitprice ), a.totalamount,
                       a.servicetype, a.chargeen, nvl( a.linktomis, 'Y' ), a.olditemid,olddescription,combcitemcode,
					   combcitemname,aFacisNo, aAcctNo,aSmartCardNo
                     from inv017 a, inv005 b
                    where b.IDENTIFYID1(+) = p_IdentifyID1
                      and b.IDENTIFYID2(+) = p_IdentifyID2
                      and b.compid(+) = p_CompId
                      and a.itemid = b.itemid(+)
                      and a.seq = aSeq  and a.billid = aBillId and a.billiditemno = aBillidItemno;
					  */
					  /*--------------------------------------------------------------------------------*/
					  /*Μ禣戈ㄖ兜ヘINV005тぃ,┮inv032A.itemidref钡穝ΘΜ禣戈ㄖ兜ヘ*/
					  /*
					 update inv032A set itemidref=(select combcitemcode from inv017
									where seq=aSeq and billid=aBillId and billiditemno=aBillidItemno) 
					   where seq is Null and billid=aBillId and billiditemno=aBillIdItemno and itemidref is null;
					   */
					  /*--------------------------------------------------------------------------------*/
						
						/*--------------------------------------------------------------------------------*/
					   /* #5439 盢セΜ禣兜ヘ穝INV032A*/
					   /*
					   update inv032A set itemid=olditemid,description=olddescription 
					    where seq is Null and billid=aBillId and billiditemno=aBillIdItemno;
					  */
						update inv017 set itemid=olditemid,description=olddescription
						where seq=aSeq and billid=aBillId and billiditemno=aBillidItemno;
					  /*--------------------------------------------------------------------------------*/
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV032A 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 糶祇布既灿郎 */

              function WriteToInv032(aId in varchar2, aSeq in number, aBillId in varchar2, aBillidItemno in number,
                aItemId in varchar2, aDesc in varchar2, aUnit in number, aTax in number, aSale in number,
                aTotal in number, aQuantity in integer) return boolean
              as
                 aResult boolean;
				 asql varchar2(2048);
              begin
                 aResult := false;
                 begin
				asql := 'insert into ' || aTempInv032 || '( identifyid1, identifyid2, invid, billid, billiditemno,' ||
						'seq, startdate, enddate, itemid, description, quantity, unitprice,' ||
                       'taxamount, totalamount, chargeen, upttime, upten, servicetype, saleamount, linktomis,' ||
                       'facisno, accountno,SmartCardNo )' ||
					   'select ''' || p_IdentifyID1 || ''',' || p_IdentifyID2 || ',''' || aId || ''',''' ||'a.billid' || ''', a.billiditemno  ,' ||
						'aSeq, a.startdate, a.enddate, aItemId, aDesc, aQuantity, aUnit,' ||
						aTax || ',' ||  aTotal || ', a.chargeen, to_date(''' || p_LogDateTime || ''','''|| 'YYYY/MM/DD HH24:MI:SS' || '''''),' ||
						'''' || p_User || ''', a.servicetype,' || aSale || ', a.linktomis,' ||
						'decode( a.itemidref, null, a.facisno, null ),' ||
                        'decode( a.itemidref, null, a.accountno, null ),' ||
					    'decode(a.itemidref,null,a.SmartCardNo,null)' ||
						' from ' || aTempInv032A || ' a' ||
						' where a.invid = ''' || aId || ''' and a.billid = ''' || aBillId || ''' and a.billiditemno = ' || aBillidItemno;
						execute immediate asql;
					/*
                    insert into inv032 ( identifyid1, identifyid2, invid, billid, billiditemno,
                       seq, startdate, enddate, itemid, description, quantity, unitprice,
                       taxamount, totalamount, chargeen, upttime, upten, servicetype, saleamount, linktomis,
                       facisno, accountno,SmartCardNo )
                    select p_IdentifyID1, p_IdentifyID2, aId, a.billid, a.billiditemno,
                       aSeq, a.startdate, a.enddate, aItemId, aDesc, aQuantity, aUnit,
                       aTax, aTotal, a.chargeen, to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ),
                       p_User, a.servicetype, aSale, a.linktomis,
                       decode( a.itemidref, null, a.facisno, null ),
                       decode( a.itemidref, null, a.accountno, null ),
					   decode(a.itemidref,null,a.SmartCardNo,null)
                    from inv032A a
                   where a.invid = aId and a.billid = aBillId and a.billiditemno = aBillidItemno;
				   */
				   
				   
                   aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV032 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* ------------------------------------------------------------------- */

              /* 穝祇布既郎 */

              function UpdateInv031(aMainInvId in varchar2, aRecord in TAutoCreate ) return boolean
              as
                 aResult boolean;
				 asql varchar2(2048);
              begin
                 aResult := false;
                 begin
					asql := 'update ' || aTempInv031 ||
						' set  maininvid = ''' || aMainInvId || ',' ||
						' saleamount = ' || aRecord.SaleAmount || ',' ||
                         ' taxamount = ' || aRecord.TaxAmount || ',' ||
                         ' invamount = ' || aRecord.InvAmount  || ',' ||
                         ' mainsaleamount = ' || aRecord.MainSaleAmount || ',' ||
                         ' maintaxamount =  ' || aRecord.MainTaxAmount || ',' ||
                          ' maininvamount = ' || aRecord.MainInvAmount ||
						  ' where invid = ''' || aRecord.InvId || '''';
                     /*
                    update inv031
                       set maininvid = aMainInvId,
                           saleamount = aRecord.SaleAmount,
                           taxamount = aRecord.TaxAmount,
                           invamount = aRecord.InvAmount,
                           mainsaleamount = aRecord.MainSaleAmount,
                           maintaxamount = aRecord.MainTaxAmount,
                           maininvamount = aRecord.MainInvAmount
                     where invid = aRecord.InvId;
					 */
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE INV031 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
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
                       p_RetMsg := 'UPDTATE INV032A 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
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
                       p_RetMsg := 'UPDTATE INV016 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* ------------------------------------------------------------------- */

              /* 穝狝╰参, 恶祇布戈 */

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
				 /*
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
				*/
				 aSql := ' UPDATE ' || p_MisDbOwner || '.' || aTableName ||
                         '    SET GUINO = ' || '''' || aId || '''' || ',' ||
                         '        PREINVOICE = ' || '''' || aPreInvoice || '''' || ',' ||
                         '        INVDATE = TO_DATE( ' || '''' || to_char( aDate, 'YYYYMMDD' ) || '''' || ', ''YYYYMMDD'' ), ' ||
                         '        INVOICETIME = SYSDATE, ' ||
                         '        INVPURPOSECODE = ' || '''' || aInvPCode || '''' || ',' ||
                         '        INVPURPOSENAME = ' || '''' || aInvPName || '''' ||
                         '  WHERE   BILLNO = ' || '''' || billid || '''' ||
                         '    AND ITEM = ' || '''' || to_char( billiditemno ) || '''';
                 begin
                    execute immediate aSql;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE ' || aTableName ||' 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* 祇布ㄖ兜ヘ */

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

              /* 砞称腹蛤Ι眀眀腹 */
			 /* #5668 糤醇紌腹 */
              function GetMisInfo(aCId in varchar2, billid in varchar2, billiditemno in number,
                aAcctNo in out varchar2, aFacisNo in out varchar2,aSmartCardNo in out varchar2) return boolean
              as
                aResult boolean;
                aSql varchar2(1024);
                aTableName varchar2(100);
                aTableName2 varchar2(100);
                aId number(1);
                aServiceType char(1);
				aCustId number(8);
              begin
                 aResult := false;
                 aTableName := GetTableNameByHowToCreate;
                 aTableName2 := 'SO002A' || GetDbLink;
                 aSql := ' SELECT A.ACCOUNTNO, A.FACISNO, A.SERVICETYPE, B.ID,A.CUSTID  ' ||
                         '   FROM ' || p_MisDbOwner || '.' || aTableName || ' A, ' || p_MisDbOwner || '.' || aTableName2 || ' B ' ||
                         '  WHERE A.CUSTID = B.CUSTID(+)  ' ||
                         '    AND A.ACCOUNTNO = B.ACCOUNTNO(+) ' ||
                         '    AND A.BILLNO = ' || '''' || billid || '''' ||
                         '    AND A.ITEM = ' || '''' || to_char( billiditemno ) || '''';
                 begin
                    execute immediate aSql into aAcctNo, aFacisNo, aServiceType, aId,aCustId;
					aTableName2 := 'SO004' || GetDbLink;
					
					
					if aFacisNo is not null then
						aSql := 'SELECT SmartCardNo FROM ' || p_MisDbOwner || '.' || aTableName2 ||
							' WHERE CUSTID = ' ||  aCustId || ' AND FACISNO = ' || '''' || aFacisNo || '''' ||
							' AND ROWNUM = 1';
						execute immediate aSql into aSmartCardNo;
					end if;
					
					
                    /* ぃ琌 P 狝叭, ぃ呼隔筿杠腹 */
					if ( p_showFaci = 0 ) then
                      if ( nvl( aServiceType, 'X' ) <> 'P' ) then aFacisNo := null; end if;
					  aSmartCardNo := null;
					end if;
                    /* 眀めぃ1, 獶獺ノ眀め, ぃ AccountNo */
                    if ( nvl( aId, 0 ) <> 1  ) then aAcctNo := null; end if;
                    /* 獺ノ腹絏, 4絏 */
                    if ( Length( Nvl( aAcctNo, 'X' ) ) > 4 ) then
                      aAcctNo := substr( aAcctNo, length( aAcctNo ) - 3, 4 );
                    end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '獺ノ腹の砞称腹 ( ' || aTableName || ' ) 祇ネ岿粇, 絪  = ' || aCId || ', SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
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
                    /* 璝莱祙玥穝璸衡 綪扳肂, 祙肂, 璝琌ㄖ筁, 肂穦Τㄇ粇畉 */
                    if ( aTaxType = 1 ) then
                       aSale := round( aTotal / ( 1 + ( aTaxRate / 100 ) ) );
                       aTax := ( aTotal - aSale );
                       /* Τㄖ筁, 穝盢虫基砞﹚Θ蛤 綪扳肂璓 */
                       if ( aCount > 1 ) then
                          aUnit := aSale;
                       end if;
                    end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '璸衡ㄖ祇布灿肂/计秖祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
							end;
							
				/*5668 тㄖ兜ヘㄤいSmartCardΤ戈 */
				function UpdInv032SmartCardNo(aId in varchar2,aSeq in number) return boolean
					as
						aResult boolean;
						aSQL varchar(2048);
					begin
						aResult := false;
						begin
							aSQL := 'UPDATE INV032 SET SMARTCARDNO = ' ||
								' (SELECT MAX(SMARTCARDNO) FROM INV032A ' ||
								' WHERE INVID ='''||  aId || ''' AND SEQ = ' || aSeq || ')' ||
								' Where seq ='||aSeq|| ' And invid='''||aid||''''||
								' And identifyid1='''||p_IdentifyID1||''''||
								' And identifyid2='||p_IdentifyID2;
								 execute immediate aSQL;
								 aResult := true;
						exception
							  when others then
							  begin
								 p_RetMsg := '恶祇布SMARTCARDNO岿粇 SQLCODE = ' || SQLCODE || ',SQLERRM = ' || SQLERRM;
							  end;		 
						end;
						return aResult;
					end;
				/*-----------------------------------------------------------------*/
				/*5358 тㄖ兜ヘ程虫沮癬﹍ら籔程虫沮ùら 恶祇布灿 */
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
								 p_RetMsg := '恶祇布癬ùら戳祇ネ岿粇, SQLCODE = ' || SQLCODE || ',SQLERRM = ' || SQLERRM;
							  end;
							end;
							return aResult;
					   end;	 
							/* ------------------------------------------------------------------- */

					function CreateTempTable  return varchar2
					as
						aTmpString varchar2(1024);
						aSQL varchar2(1024);
						aSEQ varchar2(1024);
					begin
						aSQL := 'SELECT TRIM(TO_CHAR( ' || p_MisDbOwner || '.S_TMPRPT_VIEWNAME.NEXTVAL,' || '0999999' ||'))' || ' FROM  DUAL';
						begin
							execute immediate aSQL into aSEQ;
							aTmpString := 'TMP_' || aSEQ;
							return aTmpString;
						exception
							when others then
							begin
								p_RetMsg := 'ミ既Tableア毖, SQLCODE = ' || SQLCODE || ',SQLERRM = ' || SQLERRM;
							end;
						end;
					end;
begin


    if ( p_CompId is null ) then
       p_RetMsg := '把计岿粇: そぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_PrefixString is null ) then
       p_RetMsg := '把计岿粇: ﹚秨ミ祇布セぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_ChargeStartdate is null ) or ( p_ChargeStopDate is null ) then
       p_RetMsg := '把计岿粇: 秨ミ祇布﹚Μ禣ら(癬)(ù)ぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_SystemID is null ) then
       p_RetMsg := '把计岿粇: ╰参絏(SystemID)ぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_InvDateEqualToChargeDate <> 1 ) and ( p_InvDate is null ) then
       p_RetMsg := '把计岿粇: 祇布ら戳ぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_LinkToMis = 'Y' ) and ( p_MisDbOwner is null ) then
       p_RetMsg := '把计岿粇: ﹚穝狝╰参, 戈局Τぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;

    if ( p_IdentifyID1 is null ) or ( p_IdentifyID2 is null ) then
       p_RetMsg := '把计岿粇: 醚絏(IdentifyID)ぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_OrderBy is null ) then
       p_RetMsg := '把计岿粇: 祇布秨ミ逼よΑぃ';
       p_RetCode := -1;
       return p_RetCode;
    end if;
	aTempInv031 := CreateTempTable;
	aTempInv032 := CreateTempTable;
	aTempInv032A :=CreateTempTable;


    /* 盢肚秈ㄓ舱祇布瓂╊秆 */

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
        p_RetMsg := '╊秆舱祇布瓂祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
        return p_RetCode;
      end;
    end;


    /* –セ祇布セ戈 */

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
             p_RetMsg := '弄祇布セ祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
             return p_RetCode;
          end;
       end;
    end loop;


    /* 睲 Temp Table */

    begin
       dbms_utility.exec_ddl_statement( 'truncate table inv031' );
       dbms_utility.exec_ddl_statement( 'truncate table inv032' );
       dbms_utility.exec_ddl_statement( 'truncate table inv032a' );
	  dbms_utility.exec_ddl_statement('CREATE TABLE ' || aTempInv031 || ' AS SELECT * FROM INV031 WHERE 1= 0 ');
	  dbms_utility.exec_ddl_statement('CREATE TABLE ' || aTempInv032 || ' AS SELECT * FROM INV032 WHERE 1= 0 ');
	  dbms_utility.exec_ddl_statement('CREATE TABLE ' || aTempInv032A || ' AS SELECT * FROM INV032A WHERE 1= 0 ');
	  
    exception
       when others then
       begin
          p_RetMsg := '埃 INV031/INV032/INV032Aア毖, SQLCODE = ' || SQLCODE ||', SQLERRM = ' || SQLERRM;
          p_RetCode := -99;
          return p_RetCode;
       end;
    end;


    /* 砞﹚笆传秨祇布灿掸计砞﹚ */
    if not GetInv001Param( aAutoCreateNum, aFaciCombine ) then return p_RetCode; end if;


    /* セΩ矪瞶丁, 玡狠祘Αъ Log 郎ノ */
    select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') into p_LogDateTime from dual;


    /* 秨﹍ㄏノ材セ祇布セ */
    aInv099Index := 1;


    /* ъ﹚祇布郎蹲戈, 场ъㄓ, 狦耞ぃ斗秨ミ, 玥钡糶 Log 秈 INV033 */
    for c1 in (
              select * from inv016 a where a.compid = p_CompId
                   and a.chargedate between to_date( p_ChargeStartdate, 'YYYY/MM/DD' ) and to_date( p_ChargeStopDate, 'YYYY/MM/DD' )
                   and a.beassignedinvid = 'N' and a.isvalid = 'Y'
		               and a.howtocreate = p_howtocreate
		               ---and a.invamount > 0 and a.taxtype <> '0' and a.shouldbeassigned = 'Y'
                   and a.stopflag = 0  --- PS: stopflag = 1 杠, ボ砆 user ﹚埃, 硂ぃъㄓ
                 order by
                   decode( p_OrderBy, 1, a.zipcode, to_char( a.chargedate, 'YYYYMMDD' ) ),
                   decode( p_OrderBy, 0, to_char( a.chargedate, 'YYYYMMDD' ), a.zipcode  ),
                   decode( p_OrderBy, 0, to_char( a.chargedate, 'YYYYMMDD' ), 1, a.zipcode, 2, a.zipcode, lpad( a.custid, 8, '0' ) )
               )
    loop

         aMasterCanCreate := true;
         aDetailCount := 0;
         aInv017TotalAmount := 0;

         /* 灿ъ埃砆夹ボ埃ぇ ( 癸 user ㄓ弧, 掸戈砆埃 ), 硋耞, ぃ秨ミ杠, 糶 Log  */

         for c2 in ( select * from inv017 where seq = c1.seq )
         loop

                aNotes := null;

                /* 耞掸琌斗璶秨ミ( 捌郎 ) */

                /* 郎碞夹ボぃ秨ミ */
                if ( c1.shouldbeassigned = 'N' ) then
                   aNotes := '﹚ぃ斗秨ミ';
                /* 郎肂 0, ┪灿夹ボ秨ミ, 琌灿赣掸羆肂箂 ( 猔種: Τ璽兜戈, 箂タ盽 ) */
                elsif ( c1.invamount <= 0 ) or ( ( c2.shouldbeassigned = 'Y' ) and ( c2.totalamount = 0 ) ) then
                   aNotes := '祇布肂┪单箂';
                /* 郎祙ぃ斗秨ミ, ┪灿夹ボ秨ミ, 琌祙ぃ秨ミ */
                elsif ( c1.taxtype = '0' ) or ( ( c2.shouldbeassigned = 'Y' ) and ( c2.taxtype = '0' ) ) then
                   aNotes := '祇布祙ぃ秨ミ';
                end if;


                /* 糶 Log , 铬掸 */
                if ( aNotes is not null ) then
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                     return p_RetCode;
                   end if;
                   aMasterCanCreate := false;
                   goto DoNothing;
                end if;


                /* 耞灿琌斗璶秨ミ, ( 掸灿ぃ秨ミぃ场常ぃ秨, ㄤウ灿斗璶秨ミ ) */
                if ( c2.shouldbeassigned = 'N' ) then
                   aNotes := '﹚ぃ斗秨ミ';
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                     return p_RetCode;
                   end if;
                   goto DoNothing;
                end if;



                /* 场艼浪琩秨ミ灿琌秨ミ祇布, 璶ㄤい掸ぃ才兵ン, 俱眎碞ぃぉ秨ミ */
                /* 浪–掸浪Ч, ぃㄤい掸ぃ筁碞钡铬筁ㄤウ掸浪, –掸常璶 Log */

                /* 1. 琌斗璶浪狝╰参 */
                if ( p_LinkToMis = 'Y' ) and ( c2.linktomis = 'Y' ) then
                    /* 浪狝╰参戈 */
                    if not CheckMisSystem( c1.custid, c2.billid, c2.billiditemno, aNotes ) then
                       if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                         return p_RetCode;
                       end if;
                       aMasterCanCreate := false;
                       goto DoNothing;
                    end if;
                end if;


                /* 2. 浪祇布╰参ず琌ΤΜ禣虫腹蛤兜Ω */
                if ( CheckAlreadyCreate( c2.billid, c2.billiditemno, aNotes ) ) then
                    if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                      return p_RetCode;
                    end if;
                    aMasterCanCreate := false;
                    goto DoNothing;
                end if;


                /* 3. 耞灿祙琌籔郎ぃ才 */
                if ( c1.taxtype <> c2.taxtype ) then
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, 'Μ禣灿祙瞯籔郎ぃ' ) then
                     return p_RetCode;
                   end if;
                   aMasterCanCreate := false;
                   goto DoNothing;
                end if;


                /* 4. 羆灿肂 */
                aInv017TotalAmount := nvl( aInv017TotalAmount, 0 ) + c2.totalamount;

                /* 灿絋﹚秨ミ掸计 */

                aDetailCount := ( aDetailCount + 1 );

                << DoNothing >>

                   null;

         end loop;



         /* 灿挡, 狦耞琌秨ミ祇布, 玥耞捌郎肂琌才 */
         if ( aMasterCanCreate ) and ( aDetailCount > 0 ) and  ( c1.invamount <> aInv017TotalAmount ) then
            for c2 in ( select billid, billiditemno from inv017 where seq = c1.seq and shouldbeassigned = 'Y' )
            loop
              if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, '捌郎祇布肂ぃ才' ) then
                return p_RetCode;
              end if;
            end loop;
            aMasterCanCreate := false;
         end if;


         /* 絋﹚秨ミ眎祇布 */
         if ( aMasterCanCreate ) and ( aDetailCount > 0 ) then

               /* 糶祇布╰参め郎 */
               if not CheckInv002Exists( c1.custid ) then
                  if not WriteToInv002( c1.seq ) then  return p_RetCode; end if;
               end if;


               /* 糶祇布╰参め灿郎 */
               if not CheckInv019Exists( c1.custid  ) then
                  if not WriteToInv019( c1.seq ) then return p_RetCode; end if;
               end if;


               aDetailSeq := 0;

               /* ъ秨ミ灿, 糶 INV032A 非称暗だ摸 */
               for c2 in ( select * from inv017 where seq = c1.seq and shouldbeassigned = 'Y'
                  and totalamount <> 0 and taxtype <> '0' order by billid, billiditemno )
               loop

                   /* ﹚ゅ郎蹲眀腹/腹  */
                   aAccountNo := c2.accountno;
                   /* 砞称腹絏⊿Τ眖ゅ郎蹲 */
                   aFacisNo := null;
				  aSmartCardNo :=null;
                   /* 璝琌籔狝╰参ざ钡, 玥獺ノ腹のCP呼隔筿杠腹絏(狝叭P) */
                   if ( p_LinkToMis = 'Y' ) and ( nvl( c2.linktomis, 'Y' ) = 'Y' ) then
                      if not GetMisInfo( c1.custid, c2.billid, c2.billiditemno, aAccountNo, aFacisNo,aSmartCardNo ) then
                        return p_RetCode;
                      end if;

                   end if;

                   /* 糶灿郎 INV032A */
                   /*
                      猔種: 糶 INV032A ぃ祇布腹絏, 斗璶暗だ摸, ㄖ, 矪瞶Ч絋﹚祇布腹絏
                      瑈祘: write to inv032a --> だ摸,ㄖ --> inv032
                   */

                   if not WriteToInv032A( c2.seq, 'X', c2.billid, c2.billiditemno, aAccountNo, aFacisNo,aSmartCardNo ) then
                     return p_RetCode;
                   end if;


               end loop;
			   
			   /* #5439 т ItemIdref 礚, itemid Τ才 ぃ掸 itemidref,硂璶ㄖΩ */
			   for c2 in ( select * from inv032A A where A.invid='X' and A.itemidref is null
			      and exists(select * from inv032A B Where B.invid='X' and A.itemid=B.itemidref
                  and B.itemidref is not null ) )
			   loop
			      update inv032A Set itemidref=(select distinct itemidref from inv032A 
				      where itemidref=c2.itemid   and invid='X' and seq is null)
				   where invid='X' and seq is null and itemid=c2.itemid;
			   end loop;
			   
               

               /* ㄖ祇布灿  1.だ摸  2.絪腹 (INV032A) */


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


                    /* 1.Τㄖ兜ヘΤ砞称腹 */
                    if ( c2.itemidref is not null ) and ( c2.facisno is not null ) then


                        /* 把计砞﹚ぃ砞称腹, ぃ斗璶ㄖ -> FaciCombine = 0 */
                        if ( aFaciCombine = 0 ) then
                           aDetailSeq := ( aDetailSeq + 1 );
                        /* 把计砞﹚ぃ砞称腹, 斗璶ㄖ  -> FaciCombine = 1 */
                        elsif ( nvl( aOldItemRefId, -1 ) <> c2.itemidref ) then
                           aDetailSeq := ( aDetailSeq + 1 );
                           aOldItemRefId := c2.itemidref;
                        end if;


                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref = c2.itemidref and facisno = c2.facisno
                           and seq is null;


                    /* 2. Τㄖ兜ヘ, ⊿Τ砞称腹 */
                    elsif ( c2.itemidref is not null ) and ( c2.facisno is null ) then

                        aDetailSeq := ( aDetailSeq + 1 );
                        aOldItemRefId := null;

                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref = c2.itemidref and facisno is null
                           and seq is null;

                    /* 3. ⊿Τㄖ兜ヘ */
                    else

                        aDetailSeq := ( aDetailSeq + 1 );
                        aOldItemRefId := null;

                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref is null and seq is null and rownum <= 1;

                    end if;


               end loop;


               /* だ摸Ч, 祇布腹絏, 糶 INV032 */

               aDetailSeq := 0;              --- 痷タ糶祇布灿腹
               aAutoCreateMaxBounds := 0;    --- 笆传秨癘魁 List 计秖
               aOldSeq := null;              --- ゑ耕膀非腹
               aMemo1 := null;
               for c2 in ( select * from inv032A where invid = 'X' order by seq )
               loop
			       /*------------------------------------------------------------------ */
			       /* 拜肈栋5358 狦砞﹚陪ボ砞称,玥璶盢砞称腹糶inv008.Memo1 By Kin */
				   if ( p_ShowFaci = 1) then
				     if ( c2.facisno is not null ) then
				        if ( aMemo1 is null ) then
						  aMemo1 := '砞称:' || c2.facisno;
						else
						  if ( instr(aMemo1,c2.facisno) = 0 ) then
						    aMemo1 := aMemo1 || ',' || c2.facisno;
						  end if;
						end if;
					 end if;
				   end if;
				   /*------------------------------------------------------------------ */
				   
                   if ( c2.seq <> nvl( aOldSeq, 0 ) ) then

                       /* 耞灿掸计琌砞﹚笆传秨掸计, 璝琌传秨掸计, 玥竚灿腹 */
                       if ( aAutoCreateNum > 0 ) then
                         if ( aDetailSeq >= aAutoCreateNum ) then
                            aDetailSeq := 0;
                         end if;
                       end if;

                       aDetailSeq := ( aDetailSeq + 1 );

                       /* 灿材掸, 祇布腹絏 */
                       if ( aDetailSeq = 1 ) then

                           /* 痷タ祇布ら */
                           if ( p_InvDateEqualToChargeDate = 1 ) then
                             aRealInvDate := c1.chargedate; ---祇布ら单Μ禣ら
                           else
                             aRealInvDate := to_date( p_InvDate, 'YYYY/MM/DD' );  ---祇布ら单肚祇布ら
                           end if;

                           /* 祇布腹絏 */
                           if not GetInvNumber( aInv099Index, aInvId, aRealInvDate ) then return p_RetCode;  end if;

                           /* 璸衡祇布浪琩絏 */
                           aCheckNo := CalcInvCheckSum( substr( aInvId, 3, length( aInvId ) - 2 ),
                             p_SystemID || to_char( to_number( to_char( aRealInvDate, 'YYYYMMDD' ) ) - 19110000 ) );

                           /* 糶郎, 抖獽赣眎祇布祙の祙瞯, ㄖ灿穝璸衡祙肂の綪扳肂,穦ノ */
                           if not WriteToInv031( c1.seq, aInvId, aCheckNo, aRealInvDate, aInvTaxType, aInvTaxRate ) then
                              return p_RetCode;
                           end if;

                           /* 盢セΩ秨ミ祇布腹絏癘魁ㄓ, 肂竚 */
                           aAutoCreateMaxBounds := ( aAutoCreateMaxBounds + 1 );
                           aAutoCreateList( aAutoCreateMaxBounds ).InvId := aInvId;
                           aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).InvAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainSaleAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainTaxAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainInvAmount := 0;

                       end if;

                       /* 穝祇布腹絏のタ絋腹 INV032A */
                       if not UpdateInv032A( aInvId, c2.seq, aDetailSeq ) then return p_RetCode; end if;

                       aItemDesc := null;

                       /* ㄖ把σ腹Τ, 把σ腹嘿 */
                       /* ⊿, ┪琌⊿Τ砞﹚把σ腹, ノ蹲郎秈ㄓ嘿 */
                       if ( c2.itemidref is not null ) then
                         if not GetMergeItem( c2.itemidref, aItemDesc ) then return p_RetCode; end if;
                       end if;

                       /* ㄖ赣掸灿肂, ゲ斗ノ祇布腹絏 + 穝腹 */

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


                       /* 狦ㄖ掸计禬筁 1 掸, 玥糶 INV032  计秖玥 1, 玥ㄌ酚セ计秖 */
                       if ( aMergeCount > 1 ) then
                         aMergeQuantity := 1;
                       end if;


                       /* 盢 INV032A ㄖ兜ヘ糶 INV032 灿郎 */
					   /*狦Μ禣兜ヘΤㄖ嘿INV005тぃ杠,碞ㄖ兜ヘ嘿 */
					   
					   if (aItemDesc is  Null) then
					      if (c2.combcitemcode is not null) And (c2.combcitemName is not null) then
							 aItemDesc := c2.CombCitemName;
						  end if;
					   end if;
					   
                       if not WriteToInv032( aInvId, aDetailSeq, c2.billid, c2.billiditemno,
                           nvl( c2.itemidref, c2.itemid ),   nvl( aItemDesc, c2.description ),
                           aMergeUnitPrice, aMergeTaxAmount, aMergeSaleAmount, aMergeTotalAmount, aMergeQuantity ) then
                         return p_RetCode;
                       end if;

                       aOldSeq := c2.seq;


                       /* 仓璸赣眎祇布郎灿肂 */
                       aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount :=
                         ( aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount + aMergeSaleAmount );
                       aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount :=
                         ( aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount + aMergeTaxAmount );
                       aAutoCreateList( aAutoCreateMaxBounds ).InvAmount :=
                        ( aAutoCreateList( aAutoCreateMaxBounds ).InvAmount + aMergeTotalAmount );


                       /* 盢祇布肂癘魁材眎秨ミ祇布い*/
                       aAutoCreateList( 1 ).MainSaleAmount :=
                         ( aAutoCreateList( 1 ).MainSaleAmount + aMergeSaleAmount );
                       aAutoCreateList( 1 ).MainTaxAmount :=
                         ( aAutoCreateList( 1 ).MainTaxAmount + aMergeTaxAmount );
                       aAutoCreateList( 1 ).MainInvAmount :=
                        ( aAutoCreateList( 1 ).MainInvAmount + aMergeTotalAmount );


                   end if;


               end loop;
			   /* 拜肈栋5358 狦砞﹚陪ボ砞称,璶盢砞称腹糶 Inv007 */
			   if ( p_ShowFaci = 1  And aMemo1 is Not null ) then
			     begin
			       Update inv031 set Memo1 = aMemo1 
				   Where invid = ainvid
				     and identifyid1 = p_IdentifyID1
					 and identifyid2 = p_IdentifyID2;
			     exception
			     when others then
			       begin
				     p_RetCode := -99;
				     p_RetMsg := '糶祇布郎称爹祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
			         return p_RetCode;
				   end;
				 end;
			   end if;
			   
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
                   if not  UpdInv032SmartCardNo(ainvId,C2.Seq) then
						return p_RetCode;
					end if;
                end loop;

               /* –矪瞶Ч眎郎灿, 浪琌Τ簗奔 */

               select count(1) into aCount from inv032a where invid = 'X';

               if ( aCount > 0 ) then
                  p_RetMsg := 'ㄖ祇布灿祇ネ岿粇, ﹟Τ灿ゼЧΘㄖ';
                  p_RetCode := -99;
                  return p_RetCode;
               end if;


               /* 盢┮Τ祇布祇布肂恶 */
               for aIndex in 1..aAutoCreateMaxBounds loop
                 aAutoCreateList( aIndex ).MainSaleAmount := aAutoCreateList( 1 ).MainSaleAmount;
                 aAutoCreateList( aIndex ).MainTaxAmount := aAutoCreateList( 1 ).MainTaxAmount;
                 aAutoCreateList( aIndex ).MainInvAmount := aAutoCreateList( 1 ).MainInvAmount;
               end loop;


               /* 繷穝郎 maininvid, saleamount, taxamount, invamount */
               /* PS: MainInvId 恶程眎祇布腹絏祇布腹絏, 璝琌Τ眎祇布单 invid = maininvid */
               for aIndex in 1..aAutoCreateMaxBounds loop
                      if not UpdateInv031( aInvId, aAutoCreateList( aIndex ) ) then
                      return p_RetCode;
                   end if;
               end loop;


               /* 穝祇布ㄓ方郎 */
               if not UpdateInv016( c1.Seq ) then return p_RetCode; end if;

         end if;

    end loop;



    /* 糶祇布郎 */
	/* 拜肈栋5358 狦砞﹚陪ボ砞称璶盢砞称腹糶Memo1 */
	/* 拜肈栋5668 糤糶5逆 */
    begin
       insert into inv007 ( identifyid1, identifyid2, checkno, invid, canmodify, invdate, chargedate,
           compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat,
           taxtype, taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete,
           howtocreate, upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount,
           maininvamount, invuseid, invusedesc,Memo1,InvoiceKind,Email,
		  Contmobile,GiveUnitId,GiveUnitDesc)
       select p_IdentifyID1, p_IdentifyID2, checkno, invid, canmodify, invdate, chargedate,
           compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat,
           taxtype, taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete,
           howtocreate, upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount,
           maininvamount, invuseid, invusedesc,Memo1,InvoiceKind,Email,
		   Contmobile,GiveUnitId,GiveUnitDesc
         from inv031
        where identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2
          and compid = p_CompId;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '糶祇布郎祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;


    /* 糶祇布灿郎 */
	/* 拜肈栋5358 狦砞﹚陪ボ砞称璶盢砞称腹糶Memo1 */
	/* 拜肈栋5668 糶SmartCardNo */
    begin
       insert into inv008 ( identifyid1, identifyid2, invid, billid, billiditemno, seq,
          startdate, enddate, itemid, description, quantity, unitprice,
          taxamount, totalamount, chargeen, upttime, upten, servicetype,
		  saleamount, linktomis, facisno ,SmartCardNo)
       select p_IdentifyID1, p_IdentifyID2, b.invid, b.billid, b.billiditemno, b.seq,
          b.startdate, b.enddate, b.itemid, b.description, b.quantity, b.unitprice,
          b.taxamount, b.totalamount, b.chargeen, b.upttime, b.upten, b.servicetype,
		  b.saleamount, nvl( b.LinkToMis, 'Y' ), b.facisno,b.SmartCardNo
         from inv031 a, inv032 b
        where a.identifyid1 = b.identifyid1 and a.identifyid2 = b.identifyid2
          and a.invid = b.invid
          and a.identifyid1 = p_IdentifyID1 and a.identifyid2 = p_IdentifyID2
          and a.compid = p_CompId;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '糶祇布灿郎祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;



    /* 糶祇布ㄖ郎, Τㄖ糶筁 */
	/* #5668 糤糶SmartCardNo */
    begin
       insert into inv008A ( itemidref, invid, billid, billiditemno, seq,
          startdate, enddate, itemid, description, quantity, unitprice,
          taxamount, totalamount, servicetype, saleamount, facisno, accountno,SmartCardNo )
       select b.itemidref, b.invid, b.billid, b.billiditemno, b.seq,
          b.startdate, b.enddate, b.itemid, b.description, b.quantity, b.unitprice,
          b.taxamount, b.totalamount, b.servicetype, b.saleamount, b.facisno, b.accountno,b.SmartCardNo
         from inv031 a, inv032a b
        where a.invid = b.invid
          and a.identifyid1 = p_IdentifyID1 and a.identifyid2 = p_IdentifyID2
          and a.compid = p_CompId
          and b.itemidref is not null;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '糶祇布ㄖ郎祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;



    /* 祇布矪瞶Ч拨, 穝狝╰参, ъ INV032A ㄖ灿戈ㄓ穝 */
    if ( p_LinkToMis = 'Y' ) then
		aTempSql := 'select a.invdate, a.invid, a.custid, a.invuseid, a.invusedesc, b.billid, b.billiditemno from ' || aTempInv031 || ' a,' || aTempinv032A || '  b' ||
			' where a.invid = b.invid and nvl( b.linktomis,' || '''' || 'Y' || ''' ) =' || '''' || 'Y' || ''' and a.compid = ' || '''' || p_CompId || '''' ||
			' order by b.invid, b.seq ';
       for c1 in ( aTempSql)
       loop
          if not UpdateMisSystem( c1.invid, c1.invdate, c1.custid, c1.billid, c1.billiditemno, c1.invuseid, c1.invusedesc ) then
             return p_RetCode;
          end if;
       end loop;
    end if;



    /* 穝祇布セ */
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
                p_RetMsg := 'UPDTATE INV099 祇ネ岿粇, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                p_RetCode := -99;
                return p_RetCode;
             end;
          end;
       end if;
    end loop;

    p_RetCode := 0;
    p_RetMsg := '祇布秨ミ磅︽Ч拨';
	dbms_utility.exec_ddl_statement( 'Drop Table ' || aTempInv031);
	dbms_utility.exec_ddl_statement( 'Drop Table ' || aTempInv032);
	dbms_utility.exec_ddl_statement( 'Drop Table ' || aTempInv032A);
    return p_RetCode;

end;
/

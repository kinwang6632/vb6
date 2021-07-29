CREATE OR REPLACE FUNCTION SF_ASSIGNINVID
  ( p_User                 VARCHAR2,
    p_CompCode             NUMBER,
    p_LinkToMis            VARCHAR2,
    p_DbLink               VARCHAR2,
    p_HowToCreate          NUMBER,
    p_InvDateEqualToChargeDate NUMBER,
    p_InvDate              VARCHAR2,
    p_InvYearMonth         VARCHAR2,
    p_ChargeStartdate      VARCHAR2,
    p_ChargeStopDate       VARCHAR2,
    p_IdentifyID1          NUMBER,
    p_IdentifyID2          NUMBER,
    p_SystemID             VARCHAR2,
    p_PrefixString         VARCHAR2,
    p_OrderBy              NUMBER,
    p_MisDbOwner           VARCHAR2,
    p_RetCode              OUT NUMBER,
    p_RetMsg               OUT VARCHAR2,
    p_LogDateTime          OUT VARCHAR2
  )
RETURN NUMBER AS

/*
  --@C:\Invoice\Script\SF_AssignInvID.sql;
  VARIABLE RetMsg VARCHAR2(4000)
  VARIABLE RetCode number
  VARIABLE ReturnNo NUMBER
  EXEC :ReturnNo := SF_AssignInvID('A',1,2,2,'2003/03/29','200303','2003/03/01','2003/03/31',1,0,'01','AA10000001,BB10000001,CC10000001',0 ,'v30',:RetCode ,:RetMsg)
  PRINT ReturnNo
  PRINT RetCode
  PRINT RetMsg
  --EXEC :ReturnNo := SF_AssignInvID('A',1,1,2,'2002/12/09','200211','2002/12/07','2002/12/08',1,0,'01','QW',0 ,'v30',:RetCode ,:RetMsg)


  發票開立
  檔名: SF_AssignInvID.sql

  說明:處理發票系統之發票開立(給號)


  IN
	p_User          		Varchar2 :    程式操作者
	p_CompCode      		number :      公司別代碼
	p_HowToCreate   		number :      1=>預開;   2=>後開
  p_InvDateEqualToChargeDate	number :  1=>發票日期等於收費日期;   2=>發票日期不等於收費日期
	p_InvDate			      Varchar2 :    發票日期
	p_InvYearMonth 			Varchar2 :    發票年月
	p_ChargeStartdate 	Varchar2 :    收費起始日
	p_ChargeStopDate 		Varchar2 :    收費截止日
	p_IdentifyID1 			number :      識別碼(保留), 應該等於1
	p_IdentifyID2 			number :      識別碼(保留), 應該等於0
	p_SystemID 			    Varchar2 :    System ID
  p_PrefixString 		  Varchar2 :    發票字軌
	p_OrderBy 			    number :      排序方式 : 0=>依收費日期開立 ; 1=>依郵遞區號開立
  p_MisDbOwner        Varchar2 :    回填至MIS的Table Space 的 Owner


  OUT
	p_RetMsg        		varchar2:     結果訊息 (至少200 bytes)
	p_RetCode       		number:       結果代碼


  Return
	number: 處理結果代碼, 對應訊息存於p_RetMsg中

         0:  p_RetMsg='發票開立執行完畢'
        -1:  p_RetMsg='參數錯誤'
       -99:  p_RetMsg='其他錯誤'


    By: Howard
    Date: 2002.12.23
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
    Modify By Nick   : 2005.04.22  1.抓取愈開立發票資料多加一項判斷, 稅別 TAXTYPE = 0 , 不予開立
    Modify By Nick   : 2005.09.30  針對 EMC 使用, 抓客服系統資料增加 DBLink 字串
    Modify By Nick   : 2005.10.04  增加入參數, p_Dblink, p_linkToMis,
                                   若 p_linkToMis = Y , 才更新客服系統資料,
                                   若 p_Dblink 有值, 則 sql 多加入 dblink 例如: select * from So033@catvc
    Modify By Nick   : 2005.10.20  修正因為檢查收費項目重覆而導致不能開立發票(多判斷CompCode)
    Modify By Nick   : 2006.01.16  INV033 增加 CompCode 公司別欄位, 多公司開立發票完後, 異常報表才會抓對資料

*/
  n_PreFixKind          NUMBER(2) :=0;  --前端傳來的字軌有幾種
  v_PrefixString        VARCHAR2(500);
  i                     NUMBER(2) :=0;
  v_PreFixStartNum      VARCHAR2(20);
  v_PreFix              VARCHAR2(2);
  v_StartNum            VARCHAR2(8);
  v_CurIndex            NUMBER(2) :=0; --目前是第幾組字軌


  TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  aPreFix Varchar2Ary;	-- array for PreFix
  aStartNum Varchar2Ary;	-- array for StartNum

  s_BeAssignedInvID     CHAR(1);

  s_RealInvID           VARCHAR2(10);
  aCheckNo              NUMBER(3);


  s_Useful              VARCHAR2(1);
  s_Where               VARCHAR2(4000);
  s_SQL                 VARCHAR2(4000);
  s_OrderBy             VARCHAR2(20);
  n_CurNum              NUMBER;
  n_EndNum              NUMBER;
  s_RealInvDate         VARCHAR2(10);
  s_CustSName           VARCHAR2(50);
  n_RecCount            NUMBER;
  s_Tmp1                VARCHAR2(10);
  s_Tmp2                VARCHAR2(10);
  s_CurDateTime         VARCHAR2(20);
  s_LastInvDate         VARCHAR2(10);
  s_SQLInv017           VARCHAR2(4000);

  s_SEQ                 VARCHAR2(15);
  s_CustId              VARCHAR2(8);
  s_Tel                 VARCHAR2(20);
  s_BusinessID          VARCHAR2(8);
  s_Title               VARCHAR2(50);
  s_InvAddr             VARCHAR2(60);
  s_MailAddr            VARCHAR2(60);
  s_ZipCode             VARCHAR2(5);
  s_ChargeDate          VARCHAR2(10);
  s_TaxType             VARCHAR2(1);
  n_TaxRate             NUMBER(5,2);
  n_SaleAmount          NUMBER(12);
  n_TaxAmount           NUMBER(12);
  n_InvAmount           NUMBER(12);
  s_ChargeTitle         VARCHAR2(50);
  n_Inv032SaleAmount    NUMBER(12);


  n_017Seq              NUMBER(2);
  s_017BillId           VARCHAR2(15);
  n_017BillIdItemNo     NUMBER(2);

  s_017TaxType          VARCHAR2(1);
  s_017ChargeDate       VARCHAR2(10);

  s_017ItemID           VARCHAR2(7);
  s_017Desc             VARCHAR2(40);

  n_017Quantity         NUMBER(6);
  n_017InitPrice        NUMBER(12,2);
  n_017TaxAmount        NUMBER(10);
  n_017TotalAmount      NUMBER(10);

  s_017StartDate        VARCHAR2(10);
  s_017EndDate          VARCHAR2(10);
  s_017ChargeEn         VARCHAR2(20);
  s_017ServiceType      VARCHAR2(1);

  s_Notes               VARCHAR2(255);
  s_RepeatInvID         VARCHAR2(10) := '';
  s_RepeatInvIDList     VARCHAR2(255) := '';

  s_TempSQL             VARCHAR2(4000);
  n_Check016InvAmount   NUMBER(12);
  n_Check017InvAmount   NUMBER(12);
  s_Check016TaxType     VARCHAR2(1);
  s_HaveSameTaxType     VARCHAR2(1);
  s_CheckTable          VARCHAR2(20);
  s_CheckSO033BillNo    VARCHAR2(20);
  s_CheckSO033GUINo     VARCHAR2(20);

  aDBlink               varchar2(10);


  type CurTyp is ref cursor;	--自訂cursor型態
  cDynamic CurTyp;          	--供dynamic SQL
  cDynamic2 CurTyp;          	--供dynamic SQL
  cDynamic3 CurTyp;          	--供dynamic SQL



          /* ------------------------------------------------------------------ */
          /*  計算發票檢查碼
          /* ------------------------------------------------------------------ */

          function CalcInvoiceCheckNumber(aInvId in varchar2, aInvParam in varchar2)
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

          function GetDBLink return varchar2
          as
            aResult varchar2(10);
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


BEGIN

      ---- Nick Add DBLink -----
      aDBlink := GetDBLink;


--檢查所有參數皆要有值 => 所有參數應該在呼叫端就檢查
  if p_CompCode Is Null then
    p_RetMsg := '參數錯誤:公司別不能為空值';
    p_RetCode := -1;
    Return -1;
  end if;

  if p_PrefixString Is Null then
    p_RetMsg := '參數錯誤:字軌不能為空值';
    p_RetCode := -1;
    Return -1;
  else
    --計算傳進來的字軌有幾組(AA11111111,BB22222222,CC33333333)
    --存成兩組陣列
    v_PrefixString := p_PrefixString;

    while v_PrefixString is not null loop
      n_PreFixKind := instr(v_PrefixString, ',');
      if n_PreFixKind > 0 then
  	    begin
          v_PreFixStartNum := ltrim(rtrim(substr(v_PrefixString, 1, n_PreFixKind)));

          aPreFix(i) := ltrim(rtrim(substr(v_PreFixStartNum, 1, 2)));
          aStartNum(i) := ltrim(rtrim(substr(v_PreFixStartNum, 3, 8)));

	        v_PrefixString := substr(v_PrefixString, n_PreFixKind+1);
          i:=i+1;
        exception
	      when others then
	        v_PrefixString := null;
	      end;
      else
        v_PreFixStartNum := ltrim(rtrim(v_PrefixString));

        aPreFix(i) := ltrim(rtrim(substr(v_PreFixStartNum, 1, 2)));
        aStartNum(i) := ltrim(rtrim(substr(v_PreFixStartNum, 3, 8)));
        v_PrefixString := null;
      end if;
    end loop;
  end if;



  if p_OrderBy = 0 then
    s_OrderBy := 'CHARGEDATE';  -- 收費日期
  else
    s_OrderBy := 'ZIPCODE' ; -- 郵遞區號
  end if;

  --取出第一組字軌出來

  v_CurIndex := 0;
  v_PreFix := aPreFix(v_CurIndex);
  v_StartNum := aStartNum(v_CurIndex);

  -- 取字軌資料
  begin
      SELECT TO_NUMBER( CURNUM ),
              TO_NUMBER( ENDNUM ),
              TO_CHAR( LASTINVDATE,'YYYY/MM/DD' )
         INTO n_CurNum,
               n_EndNum,
               s_LastInvDate
        FROM INV099
       WHERE COMPID = p_CompCode
         AND YEARMONTH = p_InvYearMonth
         AND PREFIX = v_Prefix
         AND USEFUL = 'Y'
         AND STARTNUM = v_StartNum;
  exception
    when others then
      ROLLBACK;
      p_RetMsg := '無法取出字軌資料:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
  end;

  -- 刪除 INV031/Inv032

  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE INV031');
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE INV032');
  exception
    when others then
      ROLLBACK;
      p_RetMsg := '刪除INV031/INV032失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
  end;


  -- select 出 INV016 之 data 且 InvAmount > 0
  s_Where := ' WHERE COMPID = ' || '''' || p_CompCode || '''' ||
             '   AND CHARGEDATE BETWEEN ' ||
             '       TO_DATE( ' || '''' || p_ChargeStartdate || '''' || ', ''YYYY/MM/DD'' ) ' ||
             '         AND    ' ||
             '       TO_DATE( ' || '''' || p_ChargeStopdate || '''' || ', ''YYYY/MM/DD'' )  ' ||
             '   AND BEASSIGNEDINVID = ''N''     ' ||
		         '   AND ISVALID = ''Y''             ' ||
		         '   AND HOWTOCREATE = ' || '''' || p_HowToCreate || '''' ||
             '   AND INVAMOUNT > 0 ' ||
             '   AND TAXTYPE <> ''0'' ' ||
             '   AND SHOULDBEASSIGNED = ''Y''    ';

  s_SQL :=  ' SELECT SEQ,                            ' ||
            '        CUSTID,                         ' ||
            '        TEL,                            ' ||
            '        BUSINESSID,                     ' ||
            '        TITLE, INVADDR, MAILADDR,       ' ||
            '        ZIPCODE,                        ' ||
            '        TO_CHAR( CHARGEDATE, ''YYYY/MM/DD'' ) CHARGEDATE,  ' ||
            '        TAXTYPE, TAXRATE, SALEAMOUNT,   ' ||
            '        TAXAMOUNT,                      ' ||
            '        INVAMOUNT,                      ' ||
            '        CHARGETITLE                     ' ||
            '   FROM INV016 ' || s_Where || ' ORDER BY ' || s_OrderBy;


  begin
    IF NOT cDynamic%ISOPEN then
      OPEN cDynamic FOR s_SQL;
    END IF;

  exception
    when others then
      ROLLBACK;
      p_RetMsg := '查詢INV016失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
  end;

  select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') into s_CurDateTime from dual;
  p_LogDateTime := s_CurDateTime;


  --Loop前先指定初值
  s_Useful := 'Y';
  --先指定第一個字軌的最後發票日期
  s_RealInvDate := s_LastInvDate;



--loop inv016,取每一筆資料來處理
  LOOP

    FETCH cDynamic INTO s_SEQ,s_CustId,s_Tel,s_BusinessID, s_Title, s_InvAddr, s_MailAddr, s_ZipCode,
                                    s_ChargeDate,s_TaxType,n_TaxRate,n_SaleAmount,n_TaxAmount,n_InvAmount, s_ChargeTitle;

    EXIT WHEN cDynamic%NOTFOUND;


    s_CustSName := s_ChargeTitle; -- 92/09/26 之新規格

    --用於主副檔金額檢核
    n_Check016InvAmount:= n_InvAmount;
    n_Check017InvAmount := 0;


    s_Check016TaxType := s_TaxType;


--------------------------------------------------------------------------------
-------down loop inv017,檢查每一筆資料是否已經開立發票
    s_SQLInv017 := 'SELECT BILLID,BILLIDITEMNO,TAXTYPE, '|| 'TO_CHAR(CHARGEDATE,' || '''' ||
                   'YYYY/MM/DD' || '''' || ') CHARGEDATE,' || 'ITEMID ' ||
		               ',DESCRIPTION,QUANTITY,UNITPRICE,TAXAMOUNT,TOTALAMOUNT,' ||
	               	'TO_CHAR(STARTDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') STARTDATE,' ||
		               'TO_CHAR(ENDDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') ENDDATE,' ||
		               'CHARGEEN,SERVICETYPE from inv017 where SEQ=' || '''' || s_SEQ || '''' ||
                   '   AND SHOULDBEASSIGNED = ''Y''    ';


    begin
      IF NOT cDynamic2%ISOPEN then
        OPEN cDynamic2 FOR s_SQLInv017;
      END IF;
    exception
    when others then
      ROLLBACK;
      CLOSE cDynamic2;
      p_RetMsg := '查詢INV017失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;


    --down loop inv017,檢查每一筆資料是否已經開立發票
    --所有的重複的發票號碼
    s_RepeatInvIDList := ' ';
    s_HaveSameTaxType := 'Y';
    loop

      FETCH cDynamic2 INTO s_017BillId,n_017BillIdItemNo,s_017TaxType, s_017ChargeDate,s_017ItemID,s_017Desc,
        		n_017Quantity, n_017InitPrice, n_017TaxAmount, n_017TotalAmount,
        		s_017StartDate, s_017EndDate, s_017ChargeEn, s_017ServiceType;

      EXIT WHEN cDynamic2%NOTFOUND;

      --用於檢核Inv016主檔與Inv017明細檔的合計金額是否一致
      n_Check017InvAmount := n_Check017InvAmount + n_017TotalAmount;



      --用於檢核是否有相同seq,但收費明細的稅率別不同應不能開立
      if s_Check016TaxType <> s_017TaxType then
        s_HaveSameTaxType := 'N';
      end if;



      --開立回填時若該筆資料不存在 so034或so033 中,或是其guino 有值時,
      --皆不能予以回填,並且不予以開立

      IF p_LinkToMis = 'Y' THEN


           if p_HowToCreate = '1' then
             s_CheckTable := 'SO033';
           else
             s_CheckTable := 'SO034';
           end if;

           s_CheckTable := ( s_CheckTable || aDBlink );

           s_CheckSO033BillNo := null;
           s_CheckSO033GUINo := null;


           BEGIN
             s_TempSQL := 'SELECT BILLNO,GUINO FROM ' || p_MisDbOwner || '.' || s_CheckTable ||' WHERE BILLNO=''' ||
                          s_017BillId || ''' AND ITEM ='||n_017BillIdItemNo;

             EXECUTE IMMEDIATE s_TempSQL INTO s_CheckSO033BillNo, s_CheckSO033GUINo;


             if ( s_CheckSO033GUINo IS NOT NULL ) then

                  begin
                    s_Notes := 'MIS此筆對應資料已有發票號碼';
                    insert into INV033( SEQ, BILLID, BILLIDITEMNO, TAXTYPE, CHARGEDATE, ITEMID,
                       DESCRIPTION, QUANTITY, UNITPRICE, TAXRATE, TAXAMOUNT,
                       TOTALAMOUNT, STARTDATE, ENDDATE, CHARGEEN,NOTES,
                       CUSTID, CUSTNAME, LOGTIME, UPTEN, COMPID )
                    values( s_SEQ, s_017BillId, n_017BillIdItemNo, s_017TaxType,
                       to_date(s_017ChargeDate, ' YYYY/MM/DD' ),
                       s_017ItemID, s_017Desc, n_017Quantity, n_017InitPrice, n_TaxRate,
                       n_017TaxAmount, n_017TotalAmount,
                       to_date( s_017StartDate, ' YYYY/MM/DD' ),
                       to_date( s_017EndDate, ' YYYY/MM/DD' ),
                       s_017ChargeEn, s_Notes, s_CustId, s_CustSName,
                       to_date( s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS' ), p_User, p_CompCode );
                  exception
                  when others then
                    close cDynamic2;
                    rollback;
                    p_RetMsg := 'INSERT INV033 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                    p_RetCode := -99;
                    return -99;
                  end;

               close cDynamic2;
               --開立回填時若該筆資料其guino 已有值時(代表已開立過),
               --則不能予以回填,並且不予以開立
               goto GO_NEXT;
             end if;

           exception
             when others then
             begin
                 s_Notes := 'MIS沒有此筆對應資料';
                 INSERT INTO INV033( SEQ, BILLID, BILLIDITEMNO, TAXTYPE, CHARGEDATE, ITEMID,
                    DESCRIPTION, QUANTITY, UNITPRICE, TAXRATE, TAXAMOUNT,
                    TOTALAMOUNT, STARTDATE, ENDDATE, CHARGEEN, NOTES,
                    CUSTID, CUSTNAME, LOGTIME, UPTEN, COMPID )
                VALUES(s_SEQ, s_017BillId, n_017BillIdItemNo, s_017TaxType,
                    to_date( s_017ChargeDate, ' YYYY/MM/DD' ),
                    s_017ItemID, s_017Desc, n_017Quantity, n_017InitPrice, n_TaxRate,
                    n_017TaxAmount, n_017TotalAmount,
                    to_date( s_017StartDate, ' YYYY/MM/DD' ),
                    to_date( s_017EndDate, ' YYYY/MM/DD' ),
                    s_017ChargeEn,s_Notes,s_CustId,s_CustSName,
                    to_date( s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS' ), p_User, p_CompCode );
                 close cDynamic2;
                 --開立回填時若該筆資料不存在 so034或so033 中,
                 --則不能予以回填,並且不予以開立
                 goto GO_NEXT;
               exception
               when others then
                 close cDynamic2;
                 rollback;
                 p_RetMsg := 'INSERT INV033 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                 p_RetCode := -99;
                 return -99;
               end;

               --開立回填時若該筆資料不存在 so034或so033 中,
               --則不能予以回填,並且不予以開立
               goto GO_NEXT;

           end;
      end if;

      --檢查是否有開立過且沒做廢的資料
      s_SQL := 'SELECT B.INVID FROM INV007 A,INV008 B WHERE B.BILLID ='|| '''' || s_017BillId || '''' ||
               ' AND B.BILLIDITEMNO=' || n_017BillIdItemNo || ' AND A.ISOBSOLETE=''N'' ' ||
               ' AND B.INVID=A.INVID AND A.COMPID=' || '''' || p_CompCode || '''' || ' ORDER BY B.INVID';




      begin
        IF NOT cDynamic3%ISOPEN then
          OPEN cDynamic3 FOR s_SQL;
        END IF;
      exception
      when others then
        ROLLBACK;
        CLOSE cDynamic3;
        p_RetMsg := '查詢INV008失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      end;



      loop

        FETCH cDynamic3 INTO s_RepeatInvID;
        EXIT WHEN cDynamic3%NOTFOUND;

        if s_RepeatInvIDList = ' ' then
          s_RepeatInvIDList := s_RepeatInvID;
        else
          s_RepeatInvIDList := s_RepeatInvIDList || ' , ' ||s_RepeatInvID;
        end if;
      end loop; --end loop inv008

      CLOSE cDynamic3;
    end loop;  --end loop inv017,檢查每一筆資料是否已經開立發票
    CLOSE cDynamic2;






    --若Inv016主檔與Inv017明細檔的合計金額不合,則 Log 至 inv033
    if n_Check016InvAmount<>n_Check017InvAmount then
      BEGIN
        s_Notes := '主副檔發票金額不符';
        INSERT INTO INV033(SEQ,BILLID,BILLIDITEMNO,TAXTYPE,CHARGEDATE,ITEMID,
           DESCRIPTION,QUANTITY,UNITPRICE,TAXRATE,TAXAMOUNT,
           TOTALAMOUNT,STARTDATE,ENDDATE,CHARGEEN,NOTES,
           CUSTID,CUSTNAME,LOGTIME,UPTEN, COMPID )
        VALUES( s_SEQ, ' ', 0,s_TaxType,
           to_date( s_ChargeDate, ' YYYY/MM/DD' ),
           ' ','', NULL, NULL, n_TaxRate,
           n_TaxAmount, n_Check016InvAmount,
           NULL, NULL, '', s_Notes, s_CustId, s_CustSName,
           to_date(s_CurDateTime, 'YYYY/MM/DD HH24:MI:SS'), p_User, p_CompCode );

      EXCEPTION
      WHEN OTHERS THEN
        CLOSE cDynamic2;
        ROLLBACK;
        p_RetMsg := 'INSERT INV033 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
        p_RetCode := -99;
        RETURN -99;
      END;


      --若Inv016主檔與Inv017明細檔的合計金額不合,
      --則 LOG  後直接換 INV016 下一筆
      goto GO_NEXT;
    end if;

   /* ---------------------------------------------------------------------------------- */



    ----用於檢核是否有相同seq,但收費明細的稅率別不同,則 Log 至 inv033
    if s_HaveSameTaxType = 'N' then
      BEGIN
        s_Notes := '收費明細的稅率別不同';
        INSERT INTO INV033(SEQ,BILLID,BILLIDITEMNO,TAXTYPE,CHARGEDATE,ITEMID,
                           DESCRIPTION,QUANTITY,UNITPRICE,TAXRATE,TAXAMOUNT,
                           TOTALAMOUNT,STARTDATE,ENDDATE,CHARGEEN,NOTES,
                           CUSTID,CUSTNAME,LOGTIME,UPTEN, COMPID )
        VALUES(s_SEQ,' ',0,s_TaxType,
               to_date(s_ChargeDate, ' YYYY/MM/DD'),
               ' ','',NULL, NULL, n_TaxRate,
               n_TaxAmount,n_Check016InvAmount,
               NULL,NULL,'',s_Notes,s_CustId,s_CustSName,
               to_date(s_CurDateTime, 'YYYY/MM/DD HH24:MI:SS'), p_User, p_CompCode );

      EXCEPTION
      WHEN OTHERS THEN
        CLOSE cDynamic2;
        ROLLBACK;
        p_RetMsg := 'INSERT INV033 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
        p_RetCode := -99;
        RETURN -99;
      END;


      --若Inv016主檔與Inv017明細檔的合計金額不合,
      --則 LOG  後直接換 INV016 下一筆
      goto GO_NEXT;
    end if;


--********************************************************************




    --若已開立發票則 Log 至 inv033
    if s_RepeatInvIDList <> ' ' then
      begin
        IF NOT cDynamic2%ISOPEN then
          OPEN cDynamic2 FOR s_SQLInv017;
        END IF;
      exception
      when others then
        ROLLBACK;
        CLOSE cDynamic2;
        p_RetMsg := '查詢INV017失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      end;


      loop
        FETCH cDynamic2 INTO s_017BillId,n_017BillIdItemNo,s_017TaxType, s_017ChargeDate,s_017ItemID,s_017Desc,
          		n_017Quantity, n_017InitPrice, n_017TaxAmount, n_017TotalAmount,
          		s_017StartDate, s_017EndDate, s_017ChargeEn;

        EXIT WHEN cDynamic2%NOTFOUND;

        --檢查是否有開立過且沒做廢的資料
        s_SQL := 'SELECT B.INVID FROM INV007 A,INV008 B WHERE B.BILLID ='|| '''' || s_017BillId || '''' ||
                 ' AND B.BILLIDITEMNO=' || n_017BillIdItemNo || ' AND ISOBSOLETE=''N'' ' ||
                 ' AND B.INVID=A.INVID AND A.COMPID=' || '''' || p_CompCode || '''' || ' ORDER BY B.INVID';

        begin
          IF NOT cDynamic3%ISOPEN then
            OPEN cDynamic3 FOR s_SQL;
          END IF;
        exception
        when others then
          ROLLBACK;
          CLOSE cDynamic3;
          p_RetMsg := '查詢INV008失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
          p_RetCode := -99;
          RETURN -99;
        end;

        s_RepeatInvIDList := ' ';

        loop
          FETCH cDynamic3 INTO s_RepeatInvID;
          EXIT WHEN cDynamic3%NOTFOUND;

          if s_RepeatInvIDList = ' ' then
            s_RepeatInvIDList := s_RepeatInvID;
          else
            s_RepeatInvIDList := s_RepeatInvIDList || ' , ' ||s_RepeatInvID;
          end if;
        end loop; --end loop inv008


        BEGIN
          s_Notes := '發票號碼:  ' || s_RepeatInvIDList;

          INSERT INTO INV033(SEQ,BILLID,BILLIDITEMNO,TAXTYPE,CHARGEDATE,ITEMID,
                             DESCRIPTION,QUANTITY,UNITPRICE,TAXRATE,TAXAMOUNT,
                             TOTALAMOUNT,STARTDATE,ENDDATE,CHARGEEN,NOTES,
                             CUSTID,CUSTNAME,LOGTIME,UPTEN, COMPID )
          VALUES(s_SEQ,s_017BillId,n_017BillIdItemNo,s_017TaxType,
                 to_date(s_017ChargeDate, ' YYYY/MM/DD'),
                 s_017ItemID,s_017Desc,n_017Quantity, n_017InitPrice, n_TaxRate,
                 n_017TaxAmount,n_017TotalAmount,
                 to_date(s_017StartDate, ' YYYY/MM/DD'),
                 to_date(s_017EndDate, ' YYYY/MM/DD'),
                 s_017ChargeEn,s_Notes,s_CustId,s_CustSName,
                 to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User, p_CompCode );


         EXCEPTION
         WHEN OTHERS THEN
           CLOSE cDynamic2;
           ROLLBACK;
           p_RetMsg := 'INSERT INV033 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
           p_RetCode := -99;
           RETURN -99;
         END;




        CLOSE cDynamic3;
      end loop;  --end loop inv017,檢查每一筆資料是否已經開立發票
      CLOSE cDynamic2;

      --若已開立過發票則 LOG  後直接換 INV016 下一筆
      goto GO_NEXT;

    end if;
--UP loop inv017,檢查每一筆資料是否已經開立發票
--------------------------------------------------------------------------------



    -- 設定真正的發票日期
    if p_InvDateEqualToChargeDate=1 then
      s_RealInvDate := s_ChargeDate; --發票日期等於收費日期
    else
      s_RealInvDate := p_InvDate; --發票日期等於傳入的發票日期
    end if;


    -- 檢查是否需要新增到 INV002(客戶主檔)
    begin
      SELECT COUNT(1) into n_RecCount FROM INV002 WHERE CUSTID=s_CustId AND COMPID =p_CompCode
             AND IDENTIFYID1=p_IdentifyID1 AND IDENTIFYID2= p_IdentifyID2;
    exception
    when others then
      ROLLBACK;
      p_RetMsg := '取出客戶主檔之數量時失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;


    if n_RecCount=0 then --表示此客戶資料不存在 INV002 中,所以要 insert 到 INV002
      begin
        INSERT INTO INV002(IDENTIFYID1,IDENTIFYID2,COMPID,CUSTID,CUSTSNAME,CUSTNAME,MZIPCODE,MAILADDR,
			         TEL1,ISSELFCREATED,UPTTIME,UPTEN)
        VALUES(p_IdentifyID1,p_IdentifyID2, p_CompCode,s_CustId,s_CustSName,s_Title,s_ZipCode,s_MailAddr,s_Tel,'N',
                         to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User);

      EXCEPTION
      WHEN OTHERS THEN
        ROLLBACK;
        p_RetMsg := 'INSERT INV002 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      END;

    end if;

    -- 檢查是否需要新增到 INV019(客戶明細檔)
    begin
      SELECT COUNT(*) COUNT into n_RecCount FROM INV019 WHERE CUSTID=s_CustId AND COMPID =p_CompCode
    AND IDENTIFYID1=p_IdentifyID1 AND IDENTIFYID2= p_IdentifyID2 AND TITLEID ='1';
    exception
    when others then
      ROLLBACK;
      p_RetMsg := '取出客戶明細檔之數量時失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;



    if n_RecCount=0 then --表示此客戶資料不存在 INV019 中,所以要 insert 到 INV019
      begin

        INSERT INTO INV019(IDENTIFYID1,IDENTIFYID2,COMPID,CUSTID,TITLEID,TITLESNAME,TITLENAME,BUSINESSID,MZIPCODE,
			         MAILADDR,INVADDR,UPTTIME,UPTEN)
        VALUES(p_IdentifyID1,p_IdentifyID2,p_CompCode,s_CustId,1,s_CustSName, s_Title,s_BusinessID,s_ZipCode,
		         	s_MailAddr, s_InvAddr,
                        to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User);
      EXCEPTION
      WHEN OTHERS THEN
        ROLLBACK;
        p_RetMsg := 'INSERT INV019 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      END;
    end if;


    -- 發票編號, 8位, 不足補 0
    s_Tmp1 := LPAD( TO_CHAR( n_CurNum ), 8, '0' );

    -- 將字串'2002/12/23' 轉換為 '911223'
    s_Tmp2 := p_SystemID || TO_CHAR( TO_NUMBER( REPLACE( s_RealInvDate, '/' ,'' ) ) - 19110000 );

    ---- 2005/01/11, Nick 修正, 發票應該為 10 碼, 不該是 LTRIM( TO_CHAR( n_CurNum ) )
    s_RealInvID := v_Prefix || LPAD( TO_CHAR( n_CurNum ), 8, '0' );

    --- 發票檢查碼
    aCheckNo := CalcInvoiceCheckNumber( s_Tmp1, s_Tmp2 );

  -- down,處理發票主暫存檔(INV031)
    begin
      INSERT INTO INV031(IDENTIFYID1,IDENTIFYID2,CHECKNO,INVID,UPTTIME,UPTEN,CANMODIFY,INVDATE,
	           CHARGEDATE,COMPID,CUSTID,CUSTSNAME,INVTITLE,ZIPCODE,INVADDR,
	           MAILADDR,BUSINESSID,INVFORMAT,TAXTYPE,TAXRATE,SALEAMOUNT,TAXAMOUNT,INVAMOUNT,
	           PRINTCOUNT,ISPAST,ISOBSOLETE,HOWTOCREATE,PRINTFUN,PRINTTIME)
      VALUES(p_IdentifyID1,p_IdentifyID2, aCheckNo, s_RealInvID,
	        	to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User,
        		'Y',to_date(s_RealInvDate,'YYYY/MM/DD'),to_date(s_ChargeDate,'YYYY/MM/DD'), p_CompCode,s_CustId,
         		s_CustSName,s_Title,s_ZipCode,s_InvAddr,s_MailAddr,s_BusinessID,'1',
        		s_TaxType,n_TaxRate,n_SaleAmount,n_TaxAmount,n_InvAmount,0,'Y','N',to_char(p_HowToCreate),0,
        		to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'));

      EXCEPTION
      WHEN OTHERS THEN
        ROLLBACK;
        p_RetMsg := 'INSERT INV031 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
        p_RetCode := -99;
        RETURN -99;
    END;
  -- up,處理發票主暫存檔(INV031)

  -- down,處理發票主暫存檔(INV032)

    s_SQL := 'SELECT BILLID,BILLIDITEMNO,TAXTYPE, '||
		'TO_CHAR(CHARGEDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') CHARGEDATE,' ||
		'ITEMID ' ||
		',DESCRIPTION,QUANTITY,UNITPRICE,TAXAMOUNT,TOTALAMOUNT,' ||
		'TO_CHAR(STARTDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') STARTDATE,' ||
		'TO_CHAR(ENDDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') ENDDATE,' ||
		'CHARGEEN,SERVICETYPE from inv017 where SEQ=' || '''' || s_SEQ || '''' ||
    '   AND SHOULDBEASSIGNED = ''Y''    ';


    begin
       IF NOT cDynamic2%ISOPEN then
       OPEN cDynamic2 FOR s_SQL;
    END IF;

    exception
    when others then
      ROLLBACK;
      p_RetMsg := '查詢INV017失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;


    --loop inv017,取每一筆資料來開立發票
    n_017Seq := 0;
    loop
     n_017Seq := n_017Seq + 1;
     FETCH cDynamic2 INTO s_017BillId,n_017BillIdItemNo,s_017TaxType, s_017ChargeDate,s_017ItemID,s_017Desc,
        	 n_017Quantity, n_017InitPrice, n_017TaxAmount, n_017TotalAmount,
        	 s_017StartDate, s_017EndDate, s_017ChargeEn, s_017ServiceType;

     EXIT WHEN cDynamic2%NOTFOUND;

     BEGIN
       n_Inv032SaleAmount := n_017Quantity * n_017InitPrice;

       -- down, insert 到 INV032
       INSERT INTO INV032(IDENTIFYID1,IDENTIFYID2,INVID,BILLIDITEMNO,SEQ,UPTTIME,UPTEN,BILLID,
          		STARTDATE,ENDDATE,ITEMID,DESCRIPTION,QUANTITY,UNITPRICE,TAXAMOUNT,TOTALAMOUNT,CHARGEEN,SERVICETYPE,SaleAmount)
                 VALUES(p_IdentifyID1,p_IdentifyID2,s_RealInvID,n_017BillIdItemNo,n_017Seq,
          		to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User,s_017BillId,
          		to_date(s_017StartDate, ' YYYY/MM/DD HH24:MI:SS'),
          		to_date(s_017EndDate, ' YYYY/MM/DD HH24:MI:SS'),
          		s_017ItemID,s_017Desc,n_017Quantity, n_017InitPrice,n_017TaxAmount,
          		n_017TotalAmount,s_017ChargeEn,s_017ServiceType,n_Inv032SaleAmount);
       -- up, insert 到 INV032

     EXCEPTION
     WHEN OTHERS THEN
       CLOSE cDynamic2;
       ROLLBACK;
       p_RetMsg := 'INSERT INV032 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
       p_RetCode := -99;
       RETURN -99;
     END;



     -- down, update Inv016
     begin
       s_BeAssignedInvID := 'Y';
       UPDATE INV016 SET BEASSIGNEDINVID = s_BeAssignedInvID WHERE SEQ=s_SEQ;
     exception
     when others then
       ROLLBACK;
       p_RetMsg := 'update INV016失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || 'SQL:' || s_SQL;
       p_RetCode := -99;
       RETURN -99;
     end;
     -- up, update Inv016


     -- down, update SO033 或 SO034

     if p_LinkToMis = 'Y' then

           if p_HowToCreate=1  then
             BEGIN
               s_TempSQL := '  UPDATE ' || p_MisDbOwner || '.SO033' || aDBLink ||
                            '    SET GUINO = ' || '''' || s_RealInvID || '''' || ', ' ||
                            '        PREINVOICE = 1, ' ||
                            '        INVDATE = TO_DATE( ' || '''' || s_RealInvDate || '''' || ', ''YYYY/MM/DD''), ' ||
                            '        INVOICETIME = SYSDATE ' ||
                            ' WHERE BILLNO = ' || '''' || s_017BillId || '''' ||
                            '   AND ITEM = ' || n_017BillIdItemNo;

               EXECUTE IMMEDIATE s_TempSQL;

             EXCEPTION
             WHEN OTHERS THEN
               ROLLBACK;
               p_RetMsg := 'UPDATE SO033 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_TempSQL;
               p_RetCode := -99;
               RETURN -99;
             END;
           else
             begin
               s_TempSQL := ' UPDATE ' || p_MisDbOwner || '.SO034' || aDBLink ||
                            '    SET GUINO = ' || '''' || s_RealInvID || '''' || ', ' ||
                            '        PREINVOICE = 0,' ||
                            '        INVDATE = TO_DATE( ' || '''' || s_RealInvDate || '''' || ', ''YYYY/MM/DD'' ), ' ||
                            '        INVOICETIME = SYSDATE ' ||
                            '  WHERE BILLNO = ' || '''' || s_017BillId || '''' ||
                            '    AND ITEM = ' || n_017BillIdItemNo;

               EXECUTE IMMEDIATE s_TempSQL;

             EXCEPTION
             WHEN OTHERS THEN
               ROLLBACK;
               p_RetMsg := 'UPDATE SO034 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_TempSQL;
               p_RetCode := -99;
               RETURN -99;
             END;
           end if;

     end if;

-- up, update SO033 或 SO034

    end loop;

    CLOSE cDynamic2;
  -- up,處理發票主暫存檔(INV032)

    n_CurNum := n_CurNum + 1;

    --每做完一筆檢查是否為該字軌最後一筆
    if n_EndNum < n_CurNum then

      --該字軌用完
      s_Useful := 'N';

      -- down, update Inv099
      begin
        UPDATE INV099
           SET CURNUM = LPAD( TO_CHAR( n_CurNum ), 8, '0' ),
               USEFUL = s_Useful,
               LASTINVDATE = TO_DATE( s_RealInvDate, 'YYYY/MM/DD' ),
    	         UPTTIME = TO_DATE( s_CurDateTime,'YYYY/MM/DD HH24:MI:SS' ),
    	         UPTEN = p_User
    	   WHERE COMPID = p_CompCode
    	     AND YEARMONTH = p_InvYearMonth
    	     AND PREFIX = v_Prefix
    	     AND STARTNUM = v_StartNum;
      exception
      when others then
        ROLLBACK;
        p_RetMsg := 'update INV099失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || 'SQL:' || s_SQL;
        p_RetCode := -99;
        RETURN -99;
      end;
      -- up, update Inv099

      --換下一字軌
      v_CurIndex := v_CurIndex + 1;
      v_PreFix := aPreFix(v_CurIndex);
      v_StartNum := aStartNum(v_CurIndex);

      -- 取字軌資料
      begin
          SELECT TO_NUMBER( CURNUM ),
                 TO_NUMBER( ENDNUM ),
                 TO_CHAR( LASTINVDATE, 'YYYY/MM/DD' )
            INTO n_CurNum,
                 n_EndNum,
                 s_LastInvDate
            FROM INV099
           WHERE COMPID = p_CompCode
             AND YEARMONTH = p_InvYearMonth
             AND PREFIX = v_Prefix
             AND USEFUL = 'Y'
             AND STARTNUM = v_StartNum;
      exception
        when others then
          ROLLBACK;
          p_RetMsg := '無法取出字軌資料:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
          p_RetCode := -99;
          RETURN -99;
      end;

      --換字軌時先指定原字軌的最後發票日期
      s_RealInvDate := s_LastInvDate;
    else
      s_Useful := 'Y';
    end if;

    --若有異常資料則 LOG  後直接換下一筆
    <<GO_NEXT>>
    NULL;

  END LOOP;

--Close dynamic cursor
  CLOSE cDynamic;








  begin
    INSERT INTO INV007 (SELECT * FROM INV031);
  exception
  when others then
    ROLLBACK;
    p_RetMsg := 'insert INV007失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
    p_RetCode := -99;
    RETURN -99;
  end;



  begin
    INSERT INTO INV008 (SELECT * FROM INV032);
  exception
  when others then
    ROLLBACK;
    p_RetMsg := 'insert INV008失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
    p_RetCode := -99;
    RETURN -99;
  end;


  -- down, update Inv099
  begin
    s_SQL :=  'AAA' || s_RealInvDate || 'BBB' || s_CurDateTime;

    UPDATE INV099
       SET CURNUM = LPAD( TO_CHAR( n_CurNum ), 8, '0' ),
           USEFUL=s_Useful,
           LASTINVDATE = TO_DATE( s_RealInvDate, 'YYYY/MM/DD' ),
	         UPTTIME = TO_DATE( s_CurDateTime, 'YYYY/MM/DD HH24:MI:SS' ),
	         UPTEN = p_User
	   WHERE COMPID = p_CompCode
	     AND YEARMONTH = p_InvYearMonth
	     AND PREFIX = v_Prefix
	     AND STARTNUM = v_StartNum;
  exception
  when others then
    ROLLBACK;
    p_RetMsg := 'update INV099失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || 'SQL:' || s_SQL;
    p_RetCode := -99;
    RETURN -99;
  end;
  -- up, update Inv099


  --- 將不開立的發票資料寫進 INV033 做 Log

  for c1 in (  select a.seq, a.custid, a.shouldbeassigned, a.invamount, a.taxtype
                 from inv016 a
                where a.compid = to_char( p_CompCode )
                  and a.chargedate between to_date( p_ChargeStartdate, 'YYYY/MM/DD' ) and to_date( p_ChargeStopdate, 'YYYY/MM//DD' )
                  and a.beassignedinvid = 'N'
                  and a.isvalid = 'Y'
                  and a.howtocreate = to_char( p_HowToCreate )
                  and ( a.shouldbeassigned = 'N' or a.invamount <= 0  or a.taxtype = '0' ) )
  loop

       s_CustSName := null;
       if ( p_LinkToMis = 'Y' ) then
           begin
             s_SQL := ' SELECT CUSTNAME FROM ' || p_MisDbOwner || '.SO001' || aDBLink ||' WHERE CUSTID = ' || c1.custid;
             execute immediate s_SQL into s_CustSName;
           exception
             when others then
             begin
                rollback;
                p_RetMsg := '未開立發票資料取客戶簡稱時有誤, 客戶代碼:' || c1.custid ||', SQLCODE = ' || SQLCODE||', SQLERRM = ' || SQLERRM;
                p_RetCode := -99;
                return -99;
             end;
           end;
       end if;

       if ( c1.shouldbeassigned = 'N' ) then
         s_Notes := '不須開立';
       elsif ( c1.invamount <= 0 ) then
         s_Notes := '金額小於或等於零';
       elsif ( c1.taxtype = '0' ) then
         s_Notes := '稅別為不開立';
       else
         s_Notes := '不須開立';
       end if;

       begin

            insert into inv033
                    (   seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        notes,
                        custid,
                        custname,
                        logtime,
                        upten,
                        compid     )
                 select seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        s_Notes,
                        c1.custid,
                        s_CustSName,
                        to_date(p_LogDateTime,'YYYY/MM/DD HH24:MI:SS'),
                        p_User,
                        p_CompCode
	               from inv017 where seq = c1.seq;

       exception
          when others then
          begin
            rollback;
            p_RetMsg := '未開立發票資料寫入 INV033 時發生錯誤: SQLCODE = ' || SQLCODE || ', SQLERRM = '||SQLERRM;
            p_RetCode := -99;
            return -99;
         end;
       end;

  end loop;

  COMMIT;


  --Dennis
  --- 將不開立的發票資料（明細部分）寫進 INV033 做 Log
  for c1 in (  select b.seq, a.custid, a.title, b.shouldbeassigned, b.totalamount, b.taxtype
                 from inv016 a, inv017 b
                where a.compid = to_char( p_CompCode )
                  and a.chargedate between to_date( p_ChargeStartdate, 'YYYY/MM/DD' ) and to_date( p_ChargeStopdate, 'YYYY/MM//DD' )
                  and a.isvalid = 'Y'
                  and a.howtocreate = to_char( p_HowToCreate )
                  and a.beassignedinvid = 'Y'
                  and a.shouldbeassigned = 'Y'
                  and b.shouldbeassigned = 'N'
                  and a.seq = b.seq
                  and ( b.shouldbeassigned = 'N' or b.totalamount <= 0  or b.taxtype = '0' ) )
  loop

       if ( c1.shouldbeassigned = 'N' ) then
         s_Notes := '不須開立';
       elsif ( c1.totalamount <= 0 ) then
         s_Notes := '金額小於或等於零';
       elsif ( c1.taxtype = '0' ) then
         s_Notes := '稅別為不開立';
       else
         s_Notes := '不須開立';
       end if;

       begin

            insert into inv033
                    (   seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        notes,
                        custid,
                        custname,
                        logtime,
                        upten,
                        compid     )
                 select seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        s_Notes,
                        c1.custid,
                        c1.title,
                        to_date(p_LogDateTime,'YYYY/MM/DD HH24:MI:SS'),
                        p_User,
                        p_CompCode
	               from inv017 where seq = c1.seq and shouldbeassigned = 'N';

       exception
          when others then
          begin
            rollback;
            p_RetMsg := '未開立發票資料寫入 INV033 時發生錯誤: SQLCODE = ' || SQLCODE || ', SQLERRM = '||SQLERRM;
            p_RetCode := -99;
            return -99;
         end;
       end;

  end loop;

  COMMIT;


  p_RetCode := 0;
  p_RetMsg := '發票開立執行完畢';

  RETURN 0;
EXCEPTION
   WHEN others THEN
     CLOSE cDynamic;
     ROLLBACK;
     p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
     p_RetCode := -99;
     RETURN -99;
END;
/

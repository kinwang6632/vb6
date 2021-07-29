CREATE OR REPLACE FUNCTION SF_CALCULATESO131
  ( p_CompCode varchar2,p_ServiceType varchar2,
    p_ComputeYM varchar2,p_StartDate varchar2,p_EndDate varchar2,
    p_ChargeItemSQL  varchar2,p_RealOrShouldDate number,
    p_RetCode OUT number,p_RetMsg OUT varchar2)
  RETURN  number AS
/*
--@D:\SF_CALCULATESO131;
VARIABLE ReturnNo NUMBER
VARIABLE SEQ NUMBER
VARIABLE MSG VARCHAR2(1000)
EXEC :ReturnNo := SF_CALCULATESO131('7','C','09201','20030101','20030131','91,92,93,94,95,96',0,:SEQ,:MSG)
PRINT ReturnNo
PRINT SEQ
PRINT MSG

  產生拆帳收費資料正項明細資料
  檔名: SF_CALCULATESO131.sql


  IN   p_CompCode          varchar2: 公司別代碼
       p_ServiceType       varchar2: 服務類別
       p_ComputeYM         varchar2: 計算年月
       p_StartDate         varchar2: 該計算年月的第一天
       p_EndDate           varchar2: 該計算年月的第後一天
       p_ChargeItemSQL     varchar2: 要查詢的收費項目別
       p_RealOrShouldDate  number: 以實收日期或應收日期為查詢對象 (0:實收金額  1:應收金額)


  OUT  p_SeqNo          number:   結果資料批號
       p_RetMsg         varchar2: 結果訊息 (至少200 bytes)

       Return           number:   處理結果代碼, 對應訊息存於p_RetMsg中
                           >=0: p_RetMsg='執行完畢, 處理筆數=<筆數>'
                            -1: p_RetMsg='參數錯誤'
                            -2: p_RetMsg='SELECT 時發生錯誤:
                            -3: p_RetMsg='INSERT 時發生錯誤:
                           -99: p_RetMsg='其他錯誤'

    By: Jackal   Date: 2003.04.09
    Modify By Jackal : 2003.04.24  計算退費資料原寫在前端,現改在後端寫
    Modify By Jackal : 2003.05.19  填至SO131時加填RefNo(1=呆帳,2=正常)
*/
  v_SQL            	varchar2(4000);		-- must long enough
  v_Column	            varchar2(1000);
  v_Where           	varchar2(1000);
  v_Table	varchar2(15);

  x_BillNo              varchar2(15);
  x_Item                Number;
  x_CustId              Number;
  x_CitemCode           Number;
  x_CitemName           varchar2(20);
  x_ShouldDate          Date;
  x_ShouldAmt           Number;
  x_RealDate            Date;
  x_RealAmt             Number;
  x_RealPeriod          Number;
  x_RealStopDate        Date;
  x_RealStartDate       Date;
  x_CompCode            Number;
  x_OrderNo             varchar2(15);
  x_SBillNo             varchar2(15);
  x_SItem               Number;

  x_MediaCode           Number;
  x_MediaName           varchar2(20);
  x_AcceptEn            varchar2(10);
  x_AcceptName          varchar2(20);
  x_IntroId             varchar2(10);
  x_IntroName           varchar2(50);
  x_STCode              Number;

  x_ReturnCitemCode       varchar2(10);
  x_ReturnCitemCodeString varchar2(200) := ' ';


  z_BillNo              varchar2(15);
  z_Item                Number;
  z_CitemCode           Number;
  z_CitemName           varchar2(20);
  z_MediaCode           Number;
  z_MediaName           varchar2(20);
  z_AcceptEn            varchar2(10);
  z_AcceptName          varchar2(20);
  z_IntroId             varchar2(10);
  z_IntroName           varchar2(50);
  z_ComputeYM           varchar2(5);
  z_SeqNo               Number;
  z_RealStopDate        Date;
  z_OrderNo             varchar2(15);


  --v_BeginMonth          Number;
  --v_EndMonth            Number;
  v_ServiceType         char(1) := UPPER(p_ServiceType);
  v_Notes               varchar2(255);
  v_TmpCitem            varchar2(255);
  v_CitemCodeCnt        Number;
  v_CurSeqNO            Number := 0;
  v_MaxSeqNO            Number := 0;
  v_IsSelectedCitemCode char(1) := 'Y';  --判斷是否為所選的收費項目
  v_RealStopDate        Date;
  v_RealStartDate       Date;

  --呆帳處理
  v_CheckSTCode         Number;
  v_RefNo               Number;


  c39                   CHAR(1) := CHR(39);

  cursor cc1 is
    SELECT CODENO,DESCRIPTION FROM CD019 WHERE  REFNO=9 AND SERVICETYPE=p_ServiceType;

  cursor cc2 is
    Select Max(SEQNO) MaxSeqNO from SO131 where ComputeYM=p_ComputeYM and
                                       RealOrShouldDate=p_RealOrShouldDate and
                                       CompCode=p_CompCode;
  cursor cc3 is
    select CodeNo FROM cd016 where refno=1;

  type CurTyp is ref cursor;	--自訂cursor型態
  cDynamic CurTyp;          	--供dynamic SQL


  I number;
  v_Index  number;

  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER;
  CitemCode NumberAry;

BEGIN
    if p_CompCode Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:公司別不能為Null值';
      Return -1;
    end if;

    if p_ServiceType Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:服務別不能為Null值';
      Return -1;
    end if;

    if p_ComputeYM Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:計算年月不能為Null值';
      Return -1;
    end if;

    if p_ChargeItemSQL Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:查詢哪些收費項目不能為Null值';
      Return -1;
    end if;

    if p_RealOrShouldDate Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:以應收或實收日期為查詢條件?';
      Return -1;
    end if;

    if p_StartDate Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:開始時間不能為Null值';
      Return -1;
    end if;

    if p_EndDate Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:結束時間不能為Null值';
      Return -1;
    end if;


    --先將SO131重複的資料刪除
    DELETE SO131 WHERE COMPUTEYM=p_ComputeYM AND RealOrShouldDate=p_RealOrShouldDate
                       AND COMPCODE=p_CompCode;

    --取出呆帳代碼
    for c3 in cc3 loop
      v_CheckSTCode := c3.CodeNo;
    end loop;


-----**********************************************************************-----
-----                               計算收費資料                           -----
-----**********************************************************************-----

    FOR i IN 1 .. 2 LOOP
    --1.組合SQL查詢字串:
      --須針對SO033及SO034查詢明細資料
      IF i = 1 THEN
        v_Table := 'SO033';
      ELSIF i = 2 THEN
        v_Table := 'SO034';
      END IF;


      v_Column := 'select BillNo,Item,CustId,CitemCode,CitemName,ShouldDate, '
                ||'NVL(ShouldAmt,0) ShouldAmt,RealDate,NVL(RealAmt,0) RealAmt, '
                ||'RealPeriod,RealStartDate,RealStopDate,CompCode,OrderNo, '
                ||'SBillNo,SItem,STCode  from '||v_Table;

      IF p_RealOrShouldDate = 0 THEN  --實收日期
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (RealDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| p_ChargeItemSQL||')';

      ELSIF p_RealOrShouldDate = 1 THEN  --應收日期
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (ShouldDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| p_ChargeItemSQL||')';

      END IF;

      --v_Where := v_Where || ' AND CancelCode IS NULL AND CancelName IS NULL';
      v_Where := v_Where || ' AND CancelFlag=0';

      v_SQL := v_Column || v_Where;


    --2.開啟Dynamic sql
      BEGIN

        IF NOT cDynamic%ISOPEN then
          OPEN cDynamic FOR v_SQL;
        END IF;

      EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetCode := -2;
          p_RetMsg := 'SELECT 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM||'      SQL='||v_SQL;
          RETURN -2;
      END;

    --3.loop取每一筆資料來處理
      LOOP
        FETCH cDynamic INTO x_BillNo,x_Item,x_CustId,x_CitemCode,x_CitemName,
                            x_ShouldDate,x_ShouldAmt,x_RealDate,x_RealAmt,x_RealPeriod,
                            x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                            x_SBillNo,x_SItem,x_STCode;
        EXIT WHEN cDynamic%NOTFOUND;

        --判定此筆資料是否為呆帳
        IF x_STCode = v_CheckSTCode THEN
          v_RefNo := 1;    --呆帳
        ELSE
          v_RefNo := 2;    --正常
        END IF;


        x_MediaCode := null;
        x_MediaName := '';
        x_AcceptEn := '';
        x_AcceptName := '';
        x_IntroId := '';
        x_IntroName := '';

        IF (x_OrderNo IS NOT NULL) OR (x_OrderNo <> '') THEN
          BEGIN
            -- Query SO105
            Select MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName
              into x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,x_IntroId,x_IntroName
              from So105 where OrderNo=x_OrderNo;


            v_Notes := '頻道收費';
          EXCEPTION
            WHEN others then
              v_Notes := '頻道收費,但訂單單號找不到對應的媒體介紹相關資料';
          END;
        ELSE
          v_Notes := '頻道收費,但訂單單號為空值,無法找到對應的媒體介紹相關資料';
        END IF;



        BEGIN
            v_CurSeqNO := v_CurSeqNO + 1;
            insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                   CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                   REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                   SBILLNO,SITEM,MEDIACODE,MEDIANAME,ACCEPTEN,ACCEPTNAME,
                   INTROID,INTRONAME,Notes,RefNo)
            Values (v_CurSeqNO,p_ComputeYM,p_RealOrShouldDate,x_BillNo,x_Item,x_CustId,
                   x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                   x_RealPeriod,x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                   x_SBillNo,x_SItem,x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,
                   x_IntroId,x_IntroName,v_Notes,v_RefNo);
        EXCEPTION
        WHEN OTHERS THEN
             ROLLBACK;
             p_RetCode := -3;
             p_RetMsg := 'INSERT時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
             RETURN -3;
        END;

      END LOOP;  --LOOP SO033及SO034的明細資料

    --4.Close dynamic cursor
      CLOSE cDynamic;

    END LOOP;    --LOOP 收費SO033及SO034

-----**********************************************************************-----
-----                               計算退費資料                           -----
-----**********************************************************************-----
    -- 取得所有的退費收費項目代碼
    for c1 in cc1 loop
      x_ReturnCitemCode := c1.CodeNo;
      if x_ReturnCitemCodeString = ' ' then
        x_ReturnCitemCodeString := x_ReturnCitemCode;
      else
        x_ReturnCitemCodeString := x_ReturnCitemCodeString || ',' || x_ReturnCitemCode;
      end if;
    end loop;

    --********* down  將CitemCode拆解成陣列 *********---
    if (p_ChargeItemSQL is not null) or (p_ChargeItemSQL <> '')then
      v_TmpCitem := p_ChargeItemSQL;
      v_Index := 1;
      I := 1;

      while v_TmpCitem is not null loop
        v_Index := instr(v_TmpCitem, ',');
        if v_Index > 0 then
  	begin
  	  CitemCode(I) := ltrim(rtrim(substr(v_TmpCitem, 1, v_Index-1)));
  	  v_TmpCitem := substr(v_TmpCitem, v_Index+1);
  	  I:=I+1;
  	exception
  	  when others then
  	    v_TmpCitem := null;
  	end;
        else
  	CitemCode(I) := to_number(rtrim(ltrim(v_TmpCitem)));
  	v_TmpCitem := null;
        end if;
      end loop;
      v_CitemCodeCnt := I;

    end if;

    FOR I IN 1 .. v_CitemCodeCnt LOOP
      v_TmpCitem := CitemCode(I);
      v_TmpCitem := v_TmpCitem;
    END LOOP;
    --********* up  將CitemCode拆解成陣列 *********---



    FOR i IN 1 .. 2 LOOP
    --1.組合SQL查詢字串:
      --須針對SO033及SO034查詢退費明細資料
      IF i = 1 THEN
        v_Table := 'SO033';
      ELSIF i = 2 THEN
        v_Table := 'SO034';
      END IF;


      v_Column := 'select BillNo,Item,CustId,CitemCode,CitemName,ShouldDate, '
                ||'NVL(ShouldAmt,0) ShouldAmt,RealDate,NVL(RealAmt,0) RealAmt, '
                ||'RealPeriod,RealStartDate,RealStopDate,CompCode,OrderNo, '
                ||'SBillNo,SItem,STCode  from '||v_Table;

      IF p_RealOrShouldDate = 0 THEN  --實收日期
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (RealDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| x_ReturnCitemCodeString ||')';

      ELSIF p_RealOrShouldDate = 1 THEN  --應收日期
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (ShouldDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| x_ReturnCitemCodeString ||')';

      END IF;

      --v_Where := v_Where || ' AND CancelCode IS NULL AND CancelName IS NULL';
      v_Where := v_Where || ' AND CancelFlag=0';

      v_SQL := v_Column || v_Where;


    --2.開啟Dynamic sql
      BEGIN

        IF NOT cDynamic%ISOPEN then
          OPEN cDynamic FOR v_SQL;
        END IF;

      EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetCode := -2;
          p_RetMsg := 'SELECT 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM||'      SQL='||v_SQL;
          RETURN -2;
      END;

    --3.loop取每一筆退費明細資料資料來處理
      LOOP
        FETCH cDynamic INTO x_BillNo,x_Item,x_CustId,x_CitemCode,x_CitemName,
                            x_ShouldDate,x_ShouldAmt,x_RealDate,x_RealAmt,x_RealPeriod,
                            x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                            x_SBillNo,x_SItem,x_STCode;
        EXIT WHEN cDynamic%NOTFOUND;


        --判定此筆資料是否為呆帳
        IF x_STCode = v_CheckSTCode THEN
          v_RefNo := 1;    --呆帳
        ELSE
          v_RefNo := 2;    --正常
        END IF;


/*
先檢查SBillNo有沒有值
<一>若 SBillNo 沒有值，直接 Append 在最後一筆
  <1>若OrderNo查出有對應的媒體介紹資料，則SO131.Notes=原始單據編號沒有值

  <2>若OrderNo查出沒有對應的媒體介紹資料，則SO131.Notes=原始單據編號沒有值 (且訂單單號找不到對應的媒體介紹相關資料)

  <3>若OrderNo根本就沒有值，則SO131.Notes=原始單據編號沒有值 (且訂單單號為空值,無法找到對應的媒體介紹相關資料)

<二>若 SBillNo 有值,SELECT 整個 SO131
  <1>若 SELECT 不到任何正項 , 直接 Append 在最後一筆 ,
      <A>若OrderNo查出有對應的媒體介紹資料，則SO131.Notes=找不到正項資料,可能早期的資料尚未計算

      <B>若OrderNo查出沒有對應的媒體介紹資料，則SO131.Notes=找不到正項資料,可能早期的資料尚未計算 (且訂單單號找不到對應的媒體介紹相關資料)

      <C>若OrderNo根本就沒有值，則SO131.Notes=找不到正項資料,可能早期的資料尚未計算 (且訂單單號為空值,無法找到對應的媒體介紹相關資料)

  <2>若 SELECT 有正項 , 檢查此正項資料是否為畫面所選的收費項目
      <A>若不是所選收費項目 , 則忽略此退費項目

      <B>若是所選收費項目 , 檢查此正項資料計算年月 , 是否為畫面所選要計算的計算年月
          <a>若是則新增一筆資料至此正項收費項目的下面 , SO131.Notes=退費
          <b>若不是則代表正項及退費項目分別屬不同年月 , 直接 Append 在最後一筆 , SO131.Notes=此月沒有此退費項目的正項資料
*/





        if (x_SBillNo <> '') OR (x_SBillNo IS NOT NULL)then
          BEGIN
            -- 找尋 SO131 全部是否有正項的資料
            Select BillNo,Item,SeqNo,ComputeYM,CitemCode,CitemName,MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName,RealStopDate,OrderNo
              into z_BillNo,z_Item,z_SeqNo,z_ComputeYM,z_CitemCode,z_CitemName,z_MediaCode,z_MediaName,z_AcceptEn,z_AcceptName,z_IntroId,z_IntroName,z_RealStopDate,z_OrderNo
              from SO131
              where RealOrShouldDate=p_RealOrShouldDate and
                    CompCode=p_CompCode and BillNo=x_SBillNo and Item=x_SItem;

            --若有找到對應的正項
            --檢查是否為所選要計算的收費項目代碼,不是的話則不計算
            FOR I IN 1 .. v_CitemCodeCnt LOOP
              v_TmpCitem := CitemCode(I);
              if z_CitemCode = v_TmpCitem then
                v_IsSelectedCitemCode := 'Y';
                --若是所選的收費項目資料
                goto GO_OUT;
              else
                v_IsSelectedCitemCode := 'N';
               end if;
            END LOOP;

            <<GO_OUT>>
            NULL;

            --CitemCode 為所選的清單中
            if v_IsSelectedCitemCode = 'Y' then
              --此筆正項年月為計算年月
              if z_ComputeYM = p_ComputeYM then
                --將此筆退費資料加到對應的正項下面
                BEGIN
                  update SO131 set SEQNO=seqno + 1 where ComputeYM=p_ComputeYM and
                                       RealOrShouldDate=p_RealOrShouldDate and
                                       CompCode=p_CompCode and SeqNo>z_SeqNo;
                EXCEPTION
                WHEN OTHERS THEN
                   ROLLBACK;
                   p_RetCode := -4;
                   p_RetMsg := 'UPDATE時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                   RETURN -4;
                END;


                v_CurSeqNO := z_SeqNo + 1;

                v_Notes := '退費';
              else
                v_Notes := '此月沒有此退費項目的收費正項';
                -- 此月沒有此退費項目的收費正項,則直接Append一筆至最後
                -- 取得最大的SEQNO
                for c2 in cc2 loop
                  v_MaxSeqNO := c2.MaxSeqNO + 1;
                end loop;

                v_CurSeqNO := v_MaxSeqNO;
              end if;


              --取得退費資料的退費開始及結束日期
              --如果退費的 RealStartDate 是 null
              --則 RealStartDate 找退費的 ShouldDate,RealStopDate 找正項的 RealStopDate
              if x_RealStartDate is not null then
                v_RealStartDate := x_RealStartDate;
                v_RealStopDate := x_RealStopDate;
              else
                v_RealStartDate := x_ShouldDate;
                v_RealStopDate := z_RealStopDate;
              end if;


              BEGIN
                  insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                         CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                         REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                         SBILLNO,SITEM,MEDIACODE,MEDIANAME,ACCEPTEN,ACCEPTNAME,
                         INTROID,INTRONAME,Notes,RefNo)
                  Values (v_CurSeqNO,p_ComputeYM,p_RealOrShouldDate,z_BillNo,z_Item,x_CustId,
                         x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                         x_RealPeriod,v_RealStartDate,v_RealStopDate,x_CompCode,x_OrderNo,
                         x_SBillNo,x_SItem,z_MediaCode,z_MediaName,z_AcceptEn,z_AcceptName,
                         z_IntroId,z_IntroName,v_Notes,v_RefNo);
              EXCEPTION
              WHEN OTHERS THEN
                   ROLLBACK;
                   p_RetCode := -3;
                   p_RetMsg := 'INSERT時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                   RETURN -3;
              END;

            end if;

          EXCEPTION
            WHEN others then
              --若 SO131 找不到對應的正項資料,代表尚有早期資料未計算,則直接Append一筆至最後
              -- 取得最大的SEQNO
              for c2 in cc2 loop
                v_MaxSeqNO := c2.MaxSeqNO + 1;
              end loop;


              x_MediaCode := null;
              x_MediaName := '';
              x_AcceptEn := '';
              x_AcceptName := '';
              x_IntroId := '';
              x_IntroName := '';

              IF (x_OrderNo IS NOT NULL) OR (x_OrderNo <> '') THEN
                BEGIN
                  -- Query SO105
                  Select MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName
                    into x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,x_IntroId,x_IntroName
                    from So105 where OrderNo=x_OrderNo;

                  v_Notes := '找不到正項資料,可能早期的資料尚未計算';
                EXCEPTION
                  WHEN others then
                    v_Notes := '找不到正項資料,可能早期的資料尚未計算 (且訂單單號找不到對應的媒體介紹相關資料)';
                END;
              ELSE
                v_Notes := '找不到正項資料,可能早期的資料尚未計算 (且訂單單號為空值,無法找到對應的媒體介紹相關資料)';
              END IF;




              BEGIN
                  insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                         CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                         REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                         SBILLNO,SITEM,MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName,Notes,RefNo)
                  Values (v_MaxSeqNO,p_ComputeYM,p_RealOrShouldDate,x_BillNo,x_Item,x_CustId,
                         x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                         x_RealPeriod,x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                         x_SBillNo,x_SItem,x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,x_IntroId,x_IntroName,v_Notes,v_RefNo);
              EXCEPTION
              WHEN OTHERS THEN
                   ROLLBACK;
                   p_RetCode := -3;
                   p_RetMsg := 'INSERT時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                   RETURN -3;
              END;
          END;


        else
          -- 取得最大的SEQNO
          for c2 in cc2 loop
            v_MaxSeqNO := c2.MaxSeqNO + 1;
          end loop;


          x_MediaCode := null;
          x_MediaName := '';
          x_AcceptEn := '';
          x_AcceptName := '';
          x_IntroId := '';
          x_IntroName := '';

          IF (x_OrderNo IS NOT NULL) OR (x_OrderNo <> '') THEN
            BEGIN
              -- Query SO105
              Select MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName
                into x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,x_IntroId,x_IntroName
                from So105 where OrderNo=x_OrderNo;

                  v_Notes := '原始單據編號沒有值';
            EXCEPTION
              WHEN others then
                v_Notes := '原始單據編號沒有值 (且訂單單號找不到對應的媒體介紹相關資料)';
            END;
          ELSE
            v_Notes := '原始單據編號沒有值 (且訂單單號為空值,無法找到對應的媒體介紹相關資料)';
          END IF;


          BEGIN
              insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                     CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                     REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                     SBILLNO,SITEM,MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName,Notes,RefNo)
              Values (v_MaxSeqNO,p_ComputeYM,p_RealOrShouldDate,x_BillNo,x_Item,x_CustId,
                     x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                     x_RealPeriod,x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                     x_SBillNo,x_SItem,x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,x_IntroId,x_IntroName,v_Notes,v_RefNo);
          EXCEPTION
          WHEN OTHERS THEN
               ROLLBACK;
               p_RetCode := -3;
               p_RetMsg := 'INSERT時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
               RETURN -3;
          END;

        end if;

      END LOOP;--退費明細資料資料來處理

      CLOSE cDynamic;

    END LOOP;    --LOOP 針對SO033及SO034退費




    COMMIT;

    p_RetCode := 0;
    p_RetMsg := '';
    RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- SF_CALCULATESO131
/
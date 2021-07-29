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

    By: Jackal
    Date: 2003.04.09
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

  --v_BeginMonth          Number;
  --v_EndMonth            Number;
  v_ServiceType         char(1) := UPPER(p_ServiceType);
  v_Notes               varchar2(255);   


  c39                   CHAR(1) := CHR(39);


  type CurTyp is ref cursor;	--自訂cursor型態
  cDynamic CurTyp;          	--供dynamic SQL

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
                ||'SBillNo,SItem  from '||v_Table;

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
          rollback;
          p_RetCode := -2;
          p_RetMsg := 'SELECT 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM||'      SQL='||v_SQL;
          RETURN -2;
      END;

    --3.loop取每一筆資料來處理
      LOOP
        FETCH cDynamic INTO x_BillNo,x_Item,x_CustId,x_CitemCode,x_CitemName,
                            x_ShouldDate,x_ShouldAmt,x_RealDate,x_RealAmt,x_RealPeriod,
                            x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                            x_SBillNo,x_SItem;
        EXIT WHEN cDynamic%NOTFOUND;

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
              p_RetCode := -2;
              p_RetMsg := 'SELECT 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
              RETURN -2;
          END;
        ELSE
          v_Notes := '訂單單號為空直,無法找到對應的媒體介紹相關資料';
        END IF;
        


        BEGIN
            insert into SO131(COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                   CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                   REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                   SBILLNO,SITEM,MEDIACODE,MEDIANAME,ACCEPTEN,ACCEPTNAME,
                   INTROID,INTRONAME,Notes)
            Values (p_ComputeYM,p_RealOrShouldDate,x_BillNo,x_Item,x_CustId,
                   x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                   x_RealPeriod,x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                   x_SBillNo,x_SItem,x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,
                   x_IntroId,x_IntroName,v_Notes);
        EXCEPTION
        WHEN OTHERS THEN
             rollback;
             p_RetCode := -3;
             p_RetMsg := 'INSERT時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
             RETURN -3;
        END;

      END LOOP;  --LOOP SO033及SO034的明細資料

    --4.Close dynamic cursor
      CLOSE cDynamic;

    END LOOP;    --LOOP SO033及SO034


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
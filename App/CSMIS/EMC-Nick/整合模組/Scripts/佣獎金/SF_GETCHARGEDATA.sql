CREATE OR REPLACE FUNCTION SF_GETCHARGEDATA
  ( p_Table varchar2,p_CompCode varchar2,p_BelongYM varchar2,p_BillNo varchar2,p_Item varchar2,
    p_RetCode OUT number,p_RetMsg OUT varchar2)
  RETURN  number AS
/*
--@D:\SF_GetChargeData;
VARIABLE ReturnNo NUMBER
VARIABLE SEQ NUMBER
VARIABLE MSG VARCHAR2(1000)
EXEC :ReturnNo := SF_GetChargeData('7','09203','3','1',:SEQ,:MSG)
PRINT ReturnNo
PRINT SEQ
PRINT MSG

  佣金計算退費資料找正項資料,找原來Table內的正項資料
  檔名: SF_GetChargeData.sql


  IN    p_Table       varchar2: Table名稱( SO122或SO134 )
        p_CompCode    varchar2: 公司別代碼
        p_BelongYM    varchar2: 歸屬年月
        p_BillNo      varchar2: 單據編號
        p_Item        varchar2: 單據項次

  OUT  p_SeqNo          number:   結果資料批號
       p_RetMsg         varchar2: 結果訊息 (至少200 bytes)

       Return           number:   處理結果代碼, 對應訊息存於p_RetMsg中
                           >=0: p_RetMsg='執行完畢,但沒有資料'
                             1; p_RetMsg= 將找到的資料串起來傳回前端
                            -1: p_RetMsg='參數錯誤'
                            -2: p_RetMsg='SELECT 時發生錯誤:
                            -3: p_RetMsg='INSERT 時發生錯誤:
                           -99: p_RetMsg='其他錯誤'

    By: Jackal   Date: 2003.05.23
    Modify By Jackal : 2003.05.30  將WHERE條件去除FirstFlag及BelongYM
    Modify By Jackal : 2003.09.05  多回傳OrderNo及CallSelf
*/
    v_SQL              	varchar2(4000);		-- must long enough
    v_Column	          varchar2(1000);
    v_Where             varchar2(1000);

    v_ReturnString     	varchar2(2000);


    x_BillNo            varchar2(15);
    x_IntAccId          varchar2(20);
    x_IntAccName        varchar2(20);
    x_IntAccOriPercent  varchar2(20);
    x_StaffCode         varchar2(20);
    x_RealStartDate     varchar2(20);
    x_RealStopDate      varchar2(20);
    x_CitemCode         varchar2(20);
    x_CitemName         varchar2(20);
    x_MediaCode         varchar2(20);
    x_MediaName         varchar2(20);
    x_PROMCODE          varchar2(20);
    x_PROMNAME          varchar2(20);
    x_CustId            varchar2(20);
    x_PeriodId          varchar2(20);
    x_FirstFlag		varchar2(20);
    x_OrderNo		varchar2(20);
    x_CallSelf		varchar2(20);

    c39                   CHAR(1) := CHR(39);

    type CurTyp is ref cursor;	--自訂cursor型態
    cDynamic CurTyp;          	--供dynamic SQL

BEGIN
    if p_Table Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:Table不能為Null值';
      Return -1;
    end if;

    if p_CompCode Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:公司別不能為Null值';
      Return -1;
    end if;


    if p_BelongYM Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:歸屬年月不能為Null值';
      Return -1;
    end if;


    if p_BillNo Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:單據編號不能為Null值';
      Return -1;
    end if;


    if p_Item Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:單據項次不能為Null值';
      Return -1;
    end if;


    --1.組合SQL查詢字串:
    v_Column := 'select BillNo,IntAccId,'||
                        'IntAccName,IntAccOriPercent,StaffCode,'||
                        'TO_CHAR(RealStartDate,'||c39||'YYYY/MM/DD'||c39||') RealStartDate,'||
                        'TO_CHAR(RealStopDate,'||c39||'YYYY/MM/DD'||c39||') RealStopDate,'||
                        'CitemCode,CitemName,MediaCode,'||
                        'MediaName,PROMCODE,PROMNAME,CustId,PeriodId,FirstFlag,OrderNo,CallSelf  from '||p_Table;

    --v_Where  := ' WHERE CompCode='||p_CompCode||' and BelongYM='||c39||p_BelongYM||c39||
    v_Where  := ' WHERE CompCode='||p_CompCode||
                ' and BillNo='||c39||p_BillNo||c39||' and Item='||p_Item;
                --' and FirstFlag='||c39||'1'||c39;

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
      x_BillNo := ' ';
      FETCH cDynamic INTO x_BillNo,
                          x_IntAccId,x_IntAccName,x_IntAccOriPercent,
                          x_StaffCode,x_RealStartDate,x_RealStopDate,x_CitemCode,
                          x_CitemName,x_MediaCode,x_MediaName,x_PromCode,x_PromName,
                          x_CustId,x_PeriodId,x_FirstFlag,x_OrderNo,x_CallSelf;
      EXIT WHEN cDynamic%NOTFOUND;

      v_ReturnString := '';

      IF (x_IntAccId = '') OR (x_IntAccId IS NULL) THEN
        x_IntAccId := 'NULL';
      END IF;

      IF (x_IntAccName = '') OR (x_IntAccName IS NULL) THEN
        x_IntAccName := 'NULL';
      END IF;

      IF (x_IntAccOriPercent = '') OR (x_IntAccOriPercent IS NULL) THEN
        x_IntAccOriPercent := 'NULL';
      END IF;

      IF (x_StaffCode = '') OR (x_StaffCode IS NULL) THEN
        x_StaffCode := 'NULL';
      END IF;

      IF (x_RealStartDate = '') OR (x_RealStartDate IS NULL) THEN
        x_RealStartDate := 'NULL';
      END IF;

      IF (x_RealStopDate = '') OR (x_RealStopDate IS NULL) THEN
        x_RealStopDate := 'NULL';
      END IF;

      IF (x_CitemCode = '') OR (x_CitemCode IS NULL) THEN
        x_CitemCode := 'NULL';
      END IF;

      IF (x_CitemName = '') OR (x_CitemName IS NULL) THEN
        x_CitemName := 'NULL';
      END IF;

      IF (x_MediaCode = '') OR (x_MediaCode IS NULL) THEN
        x_MediaCode := 'NULL';
      END IF;

      IF (x_MediaName = '') OR (x_MediaName IS NULL) THEN
        x_MediaName := 'NULL';
      END IF;

      IF (x_PromCode = '') OR (x_PromCode IS NULL) THEN
        x_PromCode := 'NULL';
      END IF;

      IF (x_PromName = '') OR (x_PromName IS NULL) THEN
        x_PromName := 'NULL';
      END IF;

      IF (x_CustId = '') OR (x_CustId IS NULL) THEN
        x_CustId := 'NULL';
      END IF;


      IF (x_PeriodId = '') OR (x_PeriodId IS NULL) THEN
        x_PeriodId := 'NULL';
      END IF;


      IF (x_FirstFlag = '') OR (x_FirstFlag IS NULL) THEN
        x_FirstFlag := 'NULL';
      END IF;


      IF (x_OrderNo = '') OR (x_OrderNo IS NULL) THEN
        x_OrderNo := 'NULL';
      END IF;


      IF (x_CallSelf = '') OR (x_CallSelf IS NULL) THEN
        x_CallSelf := 'NULL';
      END IF;





      IF (x_BillNo IS NOT NULL) OR (x_BillNo <> ' ') THEN
        v_ReturnString := x_IntAccId||','||x_IntAccName||','||x_IntAccOriPercent||','||
                          x_StaffCode||','||x_RealStartDate||','||x_RealStopDate||','||x_CitemCode||','||
                          x_CitemName||','||x_MediaCode||','||x_MediaName||','||x_PromCode||','||x_PromName||','||
                          x_CustId||','||x_PeriodId||','||x_FirstFlag||','||x_OrderNo||','||x_CallSelf;

        p_RetCode := 1;
        p_RetMsg := v_ReturnString;
        Return 1;
      END IF;

    END LOOP;


    --4.Close dynamic cursor
    CLOSE cDynamic;


    p_RetCode := 0;
    p_RetMsg := 'OK';
    RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- Function SF_GETCHARGEDATA
/
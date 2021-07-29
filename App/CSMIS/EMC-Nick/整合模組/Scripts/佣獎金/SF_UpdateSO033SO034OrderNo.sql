CREATE OR REPLACE FUNCTION SF_UPDATESO033SO034ORDERNO
  (p_CompCode varchar2,p_RetCode OUT number,p_RetMsg OUT varchar2)
  RETURN  number AS
/*
--@D:\SF_UpdateSO033SO034OrderNo.sql;
VARIABLE ReturnNo NUMBER
VARIABLE SEQ NUMBER
VARIABLE MSG VARCHAR2(1000)
EXEC :ReturnNo := SF_UpdateSO033SO034OrderNo('3',:SEQ,:MSG)
PRINT ReturnNo
PRINT SEQ
PRINT MSG

  補填SO033及SO034的OrderNo
  檔名: SF_UpdateSO033SO034OrderNo.sql


  IN   p_CompCode    varchar2: 公司別代碼


  OUT  p_SeqNo          number:   結果資料批號
       p_RetMsg         varchar2: 結果訊息 (至少200 bytes)

       Return           number:   處理結果代碼, 對應訊息存於p_RetMsg中
                           >=0: p_RetMsg='執行完畢,但沒有資料'
                             1; p_RetMsg= 將找到的資料串起來傳回前端
                            -1: p_RetMsg='參數錯誤'
                            -2: p_RetMsg='SELECT 時發生錯誤:
                            -3: p_RetMsg='UPDATE 時發生錯誤:
                           -99: p_RetMsg='其他錯誤'

    By: Jackal   Date: 2003.09.09
*/
    v_SQL              	varchar2(4000);		-- must long enough
    v_SQL2             	varchar2(4000);		-- must long enough
    v_Column	          varchar2(1000);
    v_Where             varchar2(1000);
    v_Table             varchar2(20);

    x_BillNo            varchar2(15);
    x_Item              Number;
    x_CustId            Number;
    x_CitemCode         Number;
    x_SeqNo             Number;
    x_OrderNo		        varchar2(20);

    v_Counts            Number := 0;

    c39                   CHAR(1) := CHR(39);

    type CurTyp is ref cursor;	--自訂cursor型態
    cDynamic CurTyp;          	--供dynamic SQL

BEGIN

    if p_CompCode Is Null then
      p_RetCode := -1;
      p_RetMsg := '參數錯誤:公司別不能為Null值';
      Return -1;
    end if;

    FOR i IN 1 .. 2 LOOP
      IF i = 1 THEN
        v_Table := 'SO033';
      ELSIF i = 2 THEN
        v_Table := 'SO034';
      END IF;

      --1.組合SQL查詢字串:
      v_Column := 'select A.BillNo,A.Item,A.CustId,A.CitemCode,A.SeqNo,B.OrderNo from '||v_Table||' A,SO003 B';

      v_Where  := ' WHERE A.CompCode='||p_CompCode|| ' and A.OrderNo is null' ||
                  ' and B.CustId=A.CustId and B.CitemCode=A.CitemCode'||
                  ' and B.SeqNo=A.SeqNo and B.CompCode=A.CompCode';

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
        FETCH cDynamic INTO x_BillNo,x_Item,x_CustId,x_CitemCode,x_SeqNo,x_OrderNo;
        EXIT WHEN cDynamic%NOTFOUND;


        IF (x_OrderNo IS NOT NULL) OR (x_OrderNo <> '') THEN
          BEGIN
            v_SQL2 := 'Update '||v_Table||' SET OrderNo='||c39||x_OrderNo||c39||
                      ' WHERE COMPCODE='||p_CompCode||' and BillNo='||c39||x_BillNo||c39||
                      ' and Item='||x_Item;

            EXECUTE IMMEDIATE v_SQL2;

            v_Counts := v_Counts + SQL%RowCount;
          EXCEPTION
            WHEN others then
              p_RetMsg :=''; 
          END;
        END IF;

      END LOOP;


      --4.Close dynamic cursor
      CLOSE cDynamic;

    END LOOP;  --LOOP SO033及SO034的明細資料

    COMMIT;

    p_RetCode := 0;
    p_RetMsg := '完成,共處理 ' ||v_Counts||' 筆資料';
    RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- Function SF_UpdateSO033SO034OrderNo.sql
/
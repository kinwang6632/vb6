CREATE OR REPLACE FUNCTION SF_UPDATESO122DATA
  ( p_CompCode varchar2,p_RetCode OUT number,p_RetMsg OUT varchar2)
  RETURN  number AS
/*
--@D:\SF_UpdateSO122Data;
VARIABLE ReturnNo NUMBER
VARIABLE SEQ NUMBER
VARIABLE MSG VARCHAR2(1000)
EXEC :ReturnNo := SF_UpdateSO122Data('6',:SEQ,:MSG)
PRINT ReturnNo
PRINT SEQ
PRINT MSG

  佣金頻道資料補加OrderNo及CallSelf兩個欄位
  檔名: SF_UpdateSO122Data.sql


  IN   p_CompCode    varchar2: 公司別代碼


  OUT  p_SeqNo          number:   結果資料批號
       p_RetMsg         varchar2: 結果訊息 (至少200 bytes)

       Return           number:   處理結果代碼, 對應訊息存於p_RetMsg中
                           >=0: p_RetMsg='執行完畢,但沒有資料'
                             1; p_RetMsg= 將找到的資料串起來傳回前端
                            -1: p_RetMsg='參數錯誤'
                            -2: p_RetMsg='SELECT 時發生錯誤:
                            -3: p_RetMsg='INSERT 時發生錯誤:
                           -99: p_RetMsg='其他錯誤'

    By: Jackal   Date: 2003.09.05
*/
    v_SQL              	varchar2(4000);		-- must long enough
    v_Column	        varchar2(1000);
    v_Where             varchar2(1000);
    v_Table             varchar2(20);

    x_BillNo            varchar2(15);
    x_Item              Number;
    x_ComputeYM         varchar2(15);
    x_BelongYM          varchar2(15);
    x_CustId            Number;


    x_OrderNo		        varchar2(20);
    x_CallSelf		      Number;
    x_RefNo             Number;
    
    
    
    x_StbNo             varchar2(20);
    x_Sno               varchar2(15);
    x_StaffCode         CHAR(1);
    
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

----------------------------------------------------------------------    
-------------------------   計算付費頻道   ---------------------------
----------------------------------------------------------------------
    --1.組合SQL查詢字串:(SELECT 出所有SO122付費頻道佣金資料)
    v_Column := 'select ComputeYM,BelongYM,CustId,BillNo,ITEM from SO122 ';

    v_Where  := ' WHERE CompCode='||p_CompCode|| ' and BuyOrRent is null '||
                ' and BillNo is not null';

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
      x_OrderNo := '';

      FETCH cDynamic INTO x_ComputeYM,x_BelongYM,x_CustId,x_BillNo,x_Item;
      EXIT WHEN cDynamic%NOTFOUND;

      --判別該筆付費頻道資料,是否為自來電
      BEGIN


        select A.OrderNo,C.RefNo into x_OrderNo,x_RefNo
               from SO033 A,SO105 B,CD009 C
               where A.CompCode=p_CompCode and A.BillNo=x_BillNo and
               A.Item=x_Item and B.OrderNo=A.OrderNo and
               C.CodeNo=B.MediaCode;
      EXCEPTION
        WHEN others then
          BEGIN
            select A.OrderNo,C.RefNo into x_OrderNo,x_RefNo
                   from SO034 A,SO105 B,CD009 C
                   where A.CompCode=p_CompCode and A.BillNo=x_BillNo and
                   A.Item=x_Item and B.OrderNo=A.OrderNo and
                   C.CodeNo=B.MediaCode;
          EXCEPTION
            WHEN others then
              p_RetMsg := '';
          END;
      END;

      --如果有值則判別是否為自來電,並回填SO122
      IF ((x_OrderNo<>'') OR (x_OrderNo IS NOT NULL))  AND (x_RefNo IS NOT NULL) THEN
        IF (x_RefNo=1) OR (x_RefNo=2) OR (x_RefNo=3) OR (x_RefNo=5) THEN
          x_CallSelf := 0;--非自來電
        ELSE
          x_CallSelf := 1;--自來電
        END IF;


        BEGIN
          UPDATE SO122 SET ORDERNO=x_OrderNo,CallSelf=x_CallSelf
                 WHERE CompCode=p_CompCode and ComputeYM=x_ComputeYM and
                 BelongYM=x_BelongYM and CustId= x_CustId and
                 BillNo=x_BillNo and Item= x_Item;
                 
          v_Counts := v_Counts + SQL%RowCount;
        EXCEPTION
          WHEN others then
            p_RetMsg := '';
        END;
      END IF;
    END LOOP;


    --4.Close dynamic cursor
    CLOSE cDynamic;
----------------------------------------------------------------------    
----------------------------   計算 STB   ----------------------------
----------------------------------------------------------------------
    --1.組合SQL查詢字串:(SELECT 出所有SO122付費頻道佣金資料)
    v_Column := 'select ComputeYM,BelongYM,CustId,StbNo,Sno,StaffCode from SO122 ';

    v_Where  := ' WHERE CompCode='||p_CompCode|| ' and BuyOrRent is not null '||
                ' and BillNo is null';

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
      x_OrderNo := '';
                                                       
      FETCH cDynamic INTO x_ComputeYM,x_BelongYM,x_CustId,x_StbNo,x_Sno,x_StaffCode;
      EXIT WHEN cDynamic%NOTFOUND;

      --判別該筆STB資料,是否為自來電
      IF x_StaffCode = 1 THEN
        x_CallSelf := 0;--非自來電
      ELSIF x_StaffCode = 2 THEN
        x_CallSelf := 1;--自來電
      END IF;
      
      
      
      BEGIN
        UPDATE SO122 SET CallSelf=x_CallSelf
               WHERE CompCode=p_CompCode and ComputeYM=x_ComputeYM and
               BelongYM=x_BelongYM and CustId= x_CustId and
               StbNo=x_StbNo and Sno= x_Sno;
               
           
        v_Counts := v_Counts + SQL%RowCount;    
      EXCEPTION
        WHEN others then
          p_RetMsg := '';
      END;            
    END LOOP;    


    COMMIT;
    p_RetCode := 0;
    p_RetMsg := '完成,共處理 ' ||v_Counts||' 筆資料';
    RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- Function SF_UpdateSO122Data
/
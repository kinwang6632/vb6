CREATE OR REPLACE PROCEDURE SP_GETSNAPSHOTTABLECOUNTS
AS

/*
--@D:\EMC\Snapshot\Script\SP_GetSnapShotTableCounts;
EXEC SP_GetSnapShotTableCounts


  計算要傳至 SnapShot Table 的筆數
  檔名: SP_GetSnapShotTableCounts.sql

  OUT  p_SeqNo          number:   結果資料批號
       p_RetMsg         varchar2: 結果訊息 (至少200 bytes)

       Return           number:   處理結果代碼, 對應訊息存於p_RetMsg中
                             1; p_RetMsg= 執行完畢
                            -1: p_RetMsg='SELECT 時發生錯誤:
                            -2: p_RetMsg='INSERT 時發生錯誤:
                           -99: p_RetMsg='其他錯誤'

    By: Jackal   Date: 2003.07.11
*/
    v_SQL              	varchar2(4000);		-- must long enough
    v_ErrCode           varchar2(200);
    v_ErrMsg            varchar2(1024);

    v_CurrDateTime      Date := sysdate;
    v_StrCurrDateTime   varchar2(30);
    v_MoveTableCode     varchar2(30);
    v_MoveTableName     varchar2(60);
    v_MoveTableCounts   Number := 0;
    v_StartCurrDateTime varchar2(30);
    v_StopCurrDateTime  varchar2(30);
    v_TempCounts   Number := 0;


    c39                   CHAR(1) := CHR(39);
    cursor cc1 is
      SELECT * FROM SO509;

    type CurTyp is ref cursor;	--自訂cursor型態
    cDynamic CurTyp;          	--供dynamic SQL

BEGIN
    v_StartCurrDateTime := TO_CHAR(v_CurrDateTime,'YYYY/MM/DD') || ' 00:00:01';
    v_StopCurrDateTime := TO_CHAR(v_CurrDateTime,'YYYY/MM/DD') || ' 23:59:59';
    v_StrCurrDateTime := to_char(v_CurrDateTime,'YYYY/MM/DD HH24:MI:SS');

    DELETE SO509A WHERE COMPUTEDATE BETWEEN TO_DATE(v_StartCurrDateTime,'YYYY/MM/DD HH24:MI:SS')
                                    AND TO_DATE(v_StopCurrDateTime,'YYYY/MM/DD HH24:MI:SS');

    -- 取得所有要 SnapShot 的Table
    FOR c1 IN cc1 LOOP
      v_MoveTableCode := c1.TableCode;
      v_MoveTableName := c1.TableName;

      --Query 資料庫是否有此 Table
      SELECT COUNT(*) INTO v_TempCounts from user_tables WHERE TABLE_NAME=v_MoveTableCode;


      IF v_TempCounts<>0 THEN
        --1.組合SQL查詢字串:
        v_SQL := 'SELECT COUNT(*) COUNTS FROM ' || v_MoveTableCode;

        --2.開啟Dynamic sql
        BEGIN
          IF NOT cDynamic%ISOPEN then
            OPEN cDynamic FOR v_SQL;
          END IF;
        EXCEPTION
          WHEN OTHERS THEN
            ROLLBACK;
            
        END;

        --3.loop計算每一個要 SnapShot 的 Table 筆數
        v_MoveTableCounts := 0;
        LOOP
          FETCH cDynamic INTO v_MoveTableCounts;
          EXIT WHEN cDynamic%NOTFOUND;

          --將要 SnapShot Table 的筆數儲存至SO509A
          BEGIN
              v_SQL := 'insert into SO509A(TABLECODE,TABLENAME,COMPUTEDATE,RECORDCOUNTS) Values (''' ||
                       v_MoveTableCode || ''',''' || v_MoveTableName || 
                       ''',to_date(''' || ''',''YYYY/MM/DD HH24:MI:SS''),'||
                       v_MoveTableCounts || ')';
                       
              insert into SO509A(TABLECODE,TABLENAME,COMPUTEDATE,RECORDCOUNTS)
              Values (v_MoveTableCode,v_MoveTableName,v_CurrDateTime,v_MoveTableCounts);
          EXCEPTION
          WHEN OTHERS THEN
              ROLLBACK;
              BEGIN
                  v_ErrCode := SQLCODE;
                  v_ErrMsg := SQLERRM;
                  insert into SO509C(ERRCODE,ERRMSG,DESCRIPTION,FUNCTIONNAME,LOGDATETIME)
                  Values (v_ErrCode,v_ErrMsg,v_SQL,'SP_GetSnapShotTableCounts.sql',v_CurrDateTime);
              EXCEPTION
              WHEN OTHERS THEN
                   ROLLBACK;
              END;
          END;
        END LOOP; --計算每一個要 SnapShot 的 Table 筆數

        --4.Close dynamic cursor
        CLOSE cDynamic;

      ELSE
        --若該 Table 不存在則數量顯示 0
        v_MoveTableCounts := 0;

        --將要 SnapShot Table 的筆數儲存至SO509A
        BEGIN
            v_SQL := 'insert into SO509A(TABLECODE,TABLENAME,COMPUTEDATE,RECORDCOUNTS) Values (''' ||
                     v_MoveTableCode || ''',''' || v_MoveTableName || 
                     ''',to_date(''' || ''',''YYYY/MM/DD HH24:MI:SS''),'||
                     v_MoveTableCounts || ')';
                               
            insert into SO509A(TABLECODE,TABLENAME,COMPUTEDATE,RECORDCOUNTS)
            Values (v_MoveTableCode,v_MoveTableName,v_CurrDateTime,v_MoveTableCounts);
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            BEGIN
                v_ErrCode := SQLCODE;
                v_ErrMsg := SQLERRM;
                insert into SO509C(ERRCODE,ERRMSG,DESCRIPTION,FUNCTIONNAME,LOGDATETIME)
                Values (v_ErrCode,v_ErrMsg,v_SQL,'SP_GetSnapShotTableCounts.sql',v_CurrDateTime);
            EXCEPTION
            WHEN OTHERS THEN
                 ROLLBACK;
            END;
        END;
        
        --不存在的 Table Log 起來
        BEGIN
            v_ErrCode := SQLCODE;
            v_ErrMsg := SQLERRM;
            v_SQL := 'Table ' || v_MoveTableCode || ' 不存在';
            insert into SO509C(ERRCODE,ERRMSG,DESCRIPTION,FUNCTIONNAME,LOGDATETIME)
            Values (v_ErrCode,v_ErrMsg,v_SQL,'SP_GetSnapShotTableCounts.sql',v_CurrDateTime);
        EXCEPTION
        WHEN OTHERS THEN
             ROLLBACK;
        END;          
      END IF;
    END LOOP; --取得所有要SnapShot 的Table


    COMMIT;
EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  BEGIN
      v_ErrCode := SQLCODE;
      v_ErrMsg := SQLERRM;
      insert into SO509C(ERRCODE,ERRMSG,DESCRIPTION,FUNCTIONNAME,LOGDATETIME)
      Values (v_ErrCode,v_ErrMsg,'其他錯誤','SP_GetSnapShotTableCounts.sql',v_CurrDateTime);
  EXCEPTION
  WHEN OTHERS THEN
       ROLLBACK;
  END;  
END; -- Function SP_GetSnapShotTableCounts
/
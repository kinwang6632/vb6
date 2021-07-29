CREATE OR REPLACE PROCEDURE SP_COMPUTESESSION(p_RetCode Out Number, p_RetMsg Out Varchar2)IS


/*

  --@D:\App\EMC\資料庫之session紀錄程式\SP_COMPUTESESSION.sql
    variable p_RetCode number
    variable p_RetMsg varchar2(100)
    exec  SP_COMPUTESESSION(:p_RetCode,:p_RetMsg);
    print P_Retcode
    print p_RetMsg

  目的: 計算當下資料的 session 數,並存入 table 中 
  檔名: SP_ComputeSession.sql



  OUT p_RetMsg  varchar2: 處理結果自串

      p_RetCode number: 處理結果代碼, 對應訊息存於p_RetMsg中

              0: p_RetMsg='計算完畢'
             -1: p_RetMsg='Insert至table時發生錯誤
            -99: p_RetMsg='其他錯誤'

  流程:


  定義:

  Date: 2004/02/03
  By:Howard
*/


  n_SessionCount  Number(8);
  s_SQL		VARCHAR2(100);

  s_UserName    VARCHAR2(30);
  s_Status      VARCHAR2(8);
  s_CurDateTime Varchar2(20);


  type CurTyp is ref cursor;	--自訂cursor型態
  cDynamic CurTyp;          	--供dynamic SQL




BEGIN

  s_SQL := 'select username,status,count(*) from v$session where username IS NOT NULL group by username, status';

  -- down, 執行 SQL, 取出 session 的資料...
  begin
    IF NOT cDynamic%ISOPEN then
       OPEN cDynamic FOR s_SQL;
    END IF;
  exception
    when others then
       ROLLBACK;
       CLOSE cDynamic;
       p_RetMsg := '查詢session 資料失敗:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
       p_RetCode := -99;
  end;
  -- up, 執行 SQL, 取出 session 的資料...


  select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') into s_CurDateTime from dual;

  --down, loop 處理每一筆 session的資料
  loop


    FETCH cDynamic INTO  s_UserName,s_Status,n_SessionCount;
    EXIT WHEN cDynamic%NOTFOUND;

    begin    
      INSERT INTO SessionCountTab(COMPUTEDATE,USERNAME,STATUS,SESSIONCOUNT)
      VALUES(to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), s_UserName,s_Status,n_SessionCount);

    EXCEPTION
    WHEN OTHERS THEN
      CLOSE cDynamic;
      ROLLBACK;
      p_RetMsg := 'INSERT SessionCountTab 時發生錯誤: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
      p_RetCode := -99;
    END;



  end loop;
  --up, loop 處理每一筆 session的資料

  CLOSE cDynamic;

  COMMIT;

  p_RetMsg:='計算完畢';
  p_RetCode := 0;

 EXCEPTION
    WHEN others THEN
     rollback ;
     p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
     P_RetCode:=-99;

END ; 

/

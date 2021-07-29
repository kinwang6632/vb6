CREATE OR REPLACE PROCEDURE SP_GETSNAPSHOTTABLECOUNTS
AS

/*
--@D:\EMC\Snapshot\Script\SP_GetSnapShotTableCounts;
EXEC SP_GetSnapShotTableCounts


  �p��n�Ǧ� SnapShot Table ������
  �ɦW: SP_GetSnapShotTableCounts.sql

  OUT  p_SeqNo          number:   ���G��Ƨ帹
       p_RetMsg         varchar2: ���G�T�� (�ܤ�200 bytes)

       Return           number:   �B�z���G�N�X, �����T���s��p_RetMsg��
                             1; p_RetMsg= ���槹��
                            -1: p_RetMsg='SELECT �ɵo�Ϳ��~:
                            -2: p_RetMsg='INSERT �ɵo�Ϳ��~:
                           -99: p_RetMsg='��L���~'

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

    type CurTyp is ref cursor;	--�ۭqcursor���A
    cDynamic CurTyp;          	--��dynamic SQL

BEGIN
    v_StartCurrDateTime := TO_CHAR(v_CurrDateTime,'YYYY/MM/DD') || ' 00:00:01';
    v_StopCurrDateTime := TO_CHAR(v_CurrDateTime,'YYYY/MM/DD') || ' 23:59:59';
    v_StrCurrDateTime := to_char(v_CurrDateTime,'YYYY/MM/DD HH24:MI:SS');

    DELETE SO509A WHERE COMPUTEDATE BETWEEN TO_DATE(v_StartCurrDateTime,'YYYY/MM/DD HH24:MI:SS')
                                    AND TO_DATE(v_StopCurrDateTime,'YYYY/MM/DD HH24:MI:SS');

    -- ���o�Ҧ��n SnapShot ��Table
    FOR c1 IN cc1 LOOP
      v_MoveTableCode := c1.TableCode;
      v_MoveTableName := c1.TableName;

      --Query ��Ʈw�O�_���� Table
      SELECT COUNT(*) INTO v_TempCounts from user_tables WHERE TABLE_NAME=v_MoveTableCode;


      IF v_TempCounts<>0 THEN
        --1.�զXSQL�d�ߦr��:
        v_SQL := 'SELECT COUNT(*) COUNTS FROM ' || v_MoveTableCode;

        --2.�}��Dynamic sql
        BEGIN
          IF NOT cDynamic%ISOPEN then
            OPEN cDynamic FOR v_SQL;
          END IF;
        EXCEPTION
          WHEN OTHERS THEN
            ROLLBACK;
            
        END;

        --3.loop�p��C�@�ӭn SnapShot �� Table ����
        v_MoveTableCounts := 0;
        LOOP
          FETCH cDynamic INTO v_MoveTableCounts;
          EXIT WHEN cDynamic%NOTFOUND;

          --�N�n SnapShot Table �������x�s��SO509A
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
        END LOOP; --�p��C�@�ӭn SnapShot �� Table ����

        --4.Close dynamic cursor
        CLOSE cDynamic;

      ELSE
        --�Y�� Table ���s�b�h�ƶq��� 0
        v_MoveTableCounts := 0;

        --�N�n SnapShot Table �������x�s��SO509A
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
        
        --���s�b�� Table Log �_��
        BEGIN
            v_ErrCode := SQLCODE;
            v_ErrMsg := SQLERRM;
            v_SQL := 'Table ' || v_MoveTableCode || ' ���s�b';
            insert into SO509C(ERRCODE,ERRMSG,DESCRIPTION,FUNCTIONNAME,LOGDATETIME)
            Values (v_ErrCode,v_ErrMsg,v_SQL,'SP_GetSnapShotTableCounts.sql',v_CurrDateTime);
        EXCEPTION
        WHEN OTHERS THEN
             ROLLBACK;
        END;          
      END IF;
    END LOOP; --���o�Ҧ��nSnapShot ��Table


    COMMIT;
EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  BEGIN
      v_ErrCode := SQLCODE;
      v_ErrMsg := SQLERRM;
      insert into SO509C(ERRCODE,ERRMSG,DESCRIPTION,FUNCTIONNAME,LOGDATETIME)
      Values (v_ErrCode,v_ErrMsg,'��L���~','SP_GetSnapShotTableCounts.sql',v_CurrDateTime);
  EXCEPTION
  WHEN OTHERS THEN
       ROLLBACK;
  END;  
END; -- Function SP_GetSnapShotTableCounts
/
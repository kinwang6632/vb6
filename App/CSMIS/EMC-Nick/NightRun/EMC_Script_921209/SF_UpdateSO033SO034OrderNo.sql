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

  �ɶ�SO033��SO034��OrderNo
  �ɦW: SF_UpdateSO033SO034OrderNo.sql


  IN   p_CompCode    varchar2: ���q�O�N�X


  OUT  p_SeqNo          number:   ���G��Ƨ帹
       p_RetMsg         varchar2: ���G�T�� (�ܤ�200 bytes)

       Return           number:   �B�z���G�N�X, �����T���s��p_RetMsg��
                           >=0: p_RetMsg='���槹��,���S�����'
                             1; p_RetMsg= �N��쪺��Ʀ�_�ӶǦ^�e��
                            -1: p_RetMsg='�Ѽƿ��~'
                            -2: p_RetMsg='SELECT �ɵo�Ϳ��~:
                            -3: p_RetMsg='UPDATE �ɵo�Ϳ��~:
                           -99: p_RetMsg='��L���~'

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

    type CurTyp is ref cursor;	--�ۭqcursor���A
    cDynamic CurTyp;          	--��dynamic SQL

BEGIN

    if p_CompCode Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:���q�O���ରNull��';
      Return -1;
    end if;

    FOR i IN 1 .. 2 LOOP
      IF i = 1 THEN
        v_Table := 'SO033';
      ELSIF i = 2 THEN
        v_Table := 'SO034';
      END IF;

      --1.�զXSQL�d�ߦr��:
      v_Column := 'select A.BillNo,A.Item,A.CustId,A.CitemCode,A.SeqNo,B.OrderNo from '||v_Table||' A,SO003 B';

      v_Where  := ' WHERE A.CompCode='||p_CompCode|| ' and A.OrderNo is null' ||
                  ' and B.CustId=A.CustId and B.CitemCode=A.CitemCode'||
                  ' and B.SeqNo=A.SeqNo and B.CompCode=A.CompCode';

      v_SQL := v_Column || v_Where;

      --2.�}��Dynamic sql
      BEGIN
        IF NOT cDynamic%ISOPEN then
          OPEN cDynamic FOR v_SQL;
        END IF;
      EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_RetCode := -2;
          p_RetMsg := 'SELECT �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM||'      SQL='||v_SQL;
          RETURN -2;
      END;


       --3.loop���C�@����ƨӳB�z
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

    END LOOP;  --LOOP SO033��SO034�����Ӹ��

    COMMIT;

    p_RetCode := 0;
    p_RetMsg := '����,�@�B�z ' ||v_Counts||' �����';
    RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- Function SF_UpdateSO033SO034OrderNo.sql
/
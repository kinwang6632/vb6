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

  �����W�D��Ƹɥ[OrderNo��CallSelf������
  �ɦW: SF_UpdateSO122Data.sql


  IN   p_CompCode    varchar2: ���q�O�N�X


  OUT  p_SeqNo          number:   ���G��Ƨ帹
       p_RetMsg         varchar2: ���G�T�� (�ܤ�200 bytes)

       Return           number:   �B�z���G�N�X, �����T���s��p_RetMsg��
                           >=0: p_RetMsg='���槹��,���S�����'
                             1; p_RetMsg= �N��쪺��Ʀ�_�ӶǦ^�e��
                            -1: p_RetMsg='�Ѽƿ��~'
                            -2: p_RetMsg='SELECT �ɵo�Ϳ��~:
                            -3: p_RetMsg='INSERT �ɵo�Ϳ��~:
                           -99: p_RetMsg='��L���~'

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

    type CurTyp is ref cursor;	--�ۭqcursor���A
    cDynamic CurTyp;          	--��dynamic SQL

BEGIN
    if p_CompCode Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:���q�O���ରNull��';
      Return -1;
    end if;

----------------------------------------------------------------------    
-------------------------   �p��I�O�W�D   ---------------------------
----------------------------------------------------------------------
    --1.�զXSQL�d�ߦr��:(SELECT �X�Ҧ�SO122�I�O�W�D�������)
    v_Column := 'select ComputeYM,BelongYM,CustId,BillNo,ITEM from SO122 ';

    v_Where  := ' WHERE CompCode='||p_CompCode|| ' and BuyOrRent is null '||
                ' and BillNo is not null';

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
      x_BillNo := ' ';
      x_OrderNo := '';

      FETCH cDynamic INTO x_ComputeYM,x_BelongYM,x_CustId,x_BillNo,x_Item;
      EXIT WHEN cDynamic%NOTFOUND;

      --�P�O�ӵ��I�O�W�D���,�O�_���ۨӹq
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

      --�p�G���ȫh�P�O�O�_���ۨӹq,�æ^��SO122
      IF ((x_OrderNo<>'') OR (x_OrderNo IS NOT NULL))  AND (x_RefNo IS NOT NULL) THEN
        IF (x_RefNo=1) OR (x_RefNo=2) OR (x_RefNo=3) OR (x_RefNo=5) THEN
          x_CallSelf := 0;--�D�ۨӹq
        ELSE
          x_CallSelf := 1;--�ۨӹq
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
----------------------------   �p�� STB   ----------------------------
----------------------------------------------------------------------
    --1.�զXSQL�d�ߦr��:(SELECT �X�Ҧ�SO122�I�O�W�D�������)
    v_Column := 'select ComputeYM,BelongYM,CustId,StbNo,Sno,StaffCode from SO122 ';

    v_Where  := ' WHERE CompCode='||p_CompCode|| ' and BuyOrRent is not null '||
                ' and BillNo is null';

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
      x_BillNo := ' ';
      x_OrderNo := '';
                                                       
      FETCH cDynamic INTO x_ComputeYM,x_BelongYM,x_CustId,x_StbNo,x_Sno,x_StaffCode;
      EXIT WHEN cDynamic%NOTFOUND;

      --�P�O�ӵ�STB���,�O�_���ۨӹq
      IF x_StaffCode = 1 THEN
        x_CallSelf := 0;--�D�ۨӹq
      ELSIF x_StaffCode = 2 THEN
        x_CallSelf := 1;--�ۨӹq
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
    p_RetMsg := '����,�@�B�z ' ||v_Counts||' �����';
    RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- Function SF_UpdateSO122Data
/
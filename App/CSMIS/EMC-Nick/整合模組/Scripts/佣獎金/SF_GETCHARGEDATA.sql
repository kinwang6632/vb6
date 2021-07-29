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

  �����p��h�O��Ƨ䥿�����,����Table�����������
  �ɦW: SF_GetChargeData.sql


  IN    p_Table       varchar2: Table�W��( SO122��SO134 )
        p_CompCode    varchar2: ���q�O�N�X
        p_BelongYM    varchar2: �k�ݦ~��
        p_BillNo      varchar2: ��ڽs��
        p_Item        varchar2: ��ڶ���

  OUT  p_SeqNo          number:   ���G��Ƨ帹
       p_RetMsg         varchar2: ���G�T�� (�ܤ�200 bytes)

       Return           number:   �B�z���G�N�X, �����T���s��p_RetMsg��
                           >=0: p_RetMsg='���槹��,���S�����'
                             1; p_RetMsg= �N��쪺��Ʀ�_�ӶǦ^�e��
                            -1: p_RetMsg='�Ѽƿ��~'
                            -2: p_RetMsg='SELECT �ɵo�Ϳ��~:
                            -3: p_RetMsg='INSERT �ɵo�Ϳ��~:
                           -99: p_RetMsg='��L���~'

    By: Jackal   Date: 2003.05.23
    Modify By Jackal : 2003.05.30  �NWHERE����h��FirstFlag��BelongYM
    Modify By Jackal : 2003.09.05  �h�^��OrderNo��CallSelf
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

    type CurTyp is ref cursor;	--�ۭqcursor���A
    cDynamic CurTyp;          	--��dynamic SQL

BEGIN
    if p_Table Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:Table���ରNull��';
      Return -1;
    end if;

    if p_CompCode Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:���q�O���ରNull��';
      Return -1;
    end if;


    if p_BelongYM Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:�k�ݦ~�뤣�ରNull��';
      Return -1;
    end if;


    if p_BillNo Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:��ڽs�����ରNull��';
      Return -1;
    end if;


    if p_Item Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:��ڶ������ରNull��';
      Return -1;
    end if;


    --1.�զXSQL�d�ߦr��:
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
  p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- Function SF_GETCHARGEDATA
/
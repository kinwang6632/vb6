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

  ���ͩ�b���O��ƥ������Ӹ��
  �ɦW: SF_CALCULATESO131.sql


  IN   p_CompCode          varchar2: ���q�O�N�X
       p_ServiceType       varchar2: �A�����O
       p_ComputeYM         varchar2: �p��~��
       p_StartDate         varchar2: �ӭp��~�몺�Ĥ@��
       p_EndDate           varchar2: �ӭp��~�몺�ī�@��
       p_ChargeItemSQL     varchar2: �n�d�ߪ����O���اO
       p_RealOrShouldDate  number: �H�ꦬ���������������d�߹�H (0:�ꦬ���B  1:�������B)


  OUT  p_SeqNo          number:   ���G��Ƨ帹
       p_RetMsg         varchar2: ���G�T�� (�ܤ�200 bytes)

       Return           number:   �B�z���G�N�X, �����T���s��p_RetMsg��
                           >=0: p_RetMsg='���槹��, �B�z����=<����>'
                            -1: p_RetMsg='�Ѽƿ��~'
                            -2: p_RetMsg='SELECT �ɵo�Ϳ��~:
                            -3: p_RetMsg='INSERT �ɵo�Ϳ��~:
                           -99: p_RetMsg='��L���~'

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


  type CurTyp is ref cursor;	--�ۭqcursor���A
  cDynamic CurTyp;          	--��dynamic SQL

BEGIN
    if p_CompCode Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:���q�O���ରNull��';
      Return -1;
    end if;

    if p_ServiceType Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:�A�ȧO���ରNull��';
      Return -1;
    end if;

    if p_ComputeYM Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:�p��~�뤣�ରNull��';
      Return -1;
    end if;

    if p_ChargeItemSQL Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:�d�߭��Ǧ��O���ؤ��ରNull��';
      Return -1;
    end if;

    if p_RealOrShouldDate Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:�H�����ιꦬ������d�߱���?';
      Return -1;
    end if;

    if p_StartDate Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:�}�l�ɶ����ରNull��';
      Return -1;
    end if;

    if p_EndDate Is Null then
      p_RetCode := -1;
      p_RetMsg := '�Ѽƿ��~:�����ɶ����ରNull��';
      Return -1;
    end if;


    --���NSO131���ƪ���ƧR��
    DELETE SO131 WHERE COMPUTEYM=p_ComputeYM AND RealOrShouldDate=p_RealOrShouldDate
                       AND COMPCODE=p_CompCode;

    FOR i IN 1 .. 2 LOOP
    --1.�զXSQL�d�ߦr��:
      --���w��SO033��SO034�d�ߩ��Ӹ��
      IF i = 1 THEN
        v_Table := 'SO033';
      ELSIF i = 2 THEN
        v_Table := 'SO034';
      END IF;


      v_Column := 'select BillNo,Item,CustId,CitemCode,CitemName,ShouldDate, '
                ||'NVL(ShouldAmt,0) ShouldAmt,RealDate,NVL(RealAmt,0) RealAmt, '
                ||'RealPeriod,RealStartDate,RealStopDate,CompCode,OrderNo, '
                ||'SBillNo,SItem  from '||v_Table;

      IF p_RealOrShouldDate = 0 THEN  --�ꦬ���
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (RealDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| p_ChargeItemSQL||')';

      ELSIF p_RealOrShouldDate = 1 THEN  --�������
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (ShouldDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| p_ChargeItemSQL||')';

      END IF;
      
      --v_Where := v_Where || ' AND CancelCode IS NULL AND CancelName IS NULL';
      v_Where := v_Where || ' AND CancelFlag=0';

      v_SQL := v_Column || v_Where;


    --2.�}��Dynamic sql
      BEGIN

        IF NOT cDynamic%ISOPEN then
          OPEN cDynamic FOR v_SQL;
        END IF;

      EXCEPTION
        WHEN OTHERS THEN
          rollback;
          p_RetCode := -2;
          p_RetMsg := 'SELECT �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM||'      SQL='||v_SQL;
          RETURN -2;
      END;

    --3.loop���C�@����ƨӳB�z
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
              
            
            v_Notes := '�W�D���O';
          EXCEPTION
            WHEN others then
              p_RetCode := -2;
              p_RetMsg := 'SELECT �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
              RETURN -2;
          END;
        ELSE
          v_Notes := '�q��渹���Ū�,�L�k���������C�餶�Ь������';
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
             p_RetMsg := 'INSERT�ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
             RETURN -3;
        END;

      END LOOP;  --LOOP SO033��SO034�����Ӹ��

    --4.Close dynamic cursor
      CLOSE cDynamic;

    END LOOP;    --LOOP SO033��SO034


    COMMIT;

    p_RetCode := 0;
    p_RetMsg := '';
    RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  RETURN -99;
END; -- SF_CALCULATESO131
/
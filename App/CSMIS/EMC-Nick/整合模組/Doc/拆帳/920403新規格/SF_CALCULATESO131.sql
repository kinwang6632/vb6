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

    By: Jackal   Date: 2003.04.09
    Modify By Jackal : 2003.04.24  �p��h�O��ƭ�g�b�e��,�{��b��ݼg
    Modify By Jackal : 2003.05.19  ���SO131�ɥ[��RefNo(1=�b�b,2=���`)
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
  x_STCode              Number;

  x_ReturnCitemCode       varchar2(10);
  x_ReturnCitemCodeString varchar2(200) := ' ';


  z_BillNo              varchar2(15);
  z_Item                Number;
  z_CitemCode           Number;
  z_CitemName           varchar2(20);
  z_MediaCode           Number;
  z_MediaName           varchar2(20);
  z_AcceptEn            varchar2(10);
  z_AcceptName          varchar2(20);
  z_IntroId             varchar2(10);
  z_IntroName           varchar2(50);
  z_ComputeYM           varchar2(5);
  z_SeqNo               Number;
  z_RealStopDate        Date;
  z_OrderNo             varchar2(15);


  --v_BeginMonth          Number;
  --v_EndMonth            Number;
  v_ServiceType         char(1) := UPPER(p_ServiceType);
  v_Notes               varchar2(255);
  v_TmpCitem            varchar2(255);
  v_CitemCodeCnt        Number;
  v_CurSeqNO            Number := 0;
  v_MaxSeqNO            Number := 0;
  v_IsSelectedCitemCode char(1) := 'Y';  --�P�_�O�_���ҿ諸���O����
  v_RealStopDate        Date;
  v_RealStartDate       Date;

  --�b�b�B�z
  v_CheckSTCode         Number;
  v_RefNo               Number;


  c39                   CHAR(1) := CHR(39);

  cursor cc1 is
    SELECT CODENO,DESCRIPTION FROM CD019 WHERE  REFNO=9 AND SERVICETYPE=p_ServiceType;

  cursor cc2 is
    Select Max(SEQNO) MaxSeqNO from SO131 where ComputeYM=p_ComputeYM and
                                       RealOrShouldDate=p_RealOrShouldDate and
                                       CompCode=p_CompCode;
  cursor cc3 is
    select CodeNo FROM cd016 where refno=1;

  type CurTyp is ref cursor;	--�ۭqcursor���A
  cDynamic CurTyp;          	--��dynamic SQL


  I number;
  v_Index  number;

  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER;
  CitemCode NumberAry;

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

    --���X�b�b�N�X
    for c3 in cc3 loop
      v_CheckSTCode := c3.CodeNo;
    end loop;


-----**********************************************************************-----
-----                               �p�⦬�O���                           -----
-----**********************************************************************-----

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
                ||'SBillNo,SItem,STCode  from '||v_Table;

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
          ROLLBACK;
          p_RetCode := -2;
          p_RetMsg := 'SELECT �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM||'      SQL='||v_SQL;
          RETURN -2;
      END;

    --3.loop���C�@����ƨӳB�z
      LOOP
        FETCH cDynamic INTO x_BillNo,x_Item,x_CustId,x_CitemCode,x_CitemName,
                            x_ShouldDate,x_ShouldAmt,x_RealDate,x_RealAmt,x_RealPeriod,
                            x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                            x_SBillNo,x_SItem,x_STCode;
        EXIT WHEN cDynamic%NOTFOUND;

        --�P�w������ƬO�_���b�b
        IF x_STCode = v_CheckSTCode THEN
          v_RefNo := 1;    --�b�b
        ELSE
          v_RefNo := 2;    --���`
        END IF;


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
              v_Notes := '�W�D���O,���q��渹�䤣��������C�餶�Ь������';
          END;
        ELSE
          v_Notes := '�W�D���O,���q��渹���ŭ�,�L�k���������C�餶�Ь������';
        END IF;



        BEGIN
            v_CurSeqNO := v_CurSeqNO + 1;
            insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                   CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                   REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                   SBILLNO,SITEM,MEDIACODE,MEDIANAME,ACCEPTEN,ACCEPTNAME,
                   INTROID,INTRONAME,Notes,RefNo)
            Values (v_CurSeqNO,p_ComputeYM,p_RealOrShouldDate,x_BillNo,x_Item,x_CustId,
                   x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                   x_RealPeriod,x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                   x_SBillNo,x_SItem,x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,
                   x_IntroId,x_IntroName,v_Notes,v_RefNo);
        EXCEPTION
        WHEN OTHERS THEN
             ROLLBACK;
             p_RetCode := -3;
             p_RetMsg := 'INSERT�ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
             RETURN -3;
        END;

      END LOOP;  --LOOP SO033��SO034�����Ӹ��

    --4.Close dynamic cursor
      CLOSE cDynamic;

    END LOOP;    --LOOP ���OSO033��SO034

-----**********************************************************************-----
-----                               �p��h�O���                           -----
-----**********************************************************************-----
    -- ���o�Ҧ����h�O���O���إN�X
    for c1 in cc1 loop
      x_ReturnCitemCode := c1.CodeNo;
      if x_ReturnCitemCodeString = ' ' then
        x_ReturnCitemCodeString := x_ReturnCitemCode;
      else
        x_ReturnCitemCodeString := x_ReturnCitemCodeString || ',' || x_ReturnCitemCode;
      end if;
    end loop;

    --********* down  �NCitemCode��Ѧ��}�C *********---
    if (p_ChargeItemSQL is not null) or (p_ChargeItemSQL <> '')then
      v_TmpCitem := p_ChargeItemSQL;
      v_Index := 1;
      I := 1;

      while v_TmpCitem is not null loop
        v_Index := instr(v_TmpCitem, ',');
        if v_Index > 0 then
  	begin
  	  CitemCode(I) := ltrim(rtrim(substr(v_TmpCitem, 1, v_Index-1)));
  	  v_TmpCitem := substr(v_TmpCitem, v_Index+1);
  	  I:=I+1;
  	exception
  	  when others then
  	    v_TmpCitem := null;
  	end;
        else
  	CitemCode(I) := to_number(rtrim(ltrim(v_TmpCitem)));
  	v_TmpCitem := null;
        end if;
      end loop;
      v_CitemCodeCnt := I;

    end if;

    FOR I IN 1 .. v_CitemCodeCnt LOOP
      v_TmpCitem := CitemCode(I);
      v_TmpCitem := v_TmpCitem;
    END LOOP;
    --********* up  �NCitemCode��Ѧ��}�C *********---



    FOR i IN 1 .. 2 LOOP
    --1.�զXSQL�d�ߦr��:
      --���w��SO033��SO034�d�߰h�O���Ӹ��
      IF i = 1 THEN
        v_Table := 'SO033';
      ELSIF i = 2 THEN
        v_Table := 'SO034';
      END IF;


      v_Column := 'select BillNo,Item,CustId,CitemCode,CitemName,ShouldDate, '
                ||'NVL(ShouldAmt,0) ShouldAmt,RealDate,NVL(RealAmt,0) RealAmt, '
                ||'RealPeriod,RealStartDate,RealStopDate,CompCode,OrderNo, '
                ||'SBillNo,SItem,STCode  from '||v_Table;

      IF p_RealOrShouldDate = 0 THEN  --�ꦬ���
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (RealDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| x_ReturnCitemCodeString ||')';

      ELSIF p_RealOrShouldDate = 1 THEN  --�������
        v_Where  := ' WHERE CompCode='||p_CompCode||' and ServiceType='||c39||v_ServiceType||c39||
                     ' and (ShouldDate between to_date('||p_StartDate||','||c39||'YYYYMMDD'||c39||') and' ||
                     ' to_date('||p_EndDate||','||c39||'YYYYMMDD'||c39||')) and '||
                     'CitemCode in ('|| x_ReturnCitemCodeString ||')';

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
          ROLLBACK;
          p_RetCode := -2;
          p_RetMsg := 'SELECT �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM||'      SQL='||v_SQL;
          RETURN -2;
      END;

    --3.loop���C�@���h�O���Ӹ�Ƹ�ƨӳB�z
      LOOP
        FETCH cDynamic INTO x_BillNo,x_Item,x_CustId,x_CitemCode,x_CitemName,
                            x_ShouldDate,x_ShouldAmt,x_RealDate,x_RealAmt,x_RealPeriod,
                            x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                            x_SBillNo,x_SItem,x_STCode;
        EXIT WHEN cDynamic%NOTFOUND;


        --�P�w������ƬO�_���b�b
        IF x_STCode = v_CheckSTCode THEN
          v_RefNo := 1;    --�b�b
        ELSE
          v_RefNo := 2;    --���`
        END IF;


/*
���ˬdSBillNo���S����
<�@>�Y SBillNo �S���ȡA���� Append �b�̫�@��
  <1>�YOrderNo�d�X���������C�餶�и�ơA�hSO131.Notes=��l��ڽs���S����

  <2>�YOrderNo�d�X�S���������C�餶�и�ơA�hSO131.Notes=��l��ڽs���S���� (�B�q��渹�䤣��������C�餶�Ь������)

  <3>�YOrderNo�ڥ��N�S���ȡA�hSO131.Notes=��l��ڽs���S���� (�B�q��渹���ŭ�,�L�k���������C�餶�Ь������)

<�G>�Y SBillNo ����,SELECT ��� SO131
  <1>�Y SELECT ������󥿶� , ���� Append �b�̫�@�� ,
      <A>�YOrderNo�d�X���������C�餶�и�ơA�hSO131.Notes=�䤣�쥿�����,�i�ভ������Ʃ|���p��

      <B>�YOrderNo�d�X�S���������C�餶�и�ơA�hSO131.Notes=�䤣�쥿�����,�i�ভ������Ʃ|���p�� (�B�q��渹�䤣��������C�餶�Ь������)

      <C>�YOrderNo�ڥ��N�S���ȡA�hSO131.Notes=�䤣�쥿�����,�i�ভ������Ʃ|���p�� (�B�q��渹���ŭ�,�L�k���������C�餶�Ь������)

  <2>�Y SELECT ������ , �ˬd��������ƬO�_���e���ҿ諸���O����
      <A>�Y���O�ҿ怜�O���� , �h�������h�O����

      <B>�Y�O�ҿ怜�O���� , �ˬd��������ƭp��~�� , �O�_���e���ҿ�n�p�⪺�p��~��
          <a>�Y�O�h�s�W�@����Ʀܦ��������O���ت��U�� , SO131.Notes=�h�O
          <b>�Y���O�h�N�����ΰh�O���ؤ��O�ݤ��P�~�� , ���� Append �b�̫�@�� , SO131.Notes=����S�����h�O���ت��������
*/





        if (x_SBillNo <> '') OR (x_SBillNo IS NOT NULL)then
          BEGIN
            -- ��M SO131 �����O�_�����������
            Select BillNo,Item,SeqNo,ComputeYM,CitemCode,CitemName,MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName,RealStopDate,OrderNo
              into z_BillNo,z_Item,z_SeqNo,z_ComputeYM,z_CitemCode,z_CitemName,z_MediaCode,z_MediaName,z_AcceptEn,z_AcceptName,z_IntroId,z_IntroName,z_RealStopDate,z_OrderNo
              from SO131
              where RealOrShouldDate=p_RealOrShouldDate and
                    CompCode=p_CompCode and BillNo=x_SBillNo and Item=x_SItem;

            --�Y��������������
            --�ˬd�O�_���ҿ�n�p�⪺���O���إN�X,���O���ܫh���p��
            FOR I IN 1 .. v_CitemCodeCnt LOOP
              v_TmpCitem := CitemCode(I);
              if z_CitemCode = v_TmpCitem then
                v_IsSelectedCitemCode := 'Y';
                --�Y�O�ҿ諸���O���ظ��
                goto GO_OUT;
              else
                v_IsSelectedCitemCode := 'N';
               end if;
            END LOOP;

            <<GO_OUT>>
            NULL;

            --CitemCode ���ҿ諸�M�椤
            if v_IsSelectedCitemCode = 'Y' then
              --���������~�묰�p��~��
              if z_ComputeYM = p_ComputeYM then
                --�N�����h�O��ƥ[������������U��
                BEGIN
                  update SO131 set SEQNO=seqno + 1 where ComputeYM=p_ComputeYM and
                                       RealOrShouldDate=p_RealOrShouldDate and
                                       CompCode=p_CompCode and SeqNo>z_SeqNo;
                EXCEPTION
                WHEN OTHERS THEN
                   ROLLBACK;
                   p_RetCode := -4;
                   p_RetMsg := 'UPDATE�ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                   RETURN -4;
                END;


                v_CurSeqNO := z_SeqNo + 1;

                v_Notes := '�h�O';
              else
                v_Notes := '����S�����h�O���ت����O����';
                -- ����S�����h�O���ت����O����,�h����Append�@���̫ܳ�
                -- ���o�̤j��SEQNO
                for c2 in cc2 loop
                  v_MaxSeqNO := c2.MaxSeqNO + 1;
                end loop;

                v_CurSeqNO := v_MaxSeqNO;
              end if;


              --���o�h�O��ƪ��h�O�}�l�ε������
              --�p�G�h�O�� RealStartDate �O null
              --�h RealStartDate ��h�O�� ShouldDate,RealStopDate �䥿���� RealStopDate
              if x_RealStartDate is not null then
                v_RealStartDate := x_RealStartDate;
                v_RealStopDate := x_RealStopDate;
              else
                v_RealStartDate := x_ShouldDate;
                v_RealStopDate := z_RealStopDate;
              end if;


              BEGIN
                  insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                         CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                         REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                         SBILLNO,SITEM,MEDIACODE,MEDIANAME,ACCEPTEN,ACCEPTNAME,
                         INTROID,INTRONAME,Notes,RefNo)
                  Values (v_CurSeqNO,p_ComputeYM,p_RealOrShouldDate,z_BillNo,z_Item,x_CustId,
                         x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                         x_RealPeriod,v_RealStartDate,v_RealStopDate,x_CompCode,x_OrderNo,
                         x_SBillNo,x_SItem,z_MediaCode,z_MediaName,z_AcceptEn,z_AcceptName,
                         z_IntroId,z_IntroName,v_Notes,v_RefNo);
              EXCEPTION
              WHEN OTHERS THEN
                   ROLLBACK;
                   p_RetCode := -3;
                   p_RetMsg := 'INSERT�ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                   RETURN -3;
              END;

            end if;

          EXCEPTION
            WHEN others then
              --�Y SO131 �䤣��������������,�N��|��������ƥ��p��,�h����Append�@���̫ܳ�
              -- ���o�̤j��SEQNO
              for c2 in cc2 loop
                v_MaxSeqNO := c2.MaxSeqNO + 1;
              end loop;


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

                  v_Notes := '�䤣�쥿�����,�i�ভ������Ʃ|���p��';
                EXCEPTION
                  WHEN others then
                    v_Notes := '�䤣�쥿�����,�i�ভ������Ʃ|���p�� (�B�q��渹�䤣��������C�餶�Ь������)';
                END;
              ELSE
                v_Notes := '�䤣�쥿�����,�i�ভ������Ʃ|���p�� (�B�q��渹���ŭ�,�L�k���������C�餶�Ь������)';
              END IF;




              BEGIN
                  insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                         CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                         REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                         SBILLNO,SITEM,MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName,Notes,RefNo)
                  Values (v_MaxSeqNO,p_ComputeYM,p_RealOrShouldDate,x_BillNo,x_Item,x_CustId,
                         x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                         x_RealPeriod,x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                         x_SBillNo,x_SItem,x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,x_IntroId,x_IntroName,v_Notes,v_RefNo);
              EXCEPTION
              WHEN OTHERS THEN
                   ROLLBACK;
                   p_RetCode := -3;
                   p_RetMsg := 'INSERT�ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                   RETURN -3;
              END;
          END;


        else
          -- ���o�̤j��SEQNO
          for c2 in cc2 loop
            v_MaxSeqNO := c2.MaxSeqNO + 1;
          end loop;


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

                  v_Notes := '��l��ڽs���S����';
            EXCEPTION
              WHEN others then
                v_Notes := '��l��ڽs���S���� (�B�q��渹�䤣��������C�餶�Ь������)';
            END;
          ELSE
            v_Notes := '��l��ڽs���S���� (�B�q��渹���ŭ�,�L�k���������C�餶�Ь������)';
          END IF;


          BEGIN
              insert into SO131(SEQNO,COMPUTEYM,REALORSHOULDDATE,BILLNO,ITEM,CUSTID,
                     CITEMCODE,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT,
                     REALPERIOD,REALSTARTDATE,REALSTOPDATE,COMPCODE,ORDERNO,
                     SBILLNO,SITEM,MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName,Notes,RefNo)
              Values (v_MaxSeqNO,p_ComputeYM,p_RealOrShouldDate,x_BillNo,x_Item,x_CustId,
                     x_CitemCode,x_CitemName,x_ShouldDate,NVL(x_ShouldAmt,0),x_RealDate,NVL(x_RealAmt,0),
                     x_RealPeriod,x_RealStartDate,x_RealStopDate,x_CompCode,x_OrderNo,
                     x_SBillNo,x_SItem,x_MediaCode,x_MediaName,x_AcceptEn,x_AcceptName,x_IntroId,x_IntroName,v_Notes,v_RefNo);
          EXCEPTION
          WHEN OTHERS THEN
               ROLLBACK;
               p_RetCode := -3;
               p_RetMsg := 'INSERT�ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
               RETURN -3;
          END;

        end if;

      END LOOP;--�h�O���Ӹ�Ƹ�ƨӳB�z

      CLOSE cDynamic;

    END LOOP;    --LOOP �w��SO033��SO034�h�O




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
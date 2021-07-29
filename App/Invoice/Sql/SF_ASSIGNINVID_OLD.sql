CREATE OR REPLACE FUNCTION SF_ASSIGNINVID
  ( p_User                 VARCHAR2,
    p_CompCode             NUMBER,
    p_LinkToMis            VARCHAR2,
    p_DbLink               VARCHAR2,
    p_HowToCreate          NUMBER,
    p_InvDateEqualToChargeDate NUMBER,
    p_InvDate              VARCHAR2,
    p_InvYearMonth         VARCHAR2,
    p_ChargeStartdate      VARCHAR2,
    p_ChargeStopDate       VARCHAR2,
    p_IdentifyID1          NUMBER,
    p_IdentifyID2          NUMBER,
    p_SystemID             VARCHAR2,
    p_PrefixString         VARCHAR2,
    p_OrderBy              NUMBER,
    p_MisDbOwner           VARCHAR2,
    p_RetCode              OUT NUMBER,
    p_RetMsg               OUT VARCHAR2,
    p_LogDateTime          OUT VARCHAR2
  )
RETURN NUMBER AS

/*
  --@C:\Invoice\Script\SF_AssignInvID.sql;
  VARIABLE RetMsg VARCHAR2(4000)
  VARIABLE RetCode number
  VARIABLE ReturnNo NUMBER
  EXEC :ReturnNo := SF_AssignInvID('A',1,2,2,'2003/03/29','200303','2003/03/01','2003/03/31',1,0,'01','AA10000001,BB10000001,CC10000001',0 ,'v30',:RetCode ,:RetMsg)
  PRINT ReturnNo
  PRINT RetCode
  PRINT RetMsg
  --EXEC :ReturnNo := SF_AssignInvID('A',1,1,2,'2002/12/09','200211','2002/12/07','2002/12/08',1,0,'01','QW',0 ,'v30',:RetCode ,:RetMsg)


  �o���}��
  �ɦW: SF_AssignInvID.sql

  ����:�B�z�o���t�Τ��o���}��(����)


  IN
	p_User          		Varchar2 :    �{���ާ@��
	p_CompCode      		number :      ���q�O�N�X
	p_HowToCreate   		number :      1=>�w�};   2=>��}
  p_InvDateEqualToChargeDate	number :  1=>�o��������󦬶O���;   2=>�o����������󦬶O���
	p_InvDate			      Varchar2 :    �o�����
	p_InvYearMonth 			Varchar2 :    �o���~��
	p_ChargeStartdate 	Varchar2 :    ���O�_�l��
	p_ChargeStopDate 		Varchar2 :    ���O�I���
	p_IdentifyID1 			number :      �ѧO�X(�O�d), ���ӵ���1
	p_IdentifyID2 			number :      �ѧO�X(�O�d), ���ӵ���0
	p_SystemID 			    Varchar2 :    System ID
  p_PrefixString 		  Varchar2 :    �o���r�y
	p_OrderBy 			    number :      �ƧǤ覡 : 0=>�̦��O����}�� ; 1=>�̶l���ϸ��}��
  p_MisDbOwner        Varchar2 :    �^���MIS��Table Space �� Owner


  OUT
	p_RetMsg        		varchar2:     ���G�T�� (�ܤ�200 bytes)
	p_RetCode       		number:       ���G�N�X


  Return
	number: �B�z���G�N�X, �����T���s��p_RetMsg��

         0:  p_RetMsg='�o���}�߰��槹��'
        -1:  p_RetMsg='�Ѽƿ��~'
       -99:  p_RetMsg='��L���~'


    By: Howard
    Date: 2002.12.23
    Modify By Jackal : 2003.03.11  ��@���ȯ�w��@���o���}��,�令�i�H�P�ɦh���o���}��
    Modify By Jackal : 2003.04.19  1.�z��Y�o�����B <0 �����ζ}�o���� Log ��Ʀ� INV033
                                   2.�ˬd��ƬO�_�w�}�L�o��,�Y�}�L���ζ}�o���� Log ��Ʀ� INV033
    Modify By Jackal : 2003.04.24  INV016 �h�@��� ShouldBeAssigned , �Ω�P�_�����O���جO�_�n�}�ߵo��
    Modify By Jackal : 2003.07.22  �h�Ǥ@�ӰѼ�p_MisDbOwner
    Modify By Jackal : 2004.01.16  �N�ܼ�s_CustSName��Varchar2(26)�אּVarchar2(50)
    Modify By Jackal : 2004.05.13  1.�YInv016�D�ɻPInv017�����ɪ��X�p���B���X,�h���}�ߨ�Log �� inv033
                                   2.�ˮ֬O�_���ۦPseq,�����O���Ӫ��|�v�O���P������}�ߨ�Log �� inv033
                                   3.�}�ߦ^��ɭY�ӵ���Ƥ��s�b so034��so033 ��,�h���}�ߨ�Log �� inv033
                                   4.�}�ߦ^��ɭY�ӵ���ƨ�guino �w���Ȯ�(�N��w�}�߹L),�h���}�ߨ�Log
    Modify By Jackal : 2004.05.14  INV008��INV032�U�[�@�ӾP���B���(SaleAmount),�P���B=���*�ƶq(�å|�ˤ��J�����)
    Modify By Nick   : 2005.01.17  �o���ˬd�X�p�⤣��, �o������ INV099 STARTNUM, CURNUM, ENDNUM ���令 VARCHAR2(8)
    Modify By Nick   : 2005.04.22  1.����U�}�ߵo����Ʀh�[�@���P�_, �|�O TAXTYPE = 0 , �����}��
    Modify By Nick   : 2005.09.30  �w�� EMC �ϥ�, ��ȪA�t�θ�ƼW�[ DBLink �r��
    Modify By Nick   : 2005.10.04  �W�[�J�Ѽ�, p_Dblink, p_linkToMis,
                                   �Y p_linkToMis = Y , �~��s�ȪA�t�θ��,
                                   �Y p_Dblink ����, �h sql �h�[�J dblink �Ҧp: select * from So033@catvc
    Modify By Nick   : 2005.10.20  �ץ��]���ˬd���O���ح��ЦӾɭP����}�ߵo��(�h�P�_CompCode)
    Modify By Nick   : 2006.01.16  INV033 �W�[ CompCode ���q�O���, �h���q�}�ߵo������, ���`����~�|�����

*/
  n_PreFixKind          NUMBER(2) :=0;  --�e�ݶǨӪ��r�y���X��
  v_PrefixString        VARCHAR2(500);
  i                     NUMBER(2) :=0;
  v_PreFixStartNum      VARCHAR2(20);
  v_PreFix              VARCHAR2(2);
  v_StartNum            VARCHAR2(8);
  v_CurIndex            NUMBER(2) :=0; --�ثe�O�ĴX�զr�y


  TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  aPreFix Varchar2Ary;	-- array for PreFix
  aStartNum Varchar2Ary;	-- array for StartNum

  s_BeAssignedInvID     CHAR(1);

  s_RealInvID           VARCHAR2(10);
  aCheckNo              NUMBER(3);


  s_Useful              VARCHAR2(1);
  s_Where               VARCHAR2(4000);
  s_SQL                 VARCHAR2(4000);
  s_OrderBy             VARCHAR2(20);
  n_CurNum              NUMBER;
  n_EndNum              NUMBER;
  s_RealInvDate         VARCHAR2(10);
  s_CustSName           VARCHAR2(50);
  n_RecCount            NUMBER;
  s_Tmp1                VARCHAR2(10);
  s_Tmp2                VARCHAR2(10);
  s_CurDateTime         VARCHAR2(20);
  s_LastInvDate         VARCHAR2(10);
  s_SQLInv017           VARCHAR2(4000);

  s_SEQ                 VARCHAR2(15);
  s_CustId              VARCHAR2(8);
  s_Tel                 VARCHAR2(20);
  s_BusinessID          VARCHAR2(8);
  s_Title               VARCHAR2(50);
  s_InvAddr             VARCHAR2(60);
  s_MailAddr            VARCHAR2(60);
  s_ZipCode             VARCHAR2(5);
  s_ChargeDate          VARCHAR2(10);
  s_TaxType             VARCHAR2(1);
  n_TaxRate             NUMBER(5,2);
  n_SaleAmount          NUMBER(12);
  n_TaxAmount           NUMBER(12);
  n_InvAmount           NUMBER(12);
  s_ChargeTitle         VARCHAR2(50);
  n_Inv032SaleAmount    NUMBER(12);


  n_017Seq              NUMBER(2);
  s_017BillId           VARCHAR2(15);
  n_017BillIdItemNo     NUMBER(2);

  s_017TaxType          VARCHAR2(1);
  s_017ChargeDate       VARCHAR2(10);

  s_017ItemID           VARCHAR2(7);
  s_017Desc             VARCHAR2(40);

  n_017Quantity         NUMBER(6);
  n_017InitPrice        NUMBER(12,2);
  n_017TaxAmount        NUMBER(10);
  n_017TotalAmount      NUMBER(10);

  s_017StartDate        VARCHAR2(10);
  s_017EndDate          VARCHAR2(10);
  s_017ChargeEn         VARCHAR2(20);
  s_017ServiceType      VARCHAR2(1);

  s_Notes               VARCHAR2(255);
  s_RepeatInvID         VARCHAR2(10) := '';
  s_RepeatInvIDList     VARCHAR2(255) := '';

  s_TempSQL             VARCHAR2(4000);
  n_Check016InvAmount   NUMBER(12);
  n_Check017InvAmount   NUMBER(12);
  s_Check016TaxType     VARCHAR2(1);
  s_HaveSameTaxType     VARCHAR2(1);
  s_CheckTable          VARCHAR2(20);
  s_CheckSO033BillNo    VARCHAR2(20);
  s_CheckSO033GUINo     VARCHAR2(20);

  aDBlink               varchar2(10);


  type CurTyp is ref cursor;	--�ۭqcursor���A
  cDynamic CurTyp;          	--��dynamic SQL
  cDynamic2 CurTyp;          	--��dynamic SQL
  cDynamic3 CurTyp;          	--��dynamic SQL



          /* ------------------------------------------------------------------ */
          /*  �p��o���ˬd�X
          /* ------------------------------------------------------------------ */

          function CalcInvoiceCheckNumber(aInvId in varchar2, aInvParam in varchar2)
             return varchar2
          as
            type IntegerArray is table of varchar2(2) index by binary_integer;
            aCalcResult IntegerArray;
            aLength binary_integer;
            aResult integer := 0;
          begin
            for aIndex in 1..Length( aInvId )
            loop
                aLength := aIndex;
                aCalcResult( aLength ) := Lpad( to_char(
                  substr( aInvId, aIndex, 1 ) * substr( aInvParam, aIndex, 1 ) ), 2, '0' );
            end loop;
            for aIndex in 1..aLength
            loop
               if ( aIndex = 1 ) then
                 aCalcResult( aIndex ) := substr( aCalcResult( aIndex ),
                   length( aCalcResult( aIndex ) ), 1 );
               else
                 aCalcResult( aIndex ) :=
                   ( substr( aCalcResult( aIndex ),
                      length( aCalcResult( aIndex ) ), 1 ) ) +
                   ( substr( aCalcResult( aIndex ),
                      length( aCalcResult( aIndex - 1 ) ) - 1, 1 ) );
               end if;
            end loop;
            for aIndex in 1..aLength
            loop
               aResult := ( aResult + to_number( aCalcResult( aIndex ) ) );
            end loop;
            return aResult;
          end;

          /* ------------------------------------------------------------------- */

          function GetDBLink return varchar2
          as
            aResult varchar2(10);
          begin
            aResult := null;
            if ( p_DbLink is not null ) then
              if substr( p_DbLink, 1, 1 ) <> '@' then
                aResult := '@' || p_DbLink;
              else
                aResult := p_DbLink;
              end if;
            end if;
            return aResult;
          end;

          /* ------------------------------------------------------------------- */


BEGIN

      ---- Nick Add DBLink -----
      aDBlink := GetDBLink;


--�ˬd�Ҧ��ѼƬҭn���� => �Ҧ��Ѽ����Ӧb�I�s�ݴN�ˬd
  if p_CompCode Is Null then
    p_RetMsg := '�Ѽƿ��~:���q�O���ର�ŭ�';
    p_RetCode := -1;
    Return -1;
  end if;

  if p_PrefixString Is Null then
    p_RetMsg := '�Ѽƿ��~:�r�y���ର�ŭ�';
    p_RetCode := -1;
    Return -1;
  else
    --�p��Ƕi�Ӫ��r�y���X��(AA11111111,BB22222222,CC33333333)
    --�s����հ}�C
    v_PrefixString := p_PrefixString;

    while v_PrefixString is not null loop
      n_PreFixKind := instr(v_PrefixString, ',');
      if n_PreFixKind > 0 then
  	    begin
          v_PreFixStartNum := ltrim(rtrim(substr(v_PrefixString, 1, n_PreFixKind)));

          aPreFix(i) := ltrim(rtrim(substr(v_PreFixStartNum, 1, 2)));
          aStartNum(i) := ltrim(rtrim(substr(v_PreFixStartNum, 3, 8)));

	        v_PrefixString := substr(v_PrefixString, n_PreFixKind+1);
          i:=i+1;
        exception
	      when others then
	        v_PrefixString := null;
	      end;
      else
        v_PreFixStartNum := ltrim(rtrim(v_PrefixString));

        aPreFix(i) := ltrim(rtrim(substr(v_PreFixStartNum, 1, 2)));
        aStartNum(i) := ltrim(rtrim(substr(v_PreFixStartNum, 3, 8)));
        v_PrefixString := null;
      end if;
    end loop;
  end if;



  if p_OrderBy = 0 then
    s_OrderBy := 'CHARGEDATE';  -- ���O���
  else
    s_OrderBy := 'ZIPCODE' ; -- �l���ϸ�
  end if;

  --���X�Ĥ@�զr�y�X��

  v_CurIndex := 0;
  v_PreFix := aPreFix(v_CurIndex);
  v_StartNum := aStartNum(v_CurIndex);

  -- ���r�y���
  begin
      SELECT TO_NUMBER( CURNUM ),
              TO_NUMBER( ENDNUM ),
              TO_CHAR( LASTINVDATE,'YYYY/MM/DD' )
         INTO n_CurNum,
               n_EndNum,
               s_LastInvDate
        FROM INV099
       WHERE COMPID = p_CompCode
         AND YEARMONTH = p_InvYearMonth
         AND PREFIX = v_Prefix
         AND USEFUL = 'Y'
         AND STARTNUM = v_StartNum;
  exception
    when others then
      ROLLBACK;
      p_RetMsg := '�L�k���X�r�y���:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
  end;

  -- �R�� INV031/Inv032

  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE INV031');
    DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE INV032');
  exception
    when others then
      ROLLBACK;
      p_RetMsg := '�R��INV031/INV032����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
  end;


  -- select �X INV016 �� data �B InvAmount > 0
  s_Where := ' WHERE COMPID = ' || '''' || p_CompCode || '''' ||
             '   AND CHARGEDATE BETWEEN ' ||
             '       TO_DATE( ' || '''' || p_ChargeStartdate || '''' || ', ''YYYY/MM/DD'' ) ' ||
             '         AND    ' ||
             '       TO_DATE( ' || '''' || p_ChargeStopdate || '''' || ', ''YYYY/MM/DD'' )  ' ||
             '   AND BEASSIGNEDINVID = ''N''     ' ||
		         '   AND ISVALID = ''Y''             ' ||
		         '   AND HOWTOCREATE = ' || '''' || p_HowToCreate || '''' ||
             '   AND INVAMOUNT > 0 ' ||
             '   AND TAXTYPE <> ''0'' ' ||
             '   AND SHOULDBEASSIGNED = ''Y''    ';

  s_SQL :=  ' SELECT SEQ,                            ' ||
            '        CUSTID,                         ' ||
            '        TEL,                            ' ||
            '        BUSINESSID,                     ' ||
            '        TITLE, INVADDR, MAILADDR,       ' ||
            '        ZIPCODE,                        ' ||
            '        TO_CHAR( CHARGEDATE, ''YYYY/MM/DD'' ) CHARGEDATE,  ' ||
            '        TAXTYPE, TAXRATE, SALEAMOUNT,   ' ||
            '        TAXAMOUNT,                      ' ||
            '        INVAMOUNT,                      ' ||
            '        CHARGETITLE                     ' ||
            '   FROM INV016 ' || s_Where || ' ORDER BY ' || s_OrderBy;


  begin
    IF NOT cDynamic%ISOPEN then
      OPEN cDynamic FOR s_SQL;
    END IF;

  exception
    when others then
      ROLLBACK;
      p_RetMsg := '�d��INV016����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
  end;

  select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') into s_CurDateTime from dual;
  p_LogDateTime := s_CurDateTime;


  --Loop�e�����w���
  s_Useful := 'Y';
  --�����w�Ĥ@�Ӧr�y���̫�o�����
  s_RealInvDate := s_LastInvDate;



--loop inv016,���C�@����ƨӳB�z
  LOOP

    FETCH cDynamic INTO s_SEQ,s_CustId,s_Tel,s_BusinessID, s_Title, s_InvAddr, s_MailAddr, s_ZipCode,
                                    s_ChargeDate,s_TaxType,n_TaxRate,n_SaleAmount,n_TaxAmount,n_InvAmount, s_ChargeTitle;

    EXIT WHEN cDynamic%NOTFOUND;


    s_CustSName := s_ChargeTitle; -- 92/09/26 ���s�W��

    --�Ω�D���ɪ��B�ˮ�
    n_Check016InvAmount:= n_InvAmount;
    n_Check017InvAmount := 0;


    s_Check016TaxType := s_TaxType;


--------------------------------------------------------------------------------
-------down loop inv017,�ˬd�C�@����ƬO�_�w�g�}�ߵo��
    s_SQLInv017 := 'SELECT BILLID,BILLIDITEMNO,TAXTYPE, '|| 'TO_CHAR(CHARGEDATE,' || '''' ||
                   'YYYY/MM/DD' || '''' || ') CHARGEDATE,' || 'ITEMID ' ||
		               ',DESCRIPTION,QUANTITY,UNITPRICE,TAXAMOUNT,TOTALAMOUNT,' ||
	               	'TO_CHAR(STARTDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') STARTDATE,' ||
		               'TO_CHAR(ENDDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') ENDDATE,' ||
		               'CHARGEEN,SERVICETYPE from inv017 where SEQ=' || '''' || s_SEQ || '''' ||
                   '   AND SHOULDBEASSIGNED = ''Y''    ';


    begin
      IF NOT cDynamic2%ISOPEN then
        OPEN cDynamic2 FOR s_SQLInv017;
      END IF;
    exception
    when others then
      ROLLBACK;
      CLOSE cDynamic2;
      p_RetMsg := '�d��INV017����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;


    --down loop inv017,�ˬd�C�@����ƬO�_�w�g�}�ߵo��
    --�Ҧ������ƪ��o�����X
    s_RepeatInvIDList := ' ';
    s_HaveSameTaxType := 'Y';
    loop

      FETCH cDynamic2 INTO s_017BillId,n_017BillIdItemNo,s_017TaxType, s_017ChargeDate,s_017ItemID,s_017Desc,
        		n_017Quantity, n_017InitPrice, n_017TaxAmount, n_017TotalAmount,
        		s_017StartDate, s_017EndDate, s_017ChargeEn, s_017ServiceType;

      EXIT WHEN cDynamic2%NOTFOUND;

      --�Ω��ˮ�Inv016�D�ɻPInv017�����ɪ��X�p���B�O�_�@�P
      n_Check017InvAmount := n_Check017InvAmount + n_017TotalAmount;



      --�Ω��ˮ֬O�_���ۦPseq,�����O���Ӫ��|�v�O���P������}��
      if s_Check016TaxType <> s_017TaxType then
        s_HaveSameTaxType := 'N';
      end if;



      --�}�ߦ^��ɭY�ӵ���Ƥ��s�b so034��so033 ��,�άO��guino ���Ȯ�,
      --�Ҥ��ऩ�H�^��,�åB�����H�}��

      IF p_LinkToMis = 'Y' THEN


           if p_HowToCreate = '1' then
             s_CheckTable := 'SO033';
           else
             s_CheckTable := 'SO034';
           end if;

           s_CheckTable := ( s_CheckTable || aDBlink );

           s_CheckSO033BillNo := null;
           s_CheckSO033GUINo := null;


           BEGIN
             s_TempSQL := 'SELECT BILLNO,GUINO FROM ' || p_MisDbOwner || '.' || s_CheckTable ||' WHERE BILLNO=''' ||
                          s_017BillId || ''' AND ITEM ='||n_017BillIdItemNo;

             EXECUTE IMMEDIATE s_TempSQL INTO s_CheckSO033BillNo, s_CheckSO033GUINo;


             if ( s_CheckSO033GUINo IS NOT NULL ) then

                  begin
                    s_Notes := 'MIS����������Ƥw���o�����X';
                    insert into INV033( SEQ, BILLID, BILLIDITEMNO, TAXTYPE, CHARGEDATE, ITEMID,
                       DESCRIPTION, QUANTITY, UNITPRICE, TAXRATE, TAXAMOUNT,
                       TOTALAMOUNT, STARTDATE, ENDDATE, CHARGEEN,NOTES,
                       CUSTID, CUSTNAME, LOGTIME, UPTEN, COMPID )
                    values( s_SEQ, s_017BillId, n_017BillIdItemNo, s_017TaxType,
                       to_date(s_017ChargeDate, ' YYYY/MM/DD' ),
                       s_017ItemID, s_017Desc, n_017Quantity, n_017InitPrice, n_TaxRate,
                       n_017TaxAmount, n_017TotalAmount,
                       to_date( s_017StartDate, ' YYYY/MM/DD' ),
                       to_date( s_017EndDate, ' YYYY/MM/DD' ),
                       s_017ChargeEn, s_Notes, s_CustId, s_CustSName,
                       to_date( s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS' ), p_User, p_CompCode );
                  exception
                  when others then
                    close cDynamic2;
                    rollback;
                    p_RetMsg := 'INSERT INV033 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                    p_RetCode := -99;
                    return -99;
                  end;

               close cDynamic2;
               --�}�ߦ^��ɭY�ӵ���ƨ�guino �w���Ȯ�(�N��w�}�߹L),
               --�h���ऩ�H�^��,�åB�����H�}��
               goto GO_NEXT;
             end if;

           exception
             when others then
             begin
                 s_Notes := 'MIS�S�������������';
                 INSERT INTO INV033( SEQ, BILLID, BILLIDITEMNO, TAXTYPE, CHARGEDATE, ITEMID,
                    DESCRIPTION, QUANTITY, UNITPRICE, TAXRATE, TAXAMOUNT,
                    TOTALAMOUNT, STARTDATE, ENDDATE, CHARGEEN, NOTES,
                    CUSTID, CUSTNAME, LOGTIME, UPTEN, COMPID )
                VALUES(s_SEQ, s_017BillId, n_017BillIdItemNo, s_017TaxType,
                    to_date( s_017ChargeDate, ' YYYY/MM/DD' ),
                    s_017ItemID, s_017Desc, n_017Quantity, n_017InitPrice, n_TaxRate,
                    n_017TaxAmount, n_017TotalAmount,
                    to_date( s_017StartDate, ' YYYY/MM/DD' ),
                    to_date( s_017EndDate, ' YYYY/MM/DD' ),
                    s_017ChargeEn,s_Notes,s_CustId,s_CustSName,
                    to_date( s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS' ), p_User, p_CompCode );
                 close cDynamic2;
                 --�}�ߦ^��ɭY�ӵ���Ƥ��s�b so034��so033 ��,
                 --�h���ऩ�H�^��,�åB�����H�}��
                 goto GO_NEXT;
               exception
               when others then
                 close cDynamic2;
                 rollback;
                 p_RetMsg := 'INSERT INV033 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
                 p_RetCode := -99;
                 return -99;
               end;

               --�}�ߦ^��ɭY�ӵ���Ƥ��s�b so034��so033 ��,
               --�h���ऩ�H�^��,�åB�����H�}��
               goto GO_NEXT;

           end;
      end if;

      --�ˬd�O�_���}�߹L�B�S���o�����
      s_SQL := 'SELECT B.INVID FROM INV007 A,INV008 B WHERE B.BILLID ='|| '''' || s_017BillId || '''' ||
               ' AND B.BILLIDITEMNO=' || n_017BillIdItemNo || ' AND A.ISOBSOLETE=''N'' ' ||
               ' AND B.INVID=A.INVID AND A.COMPID=' || '''' || p_CompCode || '''' || ' ORDER BY B.INVID';




      begin
        IF NOT cDynamic3%ISOPEN then
          OPEN cDynamic3 FOR s_SQL;
        END IF;
      exception
      when others then
        ROLLBACK;
        CLOSE cDynamic3;
        p_RetMsg := '�d��INV008����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      end;



      loop

        FETCH cDynamic3 INTO s_RepeatInvID;
        EXIT WHEN cDynamic3%NOTFOUND;

        if s_RepeatInvIDList = ' ' then
          s_RepeatInvIDList := s_RepeatInvID;
        else
          s_RepeatInvIDList := s_RepeatInvIDList || ' , ' ||s_RepeatInvID;
        end if;
      end loop; --end loop inv008

      CLOSE cDynamic3;
    end loop;  --end loop inv017,�ˬd�C�@����ƬO�_�w�g�}�ߵo��
    CLOSE cDynamic2;






    --�YInv016�D�ɻPInv017�����ɪ��X�p���B���X,�h Log �� inv033
    if n_Check016InvAmount<>n_Check017InvAmount then
      BEGIN
        s_Notes := '�D���ɵo�����B����';
        INSERT INTO INV033(SEQ,BILLID,BILLIDITEMNO,TAXTYPE,CHARGEDATE,ITEMID,
           DESCRIPTION,QUANTITY,UNITPRICE,TAXRATE,TAXAMOUNT,
           TOTALAMOUNT,STARTDATE,ENDDATE,CHARGEEN,NOTES,
           CUSTID,CUSTNAME,LOGTIME,UPTEN, COMPID )
        VALUES( s_SEQ, ' ', 0,s_TaxType,
           to_date( s_ChargeDate, ' YYYY/MM/DD' ),
           ' ','', NULL, NULL, n_TaxRate,
           n_TaxAmount, n_Check016InvAmount,
           NULL, NULL, '', s_Notes, s_CustId, s_CustSName,
           to_date(s_CurDateTime, 'YYYY/MM/DD HH24:MI:SS'), p_User, p_CompCode );

      EXCEPTION
      WHEN OTHERS THEN
        CLOSE cDynamic2;
        ROLLBACK;
        p_RetMsg := 'INSERT INV033 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
        p_RetCode := -99;
        RETURN -99;
      END;


      --�YInv016�D�ɻPInv017�����ɪ��X�p���B���X,
      --�h LOG  �᪽���� INV016 �U�@��
      goto GO_NEXT;
    end if;

   /* ---------------------------------------------------------------------------------- */



    ----�Ω��ˮ֬O�_���ۦPseq,�����O���Ӫ��|�v�O���P,�h Log �� inv033
    if s_HaveSameTaxType = 'N' then
      BEGIN
        s_Notes := '���O���Ӫ��|�v�O���P';
        INSERT INTO INV033(SEQ,BILLID,BILLIDITEMNO,TAXTYPE,CHARGEDATE,ITEMID,
                           DESCRIPTION,QUANTITY,UNITPRICE,TAXRATE,TAXAMOUNT,
                           TOTALAMOUNT,STARTDATE,ENDDATE,CHARGEEN,NOTES,
                           CUSTID,CUSTNAME,LOGTIME,UPTEN, COMPID )
        VALUES(s_SEQ,' ',0,s_TaxType,
               to_date(s_ChargeDate, ' YYYY/MM/DD'),
               ' ','',NULL, NULL, n_TaxRate,
               n_TaxAmount,n_Check016InvAmount,
               NULL,NULL,'',s_Notes,s_CustId,s_CustSName,
               to_date(s_CurDateTime, 'YYYY/MM/DD HH24:MI:SS'), p_User, p_CompCode );

      EXCEPTION
      WHEN OTHERS THEN
        CLOSE cDynamic2;
        ROLLBACK;
        p_RetMsg := 'INSERT INV033 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
        p_RetCode := -99;
        RETURN -99;
      END;


      --�YInv016�D�ɻPInv017�����ɪ��X�p���B���X,
      --�h LOG  �᪽���� INV016 �U�@��
      goto GO_NEXT;
    end if;


--********************************************************************




    --�Y�w�}�ߵo���h Log �� inv033
    if s_RepeatInvIDList <> ' ' then
      begin
        IF NOT cDynamic2%ISOPEN then
          OPEN cDynamic2 FOR s_SQLInv017;
        END IF;
      exception
      when others then
        ROLLBACK;
        CLOSE cDynamic2;
        p_RetMsg := '�d��INV017����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      end;


      loop
        FETCH cDynamic2 INTO s_017BillId,n_017BillIdItemNo,s_017TaxType, s_017ChargeDate,s_017ItemID,s_017Desc,
          		n_017Quantity, n_017InitPrice, n_017TaxAmount, n_017TotalAmount,
          		s_017StartDate, s_017EndDate, s_017ChargeEn;

        EXIT WHEN cDynamic2%NOTFOUND;

        --�ˬd�O�_���}�߹L�B�S���o�����
        s_SQL := 'SELECT B.INVID FROM INV007 A,INV008 B WHERE B.BILLID ='|| '''' || s_017BillId || '''' ||
                 ' AND B.BILLIDITEMNO=' || n_017BillIdItemNo || ' AND ISOBSOLETE=''N'' ' ||
                 ' AND B.INVID=A.INVID AND A.COMPID=' || '''' || p_CompCode || '''' || ' ORDER BY B.INVID';

        begin
          IF NOT cDynamic3%ISOPEN then
            OPEN cDynamic3 FOR s_SQL;
          END IF;
        exception
        when others then
          ROLLBACK;
          CLOSE cDynamic3;
          p_RetMsg := '�d��INV008����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
          p_RetCode := -99;
          RETURN -99;
        end;

        s_RepeatInvIDList := ' ';

        loop
          FETCH cDynamic3 INTO s_RepeatInvID;
          EXIT WHEN cDynamic3%NOTFOUND;

          if s_RepeatInvIDList = ' ' then
            s_RepeatInvIDList := s_RepeatInvID;
          else
            s_RepeatInvIDList := s_RepeatInvIDList || ' , ' ||s_RepeatInvID;
          end if;
        end loop; --end loop inv008


        BEGIN
          s_Notes := '�o�����X:  ' || s_RepeatInvIDList;

          INSERT INTO INV033(SEQ,BILLID,BILLIDITEMNO,TAXTYPE,CHARGEDATE,ITEMID,
                             DESCRIPTION,QUANTITY,UNITPRICE,TAXRATE,TAXAMOUNT,
                             TOTALAMOUNT,STARTDATE,ENDDATE,CHARGEEN,NOTES,
                             CUSTID,CUSTNAME,LOGTIME,UPTEN, COMPID )
          VALUES(s_SEQ,s_017BillId,n_017BillIdItemNo,s_017TaxType,
                 to_date(s_017ChargeDate, ' YYYY/MM/DD'),
                 s_017ItemID,s_017Desc,n_017Quantity, n_017InitPrice, n_TaxRate,
                 n_017TaxAmount,n_017TotalAmount,
                 to_date(s_017StartDate, ' YYYY/MM/DD'),
                 to_date(s_017EndDate, ' YYYY/MM/DD'),
                 s_017ChargeEn,s_Notes,s_CustId,s_CustSName,
                 to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User, p_CompCode );


         EXCEPTION
         WHEN OTHERS THEN
           CLOSE cDynamic2;
           ROLLBACK;
           p_RetMsg := 'INSERT INV033 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
           p_RetCode := -99;
           RETURN -99;
         END;




        CLOSE cDynamic3;
      end loop;  --end loop inv017,�ˬd�C�@����ƬO�_�w�g�}�ߵo��
      CLOSE cDynamic2;

      --�Y�w�}�߹L�o���h LOG  �᪽���� INV016 �U�@��
      goto GO_NEXT;

    end if;
--UP loop inv017,�ˬd�C�@����ƬO�_�w�g�}�ߵo��
--------------------------------------------------------------------------------



    -- �]�w�u�����o�����
    if p_InvDateEqualToChargeDate=1 then
      s_RealInvDate := s_ChargeDate; --�o��������󦬶O���
    else
      s_RealInvDate := p_InvDate; --�o���������ǤJ���o�����
    end if;


    -- �ˬd�O�_�ݭn�s�W�� INV002(�Ȥ�D��)
    begin
      SELECT COUNT(1) into n_RecCount FROM INV002 WHERE CUSTID=s_CustId AND COMPID =p_CompCode
             AND IDENTIFYID1=p_IdentifyID1 AND IDENTIFYID2= p_IdentifyID2;
    exception
    when others then
      ROLLBACK;
      p_RetMsg := '���X�Ȥ�D�ɤ��ƶq�ɥ���:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;


    if n_RecCount=0 then --��ܦ��Ȥ��Ƥ��s�b INV002 ��,�ҥH�n insert �� INV002
      begin
        INSERT INTO INV002(IDENTIFYID1,IDENTIFYID2,COMPID,CUSTID,CUSTSNAME,CUSTNAME,MZIPCODE,MAILADDR,
			         TEL1,ISSELFCREATED,UPTTIME,UPTEN)
        VALUES(p_IdentifyID1,p_IdentifyID2, p_CompCode,s_CustId,s_CustSName,s_Title,s_ZipCode,s_MailAddr,s_Tel,'N',
                         to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User);

      EXCEPTION
      WHEN OTHERS THEN
        ROLLBACK;
        p_RetMsg := 'INSERT INV002 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      END;

    end if;

    -- �ˬd�O�_�ݭn�s�W�� INV019(�Ȥ������)
    begin
      SELECT COUNT(*) COUNT into n_RecCount FROM INV019 WHERE CUSTID=s_CustId AND COMPID =p_CompCode
    AND IDENTIFYID1=p_IdentifyID1 AND IDENTIFYID2= p_IdentifyID2 AND TITLEID ='1';
    exception
    when others then
      ROLLBACK;
      p_RetMsg := '���X�Ȥ�����ɤ��ƶq�ɥ���:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;



    if n_RecCount=0 then --��ܦ��Ȥ��Ƥ��s�b INV019 ��,�ҥH�n insert �� INV019
      begin

        INSERT INTO INV019(IDENTIFYID1,IDENTIFYID2,COMPID,CUSTID,TITLEID,TITLESNAME,TITLENAME,BUSINESSID,MZIPCODE,
			         MAILADDR,INVADDR,UPTTIME,UPTEN)
        VALUES(p_IdentifyID1,p_IdentifyID2,p_CompCode,s_CustId,1,s_CustSName, s_Title,s_BusinessID,s_ZipCode,
		         	s_MailAddr, s_InvAddr,
                        to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User);
      EXCEPTION
      WHEN OTHERS THEN
        ROLLBACK;
        p_RetMsg := 'INSERT INV019 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
        p_RetCode := -99;
        RETURN -99;
      END;
    end if;


    -- �o���s��, 8��, ������ 0
    s_Tmp1 := LPAD( TO_CHAR( n_CurNum ), 8, '0' );

    -- �N�r��'2002/12/23' �ഫ�� '911223'
    s_Tmp2 := p_SystemID || TO_CHAR( TO_NUMBER( REPLACE( s_RealInvDate, '/' ,'' ) ) - 19110000 );

    ---- 2005/01/11, Nick �ץ�, �o�����Ӭ� 10 �X, ���ӬO LTRIM( TO_CHAR( n_CurNum ) )
    s_RealInvID := v_Prefix || LPAD( TO_CHAR( n_CurNum ), 8, '0' );

    --- �o���ˬd�X
    aCheckNo := CalcInvoiceCheckNumber( s_Tmp1, s_Tmp2 );

  -- down,�B�z�o���D�Ȧs��(INV031)
    begin
      INSERT INTO INV031(IDENTIFYID1,IDENTIFYID2,CHECKNO,INVID,UPTTIME,UPTEN,CANMODIFY,INVDATE,
	           CHARGEDATE,COMPID,CUSTID,CUSTSNAME,INVTITLE,ZIPCODE,INVADDR,
	           MAILADDR,BUSINESSID,INVFORMAT,TAXTYPE,TAXRATE,SALEAMOUNT,TAXAMOUNT,INVAMOUNT,
	           PRINTCOUNT,ISPAST,ISOBSOLETE,HOWTOCREATE,PRINTFUN,PRINTTIME)
      VALUES(p_IdentifyID1,p_IdentifyID2, aCheckNo, s_RealInvID,
	        	to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User,
        		'Y',to_date(s_RealInvDate,'YYYY/MM/DD'),to_date(s_ChargeDate,'YYYY/MM/DD'), p_CompCode,s_CustId,
         		s_CustSName,s_Title,s_ZipCode,s_InvAddr,s_MailAddr,s_BusinessID,'1',
        		s_TaxType,n_TaxRate,n_SaleAmount,n_TaxAmount,n_InvAmount,0,'Y','N',to_char(p_HowToCreate),0,
        		to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'));

      EXCEPTION
      WHEN OTHERS THEN
        ROLLBACK;
        p_RetMsg := 'INSERT INV031 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM;
        p_RetCode := -99;
        RETURN -99;
    END;
  -- up,�B�z�o���D�Ȧs��(INV031)

  -- down,�B�z�o���D�Ȧs��(INV032)

    s_SQL := 'SELECT BILLID,BILLIDITEMNO,TAXTYPE, '||
		'TO_CHAR(CHARGEDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') CHARGEDATE,' ||
		'ITEMID ' ||
		',DESCRIPTION,QUANTITY,UNITPRICE,TAXAMOUNT,TOTALAMOUNT,' ||
		'TO_CHAR(STARTDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') STARTDATE,' ||
		'TO_CHAR(ENDDATE,' || '''' || 'YYYY/MM/DD' || '''' || ') ENDDATE,' ||
		'CHARGEEN,SERVICETYPE from inv017 where SEQ=' || '''' || s_SEQ || '''' ||
    '   AND SHOULDBEASSIGNED = ''Y''    ';


    begin
       IF NOT cDynamic2%ISOPEN then
       OPEN cDynamic2 FOR s_SQL;
    END IF;

    exception
    when others then
      ROLLBACK;
      p_RetMsg := '�d��INV017����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || '      SQL='||s_SQL;
      p_RetCode := -99;
      RETURN -99;
    end;


    --loop inv017,���C�@����ƨӶ}�ߵo��
    n_017Seq := 0;
    loop
     n_017Seq := n_017Seq + 1;
     FETCH cDynamic2 INTO s_017BillId,n_017BillIdItemNo,s_017TaxType, s_017ChargeDate,s_017ItemID,s_017Desc,
        	 n_017Quantity, n_017InitPrice, n_017TaxAmount, n_017TotalAmount,
        	 s_017StartDate, s_017EndDate, s_017ChargeEn, s_017ServiceType;

     EXIT WHEN cDynamic2%NOTFOUND;

     BEGIN
       n_Inv032SaleAmount := n_017Quantity * n_017InitPrice;

       -- down, insert �� INV032
       INSERT INTO INV032(IDENTIFYID1,IDENTIFYID2,INVID,BILLIDITEMNO,SEQ,UPTTIME,UPTEN,BILLID,
          		STARTDATE,ENDDATE,ITEMID,DESCRIPTION,QUANTITY,UNITPRICE,TAXAMOUNT,TOTALAMOUNT,CHARGEEN,SERVICETYPE,SaleAmount)
                 VALUES(p_IdentifyID1,p_IdentifyID2,s_RealInvID,n_017BillIdItemNo,n_017Seq,
          		to_date(s_CurDateTime, ' YYYY/MM/DD HH24:MI:SS'), p_User,s_017BillId,
          		to_date(s_017StartDate, ' YYYY/MM/DD HH24:MI:SS'),
          		to_date(s_017EndDate, ' YYYY/MM/DD HH24:MI:SS'),
          		s_017ItemID,s_017Desc,n_017Quantity, n_017InitPrice,n_017TaxAmount,
          		n_017TotalAmount,s_017ChargeEn,s_017ServiceType,n_Inv032SaleAmount);
       -- up, insert �� INV032

     EXCEPTION
     WHEN OTHERS THEN
       CLOSE cDynamic2;
       ROLLBACK;
       p_RetMsg := 'INSERT INV032 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
       p_RetCode := -99;
       RETURN -99;
     END;



     -- down, update Inv016
     begin
       s_BeAssignedInvID := 'Y';
       UPDATE INV016 SET BEASSIGNEDINVID = s_BeAssignedInvID WHERE SEQ=s_SEQ;
     exception
     when others then
       ROLLBACK;
       p_RetMsg := 'update INV016����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || 'SQL:' || s_SQL;
       p_RetCode := -99;
       RETURN -99;
     end;
     -- up, update Inv016


     -- down, update SO033 �� SO034

     if p_LinkToMis = 'Y' then

           if p_HowToCreate=1  then
             BEGIN
               s_TempSQL := '  UPDATE ' || p_MisDbOwner || '.SO033' || aDBLink ||
                            '    SET GUINO = ' || '''' || s_RealInvID || '''' || ', ' ||
                            '        PREINVOICE = 1, ' ||
                            '        INVDATE = TO_DATE( ' || '''' || s_RealInvDate || '''' || ', ''YYYY/MM/DD''), ' ||
                            '        INVOICETIME = SYSDATE ' ||
                            ' WHERE BILLNO = ' || '''' || s_017BillId || '''' ||
                            '   AND ITEM = ' || n_017BillIdItemNo;

               EXECUTE IMMEDIATE s_TempSQL;

             EXCEPTION
             WHEN OTHERS THEN
               ROLLBACK;
               p_RetMsg := 'UPDATE SO033 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_TempSQL;
               p_RetCode := -99;
               RETURN -99;
             END;
           else
             begin
               s_TempSQL := ' UPDATE ' || p_MisDbOwner || '.SO034' || aDBLink ||
                            '    SET GUINO = ' || '''' || s_RealInvID || '''' || ', ' ||
                            '        PREINVOICE = 0,' ||
                            '        INVDATE = TO_DATE( ' || '''' || s_RealInvDate || '''' || ', ''YYYY/MM/DD'' ), ' ||
                            '        INVOICETIME = SYSDATE ' ||
                            '  WHERE BILLNO = ' || '''' || s_017BillId || '''' ||
                            '    AND ITEM = ' || n_017BillIdItemNo;

               EXECUTE IMMEDIATE s_TempSQL;

             EXCEPTION
             WHEN OTHERS THEN
               ROLLBACK;
               p_RetMsg := 'UPDATE SO034 �ɵo�Ϳ��~: ' ||' SQLCODE='||SQLCODE||'      '||'   SQLERRM='||SQLERRM|| '      SQL='||s_TempSQL;
               p_RetCode := -99;
               RETURN -99;
             END;
           end if;

     end if;

-- up, update SO033 �� SO034

    end loop;

    CLOSE cDynamic2;
  -- up,�B�z�o���D�Ȧs��(INV032)

    n_CurNum := n_CurNum + 1;

    --�C�����@���ˬd�O�_���Ӧr�y�̫�@��
    if n_EndNum < n_CurNum then

      --�Ӧr�y�Χ�
      s_Useful := 'N';

      -- down, update Inv099
      begin
        UPDATE INV099
           SET CURNUM = LPAD( TO_CHAR( n_CurNum ), 8, '0' ),
               USEFUL = s_Useful,
               LASTINVDATE = TO_DATE( s_RealInvDate, 'YYYY/MM/DD' ),
    	         UPTTIME = TO_DATE( s_CurDateTime,'YYYY/MM/DD HH24:MI:SS' ),
    	         UPTEN = p_User
    	   WHERE COMPID = p_CompCode
    	     AND YEARMONTH = p_InvYearMonth
    	     AND PREFIX = v_Prefix
    	     AND STARTNUM = v_StartNum;
      exception
      when others then
        ROLLBACK;
        p_RetMsg := 'update INV099����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || 'SQL:' || s_SQL;
        p_RetCode := -99;
        RETURN -99;
      end;
      -- up, update Inv099

      --���U�@�r�y
      v_CurIndex := v_CurIndex + 1;
      v_PreFix := aPreFix(v_CurIndex);
      v_StartNum := aStartNum(v_CurIndex);

      -- ���r�y���
      begin
          SELECT TO_NUMBER( CURNUM ),
                 TO_NUMBER( ENDNUM ),
                 TO_CHAR( LASTINVDATE, 'YYYY/MM/DD' )
            INTO n_CurNum,
                 n_EndNum,
                 s_LastInvDate
            FROM INV099
           WHERE COMPID = p_CompCode
             AND YEARMONTH = p_InvYearMonth
             AND PREFIX = v_Prefix
             AND USEFUL = 'Y'
             AND STARTNUM = v_StartNum;
      exception
        when others then
          ROLLBACK;
          p_RetMsg := '�L�k���X�r�y���:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
          p_RetCode := -99;
          RETURN -99;
      end;

      --���r�y�ɥ����w��r�y���̫�o�����
      s_RealInvDate := s_LastInvDate;
    else
      s_Useful := 'Y';
    end if;

    --�Y�����`��ƫh LOG  �᪽�����U�@��
    <<GO_NEXT>>
    NULL;

  END LOOP;

--Close dynamic cursor
  CLOSE cDynamic;








  begin
    INSERT INTO INV007 (SELECT * FROM INV031);
  exception
  when others then
    ROLLBACK;
    p_RetMsg := 'insert INV007����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
    p_RetCode := -99;
    RETURN -99;
  end;



  begin
    INSERT INTO INV008 (SELECT * FROM INV032);
  exception
  when others then
    ROLLBACK;
    p_RetMsg := 'insert INV008����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
    p_RetCode := -99;
    RETURN -99;
  end;


  -- down, update Inv099
  begin
    s_SQL :=  'AAA' || s_RealInvDate || 'BBB' || s_CurDateTime;

    UPDATE INV099
       SET CURNUM = LPAD( TO_CHAR( n_CurNum ), 8, '0' ),
           USEFUL=s_Useful,
           LASTINVDATE = TO_DATE( s_RealInvDate, 'YYYY/MM/DD' ),
	         UPTTIME = TO_DATE( s_CurDateTime, 'YYYY/MM/DD HH24:MI:SS' ),
	         UPTEN = p_User
	   WHERE COMPID = p_CompCode
	     AND YEARMONTH = p_InvYearMonth
	     AND PREFIX = v_Prefix
	     AND STARTNUM = v_StartNum;
  exception
  when others then
    ROLLBACK;
    p_RetMsg := 'update INV099����:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM || 'SQL:' || s_SQL;
    p_RetCode := -99;
    RETURN -99;
  end;
  -- up, update Inv099


  --- �N���}�ߪ��o����Ƽg�i INV033 �� Log

  for c1 in (  select a.seq, a.custid, a.shouldbeassigned, a.invamount, a.taxtype
                 from inv016 a
                where a.compid = to_char( p_CompCode )
                  and a.chargedate between to_date( p_ChargeStartdate, 'YYYY/MM/DD' ) and to_date( p_ChargeStopdate, 'YYYY/MM//DD' )
                  and a.beassignedinvid = 'N'
                  and a.isvalid = 'Y'
                  and a.howtocreate = to_char( p_HowToCreate )
                  and ( a.shouldbeassigned = 'N' or a.invamount <= 0  or a.taxtype = '0' ) )
  loop

       s_CustSName := null;
       if ( p_LinkToMis = 'Y' ) then
           begin
             s_SQL := ' SELECT CUSTNAME FROM ' || p_MisDbOwner || '.SO001' || aDBLink ||' WHERE CUSTID = ' || c1.custid;
             execute immediate s_SQL into s_CustSName;
           exception
             when others then
             begin
                rollback;
                p_RetMsg := '���}�ߵo����ƨ��Ȥ�²�ٮɦ��~, �Ȥ�N�X:' || c1.custid ||', SQLCODE = ' || SQLCODE||', SQLERRM = ' || SQLERRM;
                p_RetCode := -99;
                return -99;
             end;
           end;
       end if;

       if ( c1.shouldbeassigned = 'N' ) then
         s_Notes := '�����}��';
       elsif ( c1.invamount <= 0 ) then
         s_Notes := '���B�p��ε���s';
       elsif ( c1.taxtype = '0' ) then
         s_Notes := '�|�O�����}��';
       else
         s_Notes := '�����}��';
       end if;

       begin

            insert into inv033
                    (   seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        notes,
                        custid,
                        custname,
                        logtime,
                        upten,
                        compid     )
                 select seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        s_Notes,
                        c1.custid,
                        s_CustSName,
                        to_date(p_LogDateTime,'YYYY/MM/DD HH24:MI:SS'),
                        p_User,
                        p_CompCode
	               from inv017 where seq = c1.seq;

       exception
          when others then
          begin
            rollback;
            p_RetMsg := '���}�ߵo����Ƽg�J INV033 �ɵo�Ϳ��~: SQLCODE = ' || SQLCODE || ', SQLERRM = '||SQLERRM;
            p_RetCode := -99;
            return -99;
         end;
       end;

  end loop;

  COMMIT;


  --Dennis
  --- �N���}�ߪ��o����ơ]���ӳ����^�g�i INV033 �� Log
  for c1 in (  select b.seq, a.custid, a.title, b.shouldbeassigned, b.totalamount, b.taxtype
                 from inv016 a, inv017 b
                where a.compid = to_char( p_CompCode )
                  and a.chargedate between to_date( p_ChargeStartdate, 'YYYY/MM/DD' ) and to_date( p_ChargeStopdate, 'YYYY/MM//DD' )
                  and a.isvalid = 'Y'
                  and a.howtocreate = to_char( p_HowToCreate )
                  and a.beassignedinvid = 'Y'
                  and a.shouldbeassigned = 'Y'
                  and b.shouldbeassigned = 'N'
                  and a.seq = b.seq
                  and ( b.shouldbeassigned = 'N' or b.totalamount <= 0  or b.taxtype = '0' ) )
  loop

       if ( c1.shouldbeassigned = 'N' ) then
         s_Notes := '�����}��';
       elsif ( c1.totalamount <= 0 ) then
         s_Notes := '���B�p��ε���s';
       elsif ( c1.taxtype = '0' ) then
         s_Notes := '�|�O�����}��';
       else
         s_Notes := '�����}��';
       end if;

       begin

            insert into inv033
                    (   seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        notes,
                        custid,
                        custname,
                        logtime,
                        upten,
                        compid     )
                 select seq,
                        billid,
                        billiditemno,
                        taxtype,
                        chargedate,
                        itemid,
                        description,
                        quantity,
                        unitprice,
                        taxrate,
                        taxamount,
                        totalamount,
                        startdate,
                        enddate,
                        chargeen,
                        s_Notes,
                        c1.custid,
                        c1.title,
                        to_date(p_LogDateTime,'YYYY/MM/DD HH24:MI:SS'),
                        p_User,
                        p_CompCode
	               from inv017 where seq = c1.seq and shouldbeassigned = 'N';

       exception
          when others then
          begin
            rollback;
            p_RetMsg := '���}�ߵo����Ƽg�J INV033 �ɵo�Ϳ��~: SQLCODE = ' || SQLCODE || ', SQLERRM = '||SQLERRM;
            p_RetCode := -99;
            return -99;
         end;
       end;

  end loop;

  COMMIT;


  p_RetCode := 0;
  p_RetMsg := '�o���}�߰��槹��';

  RETURN 0;
EXCEPTION
   WHEN others THEN
     CLOSE cDynamic;
     ROLLBACK;
     p_RetMsg := '��L���~:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
     p_RetCode := -99;
     RETURN -99;
END;
/

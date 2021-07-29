create or replace function sf_assigninvid
  ( p_User                     in varchar2,
    p_CompId                   in integer,
    p_LinkToMis                in varchar2,
    p_DbLink                   in varchar2,
    p_HowToCreate              in integer,
    p_InvDateEqualToChargeDate in integer,
    p_InvDate                  in varchar2,
    p_InvYearMonth             in varchar2,
    p_ChargeStartdate          in varchar2,
    p_ChargeStopDate           in varchar2,
    p_IdentifyID1              in varchar2,
    p_IdentifyID2              in integer,
    p_SystemID                 in varchar2,
    p_PrefixString             in varchar2,
    p_OrderBy                  in integer,
    p_MisDbOwner               in varchar2,
    p_Test                 in integer,
    p_RetCode                 out integer,
    p_RetMsg                  out varchar2,
    p_LogDateTime             out varchar2
  )
return number as

/*

  �`�N�ƶ�: �o�� Schema �� owner �Y�O�P �ȪA�t�� Schema �� owner ���P, �ȥ� grant �v�����o���t�Ϊ� owner
            �ȪA�t�ΩҶ� table         : grant select, update on SO033 to <�o���t�Ϊ�owner>
                                         grant select, update on SO034 to <�o���t�Ϊ�owner>
                                         grant select on so002a to <�o���t�Ϊ�owner>


  �o���}��
  �ɦW: SF_AssignInvID.sql

  ����: �B�z�o���t�Τ��o���}��(����)

  �Ѽ� IN:

	p_User          		        varchar2 :  �{���ާ@��
	p_CompCode      		        number   :  ���q�O�N�X
	p_LinkToMis                 varchar2 :  �O�_��s�ȪA�t��   Y=>��s; N=>����s
	p_HowToCreate   		        number   :  1=>�w�};   2=>��}
	p_DbLink                    varchar2 :  ���Ʈw�s����, �z�L DBLink �覡�USQL
  p_InvDateEqualToChargeDate	number   :  1=>�o��������󦬶O���;   2=>�o����������󦬶O���
	p_InvDate			              varchar2 :  �o�����
	p_InvYearMonth 			        varchar2 :  �o���~��
	p_ChargeStartdate 	        varchar2 :  ���O�_�l��
	p_ChargeStopDate 		        varchar2 :  ���O�I���
	p_IdentifyID1 			        number   :  �ѧO�X(�O�d), ���ӵ���1
	p_IdentifyID2 			        number   :  �ѧO�X(�O�d), ���ӵ���0
	p_SystemID 			            varchar2 :  System ID
  p_PrefixString 		          varchar2 :  �o���r�y
	p_OrderBy 			            number   :  �ƧǤ覡 : 0=>���O���
	                                                   1=>�l���ϸ�
	                                                   2=>���O���+�l���ϸ�
	                                                   3=>���O���+�l���ϸ�+�Ƚs
  p_MisDbOwner                narchar2 :  �^���MIS��Table Space �� Owner



  �Ѽ� OUT:

	p_RetMsg        		       varchar2  :     ���G�T��
	p_RetCode       		       number    :     ���G�N�X


  Return
	number: �B�z���G�N�X, �����T���s��p_RetMsg��

         0:  p_RetMsg = '�o���}�߰��槹��'
        -1:  p_RetMsg = '�Ѽƿ��~'
       -99:  p_RetMsg = '��L���~'

*/

/*
    �������ʰO��
    First release    : 2002.12.23
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
    Modify By Nick   : 2005.04.22  ����}�ߵo����Ʀh�[�@���P�_, �|�O TAXTYPE = 0 , �����}��
    Modify By Nick   : 2005.09.30  �w�� EMC �ϥ�, ��ȪA�t�θ�ƼW�[ DBLink �r��
    Modify By Nick   : 2005.10.04  �W�[�J�Ѽ�, p_Dblink, p_linkToMis,
                                   �Y p_linkToMis = Y , �~��s�ȪA�t�θ��,
                                   �Y p_Dblink ����, �h sql �h�[�J dblink �Ҧp: select * from So033@catvc
    Modify By Nick   : 2005.10.20  �ץ��]���h���q�O�ˬd���O���ح���, �ӾɭP����}�ߵo��(�h�[�P�_ CompId )
    Modify By Nick   : 2006.01.16  INV033 �W�[ CompId ���q�O���, �h���q�}�ߵo������, ���`����~�|�����
    Modify By Nick   : 2006.01.19  1.���s��g
                                   2.Ū�����q�O�]�w���Ѽ� INV001.AutoCreateNum, �W�[�}�ߧP�_,
                                     �Y���]�w ( p_AutoCreateNum > 0 ) �N���۰ʴ��}�o���\��, ��o�����ӵ��Ƥj
                                     �󦹼Ʈɥ����۰ʴ��}��2�i�o��
    Modify By Nick   : 2006.02.09  1.�ϥ� SQL Plus �sĶ�� store function ��, �|�L�k�sĶ, �ק惡 Bug.
                                   2.��s�ȪA�t�� SO033, SO034 ��� ITEMNO ��쥴��, ���� ITEM
                                   3.�o���� INV099 �� LastInvDate �S����s��, ��s�W�h���~, �����C���}�ߴN������s
    Modify By Nick   : 2006.02.14  1.INV008, INV017, INV032 �W�[ LinkToMis ���, �O�_��s���ˮ֫ȪA�t�α����ӵ�
                                     ����, INV001 �� LinkToMis �O�`�}��
                                     Ex: INV001.LINKTOMIS = 'Y' AND INV008.LINKTOMIS = 'Y' �o�ˤ~�|�u���h�ˮ֤Χ�s
                                     �ȪA�t��, �Y INV001.LINKTOMIS <> 'Y' or INV008.LINKTOMIS <> 'Y' �h���ˮ֤Χ�s�ȪA�t��
    Modify By Nick   : 2006.04.10  1.�W�[�o�����Ӧ��O���ئX�֥\��
                                     �s�W Table INV032A, INV008A
                                     �s�W��� INV017.ACCOUNTNO ��r�ɶפJ�ɷ|�h�����(�u��CNS)
                                     �}�߮�, ���P�_���Ӹ�ƬO�_�����n�X�֦��P�@�����Ӹ��,
                                     Ex: �N��-�Ȥ�CM�W���O1M, �ƶq:1, ���B:300 --|
                                         �N��-�Ȥ�CM�W���O1M, �ƶq:2, ���B:600 --|--> �X�֦��@�� �i�N��-�Ȥ�CM�W���O1M�j, �ƶq:1, ���B: 1200
                                         �N��-�Ȥ�CM�W���O1M, �ƶq:1, ���B:300 --|
                                     �Y�����Ӭ�P�A��, �h�٭n�h�P�_���P�@CP�����~�i�H�X��, �Y���P�����h�n�֨�t�@�ө���
                                   2.�Y�����Ӭ��H�Υdú�O, �h�h���H���d��, �C�L�o���ɷ|�Ψ�

    Modify By Nick   : 2006.06.26  1.INV001�W�[�Ѽ�, ����P�]�ƪ�����, �o�����O���ӬO�_���n�X�֦��P�@���O����
                                     INV001.FACICOMBINE = 1 ��, �Y�����Ӭ�P�A��, �h���P�����@�˦X�֨�P�@����
                                   2.INV008 �W�[2�����, FACISNO, ACCOUNTNO �����X�ֶ��خ�, ��2������x�s CP �����ΫH�Υd��
                                   3.�}�ߵo���ƧǤ覡, �s�W ���O���+�l���ϸ�

    Modify By Nick   : 2006.08.28  1.�}�߮�, ���s�p��P���B�ε|�B, �P���B=(�`���B/1.05),  �|�B=(�`���B-�P���B)
                                   2.�ҥH�ϥέ��s�p��L�᪺�P���B�ε|�B, �^�Y��s�D�ɪ��o�����B, �o�i��|��쥻�ȪA�߹L�Ӫ����B���@�P
                                     �ר�O�X�ֹL�᪺

    Modify By Nick   : 2006.11.10  1. 2006.08.28 ������bug,��۰ʴ��}�o����,�D�o���`���B,�D�o���P���B,�D�o���|�B���Ҭ��ӱi�����B
    Modify By Nick   : 2007.05.24  �}�ߵo���ƧǤ覡, �s�W ���O���+�l���ϸ�+�Ƚs
    Modify By Nick   : 2008.04.07  �o���}�߼W�[ 1.�o���γ~�N�X 2.�o���γ~�N�X���� (INV016, INV031, INV007)
    Modify By Nick   : 2008.08.12  �o���}�߮�, �W�[�^�g�ȪA�t�� SO033 �� SO034 ��� INVPURPOSECODE, INVPURPOSENAME
	Modify By Kin     :  2009.12.11  ���D��5358 ���X�ֶ��ثh��ڰ_�������INV008A��Min(StartDate)�PMax(EndDate)
*/



  /* �s��e�ݶǶi�Ӫ��h�յo���r�y */

  type TInv099 is record (
       CompId       inv099.compid%type,
       YearMonth    inv099.yearmonth%type,
       Prefix       inv099.prefix%type,
       StartNum     inv099.startnum%type,
       EndNum       inv099.endnum%type,
       CurNum       inv099.curnum%type,
       Useful       inv099.useful%type,
       LastInvDate  inv099.lastinvdate%type,
       HasChange    boolean );

  type TInv099Table is table of TInv099 index by binary_integer;

  aInv099             TInv099Table;
  aInv099MaxBounds    binary_integer;
  aInv099Index        binary_integer;
  aTempStr            varchar2(2000);
  aPos                integer;
  aDetailCount        integer;
  aMasterCanCreate    boolean;
  aCheckNo            varchar2(2);
  aNotes              inv033.notes%type;
  aInv017TotalAmount  inv017.totalamount%type;
  aAutoCreateNum      inv001.autocreatenum%type;
  aFaciCombine        inv001.facicombine%type;
  aInvId              inv007.invid%type;
  aRealInvDate        inv007.invdate%type;
  aInvTaxType         inv007.taxtype%type;
  aInvTaxRate         inv007.taxrate%type;
  aDetailSeq          integer;
  aOldSeq             integer;
  aItemDesc           inv005.description%type;
  aMergeUnitPrice     inv008.unitprice%type;
  aMergeTaxAmount     inv008.taxamount%type;
  aMergeSaleAmount    inv008.saleamount%type;
  aMergeTotalAmount   inv008.totalamount%type;
  aMergeQuantity      inv008.quantity%type;
  aMergeCount         integer;
  aAccountNo          inv008a.accountno%type;
  aFacisNo            inv008a.facisno%type;
  aOldItemRefId       inv008a.facisno%type;

  type TAutoCreate is record (
       InvId          inv007.invid%type,
       SaleAmount     inv007.saleamount%type,
       TaxAmount      inv007.taxamount%type,
       InvAmount      inv007.invamount%type,
       MainSaleAmount     inv007.saleamount%type,
       MainTaxAmount      inv007.taxamount%type,
       MainInvAmount      inv007.invamount%type);

  type TAutoCreateTable is table of TAutoCreate index by binary_integer;

  aAutoCreateList       TAutoCreateTable;
  aAutoCreateMaxBounds  binary_integer;
  aCount                integer;

              /* ------------------------------------------------------------------ */

              /*  �p��o���ˬd�X */

              function CalcInvCheckSum(aInvId in varchar2, aInvParam in varchar2)
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

              /* ���X�i�Ϊ��o�����X, ���K��s�o�����i�Ϊ��A */

              function GetInvNumber(aIndex in out binary_integer, aId out varchar2, aDate in date) return boolean
              as
              begin
                aId := null;
                 << RE_GET >>
                   null;
                 if ( aIndex between 1 and aInv099MaxBounds ) then
                    /* �{�b�o���o�����O�_�i�� */
                    if ( aInv099( aIndex ).CurNum <= ( aInv099( aIndex ).EndNum ) ) then
                       /* ���o�����X */
                       aId := ( aInv099( aIndex ).prefix || Lpad( to_char( to_number( aInv099( aIndex ).CurNum ) ), 8, '0' ) );
                       aInv099( aIndex ).CurNum := lpad( to_char( to_number( aInv099( aIndex ).CurNum ) + 1 ), 8, '0' );
                       aInv099( aIndex ).LastInvDate := aDate;
                       aInv099( aIndex ).HasChange := true;
                    end if;
                    /* �P�_�O�_�Χ� */
                    if ( aInv099( aIndex ).CurNum > ( aInv099( aIndex ).EndNum ) ) then
                      aInv099( aIndex ).Useful := 'N';
                      aInv099( aIndex ).LastInvDate := aDate;
                      aInv099( aIndex ).HasChange := true;
                      aIndex := ( aIndex + 1 );
                      /* �S���o�����X, �A���@�� */
                      if ( aId is null ) then goto RE_GET; end if;
                    end if;
                 end if;
                 if ( aId is null ) then
                   p_RetMsg := '���X�i�εo�����X����, ���X���X���ŭȩάO�o�������X�w�Χ�';
                   p_RetCode := -99;
                 end if;
                 return ( aId is not null );
              end;

              /* ------------------------------------------------------------------- */


              /*  �� DBLink �r�� */

              function GetDBLink return varchar2
              as
                 aResult varchar2(30);
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

              /* HowToCreate Mapping To CCB TableName */

              function GetTableNameByHowToCreate return varchar2
              as
                 aResult varchar2(100);
              begin
                 if ( p_HowToCreate = '1' ) then
                   aResult := 'SO033';
                 else
                   aResult := 'SO034';
                 end if;
                 return ( aResult || GetDBLink );
              end;

              /* ------------------------------------------------------------------- */

              function GetInv001Param(aNum in out integer, aFaciCombine in out integer) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    select nvl( autocreatenum, 0 ), nvl( facicombine, 0 )
                      into aNum, aFaciCombine
                      from inv001
                     where identifyid1 = p_IdentifyID1
                       and identifyid2 = p_IdentifyID2
                       and compid = p_CompId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '�����q�O�Ѽ�(�۰ʴ��}�o���γ]�ƪ����O�_�X��)�ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* �ˮ֫ȪA�t�Φ����o�����Ӹ�� */

              function CheckMisSystem(aId in varchar2, aBillNo in varchar2, aItem in number, aMsg out varchar2 ) return boolean
              as
                 aSql varchar2(255);
                 aTableName varchar2(100);
                 aGuiNo varchar2(10);
              begin
                 aMsg := null;
                 aTableName := GetTableNameByHowToCreate;
                 aSql := ' SELECT GUINO FROM ' || p_MisDbOwner || '.' || aTableName ||
                         '  WHERE CUSTID = ''' || aId || ''' AND BILLNO = ''' || aBillNo || ''' AND ITEM = ''' || aItem || '''';
                 begin
                   execute immediate aSql into aGuiNo;
                   if ( aGuiNo is not null ) then
                      aMsg := '�ȪA�t�Τ��������O�渹�w���o�����X, �ȪA�t�Τ��o�����X : ' || aGuiNo;
                   end if;
                 exception
                   when others then aMsg := '�ȪA�t�εL����������ơC';
                 end;
                 return ( aMsg is null );
              end;


              /* ------------------------------------------------------------------- */

              /* �ˮֵo���t�Τ��O�_�w�������O�渹�򶵦� ( �w���o���� ), �Y��, �h��o�����X�O�U�� */

              function CheckAlreadyCreate(aBillId in inv008.billid%type, aItemNo in inv008.billiditemno%type,
                aMsg out varchar2) return boolean
              as
              begin
                 aMsg := null;
                 for cCheck in ( select a.invid from inv007 a, inv008 b where a.invid = b.invid
                    and a.compid = p_CompId  and a.isobsolete = 'N' and b.billid = aBillId and b.billiditemno = aItemNo
                    order by a.invid )
                 loop
                    if ( aMsg is null ) then
                      aMsg := '�����O�渹�w�}�ߵo��, �o�����X:' || cCheck.invid;
                    else
                      aMsg := ( aMsg || ',' || cCheck.invid );
                    end if;
                 end loop;
                 return ( aMsg is not null );
              end;


              /* ------------------------------------------------------------------- */

              /* �o���}���ˮ֬������}�߮�, ���� Log �� INV033, �P�ɵ�����] */

              function LogToInv033(aSeq in inv033.seq%type, aBillId in inv033.billid%type, aItemNo in inv033.billiditemno%type,
                 aCustId in varchar2, aCustName in varchar2, aNotes in inv033.notes%type ) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv033 ( seq, billid, billiditemno, taxtype, chargedate, itemid,
                       description, quantity, unitprice, taxrate, taxamount, totalamount, startdate,
                       enddate, chargeen, notes, custid, custname, logtime, upten, compid )
                    select a.seq, a.billid, a.billiditemno, a.taxtype, a.chargedate, a.itemid,
                       a.description, a.quantity, a.unitprice, a.taxrate, a.taxamount, a.totalamount, a.startdate,
                       a.enddate, a.chargeen, substr( aNotes, 1, 200 ), aCustId, aCustName,
                       to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User, p_CompId
                      from inv017 a
                     where a.seq = aseq and a.billid = aBillId and a.billiditemno = aItemNo;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV033 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* �ˮ֫Ȥ�D�ɬO�_�����Ȥ��� */

              function CheckInv002Exists(aCustId in varchar2) return boolean
              as
                 aCount integer;
              begin
                 select count(1) into aCount from inv002
                  where custid = aCustId and compid = p_CompId
                    and identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2;
                 return ( aCount > 0 );
              end;

              /* ------------------------------------------------------------------- */

              /* �ˮ֫Ȥ�����ɬO�_�����Ȥ��� */

              function CheckInv019Exists(aCustId in varchar2) return boolean
              as
                 aCount integer;
              begin
                 select count(1) into aCount from inv019
                  where custid = aCustId and compid = p_CompId
                    and identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2
                    and titleid = 1;
                 return ( aCount > 0 );
              end;


              /* ------------------------------------------------------------------- */

              /* �g�o���t�Ϊ��Ȥ�D�� */

              function WriteToInv002(aSeq in varchar2) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv002 ( identifyid1, identifyid2, compid, custid, custsname, custname, mzipcode,
                       mailaddr, tel1, isselfcreated, upttime, upten )
                    select p_IdentifyID1, p_IdentifyID2, p_CompId, a.custid, a.chargetitle, a.title, a.zipcode,
                       a.mailaddr, a.tel, 'N', to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User
                     from inv016 a
                    where a.seq = aSeq
                      and a.compid = p_CompId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV002 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* �g�o���t�Ϊ��Ȥ������ */

              function WriteToInv019(aSeq in varchar2) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv019 ( identifyid1, identifyid2, compid, custid, titleid, titlesname,
                       titlename, businessid, mzipcode, mailaddr, invaddr, upttime, upten )
                    select p_IdentifyID1, p_IdentifyID2, p_CompId, a.custid, 1, a.chargetitle, a.title,
                       a.businessid, a.zipcode, a.mailaddr, a.invaddr, to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User
                      from inv016 a
                     where a.seq = aSeq
                       and a.compid = p_CompId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV019 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* �g�o���Ȧs�D�� */

              function WriteToInv031(aSeq in varchar2, aId in varchar2, aNo in varchar2, aInvDate in date,
                aTaxType out varchar2, aTaxRate out number)
                 return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv031 ( identifyid1, identifyid2, checkno, invid, canmodify, invdate, chargedate,
                        compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat, taxtype,
                        taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete, howtocreate,
                        upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount, maininvamount, invuseid, invusedesc )
                    select p_IdentifyID1, p_IdentifyID2, aNo, aId, 'Y', aInvDate, a.chargedate,
                        p_CompId, a.custid, a.chargetitle, a.title, a.zipcode, a.invaddr, a.mailaddr, a.businessid, '1', a.taxtype,
                        a.taxrate, 0, 0, 0, 0, 'N', 'N', p_HowToCreate,
                        to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ), p_User, '0', aId, 0, 0, 0, a.invuseid, a.invusedesc
                    from inv016 a
                   where a.seq = aSeq
                     and a.compid = p_CompId;
                    /* �� TaxType, TaxRate �^�� */
                    aTaxType := null;
                    aTaxRate := 0;
                    select taxtype, taxrate into aTaxType, aTaxRate from inv031
                     where identifyid1 = p_IdentifyID1
                       and identifyid2 = p_IdentifyID2
                       and invid = aId;
                    if ( aTaxType = '1' ) and ( aTaxRate = 0 ) then aTaxRate := 5; end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV031 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* �g�X�ֵo�������� */

              function WriteToInv032A(aSeq in varchar2, aId in varchar2, aBillId in varchar2, aBillidItemno in number,
                aAcctNo in varchar2, aFacisNo in varchar2)
                 return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv032A ( itemidref, invid, billid, billiditemno, seq, startdate, enddate,
                       itemid, description, quantity, unitprice, taxamount, saleamount, totalamount,
                       servicetype, chargeen, linktomis, facisno, accountno )
                    select b.itemidref, aId, aBillId, aBillidItemno, null, a.startdate, a.enddate,
                       a.itemid, a.description, a.quantity, a.unitprice, a.taxamount, ( a.quantity * a.unitprice ), a.totalamount,
                       a.servicetype, a.chargeen, nvl( a.linktomis, 'Y' ), aFacisNo, aAcctNo
                     from inv017 a, inv005 b
                    where b.IDENTIFYID1(+) = p_IdentifyID1
                      and b.IDENTIFYID2(+) = p_IdentifyID2
                      and b.compid(+) = p_CompId
                      and a.itemid = b.itemid(+)
                      and a.seq = aSeq  and a.billid = aBillId and a.billiditemno = aBillidItemno;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV032A �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* �g�o���Ȧs������ */

              function WriteToInv032(aId in varchar2, aSeq in number, aBillId in varchar2, aBillidItemno in number,
                aItemId in varchar2, aDesc in varchar2, aUnit in number, aTax in number, aSale in number,
                aTotal in number, aQuantity in integer) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    insert into inv032 ( identifyid1, identifyid2, invid, billid, billiditemno,
                       seq, startdate, enddate, itemid, description, quantity, unitprice,
                       taxamount, totalamount, chargeen, upttime, upten, servicetype, saleamount, linktomis,
                       facisno, accountno )
                    select p_IdentifyID1, p_IdentifyID2, aId, a.billid, a.billiditemno,
                       aSeq, a.startdate, a.enddate, aItemId, aDesc, aQuantity, aUnit,
                       aTax, aTotal, a.chargeen, to_date( p_LogDateTime, 'YYYY/MM/DD HH24:MI:SS' ),
                       p_User, a.servicetype, aSale, a.linktomis,
                       decode( a.itemidref, null, a.facisno, null ),
                       decode( a.itemidref, null, a.accountno, null )
                    from inv032A a
                   where a.invid = aId and a.billid = aBillId and a.billiditemno = aBillidItemno;
                   aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'INSERT INV032 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* ------------------------------------------------------------------- */

              /* ��s�o���Ȧs�D�� */

              function UpdateInv031(aMainInvId in varchar2, aRecord in TAutoCreate ) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    update inv031
                       set maininvid = aMainInvId,
                           saleamount = aRecord.SaleAmount,
                           taxamount = aRecord.TaxAmount,
                           invamount = aRecord.InvAmount,
                           mainsaleamount = aRecord.MainSaleAmount,
                           maintaxamount = aRecord.MainTaxAmount,
                           maininvamount = aRecord.MainInvAmount
                     where invid = aRecord.InvId;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE INV031 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* -------------------------------------------------------------------- */


              function UpdateInv032A(aId in varchar2, aOldSeq in number, aNewSeq in number) return boolean
              as
                aResult boolean;
              begin
                 aResult := false;
                 begin
                    update inv032a set invid = aId, seq = aNewSeq
                     where invid = 'X' and seq = aOldSeq;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE INV032A �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* -------------------------------------------------------------------- */

              function UpdateInv016(aSeq in varchar2) return boolean
              as
                 aResult boolean;
              begin
                 aResult := false;
                 begin
                    update inv016 set beassignedinvid = 'Y' where seq = aSeq;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE INV016 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


              /* ------------------------------------------------------------------- */

              /* ��s�ȪA�t��, �^��o����� */

              function UpdateMisSystem(aId in varchar2, aDate in date, aCId in varchar2, billid in varchar2,
                  billiditemno in number, aInvPCode in varchar2, aInvPName in varchar2) return boolean
              as
                 aResult boolean;
                 aSql varchar2(1024);
                 aTableName varchar2(100);
                 aPreInvoice varchar2(1);
              begin
                 aResult := false;
                 aTableName := GetTableNameByHowToCreate;
                 aPreInvoice := '0';
                 if ( p_HowToCreate = 1 ) then aPreInvoice := '1'; end if;
                 aSql := ' UPDATE ' || p_MisDbOwner || '.' || aTableName ||
                         '    SET GUINO = ' || '''' || aId || '''' || ',' ||
                         '        PREINVOICE = ' || '''' || aPreInvoice || '''' || ',' ||
                         '        INVDATE = TO_DATE( ' || '''' || to_char( aDate, 'YYYYMMDD' ) || '''' || ', ''YYYYMMDD'' ), ' ||
                         '        INVOICETIME = SYSDATE, ' ||
                         '        INVPURPOSECODE = ' || '''' || aInvPCode || '''' || ',' ||
                         '        INVPURPOSENAME = ' || '''' || aInvPName || '''' ||
                         '  WHERE CUSTID = ' || '''' || aCId || '''' ||
                         '    AND BILLNO = ' || '''' || billid || '''' ||
                         '    AND ITEM = ' || '''' || to_char( billiditemno ) || '''';
                 begin
                    execute immediate aSql;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := 'UPDTATE ' || aTableName ||' �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* ���o���X�ֶ��� */

              function GetMergeItem(aId in varchar2, aDesc in out varchar2 ) return boolean
              as
                aResult  boolean;
              begin
                aResult := true;
                begin
                   select a.description into aDesc from inv005 a
                   where a.Identifyid1 = p_IdentifyID1 and a.Identifyid2 = p_IdentifyID2
                     and a.compid = p_CompId and a.itemid = aId;
                exception
                   when others then aDesc := null;
                end;
                return aResult;
              end;

              /* ------------------------------------------------------------------- */

              /* ���]�ƧǸ��򦩱b�b�� */

              function GetMisInfo(aCId in varchar2, billid in varchar2, billiditemno in number,
                aAcctNo in out varchar2, aFacisNo in out varchar2) return boolean
              as
                aResult boolean;
                aSql varchar2(1024);
                aTableName varchar2(100);
                aTableName2 varchar2(100);
                aId number(1);
                aServiceType char(1);
              begin
                 aResult := false;
                 aTableName := GetTableNameByHowToCreate;
                 aTableName2 := 'SO002A' || GetDbLink;
                 aSql := ' SELECT A.ACCOUNTNO, A.FACISNO, A.SERVICETYPE, B.ID ' ||
                         '   FROM ' || p_MisDbOwner || '.' || aTableName || ' A, ' || p_MisDbOwner || '.' || aTableName2 || ' B ' ||
                         '  WHERE A.CUSTID = B.CUSTID(+)  ' ||
                         '    AND A.ACCOUNTNO = B.ACCOUNTNO(+) ' ||
                         '    AND A.CUSTID = ' || '''' || aCId || '''' ||
                         '    AND A.BILLNO = ' || '''' || billid || '''' ||
                         '    AND A.ITEM = ' || '''' || to_char( billiditemno ) || '''';
                 begin
                    execute immediate aSql into aAcctNo, aFacisNo, aServiceType, aId;
                    /* ���O P �A�ȧO, ���������q�ܪ��� */
                    if ( nvl( aServiceType, 'X' ) <> 'P' ) then aFacisNo := null; end if;
                    /* �b��O����1, �D�H�Υd�b��, ���� AccountNo */
                    if ( nvl( aId, 0 ) <> 1  ) then aAcctNo := null; end if;
                    /* �H�Υd���X, �u����4�X */
                    if ( Length( Nvl( aAcctNo, 'X' ) ) > 4 ) then
                      aAcctNo := substr( aAcctNo, length( aAcctNo ) - 3, 4 );
                    end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '���H�Υd���γ]�ƧǸ� ( ' || aTableName || ' ) �ɵo�Ϳ��~, �Ƚs  = ' || aCId || ', SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
              end;


							/* ------------------------------------------------------------------- */


							function GetMergeAmount(
							  aId in varchar2, aSeq in number, aTaxType in varchar2, aTaxRate in number,
							  aUnit in out number, aTax in out number,
							  aSale in out number, aTotal in out number,
							  aQuan in out integer, aCount in out integer) return boolean
							as
							  aResult boolean;
							begin
                 aResult := false;
                 begin
                    select sum( unitprice ), sum( taxamount ), sum( saleamount ), sum( totalamount ), sum( quantity ), count(1)
                      into aUnit, aTax, aSale, aTotal, aQuan, aCount
                      from inv032a
                     where invid = aId and seq = aSeq;
                    /* �Y�����|�h���s�p�� �P���B, �|�B, �]���Y�O�X�ֹL��, ���B�|���ǻ~�t */
                    if ( aTaxType = 1 ) then
                       aSale := round( aTotal / ( 1 + ( aTaxRate / 100 ) ) );
                       aTax := ( aTotal - aSale );
                       /* ���X�ֹL��, ���s�N����]�w���� �P���B�@�P */
                       if ( aCount > 1 ) then
                          aUnit := aSale;
                       end if;
                    end if;
                    aResult := true;
                 exception
                    when others then
                    begin
                       p_RetMsg := '�p��X�ֵo�����Ӫ��B/�ƶq�ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                       p_RetCode := -99;
                    end;
                 end;
                 return aResult;
							end;
				/*-----------------------------------------------------------------*/
				/*5358 ��X�X�ֶ��ت��̤p����ڰ_�l��P�̤j����ڨ��� �^��ܵo������ */
				function UpdInv032Date(aId in varchar2, aSeq in number) return boolean
					   as
						 aResult boolean;
						 aSqlMin varchar2(1024);
						 aSqlMax varchar2(1024);
						 aSQL  varchar2(2048);
					   begin
							aResult := false;
							begin
								aSqlMin := '(Select Min(A.StartDate) From ' || p_MisDbOwner || '.inv032A A,' || p_MisDbOwner || '.Inv005 B ' ||
										' Where A.invid=''' || aId || ''' And A.Seq=' || aSeq || ' And A.ItemID=B.ItemId And B.Sign=''' || '+'''||
										' And b.identifyid1 =''' || p_IdentifyID1 || ''' And b.identifyid2 = ' || p_IdentifyID2 || 
										' And B.CompId=' || p_CompId || ' And A.ItemIdRef is Not Null)' ;
								
								aSqlMax := '(Select Max(A.EndDate) From ' || p_MisDbOwner || '.inv032A A,' || p_MisDbOwner || '.Inv005 B ' ||
										' Where A.invid=''' || aId || ''' And A.Seq=' || aSeq || ' And A.ItemID=B.ItemId And B.Sign=''' || '+'''||
										' And b.identifyid1 =''' || p_IdentifyID1 || ''' And b.identifyid2 = ' || p_IdentifyID2 || 
										' And B.CompId=' || p_CompId || ' And A.ItemIdRef is Not Null)';
										
								aSQL := 'Update ' || p_MisDBOwner || '.inv032' ||
									 	' set StartDate='|| aSqlMin ||',' ||
										' EndDate='|| aSqlMax ||
										' Where seq ='||aSeq|| ' And invid='''||aid||''''||
										' And identifyid1='''||p_IdentifyID1||''''||
										' And identifyid2='||p_IdentifyID2;
							  execute immediate aSQL;
							  
												 
							  aResult := true;
							exception
							  when others then
							  begin
								 p_RetMsg := '��J�o�����_������o�Ϳ��~, SQLCODE = ' || SQLCODE || ',SQLERRM = ' || SQLERRM;
							  end;
							end;
							return aResult;
					   end;	 
							/* ------------------------------------------------------------------- */


begin


    if ( p_CompId is null ) then
       p_RetMsg := '�Ѽƿ��~: ���q�O���i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_PrefixString is null ) then
       p_RetMsg := '�Ѽƿ��~: ���w�}�ߵo�������i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_ChargeStartdate is null ) or ( p_ChargeStopDate is null ) then
       p_RetMsg := '�Ѽƿ��~: �}�ߵo�����w���O��(�_)(��)���i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_SystemID is null ) then
       p_RetMsg := '�Ѽƿ��~: �t�ΥN�X(SystemID)���i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_InvDateEqualToChargeDate <> 1 ) and ( p_InvDate is null ) then
       p_RetMsg := '�Ѽƿ��~: �o��������i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_LinkToMis = 'Y' ) and ( p_MisDbOwner is null ) then
       p_RetMsg := '�Ѽƿ��~: �w���w��s�ȪA�t��, ��ƪ�֦��̤��i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;

    if ( p_IdentifyID1 is null ) or ( p_IdentifyID2 is null ) then
       p_RetMsg := '�Ѽƿ��~: �ѧO�X(IdentifyID)���i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;


    if ( p_OrderBy is null ) then
       p_RetMsg := '�Ѽƿ��~: �o���}�߱ƧǤ覡���i���ŭȡC';
       p_RetCode := -1;
       return p_RetCode;
    end if;



    /* ���N�Ƕi�Ӫ��h�յo���r�y��� */

    aTempStr := p_PrefixString;
    aInv099MaxBounds := 0;

    begin
       while ( aTempStr is not null )
       loop
          aInv099MaxBounds := ( aInv099MaxBounds + 1 );
	        aInv099( aInv099MaxBounds ).CompId := p_CompId;
	        aInv099( aInv099MaxBounds ).YearMonth := p_InvYearMonth;
	        aInv099( aInv099MaxBounds ).Prefix := substr( aTempStr, 1, 2 );
	        aInv099( aInv099MaxBounds ).StartNum := substr( aTempStr, 3, 8 );
	        aInv099( aInv099MaxBounds ).EndNum := null;
	        aInv099( aInv099MaxBounds ).CurNum := null;
	        aInv099( aInv099MaxBounds ).Useful := 'Y';
	        aInv099( aInv099MaxBounds ).LastInvDate := null;
	        aInv099( aInv099MaxBounds ).HasChange := false;
          aPos := instr( aTempStr, ',' );
          if ( aPos <= 0 ) then
            aTempStr := null;
          else
            aTempStr := substr( aTempStr, aPos + 1, length( aTempStr ) - aPos );
          end if;
       end loop;
    exception
      when others then
      begin
        p_RetCode := -99;
        p_RetMsg := '��Ѧh�յo���r�y�o�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
        return p_RetCode;
      end;
    end;


    /* ���C�@���o��������� */

    for aIndex in 1..aInv099MaxBounds
    loop
       begin
          select a.curnum, a.endnum, a.lastinvdate
            into aInv099( aIndex ).CurNum, aInv099( aIndex ).EndNum, aInv099( aIndex ).LastInvDate
            from inv099 a
           where a.compid = aInv099( aIndex ).CompId
             and a.yearmonth = aInv099( aIndex ).YearMonth
             and a.prefix = aInv099( aIndex ).Prefix
             and a.startnum = aInv099( aIndex ).StartNum
             and a.useful = 'Y';
       exception
          when others then
          begin
             p_RetCode := -99;
             p_RetMsg := 'Ū���o�����ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
             return p_RetCode;
          end;
       end;
    end loop;


    /* �M�� Temp Table */

    begin
       dbms_utility.exec_ddl_statement( 'truncate table inv031' );
       dbms_utility.exec_ddl_statement( 'truncate table inv032' );
       dbms_utility.exec_ddl_statement( 'truncate table inv032a' );
    exception
       when others then
       begin
          p_RetMsg := '�R�� INV031/INV032/INV032A����, SQLCODE = ' || SQLCODE ||', SQLERRM = ' || SQLERRM;
          p_RetCode := -99;
          return p_RetCode;
       end;
    end;


    /* ���]�w�۰ʴ��}�o�����ӵ��Ƴ]�w */
    if not GetInv001Param( aAutoCreateNum, aFaciCombine ) then return p_RetCode; end if;


    /* �����B�z�ɶ�, �e�ݵ{���� Log �ɥ� */
    select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') into p_LogDateTime from dual;


    /* �}�l�ϥΪ��Ĥ@���o���� */
    aInv099Index := 1;


    /* ������w���o���D�ɶפJ���, ��������X��, �p�G�P�_�ᤣ���}��, �h�����g Log �i INV033 */
    for c1 in (
              select * from inv016 a where a.compid = p_CompId
                   and a.chargedate between to_date( p_ChargeStartdate, 'YYYY/MM/DD' ) and to_date( p_ChargeStopDate, 'YYYY/MM/DD' )
                   and a.beassignedinvid = 'N' and a.isvalid = 'Y'
		               and a.howtocreate = p_howtocreate
		               ---and a.invamount > 0 and a.taxtype <> '0' and a.shouldbeassigned = 'Y'
                   and a.stopflag = 0  --- PS: stopflag = 1 ����, ��ܳQ user ���w�R��, �o�ؤ���X��
                 order by
                   decode( p_OrderBy, 1, a.zipcode, to_char( a.chargedate, 'YYYYMMDD' ) ),
                   decode( p_OrderBy, 0, to_char( a.chargedate, 'YYYYMMDD' ), a.zipcode  ),
                   decode( p_OrderBy, 0, to_char( a.chargedate, 'YYYYMMDD' ), 1, a.zipcode, 2, a.zipcode, lpad( a.custid, 8, '0' ) )
               )
    loop

         aMasterCanCreate := true;
         aDetailCount := 0;
         aInv017TotalAmount := 0;

         /* ���ӥ��찣�F�Q�Хܬ��R�������~ ( �]���� user �ӻ�, ������Ƥw�Q�R�� ), �v�@�P�_, ���}�ߪ���, �g Log  */

         for c2 in ( select * from inv017 where seq = c1.seq )
         loop

                aNotes := null;

                /* �P�_�����O�_���n�}��( �D���� ) */

                /* �D�ɴN�Хܬ����}�� */
                if ( c1.shouldbeassigned = 'N' ) then
                   aNotes := '���w�������}��';
                /* �D�ɪ����B�� 0, �Ω��ӼХܬ��}��, ���O���Ӹӵ����`���B���s ( �`�N: ���i�ର�t�����, �p��s�����` ) */
                elsif ( c1.invamount <= 0 ) or ( ( c2.shouldbeassigned = 'Y' ) and ( c2.totalamount = 0 ) ) then
                   aNotes := '�o�����B�p��ε���s';
                /* �D�ɪ��|�O�������}��, �Ω��ӼХܬ��}��, ���O�|�O�����}�� */
                elsif ( c1.taxtype = '0' ) or ( ( c2.shouldbeassigned = 'Y' ) and ( c2.taxtype = '0' ) ) then
                   aNotes := '�o���|�O�����}��';
                end if;


                /* �g Log ��, ���U�� */
                if ( aNotes is not null ) then
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                     return p_RetCode;
                   end if;
                   aMasterCanCreate := false;
                   goto DoNothing;
                end if;


                /* �P�_�����ӬO�_���n�}��, ( �������Ӥ��}�ߤ��N����������}, �]���䥦���ӥi�භ�n�}�� ) */
                if ( c2.shouldbeassigned = 'N' ) then
                   aNotes := '���w�������}��';
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                     return p_RetCode;
                   end if;
                   goto DoNothing;
                end if;



                /* �����J���ˬd�ݶ}�ߩ��ӬO�_�i�H�}�ߵo��, �u�n�䤤�@�����ŦX����, ��i�N�����H�}�� */
                /* �ˮ֮ɨC�@���ˮ֧�, ���i�H�]���䤤�@�����L�N�������L�䥦���ˮ�, �]���C�@�����n Log */

                /* 1. �O�_���n�ˮ֫ȪA�t�� */
                if ( p_LinkToMis = 'Y' ) and ( c2.linktomis = 'Y' ) then
                    /* �ˮ֫ȪA�t�θ�� */
                    if not CheckMisSystem( c1.custid, c2.billid, c2.billiditemno, aNotes ) then
                       if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                         return p_RetCode;
                       end if;
                       aMasterCanCreate := false;
                       goto DoNothing;
                    end if;
                end if;


                /* 2. �ˮֵo���t�Τ��O�_�w�������O�渹�򶵦� */
                if ( CheckAlreadyCreate( c2.billid, c2.billiditemno, aNotes ) ) then
                    if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, aNotes ) then
                      return p_RetCode;
                    end if;
                    aMasterCanCreate := false;
                    goto DoNothing;
                end if;


                /* 3. �P�_���ӵ|�O�O�_�P�D�ɤ��� */
                if ( c1.taxtype <> c2.taxtype ) then
                   if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, '���O���Ӫ��|�v�O�P�D�ɤ��P' ) then
                     return p_RetCode;
                   end if;
                   aMasterCanCreate := false;
                   goto DoNothing;
                end if;


                /* 4. �[�`���Ӫ��B */
                aInv017TotalAmount := nvl( aInv017TotalAmount, 0 ) + c2.totalamount;

                /* ���ӽT�w�i�H�}�ߪ����� */

                aDetailCount := ( aDetailCount + 1 );

                << DoNothing >>

                   null;

         end loop;



         /* ���ӵ���, �p�G�P�_�O�i�H�}�ߪ��o��, �h�A�P�_�D���ɪ����B�O�_�۲� */
         if ( aMasterCanCreate ) and ( aDetailCount > 0 ) and  ( c1.invamount <> aInv017TotalAmount ) then
            for c2 in ( select billid, billiditemno from inv017 where seq = c1.seq and shouldbeassigned = 'Y' )
            loop
              if not LogToInv033( c1.seq, c2.billid, c2.billiditemno, c1.custid, c1.chargetitle, '�D���ɵo�����B����' ) then
                return p_RetCode;
              end if;
            end loop;
            aMasterCanCreate := false;
         end if;


         /* �T�w�i�H�}�ߦ��i�o�� */
         if ( aMasterCanCreate ) and ( aDetailCount > 0 ) then

               /* �g�o���t�Ϊ��Ȥ�D�� */
               if not CheckInv002Exists( c1.custid ) then
                  if not WriteToInv002( c1.seq ) then  return p_RetCode; end if;
               end if;


               /* �g�o���t�Ϊ��Ȥ������ */
               if not CheckInv019Exists( c1.custid  ) then
                  if not WriteToInv019( c1.seq ) then return p_RetCode; end if;
               end if;


               aDetailSeq := 0;

               /* ��i�H�}�ߪ�����, ���g�J INV032A �ǳư����� */
               for c2 in ( select * from inv017 where seq = c1.seq and shouldbeassigned = 'Y'
                  and totalamount <> 0 and taxtype <> '0' order by billid, billiditemno )
               loop

                   /* �����w����r�ɶפJ���b��/�d��  */
                   aAccountNo := c2.accountno;
                   /* �]�Ƹ��X�èS���q��r�ɶפJ */
                   aFacisNo := null;

                   /* �Y�O�P�ȪA�t�Τ���, �h���H�Υd����CP�����q�ܸ��X(�A�ȧO��P) */
                   if ( p_LinkToMis = 'Y' ) and ( nvl( c2.linktomis, 'Y' ) = 'Y' ) then
                      if not GetMisInfo( c1.custid, c2.billid, c2.billiditemno, aAccountNo, aFacisNo ) then
                        return p_RetCode;
                      end if;

                   end if;

                   /* �g������ INV032A */
                   /*
                      �`�N: ���ɼg�� INV032A ���i�����o�����X, �᭱���n������, �X��, �B�z����T�w�~�i���o�����X
                      �y�{: write to inv032a --> ����,�X�� --> inv032
                   */

                   if not WriteToInv032A( c2.seq, 'X', c2.billid, c2.billiditemno, aAccountNo, aFacisNo ) then
                     return p_RetCode;
                   end if;


               end loop;



               /* �X�ֵo������  1.����  2.�s�Ǹ� (INV032A) */


               aDetailSeq := 0;
               aOldItemRefId := null;

               for c2 in (  select itemidref, facisno, count(1) from inv032a
                             where invid = 'X'
                               and itemidref is not null
                             group by itemidref, facisno
                          union all
                            select itemidref, facisno, 1 from inv032a
                             where invid = 'X'
                               and itemidref is null
                             order by itemidref, facisno   )
               loop


                    /* 1.���X�ֶ��إB���]�ƧǸ� */
                    if ( c2.itemidref is not null ) and ( c2.facisno is not null ) then


                        /* �ѼƳ]�w�����P�]�ƧǸ�, �����n�X�� -> FaciCombine = 0 */
                        if ( aFaciCombine = 0 ) then
                           aDetailSeq := ( aDetailSeq + 1 );
                        /* �ѼƳ]�w�����P�]�ƧǸ�, ���n�X��  -> FaciCombine = 1 */
                        elsif ( nvl( aOldItemRefId, -1 ) <> c2.itemidref ) then
                           aDetailSeq := ( aDetailSeq + 1 );
                           aOldItemRefId := c2.itemidref;
                        end if;


                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref = c2.itemidref and facisno = c2.facisno
                           and seq is null;


                    /* 2. ���X�ֶ���, �S���]�ƧǸ� */
                    elsif ( c2.itemidref is not null ) and ( c2.facisno is null ) then

                        aDetailSeq := ( aDetailSeq + 1 );
                        aOldItemRefId := null;

                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref = c2.itemidref and facisno is null
                           and seq is null;

                    /* 3. �S���X�ֶ��� */
                    else

                        aDetailSeq := ( aDetailSeq + 1 );
                        aOldItemRefId := null;

                        update inv032a set seq = aDetailSeq
                         where invid = 'X' and itemidref is null and seq is null and rownum <= 1;

                    end if;


               end loop;


               /* ������, ���o�����X, �üg�J INV032 */

               aDetailSeq := 0;              --- �u���g�J��o�����Ӫ��Ǹ�
               aAutoCreateMaxBounds := 0;    --- �۰ʴ��}�O�� List ���ƶq
               aOldSeq := null;              --- �����ǧǸ�

               for c2 in ( select * from inv032A where invid = 'X' order by seq )
               loop

                   if ( c2.seq <> nvl( aOldSeq, 0 ) ) then

                       /* ���P�_���ӵ��ƬO�_�w�j��]�w�۰ʴ��}������, �Y�O�w�j�󴫶}����, �h���m���ӧǸ� */
                       if ( aAutoCreateNum > 0 ) then
                         if ( aDetailSeq >= aAutoCreateNum ) then
                            aDetailSeq := 0;
                         end if;
                       end if;

                       aDetailSeq := ( aDetailSeq + 1 );

                       /* ���ӲĤ@��, ���o�����X */
                       if ( aDetailSeq = 1 ) then

                           /* ���u�����o���� */
                           if ( p_InvDateEqualToChargeDate = 1 ) then
                             aRealInvDate := c1.chargedate; ---�o���鵥�󦬶O��
                           else
                             aRealInvDate := to_date( p_InvDate, 'YYYY/MM/DD' );  ---�o���鵥��ǤJ���o����
                           end if;

                           /* ���o�����X */
                           if not GetInvNumber( aInv099Index, aInvId, aRealInvDate ) then return p_RetCode;  end if;

                           /* �p��o���ˬd�X */
                           aCheckNo := CalcInvCheckSum( substr( aInvId, 3, length( aInvId ) - 2 ),
                             p_SystemID || to_char( to_number( to_char( aRealInvDate, 'YYYYMMDD' ) ) - 19110000 ) );

                           /* �g�D��, ���K���^�ӱi�o�����|�O�ε|�v, �X�֩��ӭ��s�p��|�B�ξP���B��,�|�Ψ� */
                           if not WriteToInv031( c1.seq, aInvId, aCheckNo, aRealInvDate, aInvTaxType, aInvTaxRate ) then
                              return p_RetCode;
                           end if;

                           /* �N�����}�ߪ��o�����X�O���U��, ���B���m */
                           aAutoCreateMaxBounds := ( aAutoCreateMaxBounds + 1 );
                           aAutoCreateList( aAutoCreateMaxBounds ).InvId := aInvId;
                           aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).InvAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainSaleAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainTaxAmount := 0;
                           aAutoCreateList( aAutoCreateMaxBounds ).MainInvAmount := 0;

                       end if;

                       /* ����s�o�����X�Υ��T���Ǹ��^ INV032A */
                       if not UpdateInv032A( aInvId, c2.seq, aDetailSeq ) then return p_RetCode; end if;

                       aItemDesc := null;

                       /* �X�ְѦҸ�����, ���ѦҸ����W�� */
                       /* �S����, �άO�S���]�w�ѦҸ�, �ζ��ɶi�Ӫ��W�� */
                       if ( c2.itemidref is not null ) then
                         if not GetMergeItem( c2.itemidref, aItemDesc ) then return p_RetCode; end if;
                       end if;

                       /* ���X�֫�ӵ����Ӫ����B, �����εo�����X + �s���Ǹ� */

                       aMergeUnitPrice := 0;
                       aMergeTaxAmount := 0;
                       aMergeSaleAmount := 0;
                       aMergeTotalAmount := 0;
                       aMergeQuantity := 0;
                       aMergeCount := 0;

                       if not GetMergeAmount( aInvId, aDetailSeq, aInvTaxType, aInvTaxRate,
                         aMergeUnitPrice, aMergeTaxAmount, aMergeSaleAmount,
                         aMergeTotalAmount, aMergeQuantity, aMergeCount ) then
                          return p_RetCode;
                       end if;


                       /* �p�G�X�֪����ƶW�L 1 ��, �h�g�J�� INV032 �� �ƶq�h�� 1, �_�h�̷ӭ쥻���ƶq */
                       if ( aMergeCount > 1 ) then
                         aMergeQuantity := 1;
                       end if;


                       /* �N INV032A ���X�ֶ��ئ^�g�� INV032 ������ */
                       if not WriteToInv032( aInvId, aDetailSeq, c2.billid, c2.billiditemno,
                           nvl( c2.itemidref, c2.itemid ), nvl( aItemDesc, c2.description ),
                           aMergeUnitPrice, aMergeTaxAmount, aMergeSaleAmount, aMergeTotalAmount, aMergeQuantity ) then
                         return p_RetCode;
                       end if;

                       aOldSeq := c2.seq;


                       /* �֭p�ӱi�o���D�ɤU�����Ӫ��B */
                       aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount :=
                         ( aAutoCreateList( aAutoCreateMaxBounds ).SaleAmount + aMergeSaleAmount );
                       aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount :=
                         ( aAutoCreateList( aAutoCreateMaxBounds ).TaxAmount + aMergeTaxAmount );
                       aAutoCreateList( aAutoCreateMaxBounds ).InvAmount :=
                        ( aAutoCreateList( aAutoCreateMaxBounds ).InvAmount + aMergeTotalAmount );


                       /* ���N�D�o�����B�O���b�Ĥ@�i�}�ߪ��o����*/
                       aAutoCreateList( 1 ).MainSaleAmount :=
                         ( aAutoCreateList( 1 ).MainSaleAmount + aMergeSaleAmount );
                       aAutoCreateList( 1 ).MainTaxAmount :=
                         ( aAutoCreateList( 1 ).MainTaxAmount + aMergeTaxAmount );
                       aAutoCreateList( 1 ).MainInvAmount :=
                        ( aAutoCreateList( 1 ).MainInvAmount + aMergeTotalAmount );


                   end if;


               end loop;
               for c2 in ( select A.seq from inv032A  A,inv005 B where A.invid = ainvId
                                and A.ItemIdRef is not null And A.ItemId=B.ItemId And B.Sign='+'
								and b.identifyid1=p_IdentifyID1 and b.identifyid2=p_IdentifyID2
								and B.CompId=p_CompId
								group by seq
                                order by seq )
                loop
					if not UpdInv032Date(ainvId,C2.Seq) then
						return p_RetCode;
					end if;
                   
                end loop;

               /* �C�B�z���@�i�D�ɤU������, �ˮ֬O�_���|���� */

               select count(1) into aCount from inv032a where invid = 'X';

               if ( aCount > 0 ) then
                  p_RetMsg := '�X�ֵo�����Ӯɵo�Ϳ��~, �|�����ӥ������X�֡C';
                  p_RetCode := -99;
                  return p_RetCode;
               end if;


               /* �N�Ҧ��o�����D�o�����B��J */
               for aIndex in 1..aAutoCreateMaxBounds loop
                 aAutoCreateList( aIndex ).MainSaleAmount := aAutoCreateList( 1 ).MainSaleAmount;
                 aAutoCreateList( aIndex ).MainTaxAmount := aAutoCreateList( 1 ).MainTaxAmount;
                 aAutoCreateList( aIndex ).MainInvAmount := aAutoCreateList( 1 ).MainInvAmount;
               end loop;


               /* �^�Y��s�D�ɪ� maininvid, saleamount, taxamount, invamount */
               /* PS: MainInvId �^��̫�@�i�o�����X���D�o�����X, �Y�O�u���@�i�o�����P�� invid = maininvid */
               for aIndex in 1..aAutoCreateMaxBounds loop
                      if not UpdateInv031( aInvId, aAutoCreateList( aIndex ) ) then
                      return p_RetCode;
                   end if;
               end loop;


               /* ��s�o���ӷ��� */
               if not UpdateInv016( c1.Seq ) then return p_RetCode; end if;

         end if;

    end loop;



    /* �g�J�o���D�� */
    begin
       insert into inv007 ( identifyid1, identifyid2, checkno, invid, canmodify, invdate, chargedate,
           compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat,
           taxtype, taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete,
           howtocreate, upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount,
           maininvamount, invuseid, invusedesc )
       select p_IdentifyID1, p_IdentifyID2, checkno, invid, canmodify, invdate, chargedate,
           compid, custid, custsname, invtitle, zipcode, invaddr, mailaddr, businessid, invformat,
           taxtype, taxrate, saleamount, taxamount, invamount, printcount, ispast, isobsolete,
           howtocreate, upttime, upten, printfun, maininvid, mainsaleamount, maintaxamount,
           maininvamount, invuseid, invusedesc
         from inv031
        where identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2
          and compid = p_CompId;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '�g�J�o���D�ɮɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;


    /* �g�J�o�������� */
    begin
       insert into inv008 ( identifyid1, identifyid2, invid, billid, billiditemno, seq,
          startdate, enddate, itemid, description, quantity, unitprice,
          taxamount, totalamount, chargeen, upttime, upten, servicetype, saleamount, linktomis )
       select p_IdentifyID1, p_IdentifyID2, b.invid, b.billid, b.billiditemno, b.seq,
          b.startdate, b.enddate, b.itemid, b.description, b.quantity, b.unitprice,
          b.taxamount, b.totalamount, b.chargeen, b.upttime, b.upten, b.servicetype, b.saleamount, nvl( b.LinkToMis, 'Y' )
         from inv031 a, inv032 b
        where a.identifyid1 = b.identifyid1 and a.identifyid2 = b.identifyid2
          and a.invid = b.invid
          and a.identifyid1 = p_IdentifyID1 and a.identifyid2 = p_IdentifyID2
          and a.compid = p_CompId;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '�g�J�o�������ɮɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;



    /* �g�J�o���X����, �u���X�֪��~�g�L�h */
    begin
       insert into inv008A ( itemidref, invid, billid, billiditemno, seq,
          startdate, enddate, itemid, description, quantity, unitprice,
          taxamount, totalamount, servicetype, saleamount, facisno, accountno )
       select b.itemidref, b.invid, b.billid, b.billiditemno, b.seq,
          b.startdate, b.enddate, b.itemid, b.description, b.quantity, b.unitprice,
          b.taxamount, b.totalamount, b.servicetype, b.saleamount, b.facisno, b.accountno
         from inv031 a, inv032a b
        where a.invid = b.invid
          and a.identifyid1 = p_IdentifyID1 and a.identifyid2 = p_IdentifyID2
          and a.compid = p_CompId
          and b.itemidref is not null;
    exception
       when others then
       begin
          p_RetCode := -99;
          p_RetMsg := '�g�J�o���X���ɮɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
          return p_RetCode;
       end;
    end;



    /* �o���B�z����, ��s�ȪA�t��, �� INV032A ���X�֩��Ӹ�ƨӧ�s */
    if ( p_LinkToMis = 'Y' ) then
       for c1 in ( select a.invdate, a.invid, a.custid, a.invuseid, a.invusedesc, b.billid, b.billiditemno from inv031 a, inv032a b
          where a.invid = b.invid and nvl( b.linktomis, 'Y' ) = 'Y' and a.compid = p_CompId  order by b.invid, b.seq )
       loop
          if not UpdateMisSystem( c1.invid, c1.invdate, c1.custid, c1.billid, c1.billiditemno, c1.invuseid, c1.invusedesc ) then
             return p_RetCode;
          end if;
       end loop;
    end if;



    /* ��s�o���� */
    for aIndex in 1..aInv099MaxBounds
    loop
       if ( aInv099( aIndex ).HasChange ) then
          begin
             update inv099 set curnum = aInv099( aIndex ).CurNum, useful = aInv099( aIndex ).Useful,
                lastinvdate = aInv099( aIndex ).LastInvDate, upttime = sysdate, upten = p_User
              where identifyid1 = p_IdentifyID1 and identifyid2 = p_IdentifyID2 and compid = p_CompId
                and yearmonth = aInv099( aIndex ).YearMonth and prefix = aInv099( aIndex ).Prefix
                and startnum = aInv099( aIndex ).startnum;
          exception
             when others then
             begin
                p_RetMsg := 'UPDTATE INV099 �ɵo�Ϳ��~, SQLCODE = ' || SQLCODE || ', SQLERRM = ' || SQLERRM;
                p_RetCode := -99;
                return p_RetCode;
             end;
          end;
       end if;
    end loop;

    p_RetCode := 0;
    p_RetMsg := '�o���}�߰��槹��';
    return p_RetCode;

end;
/

/*
@c:\gird\v350\script\SF_RptA1_2
@c:\gird\v300\script\SF_RptA1_2
variable nn number
variable msg varchar2(500)
exec :nn := SF_RptA1_2('20000101', '20000131', '20020531', 'C', 7, null, '=', null, null, 'IN (1, 20, 30, 36)', null, null, null, null, null, null, 2, 0, null, 'IN (2,3)', :msg);
print nn
print msg

prompt ���ժ����O
exec :nn := SF_RptA1_2('19980601', '19980630', '19980731', 'C', 2, 455029, '=', null, null, null, null, null, null, null, null, null, 2, 5, null, null, :msg);

-- exec :nn := SF_RptA1_2('20000101', '20010215', '20010131', 'I', 1, 2, 5, :msg);
print nn
print msg

prompt ���������O
exec :nn := SF_RptA1_2('20000101', '20000131', '20010930', 'C', 1, 111526, '=',null ,null , null, null, null, null, null, null, null, 2, 0, null, null, :msg);
exec :nn := SF_RptA1_2('20010105', '20010630', '20010630', 'C', 3, null, null, null ,null , null, null, null, null, null, null, null, 2, 0, null, null, :msg);
print nn
print msg

  ���͹w���u������(A1����)���Ӹ��: SF_RptA1()
  �ɦW: SF_RptA1.sql

  (1) �J�`�γv������������, �i�̾ڦ������ɲ���
  (2) �ƧǤ覡������ͮ�, �A�i��Ƨ�, �������ɸ�Ƥ��ƥ��Ƨ�
  (3) �Y����|�ݨ�L���, �i�Q�Φ������ɥhjoin��L��, �H���͸���줺�e

  IN	p_RealDate1 varchar2: 	���O��_�l, 'YYYYMMDD'
	p_RealDate2 varchar2:	���O��I��, 'YYYYMMDD'
	p_RefDate varchar2:	�ѦҤ��, 'YYYYMMDD'
	p_ServiceType char(1):	�A�����O, ��: 'C', 'I' or null
	p_CompCode number:	���q�O

	p_CustId number:	�Ȥ�s��
	p_CustIdOp varchar2:	�Ȥ�s�������, ��: '=', '>', '<=', ...
	p_ShouldDate1 varchar2:	������_�l, 'YYYYMMDD'
	p_ShouldDate2 varchar2:	������I��, 'YYYYMMDD'
	p_CitemSQL varchar2:	���O����SQL����r��
	--p_CompSQL varchar2:	���q�OSQL����r��
	p_ClassSQL varchar2:	�Ȥ����OSQL����r��
	p_BillTypeSQL varchar2: ������OSQL����r��
	p_ServSQL varchar2:	�A��SQL����r��
	p_AreaSQL varchar2:	��F��SQL����r��
	p_StrtSQL varchar2:	��DSQL����r��
	p_CircuitNo varchar2:	�����s������, null=����, '-1'=�ťժ̤~�έp, ��L�r��=�ŦX���r��~�έp

	p_TaxMode number:	�|�v�p��覡, 1=���|, 2=�w�|
	p_ErrYear number:	�p�O�_�l��W�L�X�~�������~����
	p_MDUSQL varchar2:	�j�ӽs��SQL����r��
	p_Citem2SQL varchar2:	�D�g�����O����SQL����r��			<==== NEW added: 2001.11.15
  OUT 
      p_RetMsg VARCHAR2�G���G�T�� (�I�s���ܼƦܤֻ�100 bytes)

  Return NUMBER�G���G�N�X
        0: ���`����
	-1: p_RetMsg='�Ѽƿ��~'
	-2: p_RetMsg='�u���N�X�ɤ��]�w�b�b�ѦҸ����~'
        -3: p_RetMsg='���O�Ѽ��ɧ䤣��Ӥ��q�θӪA�����O�����'
        -4: p_RetMsg='SQL���~: '
	-5: p_RetMsg='�Ҧ��g���ʦ��O���ت��|�v�����]���@�P'
        -99�Gp_RetMsg='��L���~'

  By: Pierre
  Date: 2001.02.23
	2001.05.14 �[�J�j�ӽs��SQL����
	2001.08.24 �令�i�B�z�h���ѦҸ��]���b�b���u����]�N�X
	2001.11.16 ��: �[���󪫥�--�D�g���ʶ���
	2003.03.22 ��: �{���^�ǿ��~�T����, �NOracle�����~�T���[�J, �H�K�ϧO�O�{�����D,
			��ư��D,�Ϊ̬OOracle���D

	2003.04.29 ��:  1. �D�g���ʭt����ƭY�����Ĵ����d��, �h���u�ܨC��. �Y�L���Ĵ����d��, �h���䥿�����
			   �����Ĵ����ӭp��ӭt�������Ĵ���, ����A���u�ܨC��.
			2. �Y�t����ƥi�����쥿��, �h�s��SO089��, �N�䦬�O���إN�X�P�W�٧令�P�������
	2003.05.02 ��: 2�B�[upper(), �H�ѨM�e�ݶǨӤp�g'in'�����D
	2003.05.14 ��: �Y998�t�����_�l��񥿶����_�l���٦�, �o�O���X�z���t��, �N�t�����_�l��אּ������
			�_�l��ӭp��

  Modify: 2001.10.04 by Lawrence
        2001.10.04 so043�[�@�Ѽ�(Taxrate)�ѥ������t�|�p���,��"�ˬd���O���إN�X�ɤ����|�v�]�w����, �Y�Ҧ��g���ʦ��O���ت��|�v�����]���@�P"�������P�_
        2001.10.27 ��: �[�P�_�_����P�@��BSO043.Para3=1��
        2002.05.10 ��: �[�P�_�ꦬ����=0�B���g���ʪ�insert into SO090
        2002.10.07 ��: �C2000��Commit�@��


  By: Stanley
        2002.03.01 �bSQL�y�k���[�JSo014��Join
        2003.05.07 �W�[���q�O�ΪA�����O���� 
*/
/*
-- ���ժ� function header
create or replace function SF_RptA1_2
  (p_RealDate1 varchar2, p_RealDate2 varchar2, p_RefDate varchar2, 
   p_ServiceType char, p_CompCode number, p_TaxMode number, p_ErrYear number, 
   p_RetMsg out varchar2) 
  return number as
*/
-- ������ function header
create or replace function SF_RptA1_2
  (p_RealDate1 varchar2, p_RealDate2 varchar2, p_RefDate varchar2, 
   p_ServiceType char, p_CompCode number, p_CustId number, p_CustIdOp varchar2, 
   p_ShouldDate1 varchar2, p_ShouldDate2 varchar2, p_CitemSQL varchar2,
   p_ClassSQL varchar2, p_BillTypeSQL varchar2, p_ServSQL varchar2, 
   p_AreaSQL varchar2, p_StrtSQL varchar2, p_CircuitNo varchar2,
   p_TaxMode number, p_ErrYear number, p_MDUSQL varchar2, p_Citem2SQL varchar2, p_RetMsg out varchar2) 
  return number as

  v_Para3 number;
  v_Para17 number;
  v_OldP3 number;
  v_Cnt   number := 0;
  v_TaxCodeCnt number := 0;
  v_CurrTaxRate number;
  v_TmpErrCnt number := 0;
  v_BadCode  number;
  v_BadCnt   number := 0;
  v_StartExecTime date;

  c39		char(1) := chr(39);
  v_SQL1 	varchar2(4000);
  v_Column1	varchar2(1000);
  v_Where1	varchar2(500);
  v_MoreTable   varchar2(100);
  v_Join	varchar2(100);

  v_TotalDay number;
  v_PreDay number;
  v_PastAmt number;
  v_01Amt number;
  v_02Amt number;
  v_03Amt number;
  v_04Amt number;
  v_05Amt number;
  v_06Amt number;
  v_07Amt number;
  v_08Amt number;
  v_09Amt number; 
  v_10Amt number;
  v_11Amt number;
  v_12Amt number;
  v_ElseAmt number;
  v_Flag number;

  v_CurrMonth number := 0;
  v_RefMonth number := 0;
  v_CurrDate date;
  v_BeginDate date;
  v_LastDate date;

  v_CurrAmt number;
  v_LeftAmt number;
  v_LeftDay number;
  v_CurrDay number;
  v_Mndx number;
  v_ClctMonth number;
  v_TaxRate number;
  v_Count number:=0;

  type CurTyp is ref cursor;  --�ۭqcursor���A
  v_DyCursor CurTyp;          --��dynamic SQL

  x_BillNo varchar2(15);
  x_Item number;
  x_CustId number;
  x_CitemCode number;
  x_CitemName varchar2(20);
  x_RealDate date;
  x_RealAmt number;
  x_RealPeriod number;
  x_RealStartDate date;
  x_RealStopDate date;
  x_STCode number;
  x_ClctEn varchar2(10);
  x_AddrNo number;
  x_ServCode varchar2(3);
  x_ServiceType char(1);

  y_RealAmt number;
  y_PeriodFlag number;
  -- for �|�v�ˬd 
  cursor cTaxRate is
	select  B.CodeNo, B.Rate1 
	  from cd019 A, cd033 B where A.TaxCode=B.CodeNo(+) and A.PeriodFlag = 1 and (A.ServiceType=p_ServiceType or A.ServiceType is null)
	  group by b.CodeNo, B.Rate1;
  -- �b�b
  cursor cBad is
	select CodeNo from CD016 where RefNo = 1 and (ServiceType=p_ServiceType or ServiceType is null);
  v_BadCodeString varchar2(300);

  -- 2003.04.29
  s_CitemCode number;		-- �h�O���ةҹ������즬�O����
  s_BillNo varchar2(15);	-- �h�O���ةҹ�������渹
  s_Item number;		-- �h�O���ةҹ������춵��
  x_ShouldDate date;		-- ������/���h��
  v_TmpCitem varchar2(500);	-- for�g�����O���سB�z��
  I number;
  v_Index  number;
  v_CitemCodeCnt number := 0;
  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER;
  CitemCode NumberAry;

  x_TmpStartDate date;

begin

  -- Check parameters
  if p_RealDate1 is null or p_RealDate2 is null or p_RefDate is null or
    p_CompCode is null or p_ServiceType is null then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  -- �B�z���O����SQL����r��, �N���নarray, �H�K����P�_�t����Ʃҹ����������O�_���󦹱��� 2003.04.29
  DBMS_OUTPUT.PUT_LINE('1. �B�z���O����SQL����r��');
  v_CitemCodeCnt := 0;
  if p_CitemSQL is not null then
    v_TmpCitem := rtrim(ltrim(p_CitemSQL));
    if upper(substr(v_TmpCitem, 1, 2)) = 'IN' then			-- 'IN (...,..,..)'
      v_TmpCitem := substr(v_TmpCitem, 5);
      v_TmpCitem := substr(v_TmpCitem, 1, length(v_TmpCitem)-1);
    elsif upper(substr(v_TmpCitem, 1, 6)) = 'NOT IN' then		-- 'NOT IN (...,..,..)'
      v_TmpCitem := substr(v_TmpCitem, 9);
      v_TmpCitem := substr(v_TmpCitem, 1, length(v_TmpCitem)-1);
    elsif substr(v_TmpCitem, 1, 1) = '=' then			-- '=...'
      v_TmpCitem := ltrim(substr(v_TmpCitem, 2));
    elsif substr(v_TmpCitem, 1, 2) = '!=' then			-- '!=...'
      v_TmpCitem := ltrim(substr(v_TmpCitem, 3));
    end if;

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

  -- �����O�Ѽ��ɬ����Ѽ�: Para3: �����_�Ϥ�P�����I���
  -- �ثe�Τ��q�O=1����ƨӰ�, �Y�D�Ӥ��q, �h���q�O�������n���
  DBMS_OUTPUT.PUT_LINE('2. �����O�Ѽ��ɬ����Ѽ�');
  if p_ServiceType is null then
    begin
      select Para3, Para17,TaxRate into v_Para3, v_Para17, v_TaxRate from SO043 where CompCode=p_CompCode;
    exception
      when others then
        p_RetMsg := '�Y�h���q�O�B�h�A�Ⱥ������t��, �h�A�����O�����n���';
        return -3;
    end;
  else
    begin
      select Para3, Para17, TaxRate into v_Para3, v_Para17, v_TaxRate from SO043 where CompCode=p_CompCode and ServiceType=p_ServiceType;
    exception
      when others then
        p_RetMsg := '���O�Ѽ��ɧ䤣��Ӥ��q�O�θӪA�����O�����';
        return -3;
    end;
  end if;

  v_OldP3 := v_Para3;
  if v_Para3=1 then
     v_Para3 := 0;
  else
     v_Para3 := 1; 
  end if; 

  -- �ˬd�u���N�X�ɤ��]�w�b�b�ѦҸ�
/*
  begin
    select CodeNo into v_BadCode from CD016 where RefNo = 1 and (ServiceType=p_ServiceType or ServiceType is null);
  exception
    when others then
      p_RetMsg := '�u���N�X�ɤ��]�w�b�b�ѦҸ����~';
      return -2;
  end;
*/
  DBMS_OUTPUT.PUT_LINE('3. ���u���N�X');
  for cr1 in cBad loop
    v_BadCodeString := v_BadCodeString || '(' || ltrim(to_char(cr1.CodeNo,'999')) || ')';
    v_BadCnt := v_BadCnt + 1;
  end loop;
  -- DBMS_OUTPUT.PUT_LINE(v_BadCodeString);
  if v_BadCodeString is null then
    v_BadCodeString := '(-1)';
  end if;
/*  if v_BadCodeString is null then
    p_RetMsg := '�u���N�X�ɤ��]�w�b�b�ѦҸ����~(���]RefNo=1�����)';
    return -2;
  end if; */

/*
  -- �ˬd���O���إN�X�ɤ����|�v�]�w����, �Y�Ҧ��g���ʦ��O���ت��|�v�����]���@�P, �h�ˬd��
  if p_TaxMode = 1 then
    for cr2 in cTaxRate loop
      v_TaxCodeCnt := v_TaxCodeCnt + 1;
      v_CurrTaxRate := nvl(cr2.Rate1, 0);
    end loop; 
    if v_TaxCodeCnt != 1 then
      p_RetMsg := '�Ҧ��g���ʦ��O���ت��|�v�����]���@�P';
      return -5;
    end if;
  else		-- �����w�|(�|���t)�覡�ӭp��
    v_CurrTaxRate := 0;
  end if;
*/

 -- �ק� 2001.10.03 by Lawrence
  if p_TaxMode = 1 then
    v_CurrTaxRate := v_TaxRate;
  else		
    v_CurrTaxRate := 0;
  end if;

  -- ������SQL���O
  -- �զXSQL�d�ߦr��: �ثe�e��SO5A10�W���Ҧ��d�߰ѼƩ�SO033������ɤ��Ҧ�, �G����join��L��
  DBMS_OUTPUT.PUT_LINE('4. �զXSQL�d�ߦr��');
  v_Column1 := 'select A.BillNo, A.Item, A.CustId, A.CitemCode, A.CitemName, A.RealDate, A.RealAmt, ' ||
		'A.RealPeriod, A.RealStartDate, A.RealStopDate, A.STCode, A.ClctEn, A.AddrNo, A.ServCode, A.ServiceType, A.sCitemCode, A.sBillNo, A.sItem, A.ShouldDate';
  v_Where1  :=  ' A.RealDate>=' || GiPackage.QryDTString0(to_date(p_RealDate1,'YYYYMMDD')) ||
		' and A.RealDate<=' || GiPackage.QryDTString0(to_date(p_RealDate2,'YYYYMMDD')) ||
		' and A.UCCode is null and A.CancelFlag=0 and nvl(A.RealAmt,0) !=0';

  -- �Y���Ȥ�s������
  if p_CustId is not null then
    v_SQL1 := v_SQL1 || ' and A.CustId' || p_CustIdOp || to_char(p_CustId);
  end if;

  -- �Y���O���ر��󦳭�

  if p_CitemSQL is not null and p_Citem2SQL is not null then 
--    v_SQL1 := v_SQL1 || ' and (nvl(A.RealPeriod,0)>0 and A.CitemCode ' || p_CitemSQL || ' or A.CitemCode '|| p_Citem2SQL ||')';
    v_SQL1 := v_SQL1 || ' and (A.CitemCode ' || p_CitemSQL || ' or A.CitemCode '|| p_Citem2SQL ||')';
  end if;

  if p_CitemSQL is not null and p_Citem2SQL is null then 
--    v_SQL1 := v_SQL1 || ' and nvl(A.RealPeriod,0)>0 and A.CitemCode ' || p_CitemSQL;
    v_SQL1 := v_SQL1 || ' and A.CitemCode ' || p_CitemSQL;
  end if;
  if p_CitemSQL is null and p_Citem2SQL is not null then 
--    v_SQL1 := v_SQL1 || ' and (nvl(A.RealPeriod,0)>0 or A.CitemCode ' || p_Citem2SQL || ')';
    v_SQL1 := v_SQL1 || ' and  (A.CitemCode ' || p_Citem2SQL || ')';
  end if;
--  if p_CitemSQL is null and p_Citem2SQL is null then 
--    v_SQL1 := v_SQL1 || ' and nvl(A.RealPeriod,0)>0';
--  end if;
/*
  if p_CitemSQL is not null then 
     v_SQL1 := v_SQL1 || ' and A.CitemCode ' || p_CitemSQL;
  end if;
*/
  -- �Y�������������
  if p_ShouldDate1 is not null then
    v_SQL1 := v_SQL1 || ' and A.ShouldDate>=' || GiPackage.QryDTString0(to_date(p_ShouldDate1,'YYYYMMDD')) ||
		 ' and A.ShouldDate<=' || GiPackage.QryDTString0(to_date(p_ShouldDate2,'YYYYMMDD'));
  end if;

  -- �Y�����q�O����
  if p_CompCode is not null then
    v_SQL1 := v_SQL1 || ' and A.CompCode=' || to_char(p_CompCode);
  end if;

  -- �Y���A�����O����
  if p_ServiceType is not null then
    v_SQL1 := v_SQL1 || ' and A.ServiceType=' || c39 || p_ServiceType || c39;
  end if;

  -- �Y���Ȥ����O����
  if p_ClassSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.ClassCode ' || p_ClassSQL;
  end if;

  -- �Y���A�Ȱϱ���
  if p_ServSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.ServCode ' || p_ServSQL;
  end if;

  -- �Y����F�ϱ���
  if p_AreaSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.AreaCode ' || p_AreaSQL;
  end if;

  -- �Y����D����
  if p_StrtSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.StrtCode ' || p_StrtSQL;
  end if;

  -- �Y��������O����
  if p_BillTypeSQL is not null then
    v_SQL1 := v_SQL1 || ' and substr(A.BillNo,7,1) ' || p_BillTypeSQL;
  end if;

  -- �Y�������s������
  if p_CircuitNo is not null then
    v_Join := v_Join || ' A.AddrNo=B.AddrNo';
    v_MoreTable := v_MoreTable || ', SO014 B';
    if p_CircuitNo = '-1' then
      v_SQL1 := v_SQL1 || ' and B.CircuitNo is null';
    else
      v_SQL1 := v_SQL1 || ' and B.CircuitNo=' || c39 || p_CircuitNo || c39;
    end if;
  end if;

  -- �Y���j�ӱ���		2001.05.14
  if p_MduSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.MduId ' || p_MduSQL;
  end if;

  -- ���SQL��_
  if v_Join is not null then
    v_Where1 := ' and' || v_Where1;
  end if;
  v_SQL1 := v_Column1 || ' from SO033 A' || v_MoreTable || ' where ' || v_Join || v_Where1 || v_SQL1 || ' union ' ||
	    v_Column1 || ' from SO034 A' || v_MoreTable || ' where ' || v_Join || v_Where1 || v_SQL1 ;

/*************************************************
  -- Log SQL statement for checking
  delete from TmpSQL;
  insert into TmpSQL (SQLString) values (v_SQL1);
  commit;
  --return 0;
**************************************************/

/*
  -- ���ժ�SQL���O
  v_SQL1 := 'select BillNo, Item, CustId, CitemCode, CitemName, RealDate, RealAmt, RealPeriod,' ||
	    ' RealStartDate, RealStopDate, STCode, ClctEn, AddrNo, ServCode from SO033' ||
	    ' where RealDate >=' || GiPackage.QryDTString0(to_date(p_RealDate1,'YYYYMMDD')) || 
 	    ' and RealDate <=' || GiPackage.QryDTString0(to_date(p_RealDate2,'YYYYMMDD')) ||
	    ' and RealStartDate is not null and UCCode is null and CancelFlag = 0 and nvl(RealAmt,0) != 0';
	    --' and CustId = 8723';
*/

  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  p_RetMsg := '';
  v_RefMonth := to_number(substr(p_RefDate,1,6));

  DBMS_OUTPUT.PUT_LINE('5. �M���Ȧs�ɤ��e');
  delete from SO089;			-- �M���Ȧs�ɤ��e
  delete from SO090;
  commit;

  begin
    open v_DyCursor for v_SQL1;
  exception
    when others then
      rollback;
      p_RetMsg := 'SQL���~: ' || v_SQL1;
      return -4;
  end;

  DBMS_OUTPUT.PUT_LINE('6. loop�}�l�ӳ��ʧ@');
  v_BadCnt := 0;
  loop       -- ���쪺�C�@�����
    fetch v_DyCursor INTO x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealAmt, x_RealPeriod, x_RealStartDate, x_RealStopDate, x_STCode, 
		x_ClctEn, x_AddrNo, x_ServCode, x_ServiceType, s_CitemCode, s_BillNo, s_Item, x_ShouldDate;
    exit when v_DyCursor%NOTFOUND;

    v_PastAmt := 0;
    v_01Amt := 0;
    v_02Amt := 0;
    v_03Amt := 0;
    v_04Amt := 0;
    v_05Amt := 0;
    v_06Amt := 0;
    v_07Amt := 0;
    v_08Amt := 0;
    v_09Amt := 0;
    v_10Amt := 0;
    v_11Amt := 0;
    v_12Amt := 0;
    v_ElseAmt := 0;

    v_TotalDay := 0;
    v_PreDay := 0;
    v_Flag := 0;
    v_CurrAmt := 0;
    v_LeftAmt := 0;
    v_LeftDay := 0;
    v_CurrDay := 0;
    v_Mndx := 0;
    y_RealAmt := 0;

    -- �ˬd�����h�O�������N�X�O�_���b�d�߱���, �Y�_, �h���p�⦹���h�O��� 2003.04.29
    if nvl(s_CitemCode,0)!=0 then	-- �h�O����
      v_Index := 0;
      if v_CitemCodeCnt>0 then
	for I in 1..v_CitemCodeCnt loop
	  if CitemCode(I)=s_CitemCode then	-- �d��
	    v_Index := I;
	  end if;
	end loop;
	if v_Index = 0 then			-- �d����, ���p�⦹���h�O���
	  goto GO_NEXT1;
	end if;
      end if;
    end if;

    /* 
	�ݦA�L�o�b�b�P���X�޿誺���:
	. �����D�����(��JSO090��): ���Ĵ������@���ŭ�, ���Ĵ����W�L�~��
	. �b�b���: by pass
    */
    -- 2002.05.10 by Lawrence
    select PeriodFlag into y_PeriodFlag from CD019 where CodeNo=x_CitemCode and ServiceType=x_ServiceType;

    --if nvl(x_STCode, -1) = v_BadCode then
    if instr(v_BadCodeString, '('||ltrim(to_char(nvl(x_STCode, 0),'999'))||')') > 0 then
      v_BadCnt := v_BadCnt + 1;			-- by pass �b�b���
      goto GO_NEXT1;
    elsif x_RealPeriod>0 and (x_RealStopDate is null or x_RealStartDate is null) 
	or x_RealPeriod>0 and x_RealStopDate<x_RealStartDate then	-- ���X�k�����Ĵ���
      insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '���X�k�����Ĵ���');
      goto GO_NEXT1;
    elsif x_RealPeriod>0 and nvl(p_ErrYear,0)>0 and months_between(last_day(x_RealStopDate),last_day(x_RealStartDate))>p_ErrYear*12 then
      insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '���Ĵ����W�L�~��'); 
      goto GO_NEXT1;
    elsif nvl(x_RealPeriod,0)=0 and nvl(y_PeriodFlag,0)=1 then
      insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '���X�k���g���ʶ���'); 
      goto GO_NEXT1;
    end if;

    -- ******************************************************************************************
    if nvl(x_RealPeriod,0)=0 and s_CitemCode is null then	-- �X�k���, �D�g���ʦ��O, �B�D�h�O
      if p_TaxMode = 1 then			-- �Ӥ�|�����u�e, �ҳѤ��`���B, �Ҽ{�w�|/���|���v�T
        y_RealAmt := round(x_RealAmt*100/(100+v_CurrTaxRate));
      else
	y_RealAmt := x_RealAmt;
      end if;
      --v_CurrMonth := TO_NUMBER(to_char(x_RealDate, 'YYYYMM'));
      v_ClctMonth := to_number(to_char(x_RealDate, 'YYYYMM'));	-- ���O�骺���
      v_Cnt := v_Cnt + 1;
      v_TotalDay := 0;
      v_PreDay := 0;
      v_CurrAmt := y_RealAmt;

      -- �Y�Ӥ���<���O�骺���, �h�Ӥ���u���B���k�ݨ즬�O�骺����h, 
      -- �M��~�A�ݦ��O�������Ӥ���u���B, ���k�ݦb�ĴX�Ӥ��
      if v_ClctMonth < v_RefMonth then 			-- �L�h�����: �֭p��"�L�h��{���J"
	v_Mndx := 0;
      elsif v_ClctMonth = v_RefMonth then		-- �ѦҤ�Ӥ�: �֭p��"�����{���J"
	v_Mndx := 1;
      else						-- ���Ӫ���� (�����|�o��)
	-- �M�w�ӵ����u���B, �����b�ĴX�Ӥ��
	v_Mndx := months_between(to_date(to_char(v_ClctMonth)||'01', 'YYYYMMDD'),
		add_months(last_day(to_date(p_RefDate,'YYYYMMDD'))+1, -1)) + 1;
      end if;

  	  if v_Mndx = 0 then
	    v_PastAmt := v_CurrAmt;
	  elsif v_Mndx = 1 then
	    v_01Amt := v_CurrAmt;
	  elsif v_Mndx = 2 then
	    v_02Amt := v_CurrAmt;
	  elsif v_Mndx = 3 then
	    v_03Amt := v_CurrAmt;
	  elsif v_Mndx = 4 then
	    v_04Amt := v_CurrAmt;
	  elsif v_Mndx = 5 then
	    v_05Amt := v_CurrAmt;
	  elsif v_Mndx = 6 then
	    v_06Amt := v_CurrAmt;
	  elsif v_Mndx = 7 then
	    v_07Amt := v_CurrAmt;
	  elsif v_Mndx = 8 then
	    v_08Amt := v_CurrAmt;
	  elsif v_Mndx = 9 then
	    v_09Amt := v_CurrAmt;
	  elsif v_Mndx = 10 then
	    v_10Amt := v_CurrAmt;
	  elsif v_Mndx = 11 then
	    v_11Amt := v_CurrAmt;
	  elsif v_Mndx = 12 then
	    v_12Amt := v_CurrAmt;
	  else
	    v_ElseAmt := v_CurrAmt;
	  end if;

      -- �s�W�@�����ӦܼȦs��: SO089
      begin
        insert into SO089 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
	  RealStartDate, RealStopDate, RealPeriod, RealAmt, TotalDay, PreDay, 
	  MonthPast, Month01, Month02, Month03, Month04, Month05, Month06, Month07, 
	  Month08, Month09, Month10, Month11, Month12, MonthElse,
	  ClctEn, AddrNo, ServCode) values 
	  (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate,
	  null, null, 0, y_RealAmt, v_TotalDay, v_PreDay,
	  v_PastAmt, v_01Amt, v_02Amt, v_03Amt, v_04Amt, v_05Amt, v_06Amt, v_07Amt,  
	  v_08Amt, v_09Amt, v_10Amt, v_11Amt, v_12Amt, v_ElseAmt,
	  x_ClctEn, x_AddrNo, x_ServCode);
      exception
	when others then
	  v_TmpErrCnt := v_TmpErrCnt + 1;
	  insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, '�L�k�s�W��SO089'); 
	  goto GO_NEXT1;
      end;

    -- ******************************************************************************************
    elsif nvl(x_RealPeriod,0)!=0 and s_CitemCode is null then	-- �X�k���, �g���ʦ��O
      v_Cnt := v_Cnt + 1;
--      v_TotalDay := x_RealStopDate - x_RealStartDate + v_Para3;	-- ���Ĵ����`���
--90.07.13 Stanley �ק�
--90.10.04 Lawrence �ק�
      v_TotalDay := SF_Days360(x_RealStartDate,x_RealStopDate + v_Para3,v_Para17,p_RetMsg);   -- ���Ĵ����`���
      v_LastDate := x_RealStopDate;				-- ���Ĵ����I���
      v_ClctMonth := to_number(to_char(x_RealDate, 'YYYYMM'));	-- ���O�骺���

      if v_TotalDay = 0 then			-- 2001.10.27 added by Pierre
	v_TotalDay := 1;
      end if;

      -- �M�w�u��C��һݤ��@���ܼ�
      v_BeginDate := x_RealStartDate;		-- �Ӥ뭺��
      v_LeftDay := v_TotalDay;			-- �Ӥ�|�����u�e, �ҳѤ��`���

      if p_TaxMode = 1 then			-- �Ӥ�|�����u�e, �ҳѤ��`���B, �Ҽ{�w�|/���|���v�T
        v_LeftAmt := round(x_RealAmt*100/(100+v_CurrTaxRate));
      else
	v_LeftAmt := x_RealAmt;
      end if;
      y_RealAmt := v_LeftAmt;

      while v_Flag = 0 loop			-- Loop�N���Ĵ�����C�@���
        v_CurrMonth := to_number(to_char(v_BeginDate,'YYYYMM')); 	-- �Ӥ���
        if x_RealStopDate >= last_day(v_BeginDate) then		-- �Ӥ뭺��ܷ�멳�����
          --v_CurrDay := last_day(v_BeginDate)-v_BeginDate+1;
-- 90.07.13 Stanley �ק�
--90.10.04 Lawrence �ק�
          v_CurrDay := SF_Days360(v_BeginDate,last_day(v_BeginDate)+1,v_Para17,p_RetMsg);

-- DBMS_OUTPUT.PUT_LINE(TO_CHAR(v_BeginDate,'YYYYMMDD') || ', ' || TO_CHAR(last_day(v_BeginDate),'YYYYMMDD') || ', ' || v_CurrDay);
	else
	  --v_CurrDay := to_char(x_RealStopDate,'DD') - v_OldP3;
-- 90.10.04 Lawrence �ק�
          v_CurrDay := SF_Days360(to_date(to_char(x_RealStopDate,'YYYYMM')||'01','YYYYMMDD'),x_RealStopDate + v_Para3,v_Para17,p_RetMsg);

-- DBMS_OUTPUT.PUT_LINE(to_char(x_RealStopDate,'YYYYMM')||'01' || ', ' || TO_CHAR(x_RealStopDate,'YYYYMMDD') || ', ' || v_CurrDay);
	end if;
	if v_CurrDay > v_LeftDay then	-- ��: �Y��Ƥj��ҳѤ��`���, �h���ץ����Ѿl���, �_�h�L�k�B�z1�骺���D
	  v_CurrDay := v_LeftDay;
	end if;

/*	-- test code: 2001.10.27
if v_LeftDay=0 or v_CurrDay=0 then
	DBMS_OUTPUT.PUT_LINE('�Ƚs=' || x_CustId || ', �渹=' || x_BillNo || ', ����=' || x_CitemName);
	DBMS_OUTPUT.PUT_LINE('v_LeftAmt ='||v_LeftAmt||', v_LeftDay = '||v_LeftDay||', v_CurrDay = ' ||v_CurrDay);
end if;
*/
        v_CurrAmt := round(v_LeftAmt / v_LeftDay * v_CurrDay); 	 	-- �Ӥ���u���B
	
	-- �w����ƪ��w�q: �Ӥ����j��ѦҤ�Ӥ�, �h�ݤ�
        if v_CurrMonth > v_RefMonth then
	  v_PreDay := v_PreDay + v_CurrDay;
	end if;

	-- �Y�Ӥ���<���O�骺���, �h�Ӥ���u���B���k�ݨ즬�O�骺����h, 
	-- �M��~�A�ݦ��O�������Ӥ���u���B, ���k�ݦb�ĴX�Ӥ��
	if v_CurrMonth < v_ClctMonth then
	  v_CurrMonth := v_ClctMonth;
	end if;

          if v_CurrMonth < v_RefMonth then 		-- �L�h�����: �֭p��"�L�h��{���J"
	    v_Mndx := 0;
	  elsif v_CurrMonth = v_RefMonth then		-- �ѦҤ�Ӥ�: �֭p��"�����{���J"
	    v_Mndx := 1;
	  else						-- ���Ӫ����
	    -- �M�w�ӵ����u���B, �����b�ĴX�Ӥ��
	    v_Mndx := months_between(to_date(to_char(v_CurrMonth)||'01', 'YYYYMMDD'),
	  	add_months(last_day(to_date(p_RefDate,'YYYYMMDD'))+1, -1)) + 1;
	  end if;

  	  if v_Mndx = 0 then
	    v_PastAmt := v_PastAmt + v_CurrAmt;
	  elsif v_Mndx = 1 then
	    v_01Amt := v_01Amt + v_CurrAmt;
	  elsif v_Mndx = 2 then
	    v_02Amt := v_02Amt + v_CurrAmt;
	  elsif v_Mndx = 3 then
	    v_03Amt := v_03Amt + v_CurrAmt;
	  elsif v_Mndx = 4 then
	    v_04Amt := v_04Amt + v_CurrAmt;
	  elsif v_Mndx = 5 then
	    v_05Amt := v_05Amt + v_CurrAmt;
	  elsif v_Mndx = 6 then
	    v_06Amt := v_06Amt + v_CurrAmt;
	  elsif v_Mndx = 7 then
	    v_07Amt := v_07Amt + v_CurrAmt;
	  elsif v_Mndx = 8 then
	    v_08Amt := v_08Amt + v_CurrAmt;
	  elsif v_Mndx = 9 then
	    v_09Amt := v_09Amt + v_CurrAmt;
	  elsif v_Mndx = 10 then
	    v_10Amt := v_10Amt + v_CurrAmt;
	  elsif v_Mndx = 11 then
	    v_11Amt := v_11Amt + v_CurrAmt;
	  elsif v_Mndx = 12 then
	    v_12Amt := v_12Amt + v_CurrAmt;
	  else
	    v_ElseAmt := v_ElseAmt + v_CurrAmt;
	  end if;
/*
	DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_CurrMonth || ', ' || v_CurrDay || ', ' || v_CurrAmt);
        DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_PastAmt || ', ' || v_LeftDay || ', ' || v_LeftAmt);
*/
	-- ���s�M�w�u��C��һݤ��@���ܼ�
        v_BeginDate := last_day(v_BeginDate) + 1;	-- ���뭺��
        v_LeftDay := v_LeftDay - v_CurrDay;		-- ����Ѿl�`���
        v_LeftAmt := v_LeftAmt - v_CurrAmt;		-- ����Ѿl�`�B 

	-- �Y�w�W�X���ĺI���,�ΩҳѤ�Ƭ�0, �h��������X�j��
        if v_BeginDate > v_LastDate or v_LeftDay = 0 then	-- or v_LeftAmt = 0
          v_Flag := 1;
        end if;

        v_Count:=v_Count+1;
        if v_Count=2000 then
            Commit;
            v_Count:=0;
        end if;

      end loop;

/*
      DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_PastAmt || ', ' || v_01Amt || ', ' || v_02Amt || ', ' || 
	v_03Amt || ', ' || v_04Amt || ', ' || v_05Amt || ', ' || v_06Amt || ', ' || v_07Amt || ', ' || 
	v_08Amt || ', ' || v_09Amt || ', ' || v_10Amt || ', ' || v_11Amt || ', ' || v_12Amt || ', ' || v_ElseAmt);
*/
      -- �s�W�@�����ӦܼȦs��: SO089
      -- ��: �ꦬ���B���� y_RealAmt, ���O x_RealAmt
      begin
        insert into SO089 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
	  RealStartDate, RealStopDate, RealPeriod, RealAmt, TotalDay, PreDay, 
	  MonthPast, Month01, Month02, Month03, Month04, Month05, Month06, Month07, 
	  Month08, Month09, Month10, Month11, Month12, MonthElse,
	  ClctEn, AddrNo, ServCode) values 
	  (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate,
	  x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, v_TotalDay, v_PreDay,
	  v_PastAmt, v_01Amt, v_02Amt, v_03Amt, v_04Amt, v_05Amt, v_06Amt, v_07Amt,  
	  v_08Amt, v_09Amt, v_10Amt, v_11Amt, v_12Amt, v_ElseAmt,
	  x_ClctEn, x_AddrNo, x_ServCode);
      exception
	when others then
	  v_TmpErrCnt := v_TmpErrCnt + 1;
	  insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, '�L�k�s�W��SO089'); 
	  goto GO_NEXT1;
      end;

    -- ******************************************************************************************
    else		-- �h�O����	2003.04.29
      -- ���ӵ��h�O��ƪ����Ĵ���, �Y��, �h�H���d��h���u�ܨC��, 
      -- �Y�L, �h�̰h�O��ƪ�������, �Ψ������������ƪ����ĺI�������Ĵ����Ӥ��u�ܨC��
      if x_RealStartDate is null or x_RealStopDate is null then
	-- x_RealStartDate := x_ShouldDate;
	begin
	  select RealStartDate, RealStopDate, CitemName into x_RealStartDate, x_RealStopDate, x_CitemName from v_Charge 
		where BillNo=s_BillNo and Item=s_Item and nvl(CancelFlag,0)=0;
	exception
	  when others then
	    insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '���h�O��ƵL���Ī��������'); 
	    goto GO_NEXT1;
	end;
	x_RealStartDate := greatest(x_ShouldDate, x_RealStartDate);	-- 2003.05.14
      else 	-- �w���h�O��ƪ����Ĵ���, �G�u�A���䥿�����O��Ƥ����ئW��
	begin
	  select Description into x_CitemName from CD019 where CodeNo=s_CitemCode;
	exception
	  when others then
	    insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '�L������ƪ����O���ئW��, �N�X='||s_CitemCode); 
	    goto GO_NEXT1;
	end;

	-- �A���������_�l��, �H�̤j���_�l�鬰�D 2003.05.14
	begin
	  select RealStartDate into x_TmpStartDate from v_Charge where BillNo=s_BillNo and Item=s_Item and nvl(CancelFlag,0)=0;
	exception
	  when others then
	    insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, x_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, x_RealAmt, '���h�O��ƵL���Ī��������'); 
	    goto GO_NEXT1;
	end;
	x_RealStartDate := greatest(x_TmpStartDate, x_RealStartDate);
      end if;

      /**********************************************************************************************
	�ܦ�, �^�H�w�����h�O��Ƥ�����������ƪ�CitemCode, CitemName, �H�ΰh�O�ҧt�\�����Ĵ��� !!!
	���u�h�O���B�ܨC�� 
      ***********************************************************************************************/
      v_Cnt := v_Cnt + 1;
      v_TotalDay := SF_Days360(x_RealStartDate,x_RealStopDate + v_Para3,v_Para17,p_RetMsg);   -- ���Ĵ����`���
      v_LastDate := x_RealStopDate;				-- ���Ĵ����I���
      v_ClctMonth := to_number(to_char(x_RealDate, 'YYYYMM'));	-- ���O�骺���

      if v_TotalDay = 0 then			-- 2001.10.27 added by Pierre
	v_TotalDay := 1;
      end if;

      -- �M�w�u��C��һݤ��@���ܼ�
      v_BeginDate := x_RealStartDate;		-- �Ӥ뭺��
      v_LeftDay := v_TotalDay;			-- �Ӥ�|�����u�e, �ҳѤ��`���
      if p_TaxMode = 1 then			-- �Ӥ�|�����u�e, �ҳѤ��`���B, �Ҽ{�w�|/���|���v�T
        v_LeftAmt := round(x_RealAmt*100/(100+v_CurrTaxRate));
      else
	v_LeftAmt := x_RealAmt;
      end if;
      y_RealAmt := v_LeftAmt;

      while v_Flag = 0 loop			-- Loop�N���Ĵ�����C�@���
        v_CurrMonth := to_number(to_char(v_BeginDate,'YYYYMM')); 	-- �Ӥ���
        if x_RealStopDate >= last_day(v_BeginDate) then		-- �Ӥ뭺��ܷ�멳�����
          v_CurrDay := SF_Days360(v_BeginDate,last_day(v_BeginDate)+1,v_Para17,p_RetMsg);
	else
          v_CurrDay := SF_Days360(to_date(to_char(x_RealStopDate,'YYYYMM')||'01','YYYYMMDD'),x_RealStopDate + v_Para3,v_Para17,p_RetMsg);
	end if;
	if v_CurrDay > v_LeftDay then	-- ��: �Y��Ƥj��ҳѤ��`���, �h���ץ����Ѿl���, �_�h�L�k�B�z1�骺���D
	  v_CurrDay := v_LeftDay;
	end if;

/*	-- test code: 2001.10.27
if v_LeftDay=0 or v_CurrDay=0 then
	DBMS_OUTPUT.PUT_LINE('�Ƚs=' || x_CustId || ', �渹=' || x_BillNo || ', ����=' || x_CitemName);
	DBMS_OUTPUT.PUT_LINE('v_LeftAmt ='||v_LeftAmt||', v_LeftDay = '||v_LeftDay||', v_CurrDay = ' ||v_CurrDay);
end if;
*/
        v_CurrAmt := round(v_LeftAmt / v_LeftDay * v_CurrDay); 	 	-- �Ӥ���u���B
	
	-- �w����ƪ��w�q: �Ӥ����j��ѦҤ�Ӥ�, �h�ݤ�
        if v_CurrMonth > v_RefMonth then
	  v_PreDay := v_PreDay + v_CurrDay;
	end if;

	-- �Y�Ӥ���<���O�骺���, �h�Ӥ���u���B���k�ݨ즬�O�骺����h, 
	-- �M��~�A�ݦ��O�������Ӥ���u���B, ���k�ݦb�ĴX�Ӥ��
	if v_CurrMonth < v_ClctMonth then
	  v_CurrMonth := v_ClctMonth;
	end if;

          if v_CurrMonth < v_RefMonth then 		-- �L�h�����: �֭p��"�L�h��{���J"
	    v_Mndx := 0;
	  elsif v_CurrMonth = v_RefMonth then		-- �ѦҤ�Ӥ�: �֭p��"�����{���J"
	    v_Mndx := 1;
	  else						-- ���Ӫ����
	    -- �M�w�ӵ����u���B, �����b�ĴX�Ӥ��
	    v_Mndx := months_between(to_date(to_char(v_CurrMonth)||'01', 'YYYYMMDD'),
	  	add_months(last_day(to_date(p_RefDate,'YYYYMMDD'))+1, -1)) + 1;
	  end if;

  	  if v_Mndx = 0 then
	    v_PastAmt := v_PastAmt + v_CurrAmt;
	  elsif v_Mndx = 1 then
	    v_01Amt := v_01Amt + v_CurrAmt;
	  elsif v_Mndx = 2 then
	    v_02Amt := v_02Amt + v_CurrAmt;
	  elsif v_Mndx = 3 then
	    v_03Amt := v_03Amt + v_CurrAmt;
	  elsif v_Mndx = 4 then
	    v_04Amt := v_04Amt + v_CurrAmt;
	  elsif v_Mndx = 5 then
	    v_05Amt := v_05Amt + v_CurrAmt;
	  elsif v_Mndx = 6 then
	    v_06Amt := v_06Amt + v_CurrAmt;
	  elsif v_Mndx = 7 then
	    v_07Amt := v_07Amt + v_CurrAmt;
	  elsif v_Mndx = 8 then
	    v_08Amt := v_08Amt + v_CurrAmt;
	  elsif v_Mndx = 9 then
	    v_09Amt := v_09Amt + v_CurrAmt;
	  elsif v_Mndx = 10 then
	    v_10Amt := v_10Amt + v_CurrAmt;
	  elsif v_Mndx = 11 then
	    v_11Amt := v_11Amt + v_CurrAmt;
	  elsif v_Mndx = 12 then
	    v_12Amt := v_12Amt + v_CurrAmt;
	  else
	    v_ElseAmt := v_ElseAmt + v_CurrAmt;
	  end if;
/*
	DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_CurrMonth || ', ' || v_CurrDay || ', ' || v_CurrAmt);
        DBMS_OUTPUT.PUT_LINE(x_CustId || ', ' || v_PastAmt || ', ' || v_LeftDay || ', ' || v_LeftAmt);
*/
	-- ���s�M�w�u��C��һݤ��@���ܼ�
        v_BeginDate := last_day(v_BeginDate) + 1;	-- ���뭺��
        v_LeftDay := v_LeftDay - v_CurrDay;		-- ����Ѿl�`���
        v_LeftAmt := v_LeftAmt - v_CurrAmt;		-- ����Ѿl�`�B 

	-- �Y�w�W�X���ĺI���,�ΩҳѤ�Ƭ�0, �h��������X�j��
        if v_BeginDate > v_LastDate or v_LeftDay = 0 then	-- or v_LeftAmt = 0
          v_Flag := 1;
        end if;

        v_Count:=v_Count+1;
        if v_Count=2000 then
            Commit;
            v_Count:=0;
        end if;
      end loop;

      -- �s�W�@�����ӦܼȦs��: SO089
      -- ��: �ꦬ���B���� y_RealAmt, ���O x_RealAmt
      -- �`�N: "���O���إN�X/�W��"�Υ�����ƪ������N,���ΰh�O��ƥ��������O���إN�X/�W��,�H�K����i�{�P�Ƨ�
      begin
        insert into SO089 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
	  RealStartDate, RealStopDate, RealPeriod, RealAmt, TotalDay, PreDay, 
	  MonthPast, Month01, Month02, Month03, Month04, Month05, Month06, Month07, 
	  Month08, Month09, Month10, Month11, Month12, MonthElse,
	  ClctEn, AddrNo, ServCode) values 
	  (x_BillNo, x_Item, x_CustId, s_CitemCode, x_CitemName, x_RealDate,
	  x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, v_TotalDay, v_PreDay,
	  v_PastAmt, v_01Amt, v_02Amt, v_03Amt, v_04Amt, v_05Amt, v_06Amt, v_07Amt,  
	  v_08Amt, v_09Amt, v_10Amt, v_11Amt, v_12Amt, v_ElseAmt,
	  x_ClctEn, x_AddrNo, x_ServCode);
      exception
	when others then
	  v_TmpErrCnt := v_TmpErrCnt + 1;
	  insert into SO090 (BillNo, Item, CustId, CitemCode, CitemName, RealDate, 
		RealStartDate, RealStopDate, RealPeriod, RealAmt, ErrorType)
		values (x_BillNo, x_Item, x_CustId, s_CitemCode, x_CitemName, x_RealDate, 
		x_RealStartDate, x_RealStopDate, x_RealPeriod, y_RealAmt, '�L�k�s�W��SO089'); 
	  goto GO_NEXT1;
      end;

    end if;

    -- ******************************************************************************************
    <<GO_NEXT1>>
    NULL;
  end loop;
 
  commit;
  p_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||
		'��, ���ͩ��ӵ���='||to_char(v_Cnt);
  return 0;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    p_RetMsg := 'Oracle�T��: '||SQLERRM||', ���~���: �Ƚs=' || x_CustId || ', �渹=' || x_BillNo || ', ����=' || x_CitemName || ', SBillNo='||s_BillNo||', SCitemCode='||s_CitemCode;
    --rollback;
    commit;
    return -99;
end;
/
/*
@c:\emc_script\SF_DailyOp3

set serveroutput on
variable nn number
variable msg varchar2(300)
exec :nn := SF_DailyOp3('testv30', 1, '20020101', '20020131', 'IM', 'Pierre', :msg);

--exec :nn := SF_DailyOp3('EMCYMS,EMCLK,EMCKP,EMCNTP', 1, '20020902', '20020902', 'IM', 'Pierre', :msg);
print nn
print msg

  �妸�@�~__�ȪA�����u���u���O�έp���� (EMC�M��)
  ����: �|�̾ڶǨӤ��UOwner, ���UOwner��ưϪ��u��Ҫ����O��ƨӲέpCSR�Z��
  �ɦW: SF_DailyOp3.sql

  �`�N: 1.�ݫإ�Oracle����: SO503(Table), V_SO503Charge(View), S_SeqNo1(Sequence)
	2.��grant�����v�����D��ư�: cd019, cd039, cm003, so007/8/9, V_SO503Charge, SO041

  IN	p_OwnerList varchar2:	Owner list, ��, 'EMCYMS,EMCLK,EMCKP,EMCNTP'
	-- p_CompCode number:	���q�O�N�X
	-- p_ServiceType char(1):	�A�����O, ��: 'C', 'I'
	p_DateType number;	1=���u��, 2=���z��
	p_RptDate1  varchar2:	�έp�_�l��, ���'YYYYMMDD'
	p_RptDate2  varchar2:	�έp�I���, ���'YYYYMMDD'
	p_BillType varchar2:	������O, ��, 'IMP'
	p_Operator varchar2:	�ާ@�̦W��

  OUT	p_RetMsg VARCHAR2�G ���G�T��, �I�s���ܼƦܤֻ�100 bytes

  return NUMBER: ���G�N�X, ���ӧ��Ʃ�SO503����Ƨ帹
        >0: ��Ƨ帹, ��ܥ��`����
	-1: p_RetMsg='�Ѽƿ��~'
	-2: p_RetMsg='���U�ضO�Ϊ��B����, ���O����='
	-3: p_RetMsg='SQL���~: ' || v_SQL2
	-4: p_RetMsg='�� Sequence S_SeqNo1 �����@�i�έȥ���'
	-5: p_RetMsg='�s�W��ƥ���, �T��='||SQLERRM;
	-6: p_RetMsg='���ҿ��~: �~�̰Ѽ��ɤ��L�A�ȧO����C�����'
	-7: p_RetMsg='���ҿ��~: ���q�O�N�X�ɤ��L�����q�O�����, ���q�O=' || v_CompCode
        -99�Gp_RetMsg='��L���~'

  SO503 (SeqNo		number(8),		-- �帹
	CompCode 	number(3),		-- ���q�O�N�X
	CompName 	varchar2(20),		-- ���q�O�W��
	ServiceType 	char(1),		-- �A�ȧO
	AcceptEn 	varchar2(10),		-- ���z/���ФH���N�X
	AcceptName 	varchar2(20),		-- ���z/���ФH���m�W
	Cnt1		number(8) default 0,	-- '�˴_���O=�P��'�ƶq
	Cnt2		number(8) default 0,	-- '�˴_���O0<X<�P��'�ƶq
	Cnt3		number(8) default 0,	-- '�˴_���O=0'�ƶq
	Cnt4		number(8) default 0,	-- '�����˸m�O'�ƶq
	Cnt5		number(8) default 0,	-- '�����T���O'�ƶq
	CntTotal	number(8) default 0,	-- '�X�p'�ƶq
	RptDate		date,			-- �����Ƥ�
	ReportTime	date,			-- �������ɶ�
	Operator	varchar2(20)		-- �ާ@�̦W��
	MediaCode	number(3)		-- �C�餶�����O�N�X
	MediaName	varchar2(20)		-- �C�餶�����O�W��
	Cnt6		number(8) default 0,	-- 'eBox(��)'�ƶq
	Cnt7		number(8) default 0,	-- 'eBox(��)'�ƶq
	);

  By: Pierre
  Date: 2002.09.10
	2002.09.11 ��: �u���έpI��, �BI�����Ѩ��z�H���אּ�������O���ȪA��
	2002.09.12 ��: �Y���ج��h�x�������T���O, �h�n�[�W�����x��, �Ӥ��O����
	2002.09.13 ��: SO503�[�W�C�餶�����O�N�X���Ϲj, �H��MediaCode=11(TM��)����Ƥ]�[�J�έp
	2002.09.19 ��bug: 'Acceptime' ==> 'AcceptTime'
	2002.12.12 ��: �]����N�X��Ƨ���, MediaCode 1=>10, 11=>18
	2002.12.13 ��: �]����N�X��Ƨ���, CitemCode�w�ﱼ�F

	2003.03.22 ��: �]�����q�O�N�X��(CD039)��U���x�w���O�u���@��, �i��|�h��, �ҥH�����q�O
			�P�W�٪��覡�אּ�hSO041���t�ΧO(�N�O���q�O), �A�hCD039�������q�W��
			�`�N: �ݦAgrant SO041���v�����D��ưϪ��ާ@��
	2003.06.06 ��: �̾�92.06.06���W��Q�װ��ק�, �аѦ�"�Ȥ�ץ���-���u���u���O�έp��W��_920606.doc"���e

*/
--create or replace function SF_DailyOp3(p_OwnerList varchar2, p_CompCode number, p_ServiceType char, 
create or replace function SF_DailyOp3(p_OwnerList varchar2, 
	p_DateType number, p_RptDate1 varchar2, p_RptDate2 varchar2, p_BillType varchar2,
	p_Operator varchar2, p_RetMsg out varchar2) return number
as
  c39		char(1) := chr(39);
  v_RetCode 	number;
  v_RetMsg 	varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;
  v_SQL 	varchar2(500);
  v_SQL2 	varchar2(1000);
  v_SQL3	varchar2(500);
  v_SQLcsr	varchar2(500);
  v_SQLcharge	varchar2(200);
  v_Tmp 	number;
  v_DateFldName varchar2(20);
  v_TableName 	varchar2(20);
  v_AllBillType varchar2(10) := 'IMP';
  v_SeqNo	number;
  v_Now		date;

  v_InstFeeCode 	number := 101;		-- 1.�˾��O���إN�X
  v_ReInstFeeCode	number := 106;		-- 2.�_���O���إN�X
  v_InstFee2Code	number := 108;		-- 3.�����˸m�O���إN�X
  v_BCFeeCode		number := 301;		-- 4.�򥻥x�T���O���إN�X
  v_BCFee2Code1		number := 306;		-- 5.�����T���O���إN�X1(1�x)
  v_BCFee2Code2		number := 111;		-- 5.�����T���O���إN�X2(1�x)
  v_BCFee2Code3		number := 112;		-- 5.�����T���O���إN�X3(2�x)
  v_BCFee2Code4		number := 113;		-- 5.�����T���O���إN�X4(3�x)
  v_BCFee2Code5		number := 114;		-- 5.�����T���O���إN�X5(4�x)
  v_BCFee2Code6		number := 115;		-- 5.�����T���O���إN�X6(5�x)

  v_eBoxSaleFeeCode1	number := 801;		-- 6. eBox(��)�@���I�M���إN�X
  v_eBoxSaleFeeCode2	number := 802;		-- 6. eBox(��)�����I�ڶ��إN�X
  v_eBoxRentFeeCode	number := 804;		-- 7. eBox(��)�O�Ҫ����إN�X

/*
  v_InstFeeCode 	number := 3;		-- 1.�˾��O���إN�X
  v_ReInstFeeCode	number := 4;		-- 2.�_���O���إN�X
  v_InstFee2Code	number := 6;		-- 3.�����˸m�O���إN�X
  v_BCFeeCode		number := 1;		-- 4.�򥻥x�T���O���إN�X
  v_BCFee2Code1		number := 21;		-- 5.�����T���O���إN�X1(1�x)
  v_BCFee2Code2		number := 22;		-- 5.�����T���O���إN�X2(1�x)
  v_BCFee2Code3		number := 23;		-- 5.�����T���O���إN�X3(2�x)
  v_BCFee2Code4		number := 24;		-- 5.�����T���O���إN�X4(3�x)
  v_BCFee2Code5		number := 25;		-- 5.�����T���O���إN�X5(4�x)
  v_BCFee2Code6		number := 66;		-- 5.�����T���O���إN�X6(5�x)
*/
  v_InstFee 		number := 0;		-- 1.�˾��O�P��
  v_ReInstFee		number := 0;		-- 2.�_���O�P��
  v_InstFee2		number := 0;		-- 3.�����˸m�O�P��
  v_BCFee		number := 0;		-- 4.�򥻥x�T���O
  v_BCFee2		number := 0;		-- 5.�����T���O

  v_eBoxSaleFee1	number := 0;		-- 6. eBox(��)�@���I�M�P��
  v_eBoxSaleFee2	number := 0;		-- 6. eBox(��)�����I�ڵP�� ==> �������ઽ�����
  v_eBoxRentFee		number := 0;		-- 7. eBox(��)�O�Ҫ��P��

  v_Cnt1	number;		-- 1.		==> [�˴_���O=�P��]�ƭ�
  v_Cnt2	number;		-- 2.		==> [�˴_���O0<X<�P��]�ƭ�
  v_Cnt3	number;		-- 3.		==> [�˴_���O=0]�ƭ�
  v_Cnt4	number;		-- 4.		==> [�����˸m�O]�ƭ�
  v_Cnt5	number;		-- 5.		==> [�����T���O]�ƭ�
  v_Cnt6	number;		-- 6.		==> [eBox(��)]�ƭ�
  v_Cnt7	number;		-- 7.		==> [eBox(��)]�ƭ�
  v_CntTotal	number;		-- �X�p

  v_FlagA	number;		-- �O�_���Ӷ���
  v_FlagB	number;		-- 
  v_FlagC	number;		-- 
  v_FlagD	number;		-- 
  v_FlagE	number;		-- 
  v_FlagF	number;		-- 
  v_FlagG	number;		-- 

  v_Cnt0	number;		-- record count
  I		number;
  v_Index	number;
  v_OwnerList	varchar2(500);
  v_OwnerCnt	number;
  v_FirstOwner	varchar2(20);	-- �Ĥ@��ư�, �Y�ثe��ư�
  v_CompCode	number;		-- �Ӹ�ưϪ����q�O�N�X
  v_CompName 	varchar2(20);	-- �Ӹ�ưϪ����q�O�W��
  v_ServiceType char(1);

  x_EmpNo	varchar2(10);	-- ��CSR���u�N��
  x_EmpName	varchar2(30);	-- ��CSR���u�m�W
  x_MediaCode	number;		-- �Ӥu��C�餶�����O�N�X
  x_MediaName	varchar2(20);	-- �Ӥu��C�餶�����O�W��
  x_SNo		varchar2(15);	-- �Ȯɳ渹
  x_WoCode	number;		-- ���u���O
  x_CitemCode	number;		-- ���O���إN�X
  x_RealAmt	number;		-- �ꦬ���B

  TYPE BillAry IS TABLE OF char INDEX BY BINARY_INTEGER; --�ŧi�@��Table�����A(�Ϊk�M�}�C�P)
  BillType BillAry;		-- array for Bill types

  TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  OwnerName Varchar2Ary;	-- array for owners
  DataTableName Varchar2Ary;	-- array for Data tables

  type CurTyp is ref cursor;  --�ۭqcursor���A
  v_DyCursor1 CurTyp;          --��dynamic SQL
  v_DyCursor2 CurTyp;          --��dynamic SQL
  v_DyCursor3 CurTyp;          --��dynamic SQL


/*
  cursor CSR is	
    select EmpNo, EmpName from CM003 where WorkClass = 'A';    -- All CSR

  cursor Charge(v_BillNo varchar2) is
    select * from SO034 where BillNo = v_BillNo;
*/
begin

  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_Now := sysdate;

/*
  -- ��������O��JBillType array��
  for i in 1..length(p_BillType) loop
    BillType(i) := substr(p_BillType, i, 1);
  end loop;

  -- �M�w����ɺ���: �ثe��SO007(I��), SO009(P��)  ==> 92.06.06 ����
  DataTableName(1) := 'SO007';
  DataTableName(2) := 'SO009';
  DBMS_OUTPUT.PUT_LINE('['||DataTableName(1)||']');
  DBMS_OUTPUT.PUT_LINE('['||DataTableName(2)||']');
*/

  --����Owner�r���OwnerName�}�C��
  v_OwnerCnt := 0;
  if p_OwnerList is not null then
    v_OwnerList := p_OwnerList;
    v_Index := 1;
    I := 1;
    while v_OwnerList is not null loop
      v_Index := instr(v_OwnerList, ',');
      if v_Index > 0 then
	begin
          OwnerName(I) := ltrim(rtrim(substr(v_OwnerList, 1, v_Index-1)));
	  v_OwnerList := substr(v_OwnerList, v_Index+1);
          I:=I+1;
	exception
	  when others then
	    v_OwnerList := null;
	end;
      else
        OwnerName(I) := rtrim(ltrim(v_OwnerList));
        v_OwnerList := null;
      end if;
    end loop;
    v_OwnerCnt := I;
    v_FirstOwner := OwnerName(1);		-- �O��Ĥ@��owner name
  end if;

/*  for j in 1..v_OwnerCnt loop
    DBMS_OUTPUT.PUT_LINE(OwnerName(j));
  end loop; */


  -- ���U�ض��ؤ��P��
  begin
    v_Tmp := v_InstFeeCode;
    select nvl(Amount,0) into v_InstFee from CD019 where CodeNo = v_InstFeeCode;
    v_Tmp := v_ReInstFeeCode;
    select nvl(Amount,0) into v_ReInstFee from CD019 where CodeNo = v_ReInstFeeCode;
/*
    v_Tmp := v_InstFee2Code;
    select nvl(Amount,0) into v_InstFee2 from CD019 where CodeNo = v_InstFee2Code;
    v_Tmp := v_BCFeeCode;
    select nvl(Amount,0) into v_BCFee from CD019 where CodeNo = v_BCFeeCode;
    v_Tmp := v_BCFee2Code;
    select nvl(Amount,0) into v_BCFee2 from CD019 where CodeNo = v_BCFee2Code;
*/
    v_Tmp := v_eBoxSaleFeeCode1;
    select nvl(Amount,0) into v_eBoxSaleFee1 from CD019 where CodeNo = v_eBoxSaleFeeCode1;
    v_Tmp := v_eBoxRentFeeCode;
    select nvl(Amount,0) into v_eBoxRentFee from CD019 where CodeNo = v_eBoxRentFeeCode;
  exception
    when others then
      p_RetMsg := '���U�ضO�Ϊ��B����, ���O����=' || v_Tmp;
      return -2;
  end;


  -- �Y���έp��d��
  if p_DateType = 1 then
    v_DateFldName := 'FinTime';
  elsif p_DateType = 2 then
    v_DateFldName := 'AcceptTime';
  else
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;
  if p_RptDate1 is not null then
    v_SQL := v_SQL || ' and '||v_DateFldName||'>=' || GiPackage.QryDTString0(to_date(p_RptDate1,'YYYYMMDD')) ||
		' and '||v_DateFldName||'<' || GiPackage.QryDTString0(to_date(p_RptDate2,'YYYYMMDD')) || '+1';
  end if;

  v_SQL := substr(v_SQL, 5);		-- ���h�}�Y�� ' and'
  v_Cnt0 := 0;

  --�� Sequence S_SeqNo1 �����@�i�έ�
  begin
    Select S_SeqNo1.Nextval Into v_SeqNo From Dual;
  exception
    when others then
      p_RetMsg := '�� Sequence S_SeqNo1 �����@�i�έȥ���';
      return -4;
  end;
  delete from SO503 where SeqNo = v_SeqNo;	-- �M���P�@�帹�����

  for i in 1..v_OwnerCnt loop		-- Loop �C�@Owner
    v_ServiceType := 'C';
    v_CompCode := null;

    -- ���Ӹ�ưϤ��w�]���q�O
    begin
      -- v_SQL3 := 'select CodeNo, Description from '||OwnerName(i)||'.CD039 where RowNum=1';
      v_SQL3 := 'select SysId from '||OwnerName(i)||'.SO041 where ServiceType like '||c39||'%C%'||c39||' and RowNum=1';
      open v_DyCursor1 for v_SQL3;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL���~: ' || v_SQL3;
        return -3;
    end;
    fetch v_DyCursor1 INTO v_CompCode; 
    close v_DyCursor1;

    if v_CompCode is null then
      rollback;
      p_RetMsg := '���ҿ��~: �~�̰Ѽ��ɤ��L�A�ȧO����C�����';
      return -6;
    end if;

    --�ھڤ��q�O, �hCD039�������q�W��
    begin
      v_SQL3 := 'select Description from '||OwnerName(i)||'.CD039 where CodeNo='||v_CompCode;
      open v_DyCursor1 for v_SQL3;
    exception
      when others then
        rollback;
        p_RetMsg := '���ҿ��~: ���q�O�N�X�ɤ��L�����q�O�����, ���q�O=' || v_CompCode;
        return -7;
    end;
    fetch v_DyCursor1 INTO v_CompName; 
    close v_DyCursor1;
    DBMS_OUTPUT.PUT_LINE(v_CompCode||','||v_CompName);

/*
    -- ��Owner��ưϪ�CM003.WorkClass = 'C'��, ��CSR
    begin
      v_SQLcsr := 'select EmpNo, EmpName from '||OwnerName(i)||'.CM003 where WorkClass='||c39||'A'||c39;    -- All CSR
      open v_DyCursor1 for v_SQLcsr;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL���~: ' || v_SQLcsr;
	DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
        return -3;
    end;
*/

    -- �H�u�椧���дC�����ȪA����TM�շ���z�H���Ӳέp
    -- 2002.12.12 ��: MediaCode 1=>10, 11=>18
    begin
      v_SQLcsr := 'select IntroId EmpNo, IntroName EmpName, MediaCode, MediaName from '||OwnerName(i)||'.SO007 where '||
		v_SQL||' and FinTime is not null and (MediaCode=10 or MediaCode=18) and IntroId is not null'||
		' group by MediaCode, MediaName, IntroId, IntroName';    -- All CSR and TM group
      open v_DyCursor1 for v_SQLcsr;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL���~: ' || v_SQLcsr;
	DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
        return -3;
    end;

    loop 				-- Loop �C�@CSR
      fetch v_DyCursor1 INTO x_EmpNo, x_EmpName, x_MediaCode, x_MediaName; 
      exit when v_DyCursor1%NOTFOUND;

      -- initialize variables
      v_Cnt1 := 0;
      v_Cnt2 := 0;
      v_Cnt3 := 0;
      v_Cnt4 := 0;
      v_Cnt5 := 0;
      v_Cnt6 := 0;
      v_Cnt7 := 0;
      v_CntTotal := 0;

      --DBMS_OUTPUT.PUT_LINE(x_EmpNo||', '||x_EmpName);

	v_SQL2 := 'select SNo, InstCode WoCode from '||OwnerName(i)||'.SO007';
        v_SQL2 := v_SQL2||' where '||v_SQL||' and FinTime is not null and IntroId='||c39||x_EmpNo||c39||
			' and MediaCode='||x_MediaCode;
        begin
	  open v_DyCursor2 for v_SQL2;
        exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL���~: ' || v_SQL2;
	    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	    return -3;
        end;
        loop       	-- loop �C�@���˾�����
	  v_FlagA := 0;
	  v_FlagB := 0;
	  v_FlagC := 0;
	  v_FlagD := 0;
	  v_FlagE := 0;
	  v_FlagF := 0;
	  v_FlagG := 0;

	  fetch v_DyCursor2 INTO x_SNo, x_WoCode; 
	  exit when v_DyCursor2%NOTFOUND;
	  v_SQLcharge := 'select CitemCode, RealAmt from '||OwnerName(i)||'.V_SO503Charge where BillNo='||c39||x_SNo||c39;
	  begin
	    open v_DyCursor3 for v_SQLcharge;
	  exception
	    when others then
	      rollback;
	      p_RetMsg := 'SQL���~: ' || v_SQLcharge;
	      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	      return -3;
	  end;
	  loop		-- loop �Ө��z�H�����ӱi�u�椧�������O, ���H�U�B�z
	    fetch v_DyCursor3 INTO x_CitemCode, x_RealAmt; 
	    exit when v_DyCursor3%NOTFOUND;
	    v_Cnt0 := v_Cnt0 + 1;

	    -- 1. �Y���ج��˾��O�A�B���B=[�˾��O�P��]�A�ζ��ج��_���O�A�B���B=[�_���O�P��]�A�h[�˴_���O=�P��]�ƭȥ[1
	    if x_CitemCode=v_InstFeeCode and x_RealAmt=v_InstFee or 
		x_CitemCode=v_ReInstFeeCode and x_RealAmt=v_ReInstFee then
	      v_FlagA := 1;
	      v_Cnt1 := v_Cnt1 + 1;

	    -- 2. �Y���ج��˾��O�A�B0<���B<[�˾��O�P��]�A�ζ��ج��_���O�A�B0<���B<[�_���O�P��]�A�h[�˴_���O0<X<�P��]�ƭȥ[1
	    elsif x_CitemCode=v_InstFeeCode and x_RealAmt>0 and x_RealAmt<v_InstFee or 
		x_CitemCode=v_ReInstFeeCode and x_RealAmt>0 and x_RealAmt<v_ReInstFee then
	      v_FlagB := 1;
	      v_Cnt2 := v_Cnt2 + 1;

	    -- 3. �Y���ج��˾��O�A�B���B=0�A�ζ��ج��_���O�A�B���B=0�A�h[�˴_���O=0]�ƭȥ[1
	    elsif x_CitemCode=v_InstFeeCode and x_RealAmt=0 or x_CitemCode=v_ReInstFeeCode and x_RealAmt=0 then
	      v_FlagC := 1;
	      v_Cnt3 := v_Cnt3 + 1;

	    -- 4. �Y���ج������˸m�O�A�h[�����˸m�O]�ƭȥ[1
	    elsif x_CitemCode=v_InstFee2Code then
	      v_FlagD := 1;
	      v_Cnt4 := v_Cnt4 + 1;

	    -- 5. ���ج������T���O�A�B�Ӥu��|�������˸m�O����(�N�X108), �h�Ӥ����T���O���p��, �_�h[�����T���O]�ƭȥ[1
	    elsif x_CitemCode in (v_BCFee2Code1,v_BCFee2Code2,v_BCFee2Code3,v_BCFee2Code4,v_BCFee2Code5,v_BCFee2Code6) then
	      v_FlagE := 1; 

	    -- 6. ���ج�eBox(��)�@���I�M�A�B���B=[eBox��P��], ��  ���ج�eBox(��)����, �B���B>0, �h[eBox(��)]�ƭȥ[1
	    elsif x_CitemCode=v_eBoxSaleFeeCode1 and x_RealAmt=v_eBoxSaleFee1 or 
		x_CitemCode=v_eBoxSaleFeeCode1 and x_RealAmt>0 then
	      v_FlagF := 1;
	      v_Cnt6 := v_Cnt6 + 1;

	    -- 7. ���ج�eBox(��)���O�Ҫ� , �B���B>0, �h[eBox(��)]�ƭȥ[1
	    elsif x_CitemCode=v_eBoxRentFeeCode and x_RealAmt>0 then
	      v_FlagG := 1;
	      v_Cnt7 := v_Cnt7 + 1;
	    end if;

	    if v_FlagE=1 and v_FlagD=0 then	-- 5. �����u���Ӥu��L�����˸m�O��, �~�p��
	      if x_CitemCode=v_BCFee2Code1 or x_CitemCode=v_BCFee2Code2 then	-- �[�W�����x��, ���O����, 2002.09.12
		v_Cnt5 := v_Cnt5 + 1;
	      elsif x_CitemCode=v_BCFee2Code3 then
		v_Cnt5 := v_Cnt5 + 2;
	      elsif x_CitemCode=v_BCFee2Code4 then
		v_Cnt5 := v_Cnt5 + 3;
	      elsif x_CitemCode=v_BCFee2Code5 then
		v_Cnt5 := v_Cnt5 + 4;
	      elsif x_CitemCode=v_BCFee2Code6 then
		v_Cnt5 := v_Cnt5 + 5;
	      end if;
	    end if;
	
	  end loop;	-- loop �Ө��z�H�����ӱi�u�椧�������O

	end loop;	-- loop �C�@���˾�����
	close v_DyCursor2;

      -- �X�p
      v_CntTotal := v_Cnt1 + v_Cnt2 + v_Cnt3 + v_Cnt4 + v_Cnt5 + v_Cnt6 + v_Cnt7;

      -- �N���G�[�J���{���έp��ɸ�ưϤ����έp��SO503
      if v_CntTotal>0 then	-- �Y�`�Ƥj��0, �~�ݭn�[�J�έp��
	begin
	  insert into SO503 (SeqNo, CompCode, CompName, ServiceType, AcceptEn, AcceptName, 
		Cnt1, Cnt2, Cnt3, Cnt4, Cnt5, Cnt6, Cnt7, CntTotal, RptDate, ReportTime, Operator, MediaCode, MediaName) 
		values (v_SeqNo, v_CompCode, v_CompName, v_ServiceType, x_EmpNo, x_EmpName, 
		v_Cnt1, v_Cnt2, v_Cnt3, v_Cnt4, v_Cnt5, v_Cnt6, v_Cnt7, v_CntTotal, to_date(p_RptDate1,'YYYYMMDD'), 
		v_Now, p_Operator, x_MediaCode, x_MediaName);
	exception
	  when others then
    	    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	    rollback;
	    p_RetMsg := '�s�W��ƥ���, �T��='||SQLERRM;
	    return -5;
	end;
      end if;

    end loop;		-- Loop �C�@CSR
    close v_DyCursor1;

  end loop;		-- Loop �C�@Owner


  commit;
  p_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'��, v_Cnt0='||v_Cnt0;

  return v_SeqNo;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;

    p_RetMsg := SQLERRM;
    return -99;
end;
/
/*  
@C:\Gird\v400\Script\SF_DailyOp1
@c:\EMC_Script\SF_DailyOp1

variable nn number
variable TName varchar2(20)
variable msg varchar2(100)

exec :nn := SF_DailyOp1(7, 'C', '20021013', '200210141500', 'Pierre', 0, null, :TName, :msg);
print nn
print TName
print msg

  ��B�����: �C��έp�Ȥ�˾�/�����, MTD, YTD		***** V3.0 *****
  �ɦW: SF_DailyOp1.sql

  IN	p_CompCode number:	���q�O
	p_ServiceType char(1):	�A�����O, ��: 'C', 'I'
	p_StopDate  varchar2:	�έp�I���, 'YYYYMMDD'
	p_CutOffTime varchar2:	Cut off�ɶ�, 'YYYYMMDDHHMI'
	p_Operator varchar2:	�ާ@�̦W��
	p_PRFlag number:	������Ȥ�O�_�C�J���Ĥ�, 0=�_, 1=�O
	p_ReturnSQL varchar2:	�h���]SQL����

  OUT   p_TmpTableName varchar2: �έp���G�Ȧs�ɦW (�I�s���ܼƦܤֻ�20 bytes)
	p_RetMsg VARCHAR2�G���G�T�� (�I�s���ܼƦܤֻ�100 bytes)

  Return NUMBER�G���G�N�X
        0: ���`����
	-1: p_RetMsg='�Ѽƿ��~'
	-2: p_RetMsg='�Ȧs���ɦW�إ߿��~, ���ˬd�O�_���إ�Sequence����: S_DSSRPT_TableName';
	-3: p_RetMsg='SQL1���~: ' || v_SQL1
	-4: p_RetMsg='�L�k�R���Ȧs��, �i��L�H���ϥΥ�����'
	-5: p_RetMsg='�L�k�إ߼Ȧs��, �i��L�H���ϥΥ�����, �еy��A��'
	-6: p_RetMsg='�s�ɦ���B���log�ɿ��~, SQLCODE='||SQLCODE||', '||'SQLERRM='||SQLERRM
        -99�Gp_RetMsg='��L���~'

  �Ȧs�ɬ[�c:
	CompCode number(3) 			���q�O
	ServiceType char(1) 			�A�����O
	ItemCode1 number(3)			���إN�X1
	ItemName1 varchar2(40)			���ئW��1
	ItemCode2 number(3)			���إN�X2
	ItemName2 varchar2(40)			���ئW��2
	Value01 number(8)			�ƭ�1
	Value02 number(8)			�ƭ�2
	Value03 number(8)			�ƭ�3
	Value04 number(8)			�ƭ�4
	Value05 number(8)			�ƭ�5
	Value06 number(8)			�ƭ�6
	Value07 number(8)			�ƭ�7
	Value08 number(8)			�ƭ�8

  By: Pierre
  Date: 2001.03.20
	2001.03.26 �վ�: ��SO093����������log��
	2001.03.26 �[�J: �έp"��L"�������˾���
	2001.04.06 �վ�: ��L.�����/�˾���(MTD)���w�q�אּ�q�Y�ܤ���, �����˾�����YTD����
	2001.09.07 �վ�: �`�P�� --> �`�P��
	2001.09.25 �վ�: �W��(1)~(27)�������u�ɶ� --> ���z�ɶ�
	2001.11.05 �վ�: 1. �p�p.�b����/���īȤ�: �����W�U�G��, �Y�W�["���īȤ�(YTD)"�o�@��ƾ�
			 2. ��L.�����, ��L.�˾���: �έp�̾ڬ����z��d�� ==> 2001.09.25�N�令�p��
			 3. ��L.�˾���: �w�q�ѰѦҸ�=1, �אּ�ѦҸ�=1��5
			 4. ��L.�����˾���: �έp�̾ڧאּñ����d�� ==> ���ӴN�p��
			 5. ��L.�����˾���: ����A�[�J�h���]������(3 or 6)
			 6. �W�[��J�Ѽ�: ������Ȥ�O�_�C�J���Ĥ�, �h���]������

	2002.01.09 �[�J: �w�}�ѽX������21���έp���, ���s��[57]~[77] (�аѦҤ������)
	2002.01.16 �ק�: ����[77]��bug
	2002.01.23 �ק�: �����_����bug
	2002.10.14 ��: �令V3.0������, CustStatusCode����SO002��

  By: Lawrence
        2002.01.04 �վ�: 1. �˾�.�_��: �w�q�[�J�ѦҸ�=7 (�����_��)
                         2. �˾�.�X�p: �w�q�[�J�ѦҸ�=7 (�����_��)
                         3. ���.���: �w�q�[�J�ѦҸ�=1 (����)
                         4. ��L.�����: �w�q�[�J�ѦҸ�=1 (����)
                         5. ��L.�˾���: �w�q�[�J�ѦҸ�=7 (�����_��)
                         6. ��L.�����˾���: �w�q�[�J�ѦҸ�=7 (�����_��)
	2002.04.09 ��: ���Ŀ������Ȥ�O�_�C�J���Ĥ�, �p�⦳�īȤ��bug
        2003.11.18 Mark delete�Ϊk,�אּTruncate
	
*/
create or replace function SF_DailyOp1
  (p_CompCode number, p_ServiceType char, p_StopDate varchar2,
   p_CutoffTime varchar2, p_Operator varchar2,
   p_PRFlag number, p_ReturnSQL varchar2, 
   p_TmpTableName out varchar2, p_RetMsg out varchar2) 
  return number as

  -- p_StartDate varchar2

  v_StartExecTime date;
  v_StartDate	varchar2(8);
  c39		char(1) := chr(39);
  v_Cnt1	number;
  v_ErrCnt1	number;
  v_Cnt2	number;
  v_ErrCnt2	number;

  v_ThisYear	char(4);
  --v_ThisMonth	char(6);
  v_SQL1 	varchar2(2000);
  v_SQL2 	varchar2(1000);
  v_SQL3 	varchar2(500);
  v_Where1	varchar2(100);
  v_ReturnCode  varchar2(500);

  z_Type number;
  z_Amt number;

  type CurTyp is ref cursor;
  cDynamic CurTyp;

  -- ***********************
  v_0110_1 number := 0;
  v_0110_2 number := 0;
  v_0110_3 number := 0;
  v_0110_4 number := 0;
  v_0110_5 number := 0;

  v_0120_1 number := 0;
  v_0120_2 number := 0;
  v_0120_3 number := 0;
  v_0120_4 number := 0;
  v_0120_5 number := 0;

  v_0199_1 number := 0;
  v_0199_2 number := 0;
  v_0199_3 number := 0;
  v_0199_4 number := 0;
  v_0199_5 number := 0;


  -- ***********************
  v_0310_1 number := 0;
  v_0310_2 number := 0;
  v_0310_3 number := 0;
  v_0310_5 number := 0;

  v_0320_1 number := 0;
  v_0320_2 number := 0;
  v_0320_3 number := 0;
  v_0320_5 number := 0;

  v_0399_1 number := 0;
  v_0399_2 number := 0;
  v_0399_3 number := 0;
  v_0399_5 number := 0;

  -- ***********************
  v_1010_1 number := 0;
  v_1010_2 number := 0;
  v_1010_3 number := 0;
  v_1010_5 number := 0;

  v_1020_1 number := 0;
  v_1020_2 number := 0;
  v_1020_3 number := 0;
  v_1020_5 number := 0;

  v_1099_1 number := 0;
  v_1099_2 number := 0;
  v_1099_3 number := 0;
  v_1099_5 number := 0;

  -- ***********************
  v_2010_1 number := 0;
  v_2010_2 number := 0;
  v_2010_3 number := 0;
  v_2010_5 number := 0;

  v_3010_1 number := 0;
  v_3010_2 number := 0;
  v_3010_3 number := 0;
  v_3010_5 number := 0;

  v_3020_5 number := 0;				-- 2001.11.05

  -- ***********************
  --v_4010_1 number := 0;
  --v_4010_2 number := 0;
  v_4010_3 number := 0;
  --v_4010_5 number := 0;

  --v_4020_1 number := 0;
  v_4020_2 number := 0;
  v_4020_3 number := 0;
  v_4020_5 number := 0;

  v_4030_1 number := 0;
  v_4030_2 number := 0;
  v_4030_3 number := 0;
  v_4030_5 number := 0;

  -- ***********************	2002.01.09 added
  v_5010_1 number := 0;
  v_5010_2 number := 0;
  v_5010_3 number := 0;
  v_5010_5 number := 0;

  v_5020_1 number := 0;
  v_5020_2 number := 0;
  v_5020_3 number := 0;
  v_5020_5 number := 0;

  v_5030_1 number := 0;
  v_5030_2 number := 0;
  v_5030_3 number := 0;
  v_5030_5 number := 0;

  v_5040_1 number := 0;
  v_5040_2 number := 0;
  v_5040_3 number := 0;
  v_5040_5 number := 0;

  v_5099_1 number := 0;
  v_5099_2 number := 0;
  v_5099_3 number := 0;
  v_5099_5 number := 0;

  v_6010_3 number := 0;

begin
  p_TmpTableName := null;

  -- Check parameters
  if p_StopDate is null or
    p_CompCode is null or p_ServiceType is null or p_CutoffTime is null or
    p_PRFlag<0 or p_PRFlag>1 then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_StartDate := substr(p_StopDate,1,6) || '01';
  -- DELETE FROM SO093;
  -- 2003.11.18 by Lawrence Mark delete�Ϊk,�אּTruncate
  DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO093');  
  -- commit;

/* **********************************************************************************
  -- �զ��Ȧs�ɦW
  begin
    select S_DSSRPT_TableName.NextVal into v_Cnt1 from dual;
  exception
    when others then
      p_RetMsg := '�Ȧs���ɦW�إ߿��~, ���ˬd�O�_���إ�Sequence����: S_DSSRPT_TableName';
      return -2;
  end;
  p_TmpTableName := 'DSSRPT_' || ltrim(to_char(v_Cnt1, '09999999'));
  --DBMS_OUTPUT.PUT_LINE('�Ȧs�ɦW: ' || p_TmpTableName);

  --�ˬd�O�_���إ߼Ȧs��, �Y���h���R��, �A�إߤ�
  select count(*) into v_Cnt1 from User_Tables where Table_Name = p_TmpTableName;
  if v_Cnt1 > 0 then 
    v_SQL3 := 'drop table ' || p_TmpTableName;
    begin
      DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL3);
    exception
      when others then
        p_RetMsg := '�L�k�R���Ȧs��, �i��L�H���ϥΥ�����';
        return -4;
    end;
  end if;
  v_SQL3 := 'CREATE TABLE ' || p_TmpTableName || '(' ||
	'CompCode number(3), ServiceType char(1), ItemCode1 number(3), ItemName1 varchar2(40),
	ItemCode2 number(3), ItemName2 varchar2(40), Value01 number(8) default 0,
	Value02 number(8) default 0, Value03 number(8) default 0, Value04 number(8) default 0,
	Value05 number(8) default 0, Value06 number(8) default 0, Value07 number(8) default 0,
	Value08 number(8) default 0)';
  begin
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL3);
  exception
    when others then
      p_RetMsg := '�L�k�إ߼Ȧs��, �i��L�H���ϥΥ�����, �еy��A��';
      return -5;
  end;
*********************************************************************************** */

  v_ThisYear := substr(p_StopDate,1,4);
  --v_ThisMonth := substr(p_StopDate,1,6);
  v_Where1 := 'A.CompCode='||p_CompCode||' and A.ServiceType='||c39||p_ServiceType||c39;

  -- ***************************************************************************
  -- �έp�I�����: �P��(1) -- �~�ȾP�� [�s�˾�(10), �_��(20), �`��(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.AcceptTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39||' group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- �s�˾�
      v_0110_2 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- ����_�� , �����_�� 2002.01.04
      v_0120_2 := v_0120_2 + z_Amt;
    end if;
  end loop;
  v_0199_2 := v_0110_2 + v_0120_2;	-- �`��
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(1)-�~�ȾP��--�I�����: �s�˾�='||v_0110_2||', �_��='||v_0120_2||', �`��='||v_0199_2);

  -- ***************************************************************************
  -- �έp�I�����: �P��(1) -- �`�P�� [�s�˾�(10), �_��(20), �`��(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.AcceptTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- �s�˾�
      v_0110_3 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- ����_�� , �����_�� 2002.01.04
      v_0120_3 := v_0120_3 + z_Amt;
    end if;
  end loop;
  v_0199_3 := v_0110_3 + v_0120_3;	-- �`��
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(1)-�`�P��--�I�����: �s�˾�='||v_0110_3||', �_��='||v_0120_3||', �`��='||v_0199_3);

  -- ***************************************************************************
  -- �έp�I���MTD: �P��(1) [�s�˾�(10), �_��(20), �`��(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- �s�˾�
      v_0110_4 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- ����_�� , �����_�� 2002.01.04
      v_0120_4 := v_0120_4 + z_Amt;
    end if;
  end loop;
  v_0199_4 := v_0110_4 + v_0120_4;	-- �Ȥ�W�[�X�p
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(1)-�έp�I���MTD: �s�˾�='||v_0110_4||', �_��='||v_0120_4||', �`��='||v_0199_4);


  -- ***************************************************************************
  -- �έp�I���YTD: �P��(1) [�s�˾�(10), �_��(20), �`��(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- �s�˾�
      v_0110_5 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- ����_�� , �����_�� 2002.01.04
      v_0120_5 := v_0120_5 + z_Amt;
    end if;
  end loop;
  v_0199_5 := v_0110_5 + v_0120_5;	-- �Ȥ�W�[�X�p
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(1)-�έp�I���YTD: �s�˾�='||v_0110_5||', �_��='||v_0120_5||', �`��='||v_0199_5);

  -- ***************************************************************************
  -- �έp�I�����: �P��(1) -- CSR [�s�˾�(10), �_��(20), �`��(99)]
  v_0110_1 := v_0110_3 - v_0110_2;
  v_0120_1 := v_0120_3 - v_0120_2;
  v_0199_1 := v_0199_3 - v_0199_2;


  -- ***************************************************************************
  -- �έp�_�l��ܺI���e�@��: �P��(3) [�~�ȾP��(20)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39 ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0320_1 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(3)-�_�l��ܺI���e�@��: �~�ȾP��='||v_0320_1);

  -- ***************************************************************************
  -- �έp�_�l��ܺI���e�@��: �P��(3) [�`�P��(99)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0399_1 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(3)-�_�l��ܺI���e�@��: �`�P��='||v_0399_1);

  -- ***************************************************************************
  -- �έp�I���YTD: �P��(3) [�~�ȾP��(20)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39 ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0320_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(3)-�I���YTD: �~�ȾP��='||v_0320_5);

  -- ***************************************************************************
  -- �έp�I���YTD: �P��(3) [�`�P��(99)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' ;

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_0399_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�P��(3)-�I���YTD: �`�P��='||v_0399_5);

  -- ***************************************************************************
  v_0310_1 := v_0399_1 - v_0320_1;		-- �_�l��ܺI���e�@��: �P��(3) [CSR(10)]

  v_0310_2 := v_0199_1;				-- �έp�I�����: �P��(3) [CSR(10)]
  v_0320_2 := v_0199_2;				-- �έp�I�����: �P��(3) [�~�ȾP��(20)]
  v_0399_2 := v_0199_3;				-- �έp�I�����: �P��(3) [�`�P��(99)]

  v_0310_3 := v_0310_1 + v_0310_2;		-- �_�l��ܺI����MTD: �P��(3) [CSR(10)]
  v_0320_3 := v_0320_1 + v_0320_2;		-- �_�l��ܺI����MTD: �P��(3) [�~�ȾP��(20)]
  v_0399_3 := v_0399_1 + v_0399_2;		-- �_�l��ܺI����MTD: �P��(3) [�`�P��(99)]

  v_0310_5 := v_0399_5 - v_0320_5;		-- �I���YTD: �P��(3) [CSR(10)]


  -- ***************************************************************************
  -- �έp�_�l��ܺI���e�@��: �˾�(10) [�s�˾�(10), �_��(20), �Ȥ�W�[�X�p(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- �s�˾�
      v_1010_1 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- ����_�� , �����_�� 2002.01.04
      v_1020_1 := v_1020_1 + z_Amt;
    end if;
  end loop;
  v_1099_1 := v_1010_1 + v_1020_1;	-- �Ȥ�W�[�X�p
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�_�l��ܺI���e�@��: �s�˾�='||v_1010_1||', �_��='||v_1020_1||', �Ȥ�W�[�X�p='||v_1099_1);

  -- ***************************************************************************
  -- �έp�I�����: �˾�(10) [�s�˾�(10), �_��(20), �Ȥ�W�[�X�p(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.FinTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- �s�˾�
      v_1010_2 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- ����_�� , �����_�� 2002.01.04
      v_1020_2 := v_1020_2 + z_Amt;
    end if;
  end loop;
  v_1099_2 := v_1010_2 + v_1020_2;	-- �Ȥ�W�[�X�p
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�I�����: �s�˾�='||v_1010_2||', �_��='||v_1020_2||', �Ȥ�W�[�X�p='||v_1099_2);

  -- ***************************************************************************
  -- �έp����ܺI�����(MTD): �˾�(10) [�s�˾�(10), �_��(20), �Ȥ�W�[�X�p(99)]
  v_1010_3 := v_1010_1 + v_1010_2;
  v_1020_3 := v_1020_1 + v_1020_2;
  v_1099_3 := v_1099_1 + v_1099_2;

  -- ***************************************************************************
  -- �έpYTD: �˾�(10) [�s�˾�(10), �_��(20), �Ȥ�W�[�X�p(99)]
  v_SQL1 := 'select B.RefNo Type,nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by B.RefNo' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then		-- �s�˾�
      v_1010_5 := z_Amt;
    elsif z_Type=5 or z_Type=7 then		-- ����_�� , �����_�� 2002.01.04
      v_1020_5 := v_1020_5 + z_Amt;
    end if;
  end loop;
  v_1099_5 := v_1010_5 + v_1020_5;	-- �Ȥ�W�[�X�p
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('YTD: �s�˾�='||v_1010_5||', �_��='||v_1020_5||', �Ȥ�W�[�X�p='||v_1099_5);

  -- ***************************************************************************
  -- �έp�_�l��ܺI���e�@��: ���(20) [���(10)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_2010_1 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�_�l��ܺI���e�@��: ���='||v_2010_1);

  -- ***************************************************************************
  -- �έp�I�����: ���(20) [���(10)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1 || ' and trunc(A.FinTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;

  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_2010_2 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�I�����: ���='||v_2010_2);

  -- ***************************************************************************
  -- �έpMTD: ���(20) [���(10)]
  v_2010_3 := v_2010_1 + v_2010_2;

  -- ***************************************************************************
  -- �έpYTD: ���(20) [���(10)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_2010_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('YTD: ���='||v_2010_5);

  -- ***************************************************************************
  -- �έp: �p�p(30) [�b����(10)]
  v_3010_1 := v_1099_1 - v_2010_1;
  v_3010_2 := v_1099_2 - v_2010_2;
  v_3010_3 := v_1099_3 - v_2010_3;
  v_3010_5 := v_1099_5 - v_2010_5;

  -- ***************************************************************************
  -- �έp: �p�p(30) [���īȤ�(20) YTD]				2001.11.05
  /* v_3020_5: 
	������Ȥᦳ�i�ण�C�J���īȤ�, ���ѼƨM�w
  */

  --** v_SQL1 := 'select count(*) Amt from SO001 where ';
  v_SQL1 := 'select count(*) Amt from SO002 where ';
  if p_PRFlag = 0 then			-- ������Ȥᤣ��
    v_SQL1 := v_SQL1 || 'CustStatusCode=1';
  else					-- ������Ȥ�n��
    v_SQL1 := v_SQL1 || '(CustStatusCode=1 or CustStatusCode=6)';
  end if;
  v_SQL1 := v_SQL1 ||' and ServiceType='||c39||p_ServiceType||c39;

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_3020_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�p�p: ���īȤ�(YTD)='||v_3020_5);


  -- ***************************************************************************
/* -- �έp�_�l��ܺI���MTD: ��L(40) [�����(10)] 
  v_SQL1 := 'select count(*) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=2 or B.RefNo=5)'; */
  -- �έp�q�Y�ܺI���: ��L(40) [�����(10)] 
  -- 2001.04.06
  v_SQL1 := 'select count(*) Amt from SO009 A, CD007 B where A.PRCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5)';

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4010_3 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�_�l��ܺI���MTD: �����='||v_4010_3);

  -- ***************************************************************************
  -- �έp�I�����: ��L(40) [�˾���(20)]
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and trunc(A.AcceptTime)=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)'; 
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4020_2 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�I�����: �˾���='||v_4020_2);

  -- ***************************************************************************
/*  -- �έp�_�l��ܺI���MTD: ��L(40) [�˾���(20)] 
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and B.RefNo=1'; */
  -- �έp�q�Y�ܺI���: ��L(40) [�˾���(20)] 
  -- 2001.04.06
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)'; 
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4020_3 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�_�l��ܺI���MTD: �˾���='||v_4020_3);
/*
  -- ***************************************************************************
  -- �έp�_�l��ܺI���YTD: ��L(40) [�˾���(20)] 
  v_SQL1 := 'select count(*) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1||' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=5)';

  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
    v_4020_5 := z_Amt;
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�_�l��ܺI���YTD: �˾���='||v_4020_5);
*/

  -- ***************************************************************************
  v_ReturnCode := ' and A.ReturnCode is not null';	-- not null�O���n����
  if p_ReturnSQL is not null then
    v_ReturnCode := v_ReturnCode || ' and A.ReturnCode ' || p_ReturnSQL;
  end if;

  -- �έp�_�l��ܺI���e�@��: ��L(40) [�����˾���(30)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.SignDate>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.SignDate<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' || v_ReturnCode;
  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_4030_1 := z_Amt;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�_�l��ܺI���e�@��: �����˾���='||v_4030_1);

  -- ***************************************************************************
  -- �έp�I�����: ��L(40) [�����˾���(30)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B ' ||
		'where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.SignDate=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' || v_ReturnCode;
  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_4030_2 := z_Amt;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�I�����: �����˾���='||v_4030_2);

  -- ***************************************************************************
  -- �έp����ܺI�����MTD: ��L(40) [�����˾���(30)]
  v_4030_3 := v_4030_1 + v_4030_2;

  -- ***************************************************************************
  -- �έpYTD: ��L(40) [�����˾���(30)]
  v_SQL1 := 'select nvl(count(*),0) Amt from SO007 A, CD005 B where A.InstCode=B.CodeNo and ' ||
		v_Where1 || ' and A.SignDate>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.SignDate<=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7)' || v_ReturnCode;
  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_4030_5 := z_Amt;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('YTD: �����˾���='||v_4030_5);

  -- ***************************************************************************
  -- �έp: �w�}�ѽX��.�P��/�X��/�ɥX/�Ȥ�۳�.���
  v_SQL1 := 'select A.BuyCode Type,nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1 group by A.BuyCode' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then				-- �P��
      v_5010_2 := z_Amt;
    elsif z_Type=3 then				-- �X��
      v_5020_2 := z_Amt;
    elsif z_Type=4 then				-- �ɥX
      v_5030_2 := z_Amt;
    elsif z_Type=5 then				-- �Ȥ�۳�
      v_5040_2 := z_Amt;
    end if;
  end loop;
  v_5099_2 := v_5010_2 + v_5020_2 + v_5030_2 + v_5040_2;	-- �`��
  close cDynamic;

  -- ***************************************************************************
  -- �έp: �w�}�ѽX��.�P��/�X��/�ɥX/�Ȥ�۳�.���MTD
  v_SQL1 := 'select A.BuyCode Type,nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.InstDate<=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1 group by A.BuyCode' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then				-- �P��
      v_5010_3 := z_Amt;
    elsif z_Type=3 then				-- �X��
      v_5020_3 := z_Amt;
    elsif z_Type=4 then				-- �ɥX
      v_5030_3 := z_Amt;
    elsif z_Type=5 then				-- �Ȥ�۳�
      v_5040_3 := z_Amt;
    end if;
  end loop;
  v_5099_3 := v_5010_3 + v_5020_3 + v_5030_3 + v_5040_3;	-- �`��
  close cDynamic;

  -- ***************************************************************************
  -- �έp: �w�}�ѽX��.�P��/�X��/�ɥX/�Ȥ�۳�.���ܫe�@��
  v_5010_1 := v_5010_3 - v_5010_2;
  v_5020_1 := v_5020_3 - v_5020_2;
  v_5030_1 := v_5030_3 - v_5030_2;
  v_5040_1 := v_5040_3 - v_5040_2;
  v_5099_1 := v_5099_3 - v_5099_2;

  -- ***************************************************************************
  -- �έp: �w�}�ѽX��.�P��/�X��/�ɥX/�Ȥ�۳�.YTD
  v_SQL1 := 'select A.BuyCode Type,nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.InstDate<=to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		') and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1 group by A.BuyCode' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Type, z_Amt;
    exit when cDynamic%NotFound;
    if z_Type=1 then				-- �P��
      v_5010_5 := z_Amt;
    elsif z_Type=3 then				-- �X��
      v_5020_5 := z_Amt;
    elsif z_Type=4 then				-- �ɥX
      v_5030_5 := z_Amt;
    elsif z_Type=5 then				-- �Ȥ�۳�
      v_5040_5 := z_Amt;
    end if;
  end loop;
  v_5099_5 := v_5010_5 + v_5020_5 + v_5030_5 + v_5040_5;	-- �`��
  close cDynamic;

  -- ***************************************************************************
  -- �έp: �w�}�ѽX��.�ݸ˸ѽX��.MTD/YTD (MTD, YTD�ȬҤ@��)
  v_SQL1 := 'select nvl(count(*),0) Amt from SO004 A, CD022 B ' ||
		'where A.FaciCode=B.CodeNo and ' ||
		v_Where1 || ' and A.InstDate is null and A.PrDate is null ' ||
		' and (A.BuyCode=1 or A.BuyCode=3 or A.BuyCode=4 or A.BuyCode=5) and B.RefNo=1' ;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    fetch cDynamic into z_Amt;
    exit when cDynamic%NotFound;
  end loop;
  v_6010_3 := z_Amt;
  close cDynamic;


  -- ***************************************************************************
  -- �s��log��: SO093
  begin
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value04, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 1, '�P��', 10, '�s�˾�', 
	v_0110_1, v_0110_2, v_0110_3, v_0110_4, v_0110_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value04, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 1, '�P��', 20, '�_��', 
	v_0120_1, v_0120_2, v_0120_3, v_0120_4, v_0120_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value04, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 1, '�P��', 99, '�`��', 
	v_0199_1, v_0199_2, v_0199_3, v_0199_4, v_0199_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 3, '�P��', 10, '�ȪA�P��', 
	v_0310_1, v_0310_2, v_0310_3, v_0310_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 3, '�P��', 20, '�~�ȾP��', 
	v_0320_1, v_0320_2, v_0320_3, v_0320_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 3, '�P��', 99, '�`�P��', 
	v_0399_1, v_0399_2, v_0399_3, v_0399_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 10, '�˾�', 10, '�s�˾�', 
	v_1010_1, v_1010_2, v_1010_3, v_1010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 10, '�˾�', 20, '�_��', 
	v_1020_1, v_1020_2, v_1020_3, v_1020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 10, '�˾�', 99, '�Ȥ�W�[�X�p', 
	v_1099_1, v_1099_2, v_1099_3, v_1099_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 20, '���', 10, '���', 
	v_2010_1, v_2010_2, v_2010_3, v_2010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 30, '�p�p', 10, '�b����', 
	v_3010_1, v_3010_2, v_3010_3, v_3010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 30, '�p�p', 20, '���īȤ�', 
	v_3020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value03, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '��L', 10, '�����', 
	v_4010_3, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
/*    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '��L', 20, '�˾���', 
	v_4020_2, v_4020_3, v_4020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator); */
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value02, Value03, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '��L', 20, '�˾���', 
	v_4020_2, v_4020_3, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 40, '��L', 30, '�����˾���', 
	v_4030_1, v_4030_2, v_4030_3, v_4030_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    -- 2002.01.09
    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '�w�}�ѽX��', 10, '�P��', 
	v_5010_1, v_5010_2, v_5010_3, v_5010_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '�w�}�ѽX��', 20, '�X��', 
	v_5020_1, v_5020_2, v_5020_3, v_5020_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '�w�}�ѽX��', 30, '�ɥX', 
	v_5030_1, v_5030_2, v_5030_3, v_5030_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '�w�}�ѽX��', 40, '�Ȥ�۳�', 
	v_5040_1, v_5040_2, v_5040_3, v_5040_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 50, '�w�}�ѽX��', 99, '�`��', 
	v_5099_1, v_5099_2, v_5099_3, v_5099_5, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

    insert into SO093 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, Value02, Value03, Value05, StartDate, StopDate, CutoffTime, ReportTime,
	Operator) values 
	(p_CompCode, p_ServiceType, 60, '�w�}�ѽX��', 10, '�ݸ˸ѽX��', 
	null, null, v_6010_3, v_6010_3, to_date(v_StartDate,'YYYYMMDD'),
	to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator);

  exception
    when others then
      p_RetMsg := '�s�ɦ���B���log�ɿ��~, SQLCODE='||SQLCODE||', '||'SQLERRM='||SQLERRM;
      rollback;
      return -6;
  end;


  commit;
  p_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'��';
  return 0;

exception
  when others then
    close cDynamic;
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    return -99;
end;
/
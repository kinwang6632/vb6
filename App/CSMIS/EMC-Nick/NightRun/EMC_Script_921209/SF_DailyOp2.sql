/*
@c:\gird\v400\script\SF_DailyOp2
variable nn number
variable TName varchar2(20)
variable msg varchar2(200)

exec :nn := SF_DailyOp2(3, 'C', '20020320', '200203221500', 'Pierre', 0, null, :TName, :msg);
print nn
print TName
print msg

  ��B�����(��F��): �C��έp�Ȥ�˾�/�����, MTD, YTD
  �ɦW: SF_DailyOp2.sql

  IN	p_CompCode number:	���q�O
	p_ServiceType char(1):	�A�����O, ��: 'C', 'I'
	p_StopDate  varchar2:	�έp�I���, 'YYYYMMDD'
	p_CutOffTime varchar2:	Cut off�ɶ�, 'YYYYMMDDHHMI'
	p_Operator varchar2:	�ާ@�̦W��
	p_PRFlag number:	������Ȥ�O�_�C�J���Ĥ�, 0=�_, 1=�O
	p_ReturnSQL varchar2:	�h���]SQL����

  OUT   p_TmpTableName varchar2: �έp���G�Ȧs�ɦW (�I�s���ܼƦܤֻ�20 bytes)
	p_RetMsg VARCHAR2�G���G�T�� (�I�s���ܼƦܤֻ�200 bytes)

  Return NUMBER�G���G�N�X
        0: ���`����
	-1: p_RetMsg='�Ѽƿ��~'
	-2: p_RetMsg='�Ȧs���ɦW�إ߿��~, ���ˬd�O�_���إ�Sequence����: S_DSSRPT_TableName';
	-3: p_RetMsg='SQL1���~: ' || v_SQL1
	-4: p_RetMsg='�L�k�R���Ȧs��, �i��L�H���ϥΥ�����'
	-5: p_RetMsg='�L�k�إ߼Ȧs��, �i��L�H���ϥΥ�����, �еy��A��'
	-6: p_RetMsg='�s�ɦ���B���log�ɿ��~, SQLCODE='||SQLCODE||', '||'SQLERRM='||SQLERRM
        -99�Gp_RetMsg='��L���~'

  By: Pierre
  Date: 2002.03.14
	Q: �Y�Y��F�ϵL�������u���, �h�O�_�i�H�]��[�J(�ƭ�=0) ? ==> Yes
	2002.04.04 ��: �[�J��F�ϦW�����
	2002.04.09 ��: �p�⦳�īȤ��bug
	2002.04.18~22 ��: �[�J����B�w����(SO103)�ө�JSO102/SO502/SO502X��
	2002.04.26 ��: 2 bugs

   By: Lawrence
   Date: 2002.06.19
            �[: ��SO002 
         2003.11.18 Mark delete�Ϊk,�אּTruncate
*/
create or replace function SF_DailyOp2
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
  v_Month	varchar2(2);
  v_BudgetSQL	varchar2(500);
  v_BudgetAmt	number;
  v_Value01	number;
  v_Variance	number;

  v_ThisYear	char(4);
  v_SQL1 	varchar2(3000);
  v_SQL2 	varchar2(1000);
  v_SQL3 	varchar2(500);
  v_Where1	varchar2(100);
  v_ReturnCode  varchar2(500);

  z_AreaCode	varchar2(3);
  z_AreaName	varchar2(20);
  z_Amt 	number;

  type CurTyp is ref cursor;
  cDynamic CurTyp;

  -- �Ӥ��q�Ҧ���F�ϥN�X
  cursor cCD001 is
	select CodeNo, Description from CD001 where CompCode = p_CompCode and StopFlag!=1;

  -- ***********************

begin
  p_TmpTableName := null;

  --delete from AA909 where FormatName = 'SF_DailyOp2';
  --commit;

  -- Check parameters
  if p_StopDate is null or
    p_CompCode is null or p_ServiceType is null or p_CutoffTime is null or
    p_PRFlag<0 or p_PRFlag>1 then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_StartDate := substr(p_StopDate,1,6) || '01';
  -- DELETE FROM SO102;
  -- 2003.11.18 by Lawrence Mark delete�Ϊk,�אּTruncate
  DBMS_UTILITY.EXEC_DDL_STATEMENT('TRUNCATE TABLE SO102');
  -- commit;

  v_ThisYear := substr(p_StopDate,1,4);
  v_Where1 := 'A.CompCode='||p_CompCode||' and A.ServiceType='||c39||p_ServiceType||c39;
  v_Month := substr(p_StopDate,5,2);
  v_BudgetSQL := 'select nvl(Amt'||v_Month||',0) from SO103 A where '||v_Where1||' and A.Year='||v_ThisYear||' and ';
  -- DBMS_OUTPUT.PUT_LINE(v_BudgetSQL);

  -- ***************************************************************************
  -- [1] 3: 310		3.Direct Sales: 1. Actual total MTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C, SO014 D ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39||' group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 3, 'Direct Sales', 310, 'Actual total MTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp3.Direct Sales: 1. Actual total MTD');


  -- ***************************************************************************
  -- [2] 3: 340		3.Direct Sales: 4. Actual total YTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C, SO014 D ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and C.WorkClass='||c39||'D'||c39||' group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 3, 'Direct Sales', 340, 'Actual total YTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp3.Direct Sales: 4. Actual total YTD');


  -- ***************************************************************************
  -- [3] 4: 310		4.CSR Sales: 1. Actual total MTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C, SO014 D ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and (C.WorkClass is null or C.WorkClass!='||c39||'D'||c39||') group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 4, 'CSR Sales', 310, 'Actual total MTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp4.CSR Sales: 1. Actual total MTD');


  -- ***************************************************************************
  -- [4] 4: 340		4.CSR Sales: 4. Actual total YTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, CM003 C, SO014 D ' ||
		'where A.IntroId=C.EmpNo and A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) and (C.WorkClass is null or C.WorkClass!='||c39||'D'||c39||') group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 4, 'CSR Sales', 340, 'Actual total YTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp4.CSR Sales: 4. Actual total YTD');


  -- ***************************************************************************
  -- [5] 5: 110		5.Total Sales: 1. Actual total MTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, SO014 D ' ||
		'where A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 5, 'Total Sales', 110, 'Actual total MTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp5.Total Sales: 1. Actual total MTD');


  -- ***************************************************************************
  -- [6] 5: 140		5.Total Sales: 4. Actual total YTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, SO014 D ' ||
		'where A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.AcceptTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 5, 'Total Sales', 140, 'Actual total YTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp5.Total Sales: 4. Actual total YTD');


  -- ***************************************************************************
  -- [7] 6: 110		6.Subscribers/Total subscribers addition: 1. Actual total MTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, SO014 D ' ||
		'where A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 6, 'Subscribers', 110, 'Actual total MTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp6.Subscribers/Total subscribers addition: 1. Actual total MTD');


  -- ***************************************************************************
  -- [8] 6: 140		6.Subscribers/Total subscribers addition: 4. Actual total YTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, SO014 D ' ||
		'where A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 6, 'Subscribers', 140, 'Actual total YTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp6.Subscribers/Total subscribers addition: 4. Actual total YTD');


  -- ***************************************************************************
  -- [9] 6: 210		6.Subscribers/Total subscribers subtraction: 1. Actual total MTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO009 A, CD007 B, SO014 D ' ||
		'where A.PRCode=B.CodeNo and A.OldAddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_StartDate||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 6, 'Subscribers', 210, 'Actual total MTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp6.Subscribers/Total subscribers subtraction: 1. Actual total MTD');


  -- ***************************************************************************
  -- [10] 6: 240	6.Subscribers/Total subscribers subtraction: 4. Actual total YTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO009 A, CD007 B, SO014 D ' ||
		'where A.PRCode=B.CodeNo and A.OldAddrNo=D.AddrNo and ' ||
		v_Where1 || ' and A.FinTime>=to_date('||c39||v_ThisYear||'0101'||c39||','||c39||'YYYYMMDD'||c39||
		') and A.FinTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 6, 'Subscribers', 240, 'Actual total YTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp6.Subscribers/Total subscribers subtraction: 4. Actual total YTD');


  -- ***************************************************************************
  -- [11] 6: 310	[12] 6: 340	��ɨ��U��F�ϸ�Ʈ�, �@�_�p��
  -- ***************************************************************************

  -- ***************************************************************************
  -- [13] 6: 410	6.Subscribers/Active customers: 1. Actual total YTD
  -- 		������Ȥᦳ�i�ण�C�J���īȤ�, ���ѼƨM�w
  -- 2002.06.19
  v_SQL1 := 'select B.AreaName, B.AreaCode, count(*) Amt from SO001 A, SO014 B, SO002 C where A.InstAddrNo=B.AddrNo and C.CustId=A.CustId and C.ServiceType=' || chr(39) || p_ServiceType || chr(39) || ' and ';
  if p_PRFlag = 0 then			-- ������Ȥᤣ��
    v_SQL1 := v_SQL1 || 'c.CustStatusCode=1';
  else					-- ������Ȥ�n��
    v_SQL1 := v_SQL1 || '(c.CustStatusCode=1 or c.CustStatusCode=6)';
  end if;
  v_SQL1 := v_SQL1 || ' group by B.AreaName, B.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;
  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 6, 'Subscribers', 410, 'Actual total YTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);
  end loop;
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp6.Subscribers/Active customers: 1. Actual total MTD');


  -- ***************************************************************************
  -- [14] 7: 110	7.Other Activity: 1. Pending disconnect actual total MTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO009 A, CD007 B, SO014 D ' ||
		'where A.PRCode=B.CodeNo and A.OldAddrNo=D.AddrNo and ' ||
		v_Where1||' and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=2 or B.RefNo=5) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;

  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 7, 'Other Activity', 110, 'Pending disconnect actual total MTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop; 
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp7.Other Activity: 1. Pending disconnect actual total MTD');


  -- ***************************************************************************
  -- [15] 7: 120	7.Other Activity: 2. Pending installs actual total MTD
  v_SQL1 := 'select D.AreaName, D.AreaCode, nvl(count(*),0) Amt from SO007 A, CD005 B, SO014 D ' ||
		'where A.InstCode=B.CodeNo and A.AddrNo=D.AddrNo and ' ||
		v_Where1||' and A.AcceptTime<to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||
		')+1 and A.FinTime is null and A.SignDate is null and A.ReturnCode is null '||
		'and (B.RefNo=1 or B.RefNo=5 or B.RefNo=7) group by D.AreaName, D.AreaCode';

  begin		-- Open dynamic cursor
    open cDynamic for v_SQL1;
  exception
    when others then
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -3;
  end;

  loop
    z_AreaName := null;
    z_AreaCode := null;
    z_Amt := 0;
    fetch cDynamic into z_AreaName, z_AreaCode, z_Amt;
    exit when cDynamic%NotFound;

    insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	(p_CompCode, p_ServiceType, 7, 'Other Activity', 120, 'Pending installs actual total MTD', 
	z_Amt, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	v_StartExecTime, p_Operator, z_AreaCode, z_AreaName);

  end loop; 
  close cDynamic;
  DBMS_OUTPUT.PUT_LINE('�έp7.Other Activity: 2. Pending installs actual total MTD');


  -- ******************************************************************************************
  -- ******************************************************************************************
  -- @@@
  -- ���Ӥ��q�U��F�ϸɨ��@���Ӥ�Ӷ���B��� ==> �T�O�C�@��F�ϨC�@���س����@�����
  v_ErrCnt1 := 0;
  v_Cnt1 := 0;
  for cr1 in cCD001 loop
    ----- ���Ӧ�F�Ϲ�ڸ�� [1] 3: 310
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=3 and ItemCode2=310 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 3, 'Direct Sales', 310, 'Actual total MTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 3: 320		3. Direct Sales: 2. Month Budget #
    v_SQL2 := v_BudgetSQL||'ItemCode1=3 and ItemCode2=320 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;

--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 3, 'Direct Sales', 320, 'Month Budget #', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--    end if;

    -- ���͹�ڻP�w��t���� 3: 330	3. Direct Sales: 3. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 3, 'Direct Sales', 330, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [2] 3: 340
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=3 and ItemCode2=340 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 3, 'Direct Sales', 340, 'Actual total YTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 3: 350		3. Direct Sales: 5. End of Month YTD Budget
    v_SQL2 := v_BudgetSQL||'ItemCode1=3 and ItemCode2=350 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;

--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 3, 'Direct Sales', 350, 'End of Month YTD Budget', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 3: 360	3. Direct Sales: 6. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 3, 'Direct Sales', 360, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [3] 4: 310
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=4 and ItemCode2=310 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 4, 'CSR Sales', 310, 'Actual total MTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 4: 320		4. CSR Sales: 2. Month Budget #
    v_SQL2 := v_BudgetSQL||'ItemCode1=4 and ItemCode2=320 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 4, 'CSR Sales', 320, 'Month Budget #', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 4: 330	4. CSR Sales: 3. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 4, 'CSR Sales', 330, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [4] 4: 340
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=4 and ItemCode2=340 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 4, 'CSR Sales', 340, 'Actual total YTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 4: 350		4. CSR Sales: 5. End of Month YTD Budget
    v_SQL2 := v_BudgetSQL||'ItemCode1=4 and ItemCode2=350 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 4, 'CSR Sales', 350, 'End of Month YTD Budget', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 4: 360	4. CSR Sales: 6. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 4, 'CSR Sales', 360, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [5] 5: 110
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=5 and ItemCode2=110 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 5, 'Total Sales', 110, 'Actual total MTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 5: 120		5. Total Sales: 2. Month Budget #
    v_SQL2 := v_BudgetSQL||'ItemCode1=5 and ItemCode2=120 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 5, 'Total Sales', 120, 'Month Budget #', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 5: 130	5. Total Sales: 3. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 5, 'Total Sales', 130, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [6] 5: 140
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=5 and ItemCode2=140 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 5, 'Total Sales', 140, 'Actual total YTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 5: 150		5. Total Sales: 5. End of Month YTD Budget
    v_SQL2 := v_BudgetSQL||'ItemCode1=5 and ItemCode2=150 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 5, 'Total Sales', 150, 'End of Month YTD Budget', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 5: 160	5. Total Sales: 6. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 5, 'Total Sales', 160, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [7] 6: 110
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=110 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 110, 'Actual total MTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 6: 120		6. Subscribers: 2. Month Budget #
    v_SQL2 := v_BudgetSQL||'ItemCode1=6 and ItemCode2=120 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 120, 'Month Budget #', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 6: 130	6. Subscribers: 3. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 130, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [8] 6: 140
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=140 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 140, 'Actual total YTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
     exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 6: 150		6. Total Sales: 5. End of Month YTD Budget
    v_SQL2 := v_BudgetSQL||'ItemCode1=6 and ItemCode2=150 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 150, 'End of Month YTD Budget', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 6: 160	6. Total Sales: 6. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 160, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [9] 6: 210
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=210 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 210, 'Actual total MTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 6: 220		6. Subscribers: 2. Month Budget #
    v_SQL2 := v_BudgetSQL||'ItemCode1=6 and ItemCode2=220 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 220, 'Month Budget #', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 6: 230	6. Subscribers: 3. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 230, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [10] 6: 240
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=240 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 240, 'Actual total YTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
     exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 6: 250		6. Total Sales: 5. End of Month YTD Budget
    v_SQL2 := v_BudgetSQL||'ItemCode1=6 and ItemCode2=250 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 250, 'End of Month YTD Budget', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 6: 260	6. Total Sales: 6. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 260, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ---- [11] 6: 310 ���� [7] - [9]	([6:110] - [6:210])
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=110 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Cnt1 := null;
    fetch cDynamic into v_Cnt1;
    close cDynamic;

    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=210 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Cnt2 := null;
    fetch cDynamic into v_Cnt2;
    close cDynamic;

    begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 310, 'Actual total MTD', 
	  nvl(v_Cnt1,0)-nvl(v_Cnt2,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
   exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
   end;

    -- ���w���� 6: 320		6. Subscribers: 2. Month Budget #
    v_SQL2 := v_BudgetSQL||'ItemCode1=6 and ItemCode2=320 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 320, 'Month Budget #', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 6: 330	6. Subscribers: 3. Variance
    begin
      v_Variance := nvl(v_Cnt1,0)-nvl(v_Cnt2,0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 330, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ---- [12] 6: 340 ���� [8] - [10]	([6:140] - [6:240])
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=140 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Cnt1 := null;
    fetch cDynamic into v_Cnt1;
    close cDynamic;

    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=240 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Cnt2 := null;
    fetch cDynamic into v_Cnt2;
    close cDynamic;

    begin
	insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 340, 'Actual total MTD', 
	  nvl(v_Cnt1,0)-nvl(v_Cnt2,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
   exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
   end;

    -- ���w���� 6: 350		6. Total Sales: 5. End of Month YTD Budget
    v_SQL2 := v_BudgetSQL||'ItemCode1=6 and ItemCode2=350 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 350, 'End of Month YTD Budget', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 6: 360	6. Total Sales: 6. Variance
    begin
      v_Variance := nvl(v_Cnt1,0)-nvl(v_Cnt2,0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 360, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [13] 6: 410
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=6 and ItemCode2=410 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 410, 'Actual total YTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    -- ���w���� 6: 420		6. Subscribers: 2. End of Month YTD Budget
    v_SQL2 := v_BudgetSQL||'ItemCode1=6 and ItemCode2=420 and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL2;
    exception
      when others then
        p_RetMsg := 'SQL���~: ' || v_SQL2;
        return -3;
    end;
    v_BudgetAmt := null;
    fetch cDynamic into v_BudgetAmt;
    close cDynamic;
--    if v_BudgetAmt is null then				-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ��w����
       begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 420, 'End of Month YTD Budget', 
	  nvl(v_BudgetAmt,0), to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
--   end if;

    -- ���͹�ڻP�w��t���� 6: 430	6. Subscribers: 3. Variance
    begin
      v_Variance := nvl(v_Value01, 0) - nvl(v_BudgetAmt, 0);
      insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 6, 'Subscribers', 430, 'Variance', 
	  v_Variance, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
    exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
    end;
    --------------------------------------------------------------------------------

    ----- [14] 7: 110
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=7 and ItemCode2=110 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 7, 'Other Activity', 110, 'Pending disconnect actual total MTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;

    ----- [15] 7: 120
    v_SQL3 := 'select Value01 from SO102 where CompCode='||p_CompCode||' and ServiceType='||c39||p_ServiceType||c39||
		' and ItemCode1=7 and ItemCode2=120 and StopDate='||'to_date('||c39||p_StopDate||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and AreaCode='||c39||cr1.CodeNo||c39;
    begin		-- Open dynamic cursor
      open cDynamic for v_SQL3;
    exception
      when others then
	p_RetMsg := 'SQL���~: ' || v_SQL3;
	return -3;
    end;
    v_Value01 := null;
    fetch cDynamic into v_Value01;
    close cDynamic;

    if v_Value01 is null then			-- ���Ӥ��q�Ӧ�F�ϸɨ��@���Ӥ�Ӷ���B���
      begin
        insert into SO102 (CompCode, ServiceType, ItemCode1, ItemName1, ItemCode2, ItemName2, 
	  Value01, StartDate, StopDate, CutoffTime, ReportTime, Operator, AreaCode, AreaName) values 
	  (p_CompCode, p_ServiceType, 7, 'Other Activity', 120, 'Pending installs actual total MTD', 
	  0, to_date(v_StartDate,'YYYYMMDD'), to_date(p_StopDate,'YYYYMMDD'), to_date(p_CutoffTime,'YYYYMMDDHH24MI'),
	  v_StartExecTime, p_Operator, cr1.CodeNo, cr1.Description);
      exception
	when others then
	  v_ErrCnt1 := v_ErrCnt1 + 1;
      end;
    end if;


  end loop;


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
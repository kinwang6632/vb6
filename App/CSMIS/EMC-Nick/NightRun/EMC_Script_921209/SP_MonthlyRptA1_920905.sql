/*
@C:\EMC_Script\SP_MonthlyRptA1

exec SP_MonthlyRptA1

  �妸�@�~__�C��۰ʲ��ͥ[���W�D�涵���~���R��� (�I�sSF_RptA1��SF_RptA1_2)

  �ɦW: SP_MonthlyRptA1.sql 
  �y�{: 
	. ������Ѽ�
	. �I�sA1�����ݵ{��
	. �R���P�@������L�h�έp���
	. �N���G�J�����έp�θ��, �ýƻs��A1�C�����ɤ�(SO508)
	. log���浲�G

  A1�������Ѽ�:
	v_RealDate1 varchar2: 	���O��_�l, 'YYYYMMDD'		==> '20030101'
	v_RealDate2 varchar2:	���O��I��, 'YYYYMMDD'		==> �����ɤ��W��멳
	v_RefDate varchar2:	�ѦҤ��, 'YYYYMMDD'		==> �����ɤ��W��멳
	v_ServiceType char(1):	�A�����O			==> SO501A.ServiceType
	v_CompCode number:	���q�O				==> SO501A.CompCode
	v_CustId number:	�Ȥ�s��			==> null
	v_CustIdOp varchar2:	�Ȥ�s�������			==> '='
	v_ShouldDate1 varchar2:	������_�l, 'YYYYMMDD'		==> null
	v_ShouldDate2 varchar2:	������I��, 'YYYYMMDD'		==> null
	v_CitemSQL varchar2:	���O����SQL����r��		==> Refo=2 (�t�|���O)
	v_ClassSQL varchar2:	�Ȥ����OSQL����r��		==> null
	v_BillTypeSQL varchar2: ������OSQL����r��		==> null
	v_ServSQL varchar2:	�A��SQL����r��			==> null
	v_AreaSQL varchar2:	��F��SQL����r��		==> null
	v_StrtSQL varchar2:	��DSQL����r��			==> null
	v_CircuitNo varchar2:	�����s������			==> null
	v_TaxMode number:	�|�v�p��覡, 1=���|, 2=�w�|	==> 2 
	v_ErrYear number:	�p�O�_�l��W�L�X�~�������~����	==> 5
	v_MDUSQL varchar2:	�j�ӽs��SQL����r��		==> null
	v_Citem2SQL varchar2:	�D�g�����O����SQL����r��	==> '=998'

  By: Pierre
  Date: 2003.05.05
	2003.05.16 ��: �Ȥ�Ʃw�q�Ѳέp�Ȥ�Ƨאּ��뤧�Ȥ��
	2003.05.21 ��: �[�W�@������`���J�A���O���إN�X��0�A���O���ئW�٬�"�`���J"�A
			���ꦬ���B(Amt)�A�Ȥ��(Cnt)�A�]���䥭������N�i�H��Excel�ӭp�⬰Amt/Cnt

	2003.07.23 ��: �[�W�@��(����STB��)��ơA���O���إN�X��-1�A���O���ئW�٬�"����STB��"�A
			���ꦬ���B(Amt)��0�A�Ȥ��(Cnt)������STB��, �p���i��X�C��STB���������
			����STB�Ƥ��w�q: 
			. [�ѦҤ�]�Ӥ�e�˾�, �B�|����������l (�ثe���ɤ��ϥΤ���STB)
			. [�ѦҤ�]�Ӥ�e�˾�, �B����鸨�b[�ѦҤ�]�᪺���l (�Ӥ餴�ϥΤ���STB)

	2003.08.04 ��: �[�W�@��(����STB���)��ơA���O���إN�X��-2�A���O���ئW�٬�"����STB���"�A
			���ꦬ���B(Amt)��0�A�Ȥ��(Cnt)������STB��, �p���i��X�C��STB�᪺�������
			����STB��Ƥ��w�q: �ŦX�H�U����STB�����(���P�Ƚs)
			. [�ѦҤ�]�Ӥ�e�˾�, �B�|����������l (�ثe���ɤ��ϥΤ���STB)
			. [�ѦҤ�]�Ӥ�e�˾�, �B����鸨�b[�ѦҤ�]�᪺���l (�Ӥ餴�ϥΤ���STB)

	2003.09.05 ��: ���Ĥ�ƻP���ĥx�ƪ��w�q�A�[�W
			. CATV�ݩ󦳮Ĥ�(���`��), �B
			. EMC�|�ݥ[�W�Ȥ����O<700, ��LMSO�h�L����
		Q: �YCATV�Ȥ᪬�A���ݩ󦳮Ĥ�, �h���{�����ө�C��1�������, �_�h�|���~�t.
*/

create or replace procedure SP_MonthlyRptA1
as
  v_RealDate1 	varchar2(8) := '20030101';
  v_RealDate2 	varchar2(8);
  v_RefDate 	varchar2(8);
  v_ServiceType char(1);
  v_CompCode 	number;
  v_CustId 	number;
  v_CustIdOp 	varchar2(10);
  v_ShouldDate1 varchar2(8);
  v_ShouldDate2 varchar2(8);
  v_CitemSQL 	varchar2(1000);
  v_ClassSQL 	varchar2(500);
  v_BillTypeSQL varchar2(100);
  v_ServSQL 	varchar2(500);
  v_AreaSQL 	varchar2(500);
  v_StrtSQL 	varchar2(500);
  v_CircuitNo 	varchar2(500);
  v_TaxMode 	number := 2;
  v_ErrYear 	number := 5;
  v_MDUSQL 	varchar2(500);
  v_Citem2SQL 	varchar2(500) := '=998';

  v_Para16	number;
  v_RptYM	number;		-- �������k�ݤ��, ���YYYYMM
  v_Operator 	varchar2(10);
  v_JobId 	number;
  c39		char(1) := chr(39);
  v_RetCode 	number := 0;
  v_RetMsg 	varchar2(1000) := '�L���G';
  v_StartExecTime date;
  v_StopExecTime date;
  v_Second	number;
  v_Selection 	varchar2(3000);
  v_Now 	date;
  v_PrgName	varchar2(100);
  v_FuncName	varchar2(100);
  v_CitemCnt	number := 0;
  v_TotalCustCnt number := 0;
  v_PrDate1	date;
  v_PrDate2	date;
  v_FlowId	number := 0;		-- MSO�N��

  cursor cc1 is
    select CodeNo from CD019 where CodeNo>=800 and RefNO=2 order by CodeNo;

  cursor cc2 is
    select CitemCode, count(distinct CustId) CustCnt from SO089 where nvl(Month01,0)!=0 group by CitemCode;

begin
  -- ********************************************************
  -- (1) ��night-run�һݰѼ�
  -- ********************************************************
  v_JobId		:= 999990;
  v_StartExecTime 	:= sysdate;		-- �}�l����ɶ�
  v_Operator		:= 'Night-run';
  v_PrgName		:= 'SP_MonthlyRptA1';
  v_FuncName		:= '�C�벣�ͥ[���W�D�涵���~���R���';

  v_RealDate2 := to_char(last_day(add_months(v_StartExecTime,-1)), 'YYYYMMDD');
  -- ***************************************************************************************
  -- �Y�n��ʭp��U����, �h�ҥΤU��, �ק令�U����, compile�B���椧, 
  -- ���O�o���槹��, �n��^�쪬(���ѸӦ�), �B�Acompile�@��, �Ҧp
  --   v_RealDate2 := '20030131'; -- �i�p��92.01�몺�������
  -- ***************************************************************************************
  v_RefDate := v_RealDate2;
  v_RptYM := to_number(substr(v_RefDate,1,6));

  -- ��v_CompCode, v_ServiceType
  begin
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501A where Rownum<=1;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, 'SO501A: ����Ѽƥ��]�wok');
      commit;
      return;
  end;

  -- ��MSO�N��: 0=EMC, 1=TBC, ...
  begin
    select FlowId into v_FlowId from SO041 where RowNum<=1;
  exception
    when others then
      v_FlowId := 0;
  end;

  -- ��v_CitemSQL 
  v_CitemCnt := 0;
  v_CitemSQL := '';
  for cr1 in cc1 loop		-- �N���O���إN�X��','�覡��_��
    v_CitemSQL := v_CitemSQL || ',' || ltrim(to_char(cr1.CodeNo,'999'));
    v_CitemCnt := v_CitemCnt + 1;
  end loop;
  if v_CitemCnt > 0 then	-- ���h�̥��䪺','
    v_CitemSQL := substr(v_CitemSQL, 2);
  end if;
  if v_CitemCnt=1 then
    v_CitemSQL := '= ' || v_CitemSQL;
  elsif v_CitemCnt > 1 then
    v_CitemSQL := 'IN ('||v_CitemSQL||')';
  end if;
  DBMS_OUTPUT.PUT_LINE('v_CitemSQL: ' || v_CitemSQL);

  -- ��Para16, �H�K�M�w�I�s, 0=SF_RptA1, 1=SF_RptA1_2
  begin
    select Para16 into v_Para16 from SO043 where CompCode=v_CompCode and ServiceType=v_ServiceType;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '�L�k��SO043.Para16');
      commit;
      return;
  end;

  -- ********************************************************
  -- (2) �I�sA1�����ݵ{��
  -- ********************************************************
  if v_Para16 = 0 then
    v_RetCode := SF_RptA1(v_RealDate1, v_RealDate2, v_RefDate, v_ServiceType, v_CompCode, 
		v_CustId, v_CustIdOp, v_ShouldDate1, v_ShouldDate2, v_CitemSQL, v_ClassSQL, 
		v_BillTypeSQL, v_ServSQL, v_AreaSQL, v_StrtSQL, v_CircuitNo, v_TaxMode,	v_ErrYear, 
		v_MDUSQL, v_Citem2SQL, v_RetMsg);
  else
    v_RetCode := SF_RptA1_2(v_RealDate1, v_RealDate2, v_RefDate, v_ServiceType, v_CompCode, 
		v_CustId, v_CustIdOp, v_ShouldDate1, v_ShouldDate2, v_CitemSQL, v_ClassSQL, 
		v_BillTypeSQL, v_ServSQL, v_AreaSQL, v_StrtSQL, v_CircuitNo, v_TaxMode,	v_ErrYear, 
		v_MDUSQL, v_Citem2SQL, v_RetMsg);
  end if;

  -- ********************************************************
  -- (3) �N���G�J�����έp�θ��, �ýƻs��A1�C�����ɤ�(SO508)
  -- ********************************************************
  -- ���R���P�@������έp���
  delete SO508 where RptYM=v_RptYM;

  -- �NSO089���Ȧs��Ʋέp��, �ƻs��SO508
  begin
    insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, CitemCode, b.Description, 
	sum(nvl(RealAmt,0)), sum(nvl(MonthPast,0)), sum(nvl(Month01,0)), 0 
	from SO089 a, Cd019 b where a.CitemCode=b.CodeNo group by a.CitemCode, b.Description);
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '�L�k�N�έp���G�ƻs��A1�C������(SO0508)');
      commit;
      return;
  end;

  -- �A�έp��뤧�Ȥ��, �æ^��SO508����CustCnt		2003.05.16
  for cr2 in cc2 loop
    update SO508 set CustCnt=cr2.CustCnt where RptYM=v_RptYM and CitemCode=cr2.CitemCode;

    -- �]���ثeCd019����CodeNo��PK, �G���Ĳv�Ҷq, �H�U���O�ĪG�P�W, �G�ϥΤW�C���O
    --update SO508 set CustCnt=cr2.CustCnt where RptYM=v_RptYM and CompCode=v_CompCode and ServiceType=v_ServiceType
    --  and CitemCode=cr2.CitemCode;
  end loop;

  -- �[�W�@������`���J�A���O���إN�X��0�A���O���ئW�٬�'�`���J'
  -- ���ꦬ���B(Amt)�A�Ȥ��(Cnt)�A�]���䥭������N�i�H��Excel�ӭp�⬰Amt/Cnt 	2003.05.21
  begin
    insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, 0, '�`���J', 
	sum(nvl(RealAmt,0)), sum(nvl(MonthPast,0)), sum(nvl(Month01,0)), count(distinct CustId) 
	from SO089 where nvl(Month01,0)!=0 );
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '�L�k�έp����`���J���');
      commit;
      return;
  end;

  -- �έp: ��릳��STB��
  -- �[�W�@��(����STB��)��ơA���O���إN�X��-1�A���O���ئW�٬�"����STB��"�A
  -- ���ꦬ���B(Amt)��0�A�Ȥ��(Cnt)������STB��, �p���i��X�C��STB���������		2003.07.23
  -- ���Ĥ�ƻP���ĥx�ƪ��w�q�A�[�j			2003.09.01
  v_PrDate1 := to_date(substr(v_RefDate,1,6)||'01', 'YYYYMMDD');	-- �Ӥ뭺��
  v_PrDate2 := to_date(v_RefDate,'YYYYMMDD')+1;		-- �ѦҤ馸��
  begin
    if v_FlowId=0 then
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -1, '����STB��', 
	0, 0, 0, count(*) from SO004 A, CD022 B, SO002 C, SO001 D where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.CustId=D.CustId and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1 and D.ClassCode1<700);
    else
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -1, '����STB��', 
	0, 0, 0, count(*) from SO004 A, CD022 B, SO002 C where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1);
    end if;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '�L�k�έp��릳��STB�Ƹ��');
      commit;
      return;
  end;

  -- �έp: ��릳��STB���
  -- �[�W�@��(����STB��)��ơA���O���إN�X��-2�A���O���ئW�٬�"����STB���"�A
  -- ���ꦬ���B(Amt)��0�A�Ȥ��(Cnt)������STB��, �p���i��X�C��STB�᪺�������	2003.08.04
  -- ���Ĥ�ƻP���ĥx�ƪ��w�q�A�[�j			2003.09.01
  begin
    if v_FlowId=0 then
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -2, '����STB���', 
	0, 0, 0, count(distinct A.CustId) from SO004 A, CD022 B, SO002 C, SO001 D where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.CustId=D.CustId and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1 and D.ClassCode1<700);
    else
      insert into SO508 (RptYM, CompCode, ServiceType, CitemCode, CitemName, RealAmt, MonthPast, 
	Month01, CustCnt) (select v_RptYM, v_CompCode, v_ServiceType, -2, '����STB���', 
	0, 0, 0, count(distinct A.CustId) from SO004 A, CD022 B, SO002 C where A.FaciCode=B.CodeNo and B.RefNo=3 and 
	A.CustId=C.CustId and A.ServiceType=v_ServiceType and A.InstDate<v_PrDate2 and 
	(A.PrDate is null or A.PrDate>=v_PrDate1) and C.CustStatusCode=1);
    end if;
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '�L�k�έp��릳��STB��Ƹ��');
      commit;
      return;
  end;

  -- �NSO090�����`��ƽƻs��SO508A
  begin
    insert into SO508A (BillNo, Item, CustId, CitemCode, CitemName, RealDate, RealStartDate, RealStopDate, 
	RealPeriod, RealAmt, ErrorType) 
	(select BillNo, Item, CustId, CitemCode, CitemName, RealDate, RealStartDate, RealStopDate, 
	RealPeriod, RealAmt, ErrorType from SO090);
  exception
    when others then
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '�L�k�NSO090�����`��ƽƻs��SO508A');
      commit;
      return;
  end;

  -- ********************************************************
  -- (4) Log into operation log file: so030a, so054
  -- ********************************************************
  v_StopExecTime := sysdate;
  v_Second := trunc(86400*(sysdate-v_StartExecTime));

  if v_RetCode < 0 then
    DBMS_OUTPUT.PUT_LINE(v_FuncName||'--����A1��ݵ{�����~, Return code: ' || v_RetCode);
  else
    DBMS_OUTPUT.PUT_LINE(v_FuncName||'--����A1��ݵ{������, Return code: ' || v_RetCode);
    v_RetMsg := '���槹��';
    v_Selection := '���O��_�l:'||v_RealDate1||';���O��I��:'||v_RealDate2||';�ѦҤ��:'||v_RefDate||
		';���O����SQL����r��:'||v_CitemSQL||';�|�v�p��覡:'||v_TaxMode||
		';�D�g�����O����SQL����r��:'||v_Citem2SQL;
    insert into SO054 (PRGNAME, EMPNO, RUNDATETIME, SECOND, SELECTION, STOPDATETIME)
	values (v_PrgName, v_Operator, v_StartExecTime, v_Second, v_Selection, v_StopExecTime);
  end if;

  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, v_StopExecTime, 
	v_Second, v_PrgName, v_FuncName, v_RetCode, v_RetMsg);


  commit;

  <<GO_NEXT1>>
  NULL;


exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;

    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, -99, '��L���~');
    commit;

end;
/
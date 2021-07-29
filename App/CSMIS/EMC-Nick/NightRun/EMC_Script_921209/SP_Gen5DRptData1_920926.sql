/*
@c:\EMC_script\SP_Gen5DRptData1 
set serveroutput on
exec SP_Gen5DRptData1;  

  �妸�@�~: EMC��B�޲z������(SO5D01/02/03/04/07)���ͧ@�~
  �ɦW: SP_Gen5DRptData1.sql

  ����: 
	1. ���{���i��Night-run�Ӧ۰ʨC��5�������@��
	2. ���{�������涶�����ө�b[EMC�C�Ӥ�ƥ����n��Ƶ{��(SP_BackupCustData)]����
	3. ���{�����έp��ƨӷ��ɬ�EMCxxx_SO001ALL_yyyymm(�C��CSO�@��), �θ˾����u��
	   ����/��������u�����
	4. �έp���G�s���"��B�޲z��������"(SO511A/B/C), �Τ��q�O�θ�Ƥ���ӰϧO, �óz
	   �L�ƥ�����(Snapshot), �N�USO�C��έp��ƶǰe���`���q���ƥ��D����, ��Web report
	   �Ϋe�ݳ���{������
	5. �]��"���Ĥ��"���έp��ɪ���, �ä��O�W�멳����, �G���{���|�έp�W�멳����
	   ���˾���ƻP������, �ä�����έp�ȥh
	6. �p�����Ҥw���WOracle night-run?
		���� select * from User_Jobs���O, �ݬݬO�_����SP_Gen5DRptData1��job
	7. �p���ˬd�C��night-run�����G?
		���� select * from SO030A where JobId=999987

  ���G�ɮ׵��c: 
  1. 5D01��B�޲z��������: SO511A
	. ���q�O		CompCode	N(3)
	. ��Ƥ��		RptYM		N(6), ���~��
	. �Ȥ�j���O�N�X	ClassCode	N(3), 1~7
	. ú�O���O�N�X		PeriodCode	N(3), 1~5
	. �U�����		NextYM		N(6), ���~��
	. �Ȥ��		CustCount	N(8)
	. ��ƺ���		DataType	N(1), 0=A��, 1=D��
     ����:
	. RptYm
	. CompCode, RptYM, ClassCode, PeriodCode, NextYM

  2. 5D02���O��ƹw���έp��: SO511B
	. ���q�O		CompCode	N(3)
	. ��Ƥ��		RptYM		N(6), ���~��
	. �Ȥ�j���O�N�X	ClassCode	N(3), 1~7
	. ú�O���O�N�X		PeriodCode	N(3), 1~5
	. �U�����		NextYM		N(6), ���~��
	. �Ȥ��		CustCount	N(8)
     ����:
	. RptYm
	. CompCode, RptYM, ClassCode, PeriodCode, NextYM

  3. 5D03���O���B�w���έp��: SO511C
	. ���q�O		CompCode	N(3)
	. ��Ƥ��		RptYM		N(6), ���~��
	. �Ȥ�j���O�N�X	ClassCode	N(3), 1~7
	. ú�O���O�N�X		PeriodCode	N(3), 1~5
	. �U�����		NextYM		N(6), ���~��
	. ���B			Amount		N(10)
     ����:
	. RptYm
	. CompCode, RptYM, ClassCode, PeriodCode, NextYM

  4. �Uú�O���O���P����: SO511Z
	. ���q�O		CompCode	N(3)
	. �Ȥ�j���O�N�X	ClassCode	N(3), 1~7
	. ú�O���O�N�X		PeriodCode	N(3), 1~5
	. �P�����B		UnitPrice	N(8)
     ����:
	. CompCode, ClassCode, PeriodCode

  By: Pierre
  Date: 2003.09.04
	2003.09.05 ��:	1. �Ȥ����O702�令701
			2. A���Ȥ�ƭn�p�ⵧ��, ���OCustCount�[�_��

	2003.09.26 ��: �W�멳���᪺�˾���Ƴ����Ƨ令��A�����Ʀr��h, �����D���������O���諸
*/

CREATE OR REPLACE PROCEDURE SP_Gen5DRptData1  IS
  v_RetMsg	varchar2(100);
  v_RetCode	number	:= -99;
  v_CompCode	number;
  v_OwnerName	varchar2(20);
  v_JobID	number;
  v_StartExecTime date;
  v_StopExecTime  date;
  v_ServiceType     char(1);
  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);
  v_SQL		varchar2(4000); -- for SQL statement
  v_Time	date;
  v_RptYM	number;		-- ��Ƥ��
  v_StartDate	date;		-- ��ƺI��餧�W�뭺��
  v_StopDate	date;		-- ��ƺI���
  v_TableName	varchar2(100);	-- �έp��ƨӷ����ɦW
  v_cnt1	number;		-- counter 1
  c39		char(1) := chr(39);
  I		number;
  J		number;
  K		number;
  v_StartMonth	number;		-- ��ƺI���W�뤧�`���, ��: 200307==>2003*12+7
  v_CntA	number;
  v_CntD	number;
  v_Delta	number;
  v_UnitPrice	number;

  type CurTyp is ref cursor;	-- �ۭqcursor���A
  v_DyCursor1 CurTyp;		-- ��dynamic SQL
  v_DyCursor2 CurTyp;		-- ��dynamic SQL
  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER; --�ŧi�@��Table�����A(�Ϊk�M�}�C�P)
  ACustCount NumberAry;		-- A�����G�}�C
  DCustCount NumberAry;		-- D�����G�}�C
  MonthAry NumberAry;		-- �U������}�C, ���sv_StartDate�_��4�Ӥ��

  TYPE VC2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER; --�ŧi�@��Table�����A(�Ϊk�M�}�C�P)
  TNameAry VC2Ary;		-- array for table names

  x_ClassCode	number;		-- �q[�έp��ƨӷ���]���X���Ȥ����O�N�X
  x_Period	number;		-- �q[�έp��ƨӷ���]���X������
  x_CustCount	number;		-- �q[�έp��ƨӷ���]���X���Ȥ��
  x_ClctDate	date;		-- �q[�έp��ƨӷ���]���X���U����
  z_ClassCode	number;		-- �Ȥ�j���O�N�X
  z_PeriodCode	number;		-- ú�O���O�N�X
  z_MonthIndex	number;		-- �U����Ҹ������index, �i��Ȭ�1~14

begin
  -- ********************************************************
  -- �]�w�Ѽƪ��
  -- ********************************************************
  v_JobID	:= 999987;		-- �}�դ��wJob�s��
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_ServiceType := 'C';
  v_FuncName	:= 'EMC��B�޲z������(SO5D01/02/03/04/07)���ͧ@�~';	-- �\��W��
  v_PrgName	:= 'SP_GEN5DRPTDATA1';	-- �{���W��

  -- ********************************************************
  -- �M�w[��ƺI���], [��Ƥ��]�ΦU�ܼƪ��
  -- ********************************************************
  v_StopDate	:= last_day(add_months(v_StartExecTime,-1));
  --v_StopDate	:= to_date('20030731', 'YYYYMMDD');	-- for testing
  v_RptYM	:= to_number(to_char(v_StopDate, 'YYYYMM'));
  v_StartDate	:= last_day(add_months(v_StartExecTime, -3))+1;
  v_StartMonth	:= to_number(to_char(v_StartDate,'YYYY'))*12 + to_number(to_char(v_StartDate,'MM'));
  for I in 1..14 loop
    MonthAry(I) := to_number(to_char(add_months(v_StartDate,I-1),'YYYYMM'));
  end loop;

  -- ********************************************************
  -- �����q�O�N�X
  -- ********************************************************
  begin 
    select CompCode into v_CompCode from SO501A where RowNum<=1;
  exception
    when others then
      v_RetMsg := '�����q�N�X�ɵo�Ϳ��~, '||SQLERRM;
      v_RetCode := -1;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  -- ********************************************************
  -- ���Ӹ�ưϥD�nOwner name
  -- ********************************************************
  begin
    select upper(Description) into v_OwnerName from SO507 where CodeNo=v_CompCode;
    ---- DBMS_OUTPUT.PUT_LINE('Owner name = '||v_OwnerName);
  exception
    when others then
      v_RetMsg := '���Ӹ�ưϥD�nOwner name�ɵo�Ϳ��~, '||SQLERRM;
      v_RetCode := -2;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  -- ********************************************************
  -- �զ�[�έp��ƨӷ���]�ɦW
  -- ********************************************************
  v_TableName := v_OwnerName||'_SO001ALL_'||ltrim(to_char(v_RptYM,'999999'));
  ---- DBMS_OUTPUT.PUT_LINE('Table name = '||v_TableName);

  -- ********************************************************
  -- �ˬd��ưϤ��O�_�w����[�έp��ƨӷ���], �Y�L, �hlog���~�õ����{��
  -- ********************************************************
  select count(*) into v_cnt1 from user_tables where table_name=v_TableName;
  if v_cnt1 = 0 then 
    v_RetMsg := '����ưϵL�έp��ƨӷ���, �ɦW: '||v_TableName;
    v_RetCode := -3;
    DBMS_OUTPUT.PUT_LINE(v_RetMsg);
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
    return;
  end if; 

  -- ********************************************************
  -- �ˬd��ưϤ��O�_�w��[��B�޲z��������], �Y�L, �hlog���~�õ����{��
  -- ********************************************************
  TNameAry(1) := 'SO511A';
  TNameAry(2) := 'SO511B';
  TNameAry(3) := 'SO511C';
  TNameAry(4) := 'SO511Z';
  for I in 1..4 loop
    select count(*) into v_cnt1 from user_tables where table_name=TNameAry(I);
    if v_cnt1 = 0 then 
      v_RetMsg := '����ưϵL��B�޲z��������, �ɦW: '||TNameAry(I);
      v_RetCode := -4;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
    end if;
  end loop;

  -- ********************************************************
  -- �M��[��B�޲z��������]���P�@[��Ƥ��]���έp���, �H�K�i���ƭp�⦹�{��
  -- ********************************************************
  delete from SO511A where CompCode=v_CompCode and RptYM=v_RptYM;
  delete from SO511B where CompCode=v_CompCode and RptYM=v_RptYM;
  delete from SO511C where CompCode=v_CompCode and RptYM=v_RptYM;
  -- commit;	-- ���ncommit, �H�Krollback segment����

  -- ********************************************************
  -- �]�w[A�����G�}�C], [D�����G�}�C]���
  -- ********************************************************
  for I in 1..7*5*14 loop
    ACustCount(I) := 0;
    DCustCount(I) := 0;
  end loop;

  -----------------------------------------------------------
  -- ����5D01��B�޲z��������: SO511A
  -----------------------------------------------------------
  -- ********************************************************
  -- ��[�έp��ƨӷ���]���p��A���Ȥ��
  -- ********************************************************
  begin
    v_SQL := 'select ClassCode1, nvl(Period,0), nvl(CustCount,0), ClctDate from '||v_TableName||
	' where CompCode='||v_CompCode||' and ClctDate>='||GiPackage.QryDTString0(v_StartDate)||
	' and CustStatusCode=1 and ClassCode1>=100 and ClassCode1<=799 and ClassCode1!=701';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor1 for v_SQL;
  exception
    when others then
      v_RetMsg := '��[�έp��ƨӷ���]���p��A���Ȥ�Ʈɵo�Ϳ��~, '||SQLERRM;
      v_RetCode := -5;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop�C�@���ŦX���󪺫Ȥ���
    fetch v_DyCursor1 INTO x_ClassCode, x_Period, x_CustCount, x_ClctDate; 
    exit when v_DyCursor1%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- �Ȥ�j���O�N�X
    if x_Period<=1 then				-- ú�O���O�N�X
      z_PeriodCode := 1;
    elsif x_Period=2 then
      z_PeriodCode := 2;
    elsif x_Period>=3 and x_Period<=5 then
      z_PeriodCode := 3;
    elsif x_Period>=6 and x_Period<=11 then
      z_PeriodCode := 4;
    else
      z_PeriodCode := 5;
    end if;

    if x_CustCount = 0 then			-- �Y�Ȥ�Ƭ�0��null, �h�ন1
      x_CustCount := 1;
    end if;

    -- �p��[�U�����]���k�ݪ����, �i��Ȭ�1~14
    z_MonthIndex := to_number(to_char(x_ClctDate,'YYYY'))*12 + to_number(to_char(x_ClctDate,'MM')) - v_StartMonth + 1;
    if z_MonthIndex > 14 then
      z_MonthIndex := 14;
    end if;

    -- �N[�Ȥ��]�֭p��[A�����G�}�C]������������ ==> 2003.09.05 �O����, ���O�Ȥ��
    ACustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) := 
	ACustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) + 1;

  end loop;		-- Loop�C�@���ŦX���󪺫Ȥ���
  close v_DyCursor1;

  -- ********************************************************	2003.09.26
  -- �έp[��ƺI���]�����ᤧ��˾������(�U�ثȤ����O�����), 
  -- �ä�����[A�����G�}�C]���������(����)
  -- ����: SO007�����u���O�N�X���ѦҸ��O1��5, �B���u�ɶ��ߩ�[��ƺI���]���U���O�Ȥ��
  -- ********************************************************
  begin
    v_SQL := 'select B.ClassCode1, count(A.CustId) from SO007 A, SO001 B, CD005 C'||
	' where A.CustId=B.CustId and A.InstCode=C.CodeNo and A.FinTime>'||GiPackage.QryDTString0(v_StopDate)||
	' and (C.RefNo=1 or C.RefNo=5) and B.ClassCode1>=100 and B.ClassCode1<=799 and B.ClassCode1!=701'||
	' group by B.ClassCode1';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor1 for v_SQL;
  exception
    when others then
      v_RetMsg := '�έp[��ƺI���]�����ᤧ��˾�����Ʈɵo�Ϳ��~, '||SQLERRM;
      v_RetCode := -7;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop�C�@���ŦX���󪺫Ȥ���
    fetch v_DyCursor1 INTO x_ClassCode, x_CustCount; 
    exit when v_DyCursor1%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- �Ȥ�j���O�N�X

    -- �Y�Ȥ����O��400~499(�j�ӲΦ�), �h�N��ƥ�[�Ȥ�j���O�N�X]��4����1�Ӥ뤧��ú��Ƥ�����
    -- �_�h, �N��ƥ�[�Ȥ�j���O�N�X]��������14�Ӥ뤧�~ú��Ƥ�����
    if z_ClassCode = 4 then
      ACustCount((4-1)*70+1) := ACustCount((4-1)*70+1) - x_CustCount;
    else
      ACustCount((z_ClassCode-1)*70+(5-1)*14+14) := ACustCount((z_ClassCode-1)*70+(5-1)*14+14) - x_CustCount;
    end if;
  end loop;
  close v_DyCursor1;

  -- ********************************************************	2003.09.26
  -- �έp[��ƺI���]�����ᤧ�����������(�U�ثȤ����O�����), 
  -- �ä�����[A�����G�}�C]���������(�[�^)
  -- ����: SO009�����u���O�N�X���ѦҸ��O2��5, �B���z�ɶ��ߩ�[��ƺI���]���U���O�Ȥ��
  -- ********************************************************
  begin
    v_SQL := 'select B.ClassCode1, count(A.CustId) from SO009 A, SO001 B, CD007 C'||
	' where A.CustId=B.CustId and A.PrCode=C.CodeNo and A.AcceptTime>'||GiPackage.QryDTString0(v_StopDate)||
	' and (C.RefNo=2 or C.RefNo=5) and B.ClassCode1>=100 and B.ClassCode1<=799 and B.ClassCode1!=701'||
	' group by B.ClassCode1';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor1 for v_SQL;
  exception
    when others then
      v_RetMsg := '�έp[��ƺI���]�����ᤧ����������Ʈɵo�Ϳ��~, '||SQLERRM;
      v_RetCode := -8;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop�C�@���ŦX���󪺫Ȥ���
    fetch v_DyCursor1 INTO x_ClassCode, x_CustCount; 
    exit when v_DyCursor1%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- �Ȥ�j���O�N�X

    -- �Y�Ȥ����O��400~499(�j�ӲΦ�), �h�N��Ʋ֭p��[�Ȥ�j���O�N�X]��4����1�Ӥ뤧��ú���
    -- �_�h, �N��Ʋ֭p��[�Ȥ�j���O�N�X]��������14�Ӥ뤧�~ú���
    if z_ClassCode = 4 then
      ACustCount((4-1)*70+1) := ACustCount((4-1)*70+1) + x_CustCount;
    else
      ACustCount((z_ClassCode-1)*70+(5-1)*14+14) := ACustCount((z_ClassCode-1)*70+(5-1)*14+14) + x_CustCount;
    end if;
  end loop;
  close v_DyCursor1;

  -- ********************************************************
  -- ��[�έp��ƨӷ���]���p��D���Ȥ��
  -- ********************************************************
  begin
    v_SQL := 'select ClassCode1, nvl(Period,0), nvl(CustCount,0), ClctDate from '||v_TableName||
	' where CompCode='||v_CompCode||' and (RealStopDate+1)>='||GiPackage.QryDTString0(v_StartDate)||
	' and CustStatusCode=1 and ClassCode1>=100 and ClassCode1<=799 and ClassCode1!=701';
    --DBMS_OUTPUT.PUT_LINE(v_SQL);
    open v_DyCursor2 for v_SQL;
  exception
    when others then
      v_RetMsg := '��[�έp��ƨӷ���]���p��D���Ȥ�Ʈɵo�Ϳ��~, '||SQLERRM;
      v_RetCode := -6;
      DBMS_OUTPUT.PUT_LINE(v_RetMsg);
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  loop 			-- Loop�C�@���ŦX���󪺫Ȥ���
    fetch v_DyCursor2 INTO x_ClassCode, x_Period, x_CustCount, x_ClctDate; 
    exit when v_DyCursor2%NOTFOUND;
    
    z_ClassCode := trunc(x_ClassCode/100);	-- �Ȥ�j���O�N�X
    if x_Period<=1 then				-- ú�O���O�N�X
      z_PeriodCode := 1;
    elsif x_Period=2 then
      z_PeriodCode := 2;
    elsif x_Period>=3 and x_Period<=5 then
      z_PeriodCode := 3;
    elsif x_Period>=6 and x_Period<=11 then
      z_PeriodCode := 4;
    else
      z_PeriodCode := 5;
    end if;

    if x_CustCount = 0 then			-- �Y�Ȥ�Ƭ�0��null, �h�ন1
      x_CustCount := 1;
    end if;

    -- �p��[�U�����]���k�ݪ����, �i��Ȭ�1~14
    z_MonthIndex := to_number(to_char(x_ClctDate,'YYYY'))*12 + to_number(to_char(x_ClctDate,'MM')) - v_StartMonth + 1;
    if z_MonthIndex > 14 then
      z_MonthIndex := 14;
    end if;

    -- �N[�Ȥ��]�֭p��[D�����G�}�C]������������
    DCustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) := 
	DCustCount((z_ClassCode-1)*70+(z_PeriodCode-1)*14+z_MonthIndex) + x_CustCount;
  end loop;		-- Loop�C�@���ŦX���󪺫Ȥ���
  close v_DyCursor2;

  -- ********************************************************
  /* �p��D���PA������[�Ȥ�j���O�N�X]�p�p���Ȥ�Ʈt��, �ä�����[D�����G�}�C]���������, 
     �Ϩ�̭Y��[�Ȥ�j���O�N�X]�p�p���`��Ƨ����@�P
	.. Loop 7��[�Ȥ�j���O]
		... �q[A�����G�}�C]���p�p14�Ӥ몺A���Ȥ�� ==> A
		... �q[D�����G�}�C]���p�p14�Ӥ몺D���Ȥ�� ==> B
		... �p��W�z��Ȫ��t��(A-B) ==> C
		... �Y�Ȥ����O��400~499(�j�ӲΦ�), �h�NC�Ȳ֭p�ܸ�[�Ȥ�j���O�N�X]����1�Ӥ뤧��ú���
		... �_�h, �NC�Ȳ֭p�ܸ�[�Ȥ�j���O�N�X]����14�Ӥ뤧�~ú���
  */
  -- ********************************************************
  v_Delta := 0;
  for I in 1..7 loop
    v_CntA := 0;
    v_CntD := 0;
    for J in 1..5 loop
      for K in 1..14 loop
	v_CntA := v_CntA + ACustCount((I-1)*70+(J-1)*14+K);	-- �ӫȤ�j���O��A���Ȥ�Ƥp�p
	v_CntD := v_CntD + DCustCount((I-1)*70+(J-1)*14+K);	-- �ӫȤ�j���O��D���Ȥ�Ƥp�p
      end loop;
    end loop;

    v_Delta := v_Delta + v_CntA-v_CntD;
    if I=4 then
      DCustCount((4-1)*70+1) := DCustCount((4-1)*70+1) + (v_CntA-v_CntD);
    else
      DCustCount((I-1)*70+(5-1)*14+14) := DCustCount((I-1)*70+(5-1)*14+14) + (v_CntA-v_CntD);
    end if;
  end loop;

  -- ********************************************************
  -- �N[A�����G�}�C]�x�s��[��B�޲z��������SO511A]	2003.09.26
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511A (CompCode, RptYM, ClassCode, PeriodCode, NextYM, CustCount, DataType)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), ACustCount((I-1)*70+(J-1)*14+K), 0);
      end loop;
    end loop;
  end loop;

  -- ********************************************************
  -- �N[D�����G�}�C]�x�s��[��B�޲z��������SO511A]
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511A (CompCode, RptYM, ClassCode, PeriodCode, NextYM, CustCount, DataType)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), DCustCount((I-1)*70+(J-1)*14+K), 1);
      end loop;
    end loop;
  end loop;

  -----------------------------------------------------------
  -- ����5D02���O��ƹw���έp��: SO511B
  ------------------------------------------------------------
  /* 
  ���k: �q[D�����G�}�C]����X�C�Ӥ���֭p�����O���
  . Loop I.�C�ثȤ�j���O
	.. Loop J.�C��ú�O���O
		... Loop K.�C�ӤU��������Ȥ��
			* �Y�����Ȥ�Ƥ��O0, �h�p��[�s�U�����]
				[�s�U�����] = �Ӥ���[�Wú�O���O���
			* �Y[�s�U�����]���W�L14�Ӥ�, �h�N�ӫȤ�Ʋ֥[��[�s�U�����]���Ȥ��
		... end of Loop K
	.. end of Loop J
  . end of Loop I
  */
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	if DCustCount((I-1)*70+(J-1)*14+K) != 0 then
	  if J <=3 then
	    z_MonthIndex := k + j;	-- 1=��ú, 2=����ú, 3=�uú
	  elsif J = 4 then
	    z_MonthIndex := k + 6;	-- 4=�b�~ú
	  else
	    z_MonthIndex := k + 12;	-- 5=�~ú
	  end if;
	  if z_MonthIndex <= 14 then
	    DCustCount((I-1)*70+(J-1)*14+z_MonthIndex) := DCustCount((I-1)*70+(J-1)*14+z_MonthIndex) + 
		DCustCount((I-1)*70+(J-1)*14+K);
	  end if;
	end if;
      end loop;
    end loop;
  end loop;

  -- ********************************************************
  -- �N[D�����G�}�C]�x�s��[���O��ƹw���έp��SO511B]
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511B (CompCode, RptYM, ClassCode, PeriodCode, NextYM, CustCount)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), DCustCount((I-1)*70+(J-1)*14+K));
      end loop;
    end loop;
  end loop;


  ------------------------------------------------------------
  -- ����5D03���O���B�w���έp��: SO511C
  ------------------------------------------------------------
  /*
  ���k: �q5D02�ҭp�⤧[D�����G�}�C], ���W��[�Ȥ�j���O]��[ú�O���O]
	  �ҹ������P��, �i����X�C�Ӥ���֭p�����O���B
  . Loop I.�C�ثȤ�j���O
	.. Loop J.�C��ú�O���O
		... ��[�Ȥ�j���O]��[ú�O���O]�ҹ�����[�P��]
		... �Y������[�P��], �hlog�@�����~���, �ñN[�P��]�]��0
		... Loop K.�C�ӤU�����
			* �N�Ӥ�����֭p���O��ƭ��W[�P��]
			* �A�s�J������[D�����G�}�C]�����
		... end of Loop K
	.. end of Loop J
  . end of Loop I
  */
  for I in 1..7 loop
    for J in 1..5 loop
      -- ��[�Ȥ�j���O]��[ú�O���O]�ҹ�����[�P��]
      begin
	select UnitPrice into v_UnitPrice from SO511Z where ClassCode=I and PeriodCode=J;
      exception
	when others then
	  v_UnitPrice := 0;
	  v_RetMsg := '�P����(SO511Z)���L�Ȥ�j���O'||I||'�Pú�O���O'||J||'�ҹ����쪺�P��';
	  v_RetCode := -9;
	  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	    TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	    values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	    trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      end;

      for K in 1..14 loop
	if DCustCount((I-1)*70+(J-1)*14+K) != 0 then
	  -- �N�Ӥ�����֭p���O��ƭ��W[�P��], �A�s�J������[D�����G�}�C]�����
	  DCustCount((I-1)*70+(J-1)*14+K) := DCustCount((I-1)*70+(J-1)*14+K) * v_UnitPrice;
	end if;
      end loop;
    end loop;
  end loop;

  -- ********************************************************
  -- �N[D�����G�}�C]�x�s��[���O���B�w���έp��SO511C]
  -- ********************************************************
  for I in 1..7 loop
    for J in 1..5 loop
      for K in 1..14 loop
	insert into SO511C (CompCode, RptYM, ClassCode, PeriodCode, NextYM, Amount)
	  values (v_CompCode, v_RptYM, I, J, MonthAry(K), DCustCount((I-1)*70+(J-1)*14+K));
      end loop;
    end loop;
  end loop;


  -- ********************************************************
  -- log���榨�\�O����[Night-run���Glog��](SO030A)
  -- ********************************************************
  v_StopExecTime := sysdate;		-- ��������ɶ�
  v_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(v_StopExecTime-v_StartExecTime)))||'��';
  v_RetCode := 0;
  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);

  commit;
  DBMS_OUTPUT.PUT_LINE(v_RetMsg);

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code='||SQLCODE||' '|| 'Error Message='||SQLERRM);
    v_RetMsg := '��L���~ :'||SQLERRM;
    v_RetCode := -99;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
    return;
END;
/
/*
-- @c:\emc_script\SP_CntStbLog2
-- @C:\Gird\V400\Script\SP_CntStbLog2
set serveroutput on
exec SP_CntStbLog2;

  �έp�C��STB�}������/���\�v/���Ѳv�A�æs��TABLE(SO510)
  �ɦW: SP_CntStbLog2.sql

  �`�N:
	1. ���]Send_Nagra�ɮ׬O�b�U��Ʈw��Com��ưϤ�, �Y���O, �h�{������'COM.'�r��n����
	2. �έp���G�ɵ��c(SO510, SO510A): �аѦҬ������

  SO030A���~�T��:
	 0:  '�έp����'
	-1:  '���ӫ��O���\���Ʀ����~: '||SQLERRM;
	-2:  '���ӫ��O���Ѧ��Ʀ����~: '||SQLERRM;
	-3:  '�O���έp���G��SO510�����~: '||SQLERRM;
	-4:  'Night-run�Ѽ���SO501A�L���'
	-5:  '�O���έp���G��SO510A�����~: '||SQLERRM;
	-99: '��L���~: '||SQLERRM;

  By: Pierre
  Date: 2003.08.28
	2003.11.03 ��: EMC��T��92.10.31�ݨD�A�Ʊ榹�{����O��STB�}�����O������~�����Ӹ��
*/

create or replace procedure SP_CntStbLog2
as
  c39 char(1) := chr(39);
  v_CmdId varchar2(5);
  v_CmdName varchar2(20);
  v_CompCode number;
  v_Tmp1 varchar2(20);
  v_Tmp2 varchar2(20);
  v_RptDate date;

  v_OkCnt number := 0;
  v_ErrCnt number := 0;
  v_OkRate number := 0;
  v_ErrRate number := 0;
  v_TotalCnt number := 0;

  v_Operator 	varchar2(10);
  v_JobId 	number;
  v_RetCode 	number := 0;
  v_RetMsg 	varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;
  v_PrgName	varchar2(100);
  v_FuncName	varchar2(100);
  x_CompCode number := 0;
  v_CompCnt	number := 0;

  v_SQL varchar2(4000);
  type CurTyp is ref cursor;  --�ۭqcursor���A
  v_DyCursor CurTyp;          --��dynamic SQL

  cursor cc1 is 
    select distinct CompCode from so501A;

begin
  -- ********************************************************
  -- (1) �]�wnight-run�һݰѼ�
  -- ********************************************************
  v_JobId		:= 999988;		-- �}�դ��wJob�s��
  v_StartExecTime 	:= sysdate;		-- �}�l����ɶ�
  v_Operator		:= 'Night-run';		-- Night-run���Φ��W��
  v_PrgName		:= 'SP_CntStbLog2';	-- �{���W��
  v_FuncName		:= '�έp�C��STB�}������/���\�v/���Ѳv';	-- �\��W��

  v_CmdId := 'A1';			-- �}�����O�N��
  v_RptDate := trunc(sysdate)-1;	-- ��Ƥ�: ���ثe�ɶ����e�@��
  --v_RptDate := to_date('20030329', 'YYYYMMDD'); -- ���ե�
  v_Tmp1 := GiPackage.ED2CDString(v_RptDate);	-- �C���έp�@��Ѫ�v_CmdId���O�����\/����...�����
  v_Tmp2 := GiPackage.ED2CDString(v_RptDate+1);

  -- ���������O�W��			-- ���q���O���n��
  if v_CmdId = 'A1' then
    v_CmdName := '�}��';
  elsif v_CmdId = 'A2' then
    v_CmdName := '����';
  elsif v_CmdId = 'A3' then
    v_CmdName := '����';
  elsif v_CmdId = 'A4' then
    v_CmdName := '�_��';
  elsif v_CmdId = 'E1' then
    v_CmdName := '�]�w�ˤl�K�X';
  else
    DBMS_OUTPUT.PUT_LINE('�L���������O�N��: '||v_CmdId);
    return;
  end if;

  -- ********************************************************
  -- (2) Loop�C�@���q, �Ӳέp�Ӥ��q���ӫ��O���污��
  -- ********************************************************
  for cr1 in cc1 loop
    x_CompCode := cr1.CompCode;
    v_CompCnt := v_CompCnt + 1;		-- �֭p���q�O
    v_TotalCnt := 0;
    v_OkCnt := 0;
    v_ErrCnt := 0;

    -- ���ӫ��O���\����
    begin
      select count(*) into v_OkCnt from SOAC0202 where UpdTime>=v_Tmp1 and UpdTime<v_Tmp2
	 and ModeType=v_CmdId and CompCode=x_CompCode;
    exception
      when others then
	v_RetCode := -1;
	v_RetMsg := '���ӫ��O���\���Ʀ����~: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;

    -- ���ӫ��O���Ѧ���
    -- ���]Send_Nagra�ɮ׬O�b�U��Ʈw��Com��ưϤ�, �Y���O, �h�{������'COM.'�r��n����
    begin
      --select count(*) into v_ErrCnt from Send_Nagra where CompCode=x_CompCode and 
      select count(*) into v_ErrCnt from COM.Send_Nagra where CompCode=x_CompCode and 
	UpdTime>=v_RptDate and UpdTime<v_RptDate+1 and HIGH_LEVEL_CMD_ID=v_CmdId AND Cmd_Status='E';
    exception
      when others then
	v_RetCode := -2;
	v_RetMsg := '���ӫ��O���Ѧ��Ʀ����~: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;

    -- �p����O�`��, ���\�v, ���Ѳv
    v_TotalCnt := v_OkCnt + v_ErrCnt;			-- ���O�`��
    if v_TotalCnt = 0 then
      v_OkRate := 0;					-- ���\�v
      v_ErrRate := 0;					-- ���Ѳv
    else
      v_OkRate := round(v_OkCnt*100/v_TotalCnt, 2);	-- �ܤp���I���|�ˤ��J
      v_ErrRate := 100 - v_OkRate;
    end if;

    -- �O���έp���G��SO510
    delete SO510 where CompCode=x_CompCode and RptDate=v_RptDate; -- ���R���P���q�P�@�Ѫ��έp���
    begin
      insert into SO510 (RptDate, CompCode, ModeType, OkCnt, ErrCnt, TotalCnt, OkRate, ErrRate) values
	(v_RptDate, x_CompCode, v_CmdId, v_OkCnt, v_ErrCnt, v_TotalCnt, v_OkRate, v_ErrRate);
    exception
      when others then
	v_RetCode := -3;
	v_RetMsg := '�O���έp���G��SO510�����~: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;


    -- �O�����O������~���Ӹ�Ʃ�SO510A		2003.11.03
    -- ���]Send_Nagra�ɮ׬O�b�U��Ʈw��Com��ưϤ�, �Y���O, �h�{������'COM.'�r��n����
    delete SO510A where CompCode=x_CompCode and RptDate=v_RptDate; -- ���R���P���q�P�@�Ѫ��έp���
    begin
      insert into SO510A (RPTDATE, COMPCODE, MODETYPE, SEQNO, ICC_NO, STB_NO, OPERATOR, UPDTIME, ERR_CODE, ERR_MSG)
	(select trunc(UPDTIME), COMPCODE, HIGH_LEVEL_CMD_ID, SEQNO, ICC_NO, STB_NO, OPERATOR, UPDTIME, 
	ERR_CODE, ERR_MSG from COM.Send_Nagra where CompCode=x_CompCode and UpdTime>=v_RptDate and UpdTime<v_RptDate+1 and 
	HIGH_LEVEL_CMD_ID=v_CmdId AND Cmd_Status='E');
	--from Send_Nagra where CompCode=x_CompCode and UpdTime>=v_RptDate and UpdTime<v_RptDate+1 and 
    exception
      when others then
	v_RetCode := -5;
	v_RetMsg := '�O���έp���G��SO510A�����~: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;


    -- log�@���O����Night-run�O����SO030A
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, 0, '�έp����');
    commit;

  end loop;		-- Loop�C�@���q


  if v_CompCnt = 0 then -- ���Ҥ���
    v_RetCode := -4;
    v_RetMsg := 'Night-run�Ѽ���SO501A�L���';
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
      TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
      values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
      trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
  else
    DBMS_OUTPUT.PUT_LINE('�έp����');
  end if;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    v_RetCode := -99;
    v_RetMsg := '��L���~: '||SQLERRM;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
      TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
      values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
      trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
end;
/
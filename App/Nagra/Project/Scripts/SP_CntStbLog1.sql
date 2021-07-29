/*
@C:\Gird\V300\Script\SP_CntStbLog1

set serveroutput on
exec SP_CntStbLog1(99, 'A1', '20030201', '20030218', 0);

--exec SP_CntStbLog1(99, 'A1', '20030201', '20030218', 1);

  �έp�Y�t�Υx��Y�q�ɶ���, �}�����������Y�ذ������O������, ���\�v, ���Ѳv
  �ɦW: SP_CntStbLog1.sql

  IN	p_CompCode number: ���q�O
	p_CmdId varchar2: �������O�N��, 
		'A1'=�}��, 'A2'=����, 'A3'=����, 'A4'=�_��, 'E1'=�]�w�ˤl�K�X
	p_D1 varchar2: �έp�_�l���, ���'YYYYMMDD'
	p_D2 varchar2: �έp�I����, ���'YYYYMMDD'
	p_Option number: 0=�����_������d��, �έp�Ҧ�log���. 1=�̤���d��έp

  �`�N:
  1. ����login�ܦU�aSO��ưϤ����榹�{���Y�i, ���ݪ`�N���q�O�N�X�n�ǤJ
  2. �ݱҰ�SQL*Plus����ܶ}��: set serveroutput on
  3. ���]Send_Nagra�ɮ׬O�b�U��Ʈw��Com��ưϤ�, �Y���O, �h�{������'COM.'�r��n����

  By: Pierre
  Date: 2003.02.19
	2003.02.20 �[: �]�w�ˤl�K�X(���OE1)���έp
	2003.03.13 �[: �h�@�Ѽ�"���q�O"

*/

create or replace procedure SP_CntStbLog1(p_CompCode number, p_CmdId varchar2, p_D1 varchar2, p_D2 varchar2,
	p_Option number)
as
  v_CmdName varchar2(20);

  v_CompCode number;
  v_CompName varchar2(20);

  v_OkCnt number := 0;
  v_ErrCnt number := 0;
  v_OkRate number := 0;
  v_ErrRate number := 0;
  v_TotalCnt number := 0;

  v_Tmp1 varchar2(20);
  v_Tmp2 varchar2(20);

begin
  if p_Option!=0 and p_Option!=1 then
    DBMS_OUTPUT.PUT_LINE('�Ѽ�p_Option�Ȼݬ�0��1');
    return;
  end if;

  -- ���������O�W��
  if p_CmdId = 'A1' then
    v_CmdName := '�}��';
  elsif p_CmdId = 'A2' then
    v_CmdName := '����';
  elsif p_CmdId = 'A3' then
    v_CmdName := '����';
  elsif p_CmdId = 'A4' then
    v_CmdName := '�_��';
  elsif p_CmdId = 'E1' then
    v_CmdName := '�]�w�ˤl�K�X';
  else
    DBMS_OUTPUT.PUT_LINE('�L���������O�N��: '||p_CmdId);
    return;
  end if;

  -- �����q�O�P�W��
  v_CompCode := p_CompCode;
  begin
    select Description into v_CompName from CD039 where CodeNo=v_CompCode;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('CD039�L�����q�O�N�X: '||v_CompCode);
      return;
  end;

  -- ���ӫ��O���\���� (��: ����N�ন����r��)
  begin
    if p_Option=0 then
      select count(*) into v_OkCnt from SOAC0202 where CompCode=v_CompCode and 
	ModeType=p_CmdId;
    else
      v_Tmp1 := GiPackage.ED2CDString(to_date(p_D1,'YYYYMMDD'));
      v_Tmp2 := GiPackage.ED2CDString(to_date(p_D2,'YYYYMMDD'))||' 23:59:59';
      select count(*) into v_OkCnt from SOAC0202 where CompCode=v_CompCode and 
	UpdTime>=v_Tmp1 and UpdTime<=v_Tmp2 and ModeType=p_CmdId;
    end if;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('���ӫ��O���\���Ʀ����~: ');
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      return;
  end;

  -- ���ӫ��O���Ѧ���
  -- ���]Send_Nagra�ɮ׬O�b�U��Ʈw��Com��ưϤ�, �Y���O, �h�{������'COM.'�r��n����
  begin
    if p_Option=0 then
      select count(*) into v_ErrCnt from COM.Send_Nagra where CompCode=v_CompCode and 
	HIGH_LEVEL_CMD_ID=p_CmdId AND Cmd_Status='E';
    else
      select count(*) into v_ErrCnt from COM.Send_Nagra where CompCode=v_CompCode and 
	UpdTime>=to_date(p_D1,'YYYYMMDD') and UpdTime<to_date(p_D2,'YYYYMMDD')+1 and 
	HIGH_LEVEL_CMD_ID=p_CmdId AND Cmd_Status='E';
    end if;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('���ӫ��O���Ѧ��Ʀ����~: ');
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      return;
  end;

  -- �p����O�`��, ���\�v, ���Ѳv
  v_TotalCnt := v_OkCnt + v_ErrCnt;			-- ���O�`��
  if v_TotalCnt = 0 then
    v_OkRate := 0;					-- ���\�v
    v_ErrRate := 0;					-- ���Ѳv
  else
    v_OkRate := round(v_OkCnt*100/v_TotalCnt, 2);
    v_ErrRate := 100 - v_OkRate;
  end if;

  -- ��ܵ��G
  DBMS_OUTPUT.PUT_LINE('*** �������O�έp���G ***');
  if p_Option=0 then
    DBMS_OUTPUT.PUT_LINE('�έp����: �q�}�llog��{�b '||to_char(sysdate,'YYYY/MM/DD HH24:MI:SS'));
  else
    DBMS_OUTPUT.PUT_LINE('�έp����: '||p_D1||' �� '||p_D2);
  end if;
  DBMS_OUTPUT.PUT_LINE('���O�N��: '||p_CmdId||'  ('||v_CmdName||')');
  DBMS_OUTPUT.PUT_LINE('���q�O  : '||v_CompCode||', ���q�W��: '||v_CompName);
  DBMS_OUTPUT.PUT_LINE('���\����: '||v_OkCnt);
  DBMS_OUTPUT.PUT_LINE('���Ѧ���: '||v_ErrCnt);
  DBMS_OUTPUT.PUT_LINE('���O�`��: '||v_TotalCnt);
  DBMS_OUTPUT.PUT_LINE('���\�v  : '||v_OkRate);
  DBMS_OUTPUT.PUT_LINE('���Ѳv  : '||v_ErrRate);

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
end;
/
/*
  -- compile�P���ի��O
  @c:\emc_script\SP_FillChEvenCode
  set serveroutput on
  variable p_RetCode number
  variable p_RetMsg varchar2(500)
  exec  SP_FillChEvenCode(:p_RetCode,:p_RetMsg);
  print P_Retcode
  print p_RetMsg

  �N���O��ƪ��Ƶ��椤�W�D���M��]���X, ��J���������M��]�N�X��줤.
  �ɦW: SP_FillEvenCode.sql

  ����: V4.0�q92.08���_�󦬶O����ɥ[�W�W�D���M��]���(ChEvenCode, ChEvenName), �W�D���M�ɷ|�N
	��]��J���G���. ���ª����W�D���M�{���|�N��]��ƶ�b���O��ƪ��Ƶ���, �G�ݼg�@��{���N
	�¸�ƪ��Ƶ��椤�W�D���M��]���X, ��J�s��줤, �H�K����έp���R.

  �y�{:
  Loop �Ҧ��Ƶ������e�����O���(SO033, SO034, SO035)
	. ���X�W�D���M��]�r��: ��"���M��] : "�P", �춵�� :"�������r��
	. ���D���M��]�r��ҹ�����CD060���N�X
	. �Y���N�X, �h��s�ӵ����O��ƪ��W�D���M��]���
	. �Y�䤣��N�X, �hError count�[1


  OUT	p_RetCode number: ���G�N�X
	p_RetMsg VARCHAR2�G ���G�T��, �I�s���ܼƦܤֻ�300 bytes

        >0: ��Ƨ帹, ��ܥ��`����
	-1: p_RetMsg='�Ѽƿ��~'
        -99�Gp_RetMsg='��L���~'

  By: Pierre
  Date: 2002.08.11
*/
create or replace procedure SP_FillChEvenCode(p_RetCode OUT NUMBER, P_RetMsg OUT Varchar2)
is
--  p_RetCode number;
--  p_RetMsg varchar2(500);

  v_StartExecTime	date;
  v_EndExecTime		date;
  c39			char(1) := chr(39);
  v_SQL			varchar2(1000);
  v_ErrorCnt		number := 0;
  v_OkCnt		number := 0;
  v_ChEvenCode		number;
  v_ChEvenName		varchar2(20);

  s1 varchar2(100);
  s2 varchar2(100);
  x1 number;
  x2 number;
  len1 number;
  len2 number;

  cursor cc33 is 
    select RowId, Note from SO033 where Note is not null;
  cursor cc34 is 
    select RowId, Note from SO034 where Note is not null;
  cursor cc35 is 
    select RowId, Note from SO035 where Note is not null;


begin
  v_StartExecTime := sysdate;		-- �}�l����ɶ�

  s1 := '���M��] : ';
  s2 := ', �춵�� :';
  len1 := lengthb(s1);
  len2 := lengthb(s2);

  -- �B�zSO033
  for cr1 in cc33 loop
    -- ���X�W�D���M��]�r��: ��"���M��] : "�P", �춵�� :"�������r��
    x1 := instrb(cr1.Note, s1);
    if x1 > 0 then		-- ���]�ts1�r��, �~���U��
      ---** DBMS_OUTPUT.PUT_LINE('x1='||x1);
      x2 := instrb(cr1.Note, s2);
    else
      x2 := 0;
    end if;
    if x2 > x1 then
      ---** DBMS_OUTPUT.PUT_LINE('x2='||x2);
      v_ChEvenName := trim(substrb(cr1.Note, x1+len1, x2-x1-len1));
      ---** DBMS_OUTPUT.PUT_LINE('v_ChEvenName=['||v_ChEvenName||']');
    else
      v_ChEvenName := null;
    end if;

    if v_ChEvenName is not null then
      begin
	-- ���D���M��]�r��ҹ�����CD060���N�X
        select CodeNo into v_ChEvenCode from CD060 where Description = trim(v_ChEvenName);

	-- �Y���N�X, �h��s�ӵ����O��ƪ��W�D���M��]���
	update SO033 set ChEvenCode=v_ChEvenCode, ChEvenName=v_ChEvenName where RowId=cr1.RowId and ChEvenCode is null;
	if SQL%rowcount > 0 then
	  v_OkCnt := v_OkCnt + 1;
	end if;
      exception
	when others then
	  v_ErrorCnt := v_ErrorCnt + 1;
      end;
    end if;
  end loop;


  -- �B�zSO034
  for cr1 in cc34 loop
    -- ���X�W�D���M��]�r��: ��"���M��] : "�P", �춵�� :"�������r��
    x1 := instrb(cr1.Note, s1);
    if x1 > 0 then		-- ���]�ts1�r��, �~���U��
      x2 := instrb(cr1.Note, s2);
    else
      x2 := 0;
    end if;
    if x2 > x1 then
      v_ChEvenName := trim(substrb(cr1.Note, x1+len1, x2-x1-len1));
    else
      v_ChEvenName := null;
    end if;
    if v_ChEvenName is not null then
      begin
	-- ���D���M��]�r��ҹ�����CD060���N�X
        select CodeNo into v_ChEvenCode from CD060 where Description = trim(v_ChEvenName);

	-- �Y���N�X, �h��s�ӵ����O��ƪ��W�D���M��]���
	update SO034 set ChEvenCode=v_ChEvenCode, ChEvenName=v_ChEvenName where RowId=cr1.RowId and ChEvenCode is null;
	if SQL%rowcount > 0 then
	  v_OkCnt := v_OkCnt + 1;
	end if;
      exception
	when others then
	  v_ErrorCnt := v_ErrorCnt + 1;
      end;
    end if;
  end loop;


  -- �B�zSO035
  for cr1 in cc35 loop
    -- ���X�W�D���M��]�r��: ��"���M��] : "�P", �춵�� :"�������r��
    x1 := instrb(cr1.Note, s1);
    if x1 > 0 then		-- ���]�ts1�r��, �~���U��
      x2 := instrb(cr1.Note, s2);
    else
      x2 := 0;
    end if;
    if x2 > x1 then
      v_ChEvenName := trim(substrb(cr1.Note, x1+len1, x2-x1-len1));
    else
      v_ChEvenName := null;
    end if;
    if v_ChEvenName is not null then
      begin
	-- ���D���M��]�r��ҹ�����CD060���N�X
        select CodeNo into v_ChEvenCode from CD060 where Description = trim(v_ChEvenName);

	-- �Y���N�X, �h��s�ӵ����O��ƪ��W�D���M��]���
	update SO035 set ChEvenCode=v_ChEvenCode, ChEvenName=v_ChEvenName where RowId=cr1.RowId and ChEvenCode is null;
	if SQL%rowcount > 0 then
	  v_OkCnt := v_OkCnt + 1;
	end if;
      exception
	when others then
	  v_ErrorCnt := v_ErrorCnt + 1;
      end;
    end if;
  end loop;

  v_EndExecTime := sysdate;		-- ��������ɶ�
  p_RetMsg := '�վ㧹��, �@��O'||to_char(trunc(86400*(v_EndExecTime-v_StartExecTime)))||'��, �վ㵧��: '||v_OkCnt;
  p_RetCode := 0;
  commit;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    p_RetMsg := '��L���~: '||SQLERRM;
    p_RetCode := -99;
end;
/
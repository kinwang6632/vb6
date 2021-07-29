/*
@C:\Transform\EMC_Transform\EMC_Check\CntNewRcd.sql

set serveroutput on
variable nn number
variable msg varchar2(200)
exec :nn := CntNewRcd(:msg);

print nn
print msg
  
  �p���ഫ�N�X��"���ɫ�"���U�ӥN�X��, ��U�ӽվ�����ɤ�����Ƽƥ�.
  �`�N: 1. ���{���󥿦����Ƶ{��"��"����
	2. ���榹�{���e, �]�n���إߦU����ɪ�index

  1. �p��̾ڬ�Code_Map1, ���c�����U�C���:
	CodeTable	varchar2(30)
	DataTable	varchar2(30)
	DataColumn	varchar2(30)

  2. �p�⵲�G�x�s��Code_Map2, ���c�����U�C���:
	CodeTable	varchar2(30)
	DataTable	varchar2(30)
	DataColumn	varchar2(30)
	OldCodeNo	number(3)
	OldRcdCnt	number(8)
	NewCodeNo	number(3)
	NewRcdCnt	number(8)	

  OUT p_RetMsg VARCHAR2�G���G�T��
  Return NUMBER�G���G�N�X
        0: ���`����
	-1: p_RetMsg='�Ѽƿ��~'
        -99�Gp_RetMsg='��L���~'

  By: Pierre
  Date: 2002.11.22 ��Z
	2002.11.25 �[�W�良�ഫ�N�X���p��
	2002.11.26 ��: �[�W�O��"�s"�N�X���W�� (����, ��b���ɫe�N�έp�F)

*/

create or replace function CntNewRcd(p_RetMsg out varchar2) return number
as
  v_SQL1 varchar2(2000);
  v_SQL2 varchar2(2000);
  v_SQL3 varchar2(2000);
  v_CodeNo number;
  v_RcdCnt number;
  v_NewCodeNo number;
  v_OldName varchar2(20);
  v_NewName varchar2(20);

  TYPE CurTyp IS REF CURSOR;  --�ۭqcursor���A, �ѷs��dynamic SQL
  v_DyCursor1 CurTyp;
  v_DyCursor2 CurTyp;
  v_DyCursor3 CurTyp;

  cursor cc1 is
	select * from Code_Map1;

begin

  -- �}�l�p��
  for cr1 in cc1 loop	-- loop Code_Map1 ���쪺�C�@�����
    v_SQL1 := 'select CodeNo, Description from '||cr1.CodeTable||' order by CodeNo';
    begin
      OPEN v_DyCursor1 FOR v_SQL1;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL1���~: ' || v_SQL1;
        return -2;
    end;

    loop       -- loop cr1.CodeTable �����쪺�C�@����ƥN�X
      v_CodeNo := null;
      v_NewName := null;

      FETCH v_DyCursor1 INTO v_CodeNo, v_NewName;
      EXIT WHEN v_DyCursor1%NOTFOUND;
      -- DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

      -- ����CodeTable��DataTable��DataColumn�����CodeNo����Ƽ�
      if v_CodeNo is not null then
	-- ����Ƽ� v_RcdCnt
	v_SQL2 := 'select count(*) from '||cr1.DataTable||' where '||cr1.DataColumn||'='||v_CodeNo;
	begin
	  OPEN v_DyCursor2 FOR v_SQL2;
	exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL2���~: ' || v_SQL2;
	  return -5;
	end;

	v_RcdCnt := 0;
	FETCH v_DyCursor2 INTO v_RcdCnt;  
	-- ��s�ܵ��G��
	update Code_Map2 set NewRcdCnt=nvl(NewRcdCnt,0)+v_RcdCnt where CodeTable=cr1.CodeTable and 
		DataTable=cr1.DataTable and DataColumn=cr1.DataColumn and NewCodeNo=v_CodeNo;

	-- �Y�䤣��, �h�i��O�s�¥N�X�@��, �G��s��
	if SQL%ROWCOUNT = 0 then
	  update Code_Map2 set NewRcdCnt=nvl(NewRcdCnt,0)+v_RcdCnt, NewCodeNo=v_CodeNo, NewName=v_NewName
		where CodeTable=cr1.CodeTable and DataTable=cr1.DataTable and DataColumn=cr1.DataColumn 
		and OldCodeNo=v_CodeNo and (NewCodeNo=v_CodeNo or NewCodeNo is null);
	  if SQL%ROWCOUNT = 0 then	-- �Y�٧䤣��, �h�A�s�W�@�� (�����o�ͦ��ر���)
	    insert into Code_Map2 (CodeTable, DataTable, DataColumn, OldCodeNo, OldRcdCnt, NewCodeNo, NewRcdCnt, NewName) 
	      values (cr1.CodeTable, cr1.DataTable, cr1.DataColumn, null, 0, v_CodeNo, v_RcdCnt, v_NewName);
	  end if;
	end if;
	CLOSE v_DyCursor2;
      end if;

    end loop;   -- Cursor �j��
    CLOSE v_DyCursor1;  -- close cursor

  end loop;

   commit;
  p_RetMsg := '���`����';
  return 0;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    p_RetMsg := '��L���~';
    return -99;
end;
/

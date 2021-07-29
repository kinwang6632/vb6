/*
@C:\Transform\EMC_Transform\EMC_Check\CntOldRcd.sql

set serveroutput on
variable nn number
variable msg varchar2(200)
exec :nn := CntOldRcd(:msg);

print nn
print msg
  
  �p����ഫ�N�X��"���ɫe"���U�ӥN�X��, ��{���U�Ӹ���ɤ�����Ƽƥ�.
  �`�N: 1. ���{���󥿦����Ƶ{��"�e"����
	2. �������{����, �~�i�Hdrop������ɪ�index
	3. ���]�Ҧ��N�X��Ӫ�R�W��: CDxxx_MAP

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
  Date: 2002.11.22
	2002.11.25 ��: �[�W�N�X�����ഫ���B�z
	2002.11.26 ��: �[�W�O��"�s��"�N�X���W��
	2002.11.27 ��: �ɱj����ɤ����Ψ쪺�N�X, ���¥N�X���o�L�ӥN�X���B�z--
			�A�ܹ�Ӫ����¥N�X�W�٤ηs�N�X
	2002.12.03 ��: �Y��Ӫ��s�¥N�X���@�˪�, �h���{���A�[�W������Ӫ�
			"������"�s�³��@�˪�(�s�N�X�W�ٶ�null).

*/

create or replace function CntOldRcd(p_RetMsg out varchar2) return number
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
  -- ���M�����G��
  delete from Code_Map2;

  -- �}�l�p��
  for cr1 in cc1 loop	-- loop Code_Map1 ���쪺�C�@�����

    -- �ˬd����ɤ����Ψ쪺�N�X, ���¥N�X���o�L�ӥN�X, 
    -- �h�A�ˬd�ӥN�X����Ӫ�, �O�_���[�J��Ӹ��, �Y��, �h�H�ӹ�Ӫ��¦W�٤ηs�N�X���D,
    -- �Y�٧䤣��, �h��ܬ����`���, 
    -- �Y�����إN�X, �@�˭n�p�����ɤ�����Ƽ�, �ðO����==>�n�аݫȤ�p��B�z?
    v_SQL3 := 'select distinct '||cr1.DataColumn||' from '||cr1.DataTable||' minus (select CodeNo from '||
		cr1.CodeTable||')';
    begin
      OPEN v_DyCursor3 FOR v_SQL3;
    exception
      when others then
	rollback;
	p_RetMsg := 'SQL3���~: ' || v_SQL3;
      return -3;
    end;
    loop
      FETCH v_DyCursor3 INTO v_CodeNo;  
      EXIT WHEN v_DyCursor3%NOTFOUND;
      --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

      if v_CodeNo is not null then
        -- �A�ˬd�ӥN�X����Ӫ�, �O�_���[�J��Ӹ��, �Y��, �h�H�ӹ�Ӫ��¦W�٤ηs�N�X���D
        v_SQL2 := 'select OldName, NewCode from '||cr1.CodeTable||'_Map where OldCode='||v_CodeNo;
        begin
	  OPEN v_DyCursor2 FOR v_SQL2;
        exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL2���~: ' || v_SQL2;
	  return -7;
        end;
        v_NewCodeNo := null;
        v_OldName := null;
        FETCH v_DyCursor2 INTO v_OldName, v_NewCodeNo; 
        CLOSE v_DyCursor2;


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
	CLOSE v_DyCursor2;

	-- �s�W�ܵ��G��, ���ظ�ƥi��L�¥N�X�W��OldName ==> �n�аݫȤ�p��B�z?
	insert into Code_Map2 (CodeTable, DataTable, DataColumn, OldCodeNo, OldRcdCnt, OldName, NewCodeNo) values
	    (cr1.CodeTable, cr1.DataTable, cr1.DataColumn, v_CodeNo, v_RcdCnt, v_OldName, v_NewCodeNo);
      end if;
    end loop;
    CLOSE v_DyCursor3;

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
      FETCH v_DyCursor1 INTO v_CodeNo, v_OldName;
      EXIT WHEN v_DyCursor1%NOTFOUND;
      -- DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

      -- ����CodeTable��DataTable��DataColumn�����CodeNo����Ƽ�
      if v_CodeNo is not null then
	-- ���������s�N�X�� v_NewCodeNo, ���i��Onull(�L�����s�N�X)
	-- ���]�Ҧ��N�X��Ӫ�R�W��: CDxxx_MAP
	-- ���]�Ҧ��s�N�X�Ȧs�ɩR�W��: CDxxx_NEW
	v_NewCodeNo := null;
	v_NewName := null;
	v_SQL3 := 'select NewCode, NewName from '||cr1.CodeTable||'_Map where OldCode='||v_CodeNo;
	begin
	  OPEN v_DyCursor3 FOR v_SQL3;
	exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL3���~: ' || v_SQL3;
	  return -4;
	end;
	FETCH v_DyCursor3 INTO v_NewCodeNo, v_NewName;  
	CLOSE v_DyCursor3;

	-- ����Ƽ� v_RcdCnt
	v_SQL2 := 'select count(*) from '||cr1.DataTable||' where '||cr1.DataColumn||'='||v_CodeNo;
	begin
	  OPEN v_DyCursor2 FOR v_SQL2;
	exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL2���~: ' || v_SQL2;
	  return -6;
	end;

	v_RcdCnt := 0;
	FETCH v_DyCursor2 INTO v_RcdCnt;  

	-- �s�W�ܵ��G��
	if v_NewCodeNo is null then	-- �Yv_NewCodeNo����, �h��ܸӥN�X�����ഫ
	  v_NewCodeNo := v_CodeNo;
	  v_NewName := null;		-- 2002.12.03 ���ة��Ӫ����C���N�X, �N��s�N�X�W�ٶ�null, �H�K�ˬd.
	  -- v_NewName := v_OldName;
	end if;
	insert into Code_Map2 (CodeTable, DataTable, DataColumn, OldCodeNo, OldRcdCnt, OldName, NewCodeNo, NewName) values
	    (cr1.CodeTable, cr1.DataTable, cr1.DataColumn, v_CodeNo, v_RcdCnt, v_OldName, v_NewCodeNo, v_NewName);

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

/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD020.SQL
  EXEC SP_UpdData_CD020
  ��s�ϥΨ�I�O�N�@�N�X��(CD020)�������ɮ�: 
	SO001(PWCode, PWName)
	SO044(PWCode)
	CD020(CodeNo,Description)


  By: Jackal
  Date: 2002.11.22
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD020
  AS
--declare
  v_ChgCnt number := 0;
  v_ServiceType  char(1) := 'C';


  v_OldCode number;
  v_OldName varchar2(20);
  v_NewCode number;
  v_NewName varchar2(20);
  v_Mark    number;

  cursor cc2 is
    select * from CD020_New order by NewCode;

  cursor ccCD020 is
    select Rowid,CodeNo from CD020 where CodeNo is not null;

  cursor ccSO001 is
    select Rowid,PWCode from SO001 where PWCode is not null;

  cursor ccSO044 is
    select Rowid,PWCode from SO044 where PWCode is not null;

begin

  DBMS_OUTPUT.PUT_LINE('*** ����: UpdData_CD020.SQL');
--�R��CD020_Map �ɤ����ʪ����
  DELETE CD020_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--��s CD020(CodeNo,Description)
  for crCD020 in ccCD020 loop  --��X�Ҧ����N�X
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD020_Map where OldCode=crCD020.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD020 WHERE Rowid=crCD020.Rowid;
        else
          update CD020 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD020.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('���~1: CD020�N�X�L�k��s��CD020.CodeNo, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
	  RETURN;
      END;
    exception
      when others then
	NULL;
    end;
  end loop;


--�s�W�¥N�X�ɨS�����s�N�X���
  for cr2 in cc2 loop
    begin
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD020 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into CD020(CODENO,DESCRIPTION,SERVICETYPE,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('���~2: CD020�N�X�L�k�s�W��CD020.CodeNo, �s�N�X='||cr2.NewCode);
	    NULL;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('���~3: CD020�N�X�L�k��s��CD020.CodeNo, �s�N�X='||cr2.NewCode);
  	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	ROLLBACK;
	RETURN;
    end;
  end loop;


--��s SO001(PWCode, PWName)
  for crSO001 in ccSO001 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD020_Map where OldCode=crSO001.PWCode;      
      begin
        update SO001 set PWCode=v_NewCode, PWName=v_NewName where Rowid=crSO001.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~4: CD020�N�X�L�k��s��SO001.PWCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO044(PWCode)
  for crSO044 in ccSO044 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD020_Map where OldCode=crSO044.PWCode;      
      begin
        update SO044 set PWCode=v_NewCode where Rowid=crSO044.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~5: CD020�N�X�L�k��s��SO044.PWCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


  -- commit;
  DBMS_OUTPUT.PUT_LINE('*** OK: ��s�ϥΨ�I�O�N�@�N�X��(CD020)�������ɮ� ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


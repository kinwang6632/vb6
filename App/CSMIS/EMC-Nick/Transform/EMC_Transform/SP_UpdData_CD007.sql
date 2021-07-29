/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD007.SQL
  EXEC SP_UpdData_CD007
  ��s�ϥΨ찱������O�N�X��(CD007)�������ɮ�: 
	SO002(PRCode, PRName) , 
	SO009(PRCode, PRName) , 
	SO073(PRCode) , 


  By: Jackal
  Date: 2002.11.22
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD007
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
    select * from CD007_New order by NewCode;

  cursor ccCD007 is
    select Rowid,CodeNo from CD007 where CodeNo is not null;

  cursor ccSO009 is
    select Rowid,PRCode from SO009 where PRCode is not null;

  cursor ccSO073 is
    select Rowid,PRCode from SO073 where PRCode is not null;

begin

  DBMS_OUTPUT.PUT_LINE('*** ����: UpdData_CD007.SQL');
--�R��CD007_Map �ɤ����ʪ����
  DELETE CD007_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--��s CD007(CodeNo,Description)
  for crCD007 in ccCD007 loop  --��X�Ҧ����N�X
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD007_Map where OldCode=crCD007.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD007 WHERE Rowid=crCD007.Rowid;
        else
          update CD007 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD007.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('���~1: CD007�N�X�L�k��s��CD007.CodeNo, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD007 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into cd007(CODENO,DESCRIPTION,SERVICETYPE,PRINTWONOW,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,0,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('���~2: CD007�N�X�L�k�s�W��CD007.CodeNo, �s�N�X='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('���~3: CD007�N�X�L�k��s��CD007.CodeNo, �s�N�X='||cr2.NewCode);
  	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    	ROLLBACK;
    	RETURN;
    end;
  end loop;



--��s SO009(PRCode, PRName)
  for crSO009 in ccSO009 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD007_Map where OldCode=crSO009.PRCode;      
      begin
        update SO009 set PRCode=v_NewCode, PRName=v_NewName where Rowid=crSO009.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~5: CD007�N�X�L�k��s��SO009.PRCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO073(PRCode)
  for crSO073 in ccSO073 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD007_Map where OldCode=crSO073.PRCode;      
      begin
        update SO073 set PRCode=v_NewCode where Rowid=crSO073.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~6: CD007�N�X�L�k��s��SO073.PRCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: ��s�ϥΨ찱������O�N�X��(CD007)�������ɮ� ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


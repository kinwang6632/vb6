/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD004.SQL
  EXEC SP_UpdData_CD004
  ��s�ϥΨ�Ȥ����O�N�X��(CD004)�������ɮ�: 
	CD004(CodeNo,Description)
	SO001(ClassCode1,ClassName1, ClassCode2,ClassName2, ClassCode3,ClassName3)
	CD019CD004(ClassCode)
	SO032(ClassCode)
	SO033(ClassCode)
	SO034(ClassCode)
	SO035(ClassCode)
	SO036(ClassCode)	
	SO044(ClassCode)
	SO019(ClassCode, ClassName)


  By: Jackal
  Date: 2002.11.21
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD004
  AS
--declare
  v_ChgCnt number := 0;
  v_ServiceType  char(1) := 'C';
  c39		char(1) := chr(39);

  v_OldCode number;
  v_OldName varchar2(20);
  v_NewCode number;
  v_NewName varchar2(20);
  v_Mark    number;

  cursor cc2 is
    select * from CD004_New order by NewCode;

  cursor ccCD004 is
    select Rowid,CodeNo from CD004 where CodeNo is not null;

  cursor ccSo001_1 is
    select Rowid,ClassCode1 from SO001 where ClassCode1 is not null;

  cursor ccSo001_2 is
    select Rowid,ClassCode2 from SO001 where ClassCode2 is not null;

  cursor ccSo001_3 is
    select Rowid,ClassCode3 from SO001 where ClassCode3 is not null;

  cursor ccCD019CD004 is
    select Rowid,ClassCode from CD019CD004 where ClassCode is not null;

  cursor ccSO032 is
    select Rowid,ClassCode from SO032 where ClassCode is not null;

  cursor ccSO033 is
    select Rowid,ClassCode from SO033 where ClassCode is not null;

  cursor ccSO034 is
    select Rowid,ClassCode from SO034 where ClassCode is not null;

  cursor ccSO035 is
    select Rowid,ClassCode from SO035 where ClassCode is not null;

  cursor ccSO036 is
    select Rowid,ClassCode from SO036 where ClassCode is not null;

  cursor ccSO044 is
    select Rowid,ClassCode from SO044 where ClassCode is not null;

  cursor ccSO019 is
    select Rowid,ClassCode from SO019 where ClassCode is not null;
begin

  DBMS_OUTPUT.PUT_LINE('*** ����: UpdData_CD004.SQL');
--�R��CD004_Map �ɤ����ʪ����
  DELETE CD004_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--��s CD004(CodeNo,Description)
  for crCD004 in ccCD004 loop  --��X�Ҧ����N�X
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD004_Map where OldCode=crCD004.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD004 WHERE Rowid=crCD004.Rowid;
        else
          update CD004 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD004.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('���~1: CD004�N�X�L�k��s��CD004.CodeNo, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD004 WHERE CodeNo=cr2.NewCode;
      if v_ChgCnt = 0 then
        begin
            insert into CD004(CODENO,DESCRIPTION,SERVICETYPE,STOPFLAG) 
              values(cr2.NewCode,cr2.NewName,v_ServiceType,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('���~2: CD004�N�X�L�k�s�W��CD004.CodeNo, �s�N�X='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('���~3: CD004�N�X�L�k��s��CD004.CodeNo, �s�N�X='||cr2.NewCode);
  	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
        ROLLBACK;
        RETURN;
    end;
  end loop;



--��s SO001(ClassCode1,ClassName1)
  for crSO001_1 in ccSO001_1 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO001_1.ClassCode1;      
      begin
        update SO001 set ClassCode1=v_NewCode, ClassName1=v_NewName where Rowid=crSO001_1.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~4: CD004�N�X�L�k��s��SO001.ClassCode1, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
   	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO001(ClassCode2,ClassName2)
  for crSO001_2 in ccSO001_2 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO001_2.ClassCode2;      
      begin
        update SO001 set ClassCode2=v_NewCode, ClassName2=v_NewName where Rowid=crSO001_2.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~5: CD004�N�X�L�k��s��SO001.ClassCode2, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
 	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO001(ClassCode3,ClassName3)
  for crSO001_3 in ccSO001_3 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO001_3.ClassCode3;      
      begin
        update SO001 set ClassCode3=v_NewCode, ClassName3=v_NewName where Rowid=crSO001_3.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~6: CD004�N�X�L�k��s��SO001.ClassCode3, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
   	  ROLLBACK;
          RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD019CD004(ClassCode)
  for crCD019CD004 in ccCD019CD004 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crCD019CD004.ClassCode;      
      begin
        update CD019CD004 set ClassCode=v_NewCode where Rowid=crCD019CD004.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~7: CD004�N�X�L�k��s��CD019CD004.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
   	  ROLLBACK;
    	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO032(ClassCode)
  for crSO032 in ccSO032 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO032.ClassCode;      
      begin
        update SO032 set ClassCode=v_NewCode where Rowid=crSO032.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~8: CD004�N�X�L�k��s��SO032.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
    	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO033(ClassCode)
  for crSO033 in ccSO033 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO033.ClassCode;      
      begin
        update SO033 set ClassCode=v_NewCode where Rowid=crSO033.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~9: CD004�N�X�L�k��s��SO033.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO034(ClassCode)
  for crSO034 in ccSO034 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO034.ClassCode;      
      begin
        update SO034 set ClassCode=v_NewCode where Rowid=crSO034.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~10: CD004�N�X�L�k��s��SO034.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO035(ClassCode)
  for crSO035 in ccSO035 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO035.ClassCode;      
      begin
        update SO035 set ClassCode=v_NewCode where Rowid=crSO035.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~11: CD004�N�X�L�k��s��SO035.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
   	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO036(ClassCode)
  for crSO036 in ccSO036 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO036.ClassCode;      
      begin
        update SO036 set ClassCode=v_NewCode where Rowid=crSO036.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~12: CD004�N�X�L�k��s��SO036.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO044(ClassCode)
  for crSO044 in ccSO044 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO044.ClassCode;      
      begin
        update SO044 set ClassCode=v_NewCode where Rowid=crSO044.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~13: CD004�N�X�L�k��s��SO044.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO019(ClassCode,ClassName)
  for crSO019 in ccSO019 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD004_Map where OldCode=crSO019.ClassCode;      
      begin
        update SO019 set ClassCode=v_NewCode, ClassName=v_NewName where Rowid=crSO019.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~14: CD004�N�X�L�k��s��SO019.ClassCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



  --commit;
  --rollback;
  DBMS_OUTPUT.PUT_LINE('*** OK: ��s�ϥΨ�Ȥ����O�N�X��(CD004)�������ɮ� ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


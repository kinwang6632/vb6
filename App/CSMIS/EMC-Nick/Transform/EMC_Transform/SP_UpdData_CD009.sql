/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD009.SQL
  EXEC SP_UpdData_CD009
  ��s�ϥΨ�C�餶�ХN�X��(CD009)�������ɮ�: 
	SO002(MediaCode,MediaName)
	SO004(MediaCode,MediaName)
	SO007(MediaCode,MediaName)
	SO055(MediaCode,MediaName, MediaCodeB,MediaNameB)
	SO006(MediaCode,MediaName)
	SO098(MediaCode,MediaName)
	SO105(MediaCode,MediaName)
	SO106(MediaCode,MediaName)
	SO503(MediaCode,MediaName)
	SO013(MediaCode,MediaName)
	CD009(CodeNo,Description) 


  By: Jackal
  Date: 2002.11.22
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD009
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
    select * from CD009_New order by NewCode;

  cursor ccCD009 is
    select Rowid,CodeNo from CD009 where CodeNo is not null;

  cursor ccSO002 is
    select Rowid,MediaCode from SO002 where MediaCode is not null;

  cursor ccSO004 is
    select Rowid,MediaCode from SO004 where MediaCode is not null;

  cursor ccSO007 is
    select Rowid,MediaCode from SO007 where MediaCode is not null;

  cursor ccSO055_1 is
    select Rowid,MediaCode from SO055 where MediaCode is not null;

  cursor ccSO055_2 is
    select Rowid,MediaCodeB from SO055 where MediaCodeB is not null;

  cursor ccSO006 is
    select Rowid,MediaCode from SO006 where MediaCode is not null;

  cursor ccSO098 is
    select Rowid,MediaCode from SO098 where MediaCode is not null;

  cursor ccSO105 is
    select Rowid,MediaCode from SO105 where MediaCode is not null;

  cursor ccSO106 is
    select Rowid,MediaCode from SO106 where MediaCode is not null;

  cursor ccSO503 is
    select Rowid,MediaCode from SO503 where MediaCode is not null;

  cursor ccSO013 is
    select Rowid,MediaCode from SO013 where MediaCode is not null;
begin

  DBMS_OUTPUT.PUT_LINE('*** ����: UpdData_CD009.SQL');
--�R��CD009_Map �ɤ����ʪ����
  DELETE CD009_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--��s CD009(CodeNo,Description)
  for crCD009 in ccCD009 loop  --��X�Ҧ����N�X
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD009_Map where OldCode=crCD009.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD009 WHERE Rowid=crCD009.Rowid;
        else
          update CD009 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD009.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('���~1: CD009�N�X�L�k��s��CD009.CodeNo, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD009 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into CD009(CODENO,DESCRIPTION,SERVICETYPE,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('���~2: CD009�N�X�L�k�s�W��CD009.CodeNo, �s�N�X='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('���~3: CD009�N�X�L�k��s��CD009.CodeNo, �s�N�X='||cr2.NewCode);
   	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	ROLLBACK;
  	RETURN;
    end;
  end loop;


--��s SO002(MediaCode,MediaName)
  for crSO002 in ccSO002 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO002.MediaCode;      
      begin
        update SO002 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO002.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~4: CD009�N�X�L�k��s��SO002.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO004(MediaCode,MediaName)
  for crSO004 in ccSO004 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO004.MediaCode;      
      begin
        update SO004 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO004.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~5: CD009�N�X�L�k��s��SO004.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO007(MediaCode,MediaName)
  for crSO007 in ccSO007 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO007.MediaCode;      
      begin
        update SO007 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO007.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~6: CD009�N�X�L�k��s��SO007.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO055(MediaCode,MediaName)
  for crSO055_1 in ccSO055_1 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO055_1.MediaCode;      
      begin
        update SO055 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO055_1.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~7: CD009�N�X�L�k��s��SO055.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO055(MediaCodeB,MediaNameB)
  for crSO055_2 in ccSO055_2 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO055_2.MediaCodeB;      
      begin
        update SO055 set MediaCodeB=v_NewCode, MediaNameB=v_NewName where Rowid=crSO055_2.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~8: CD009�N�X�L�k��s��SO055.MediaCodeB, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO006(MediaCode,MediaName)
  for crSO006 in ccSO006 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO006.MediaCode;      
      begin
        update SO006 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO006.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~9: CD009�N�X�L�k��s��SO006.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO098(MediaCode,MediaName)
  for crSO098 in ccSO098 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO098.MediaCode;      
      begin
        update SO098 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO098.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~10: CD009�N�X�L�k��s��SO098.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO105(MediaCode,MediaName) 
  for crSO105 in ccSO105 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO105.MediaCode;      
      begin
        update SO105 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO105.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~11: CD009�N�X�L�k��s��SO105.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO106(MediaCode,MediaName)
  for crSO106 in ccSO106 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO106.MediaCode;      
      begin
        update SO106 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO106.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~12: CD009�N�X�L�k��s��SO106.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO503(MediaCode,MediaName)
  for crSO503 in ccSO503 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO503.MediaCode;      
      begin
        update SO503 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO503.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~13: CD009�N�X�L�k��s��SO503.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO013(MediaCode,MediaName)
  for crSO013 in ccSO013 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD009_Map where OldCode=crSO013.MediaCode;      
      begin
        update SO013 set MediaCode=v_NewCode, MediaName=v_NewName where Rowid=crSO013.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~14: CD009�N�X�L�k��s��SO013.MediaCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: ��s�ϥΨ�C�餶�ХN�X��(CD009)�������ɮ� ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


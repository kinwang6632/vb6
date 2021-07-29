/*

  @C:\Transform\EMC_Transform\SP_UpdData_CD019.SQL
  EXEC SP_UpdData_CD019;
  ��s�ϥΨ즬�O���إN�X��(CD019)�������ɮ�: 
	CD019CD004(CitemCode)
	CD019A(CitemCode)
	CD019SO017(CitemCode)
	CD024(CitemCode, CitemName)
	CD022(DefCitemCode, DefCitemName)
	CD022CD019(CitemCode)
	SO003(CitemCode, CitemName)
	SO005(CitemCode)
	SO032 & SO033 & SO034 & SO035 & SO036(CitemCode, CitemName)
	SOAC0102(CitemCode, CitemName)
	SO043(CitemCode)
	SO301A(CitemCode)
	SO301B(CitemCode)
	SO302A(CitemCode)
	SO302B(CitemCode)
	SO303(CitemCode)
	CD005 & CD006 & CD007(CitemCode1,CitemName1,CitemCode2,CitemName2,CitemCode3,CitemName3,CitemCode4,CitemName4,CitemCode5,CitemName5)
	CD019(CodeNo,Description ,ReturnCode, ReturnName) 
  By: Jackal
  Date: 2002.11.22
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD019
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
    select * from CD019_New order by NewCode;

  cursor ccCD019 is
    select Rowid,CodeNo from CD019 where CodeNo is not null;

  cursor ccCD019CD004 is
    select Rowid,CitemCode from CD019CD004 where CitemCode is not null;

  cursor ccCD019A is
    select Rowid,CitemCode from CD019A where CitemCode is not null;

  cursor ccCD019SO017 is
    select Rowid,CitemCode from CD019SO017 where CitemCode is not null;

  cursor ccCD024 is
    select Rowid,CitemCode from CD024 where CitemCode is not null;

  cursor ccCD022 is
    select Rowid,DefCitemCode from CD022 where DefCitemCode is not null;

  cursor ccCD022CD019 is
    select Rowid,CitemCode from CD022CD019 where CitemCode is not null;

  cursor ccSO003 is
    select Rowid,CitemCode from SO003 where CitemCode is not null;

  cursor ccSO005 is
    select Rowid,CitemCode from SO005 where CitemCode is not null;

  cursor ccSO032 is
    select Rowid,CitemCode from SO032 where CitemCode is not null;

  cursor ccSO033 is
    select Rowid,CitemCode from SO033 where CitemCode is not null;

  cursor ccSO034 is
    select Rowid,CitemCode from SO034 where CitemCode is not null;

  cursor ccSO035 is
    select Rowid,CitemCode from SO035 where CitemCode is not null;

  cursor ccSO036 is
    select Rowid,CitemCode from SO036 where CitemCode is not null;

  cursor ccSOAC0102 is
    select Rowid,CitemCode from SOAC0102 where CitemCode is not null;

  cursor ccSO043 is
    select Rowid,CitemCode from SO043 where CitemCode is not null;

  cursor ccSO301A is
    select Rowid,CitemCode from SO301A where CitemCode is not null;

  cursor ccSO301B is
    select Rowid,CitemCode from SO301B where CitemCode is not null;

  cursor ccSO302A is
    select Rowid,CitemCode from SO302A where CitemCode is not null;

  cursor ccSO302B is
    select Rowid,CitemCode from SO302B where CitemCode is not null;

  cursor ccSO303 is
    select Rowid,CitemCode from SO303 where CitemCode is not null;

  cursor ccCD005_1 is
    select Rowid,CitemCode1 from CD005 where CitemCode1 is not null;

  cursor ccCD005_2 is
    select Rowid,CitemCode2 from CD005 where CitemCode2 is not null;

  cursor ccCD005_3 is
    select Rowid,CitemCode3 from CD005 where CitemCode3 is not null;

  cursor ccCD005_4 is
    select Rowid,CitemCode4 from CD005 where CitemCode4 is not null;

  cursor ccCD005_5 is
    select Rowid,CitemCode5 from CD005 where CitemCode5 is not null;

  cursor ccCD006_1 is
    select Rowid,CitemCode1 from CD006 where CitemCode1 is not null;

  cursor ccCD006_2 is
    select Rowid,CitemCode2 from CD006 where CitemCode2 is not null;

  cursor ccCD006_3 is
    select Rowid,CitemCode3 from CD006 where CitemCode3 is not null;

  cursor ccCD006_4 is
    select Rowid,CitemCode4 from CD006 where CitemCode4 is not null;

  cursor ccCD006_5 is
    select Rowid,CitemCode5 from CD006 where CitemCode5 is not null;

  cursor ccCD007_1 is
    select Rowid,CitemCode1 from CD007 where CitemCode1 is not null;

  cursor ccCD007_2 is
    select Rowid,CitemCode2 from CD007 where CitemCode2 is not null;

  cursor ccCD007_3 is
    select Rowid,CitemCode3 from CD007 where CitemCode3 is not null;

  cursor ccCD007_4 is
    select Rowid,CitemCode4 from CD007 where CitemCode4 is not null;

  cursor ccCD007_5 is
    select Rowid,CitemCode5 from CD007 where CitemCode5 is not null;

  cursor ccCD019_1 is
    select Rowid,ReturnCode from CD019 where ReturnCode is not null;

begin

  DBMS_OUTPUT.PUT_LINE('*** ����: UpdData_CD019.SQL');
--�R��CD019_Map �ɤ����ʪ����
  DELETE CD019_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;



--��s CD019(CodeNo,Description)
  for crCD019 in ccCD019 loop  --��X�Ҧ����N�X
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD019_Map where OldCode=crCD019.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD019 WHERE Rowid=crCD019.Rowid;
        else
          update CD019 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD019.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('���~1: CD019�N�X�L�k��s��CD019.CodeNo, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      END;
    exception
      when others then
	NULL;
    end;
  end loop;


--��s CD019(ReturnCode,ReturnName)
  for crCD019_1 in ccCD019_1 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD019_1.ReturnCode;      
      begin
        update CD019 set ReturnCode=v_NewCode, ReturnName=v_NewName where Rowid=crCD019_1.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~2: CD019�N�X�L�k��s��CD019.ReturnCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--�s�W�¥N�X�ɨS�����s�N�X���
  for cr2 in cc2 loop
    SELECT COUNT(CODENO) INTO v_ChgCnt FROM CD019 WHERE CodeNo=cr2.NewCode;

    if v_ChgCnt = 0 then
      begin
        insert into CD019(CODENO,DESCRIPTION,SERVICETYPE,STOPFLAG,Sign,PeriodFlag,Amount,BackAmount) 
          values(cr2.NewCode,cr2.NewName,v_ServiceType,cr2.Mark,'+',1,0,0);
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~3: CD019�N�X�L�k�s�W��CD019.CodeNo, �s�N�X='||cr2.NewCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    end if; 	
  end loop;


--��s CD019CD004(CitemCode)
  for crCD019CD004 in ccCD019CD004 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD019CD004.CitemCode;      
      begin
        update CD019CD004 set CitemCode=v_NewCode where Rowid=crCD019CD004.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~5: CD019�N�X�L�k��s��CD019CD004.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD019A(CitemCode)
  for crCD019A in ccCD019A loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD019A.CitemCode;      
      begin
        update CD019A set CitemCode=v_NewCode where Rowid=crCD019A.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~6: CD019�N�X�L�k��s��CD019A.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD019SO017(CitemCode)
  for crCD019SO017 in ccCD019SO017 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD019SO017.CitemCode;      
      begin
        update CD019SO017 set CitemCode=v_NewCode where Rowid=crCD019SO017.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~7: CD019�N�X�L�k��s��CD019SO017.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD024(CitemCode,CitemName)
  for crCD024 in ccCD024 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD024.CitemCode;      
      begin
        update CD024 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crCD024.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~8: CD019�N�X�L�k��s��CD024.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD022(DefCitemCode,DefCitemName)
  for crCD022 in ccCD022 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD022.DefCitemCode;      
      begin
        update CD022 set DefCitemCode=v_NewCode,DefCitemName=v_NewName where Rowid=crCD022.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~9: CD019�N�X�L�k��s��CD022.DefCitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s CD022CD019(CitemCode)
  for crCD022CD019 in ccCD022CD019 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD022CD019.CitemCode;      
      begin
        update CD022CD019 set CitemCode=v_NewCode where Rowid=crCD022CD019.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~10: CD019�N�X�L�k��s��CD022CD019.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO003(CitemCode,CitemName)
  for crSO003 in ccSO003 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO003.CitemCode;      
      begin
        update SO003 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crSO003.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~11: CD019�N�X�L�k��s��SO003.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO005(CitemCode)
  for crSO005 in ccSO005 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO005.CitemCode;      
      begin
        update SO005 set CitemCode=v_NewCode where Rowid=crSO005.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~12: CD019�N�X�L�k��s��SO005.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO032(CitemCode,CitemName)
  for crSO032 in ccSO032 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO032.CitemCode;      
      begin
        update SO032 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crSO032.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~13: CD019�N�X�L�k��s��SO032.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO033(CitemCode,CitemName)
  for crSO033 in ccSO033 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO033.CitemCode;      
      begin
        update SO033 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crSO033.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~14: CD019�N�X�L�k��s��SO033.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO034(CitemCode,CitemName)
  for crSO034 in ccSO034 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO034.CitemCode;      
      begin
        update SO034 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crSO034.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~15: CD019�N�X�L�k��s��SO034.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO035(CitemCode,CitemName)
  for crSO035 in ccSO035 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO035.CitemCode;      
      begin
        update SO035 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crSO035.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~16: CD019�N�X�L�k��s��SO035.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO036(CitemCode,CitemName)
  for crSO036 in ccSO036 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO036.CitemCode;      
      begin
        update SO036 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crSO036.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~17: CD019�N�X�L�k��s��SO036.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SOAC0102(CitemCode,CitemName)
  for crSOAC0102 in ccSOAC0102 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSOAC0102.CitemCode;      
      begin
        update SOAC0102 set CitemCode=v_NewCode,CitemName=v_NewName where Rowid=crSOAC0102.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~18: CD019�N�X�L�k��s��SOAC0102.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO043(CitemCode)
  for crSO043 in ccSO043 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO043.CitemCode;      
      begin
        update SO043 set CitemCode=v_NewCode where Rowid=crSO043.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~19: CD019�N�X�L�k��s��SO043.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO301A(CitemCode)
  for crSO301A in ccSO301A loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO301A.CitemCode;      
      begin
        update SO301A set CitemCode=v_NewCode where Rowid=crSO301A.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~20: CD019�N�X�L�k��s��SO301A.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s SO301B(CitemCode)
  for crSO301B in ccSO301B loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO301B.CitemCode;      
      begin
        update SO301B set CitemCode=v_NewCode where Rowid=crSO301B.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~21: CD019�N�X�L�k��s��SO301B.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;




--��s SO302A(CitemCode)
  for crSO302A in ccSO302A loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO302A.CitemCode;      
      begin
        update SO302A set CitemCode=v_NewCode where Rowid=crSO302A.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~22: CD019�N�X�L�k��s��SO302A.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;




--��s SO302B(CitemCode)
  for crSO302B in ccSO302B loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO302B.CitemCode;      
      begin
        update SO302B set CitemCode=v_NewCode where Rowid=crSO302B.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~23: CD019�N�X�L�k��s��SO302B.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s SO303(CitemCode)
  for crSO303 in ccSO303 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO303.CitemCode;      
      begin
        update SO303 set CitemCode=v_NewCode where Rowid=crSO303.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~24: CD019�N�X�L�k��s��SO303.CitemCode, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s CD005(CitemCode1,CitemName1)
  for crCD005_1 in ccCD005_1 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD005_1.CitemCode1;      
      begin
        update CD005 set CitemCode1=v_NewCode,CitemName1=v_NewName where Rowid=crCD005_1.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~25: CD019�N�X�L�k��s��CD005.CitemCode1, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD005(CitemCode2,CitemName2)
  for crCD005_2 in ccCD005_2 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD005_2.CitemCode2;      
      begin
        update CD005 set CitemCode2=v_NewCode,CitemName2=v_NewName where Rowid=crCD005_2.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~26: CD019�N�X�L�k��s��CD005.CitemCode2, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s CD005(CitemCode3,CitemName3)
  for crCD005_3 in ccCD005_3 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD005_3.CitemCode3;      
      begin
        update CD005 set CitemCode3=v_NewCode,CitemName3=v_NewName where Rowid=crCD005_3.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~27: CD019�N�X�L�k��s��CD005.CitemCode3, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD005(CitemCode4,CitemName4)
  for crCD005_4 in ccCD005_4 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD005_4.CitemCode4;      
      begin
        update CD005 set CitemCode4=v_NewCode,CitemName4=v_NewName where Rowid=crCD005_4.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~28: CD019�N�X�L�k��s��CD005.CitemCode4, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD005(CitemCode5,CitemName5)
  for crCD005_5 in ccCD005_5 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD005_5.CitemCode5;      
      begin
        update CD005 set CitemCode5=v_NewCode,CitemName5=v_NewName where Rowid=crCD005_5.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~29: CD019�N�X�L�k��s��CD005.CitemCode5, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD006(CitemCode1,CitemName1)
  for crCD006_1 in ccCD006_1 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD006_1.CitemCode1;      
      begin
        update CD006 set CitemCode1=v_NewCode,CitemName1=v_NewName where Rowid=crCD006_1.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~30: CD019�N�X�L�k��s��CD006.CitemCode1, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD006(CitemCode2,CitemName2)
  for crCD006_2 in ccCD006_2 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD006_2.CitemCode2;      
      begin
        update CD006 set CitemCode2=v_NewCode,CitemName2=v_NewName where Rowid=crCD006_2.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~31: CD019�N�X�L�k��s��CD006.CitemCode2, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s CD006(CitemCode3,CitemName3)
  for crCD006_3 in ccCD006_3 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD006_3.CitemCode3;      
      begin
        update CD006 set CitemCode3=v_NewCode,CitemName3=v_NewName where Rowid=crCD006_3.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~32: CD019�N�X�L�k��s��CD006.CitemCode3, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD006(CitemCode4,CitemName4)
  for crCD006_4 in ccCD006_4 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD006_4.CitemCode4;      
      begin
        update CD006 set CitemCode4=v_NewCode,CitemName4=v_NewName where Rowid=crCD006_4.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~33: CD019�N�X�L�k��s��CD006.CitemCode4, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD006(CitemCode5,CitemName5)
  for crCD006_5 in ccCD006_5 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD006_5.CitemCode5;      
      begin
        update CD006 set CitemCode5=v_NewCode,CitemName5=v_NewName where Rowid=crCD006_5.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~34: CD019�N�X�L�k��s��CD006.CitemCode5, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;




--��s CD007(CitemCode1,CitemName1)
  for crCD007_1 in ccCD007_1 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD007_1.CitemCode1;      
      begin
        update CD007 set CitemCode1=v_NewCode,CitemName1=v_NewName where Rowid=crCD007_1.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~35: CD019�N�X�L�k��s��CD007.CitemCode1, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD007(CitemCode2,CitemName2)
  for crCD007_2 in ccCD007_2 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD007_2.CitemCode2;      
      begin
        update CD007 set CitemCode2=v_NewCode,CitemName2=v_NewName where Rowid=crCD007_2.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~36: CD019�N�X�L�k��s��CD007.CitemCode2, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--��s CD007(CitemCode3,CitemName3)
  for crCD007_3 in ccCD007_3 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD007_3.CitemCode3;      
      begin
        update CD007 set CitemCode3=v_NewCode,CitemName3=v_NewName where Rowid=crCD007_3.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~37: CD019�N�X�L�k��s��CD007.CitemCode3, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD007(CitemCode4,CitemName4)
  for crCD007_4 in ccCD007_4 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD007_4.CitemCode4;      
      begin
        update CD007 set CitemCode4=v_NewCode,CitemName4=v_NewName where Rowid=crCD007_4.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~38: CD019�N�X�L�k��s��CD007.CitemCode4, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
   	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--��s CD007(CitemCode5,CitemName5)
  for crCD007_5 in ccCD007_5 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crCD007_5.CitemCode5;      
      begin
        update CD007 set CitemCode5=v_NewCode,CitemName5=v_NewName where Rowid=crCD007_5.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('���~39: CD019�N�X�L�k��s��CD007.CitemCode5, �s�N�X='||v_NewCode||', �¥N�X='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: ��s�ϥΨ즬�O���اO�N�X��(CD019)�������ɮ� ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


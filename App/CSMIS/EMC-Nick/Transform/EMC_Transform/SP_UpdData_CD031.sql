/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD031.SQL
  EXEC SP_UpdData_CD031
  更新使用到收費方式代碼檔(CD031)的相關檔案: 
	SO032&SO033&SO034&SO035&SO036(CMCode, CMName)
	SO002(CMCode, CMName)
	SO044(CMCode)
	SO106(CMCode, CMName)
	CD031(CodeNo,Description)


  By: Jackal
  Date: 2002.11.22
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD031
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
    select * from CD031_New order by NewCode;

  cursor ccCD031 is
    select Rowid,CodeNo from CD031 where CodeNo is not null;

  cursor ccSO032 is
    select Rowid,CMCode from SO032 where CMCode is not null;

  cursor ccSO033 is
    select Rowid,CMCode from SO033 where CMCode is not null;

  cursor ccSO034 is
    select Rowid,CMCode from SO034 where CMCode is not null;

  cursor ccSO035 is
    select Rowid,CMCode from SO035 where CMCode is not null;

  cursor ccSO036 is
    select Rowid,CMCode from SO036 where CMCode is not null;

  cursor ccSO002 is
    select Rowid,CMCode from SO002 where CMCode is not null;

  cursor ccSO044 is
    select Rowid,CMCode from SO044 where CMCode is not null;

  cursor ccSO106 is
    select Rowid,CMCode from SO106 where CMCode is not null;

begin

  DBMS_OUTPUT.PUT_LINE('*** 執行: UpdData_CD031.SQL');
--刪除CD031_Map 檔不異動的資料
  DELETE CD031_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;



--更新 CD031(CodeNo,Description)
  for crCD031 in ccCD031 loop  --選出所有的代碼
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD031_Map where OldCode=crCD031.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD031 WHERE Rowid=crCD031.Rowid;
        else
          update CD031 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD031.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤1: CD031代碼無法更新至CD031.CodeNo, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      END;
    exception
      when others then
	NULL;
    end;
  end loop;


--新增舊代碼檔沒有的新代碼欄位
  for cr2 in cc2 loop
    begin

      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD031 WHERE CodeNo=cr2.NewCode;
      if v_ChgCnt = 0 then
        begin
          insert into CD031(CODENO,DESCRIPTION,SERVICETYPE,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,cr2.Mark);

        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('錯誤2: CD031代碼無法新增至CD031.CodeNo, 新代碼='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('錯誤3: CD031代碼無法更新至CD031.CodeNo, 新代碼='||cr2.NewCode);
   	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	ROLLBACK;
  	RETURN;
    end;
  end loop;



--更新 SO032(CMCode, CMName)
  for crSO032 in ccSO032 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO032.CMCode;      
      begin
        update SO032 set CMCode=v_NewCode, CMName=v_NewName where Rowid=crSO032.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤4: CD031代碼無法更新至SO032.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO033(CMCode, CMName)
  for crSO033 in ccSO033 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO033.CMCode;      
      begin
        update SO033 set CMCode=v_NewCode, CMName=v_NewName where Rowid=crSO033.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤5: CD031代碼無法更新至SO033.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO034(CMCode, CMName)
  for crSO034 in ccSO034 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO034.CMCode;      
      begin
        update SO034 set CMCode=v_NewCode, CMName=v_NewName where Rowid=crSO034.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤6: CD031代碼無法更新至SO034.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO035(CMCode, CMName)
  for crSO035 in ccSO035 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO035.CMCode;      
      begin
        update SO035 set CMCode=v_NewCode, CMName=v_NewName where Rowid=crSO035.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤7: CD031代碼無法更新至SO035.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO036(CMCode, CMName)
  for crSO036 in ccSO036 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO036.CMCode;      
      begin
        update SO036 set CMCode=v_NewCode, CMName=v_NewName where Rowid=crSO036.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤8: CD031代碼無法更新至SO036.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO002(CMCode, CMName)
  for crSO002 in ccSO002 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO002.CMCode;      
      begin
        update SO002 set CMCode=v_NewCode, CMName=v_NewName where Rowid=crSO002.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤9: CD031代碼無法更新至SO002.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO044(CMCode)
  for crSO044 in ccSO044 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO044.CMCode;      
      begin
        update SO044 set CMCode=v_NewCode where Rowid=crSO044.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤10: CD031代碼無法更新至SO044.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO106(CMCode, CMName)
  for crSO106 in ccSO106 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD031_Map where OldCode=crSO106.CMCode;      
      begin
        update SO106 set CMCode=v_NewCode, CMName=v_NewName where Rowid=crSO106.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤11: CD031代碼無法更新至SO106.CMCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: 更新使用到收費方式代碼檔(CD031)的相關檔案 ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


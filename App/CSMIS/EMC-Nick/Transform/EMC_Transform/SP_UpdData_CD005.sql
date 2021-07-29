/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD005.SQL
  EXEC SP_UpdData_CD005
  更新使用到裝機類別代碼檔(CD005)的相關檔案: 
	SO007(InstCode, InstName)
	SO071(InstCode)
	SO100(InstCode, InstName)
	SO104(InstCode)
	CD005(CodeNo,Description) 


  By: Jackal
  Date: 2002.11.22
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD005
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
    select * from CD005_New order by NewCode;

  cursor ccCD005 is
    select Rowid,CodeNo from CD005 where CodeNo is not null;

  cursor ccSO007 is
    select Rowid,InstCode from SO007 where InstCode is not null;

  cursor ccSO071 is
    select Rowid,InstCode from SO071 where InstCode is not null;

  cursor ccSO100 is
    select Rowid,InstCode from SO100 where InstCode is not null;

  cursor ccSO104 is
    select Rowid,InstCode from SO104 where InstCode is not null;

begin

  DBMS_OUTPUT.PUT_LINE('*** 執行: UpdData_CD005.SQL');
--刪除CD005_Map 檔不異動的資料
  DELETE CD005_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--更新 CD005(CodeNo,Description)
  for crCD005 in ccCD005 loop  --選出所有的代碼
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD005_Map where OldCode=crCD005.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD005 WHERE Rowid=crCD005.Rowid;
        else
          update CD005 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD005.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤1: CD005代碼無法更新至CD005.CodeNo, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD005 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into CD005(CODENO,DESCRIPTION,SERVICETYPE,PRINTWONOW,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,0,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('錯誤2: CD005代碼無法新增至CD005.CodeNo, 新代碼='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('錯誤3: CD005代碼無法更新至CD005.CodeNo, 新代碼='||cr2.NewCode);
  	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
   	ROLLBACK;
    	RETURN;
    end;
  end loop;



--更新 SO007(InstCode, InstName)
  for crSO007 in ccSO007 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD005_Map where OldCode=crSO007.InstCode;      
      begin
        update SO007 set InstCode=v_NewCode, InstName=v_NewName where Rowid=crSO007.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤4: CD005代碼無法更新至SO007.InstCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO071(InstCode)
  for crSO071 in ccSO071 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD005_Map where OldCode=crSO071.InstCode;      
      begin
        update SO071 set InstCode=v_NewCode where Rowid=crSO071.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤5: CD005代碼無法更新至SO071.InstCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO100(InstCode, InstName)
  for crSO100 in ccSO100 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD005_Map where OldCode=crSO100.InstCode;      
      begin
        update SO100 set InstCode=v_NewCode, InstName=v_NewName where Rowid=crSO100.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤6: CD005代碼無法更新至SO100.InstCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO104(InstCode)
  for crSO104 in ccSO104 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD005_Map where OldCode=crSO104.InstCode;      
      begin
        update SO104 set InstCode=v_NewCode where Rowid=crSO104.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤7: CD005代碼無法更新至SO104.InstCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: 更新使用到裝機類別代碼檔(CD005)的相關檔案 ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


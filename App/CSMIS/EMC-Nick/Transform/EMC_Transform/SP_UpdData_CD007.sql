/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD007.SQL
  EXEC SP_UpdData_CD007
  更新使用到停拆移機類別代碼檔(CD007)的相關檔案: 
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

  DBMS_OUTPUT.PUT_LINE('*** 執行: UpdData_CD007.SQL');
--刪除CD007_Map 檔不異動的資料
  DELETE CD007_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--更新 CD007(CodeNo,Description)
  for crCD007 in ccCD007 loop  --選出所有的代碼
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
          DBMS_OUTPUT.PUT_LINE('錯誤1: CD007代碼無法更新至CD007.CodeNo, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD007 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into cd007(CODENO,DESCRIPTION,SERVICETYPE,PRINTWONOW,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,0,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('錯誤2: CD007代碼無法新增至CD007.CodeNo, 新代碼='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('錯誤3: CD007代碼無法更新至CD007.CodeNo, 新代碼='||cr2.NewCode);
  	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    	ROLLBACK;
    	RETURN;
    end;
  end loop;



--更新 SO009(PRCode, PRName)
  for crSO009 in ccSO009 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD007_Map where OldCode=crSO009.PRCode;      
      begin
        update SO009 set PRCode=v_NewCode, PRName=v_NewName where Rowid=crSO009.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤5: CD007代碼無法更新至SO009.PRCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--更新 SO073(PRCode)
  for crSO073 in ccSO073 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD007_Map where OldCode=crSO073.PRCode;      
      begin
        update SO073 set PRCode=v_NewCode where Rowid=crSO073.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤6: CD007代碼無法更新至SO073.PRCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: 更新使用到停拆移機類別代碼檔(CD007)的相關檔案 ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


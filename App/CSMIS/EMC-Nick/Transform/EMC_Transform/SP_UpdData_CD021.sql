/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD021.SQL
  EXEC SP_UpdData_CD021
  更新使用到建築型態代碼檔(CD021)的相關檔案: 
	SO014(BTCode,BTName)
	SO016(BTCode,BTName)
	CD021(CodeNo,Description) 


  By: Jackal
  Date: 2002.11.21
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD021
  AS
--declare
  v_ChgCnt number := 0;

  v_OldCode number;
  v_OldName varchar2(20);
  v_NewCode number;
  v_NewName varchar2(20);
  v_Mark    number;

  cursor cc2 is
    select * from CD021_New order by NewCode;

  cursor ccCD021 is
    select Rowid,CodeNo from CD021 where CodeNo is not null;

  cursor ccSO014 is
    select Rowid,BTCode from SO014 where BTCode is not null;

  cursor ccSO016 is
    select Rowid,BTCode from SO016 where BTCode is not null;

begin

  DBMS_OUTPUT.PUT_LINE('*** 執行: UpdData_CD021.SQL');
--刪除CD021_Map 檔不異動的資料
  DELETE CD021_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--更新 CD021(CodeNo,Description)
  for crCD021 in ccCD021 loop  --選出所有的代碼
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD021_Map where OldCode=crCD021.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD021 WHERE Rowid=crCD021.Rowid;
        else
          update CD021 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD021.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤1: CD021代碼無法更新至CD021.CodeNo, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD021 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into CD021(CODENO,DESCRIPTION,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('錯誤2: CD021代碼無法新增至CD021.CodeNo, 新代碼='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('錯誤3: CD021代碼無法更新至CD021.CodeNo, 新代碼='||cr2.NewCode);
        DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	ROLLBACK;
  	RETURN;
    end;
  end loop;


--更新 SO014(BTCode,BTName)
  for crSO014 in ccSO014 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD021_Map where OldCode=crSO014.BTCode;      
      begin
        update SO014 set BTCode=v_NewCode, BTName=v_NewName where Rowid=crSO014.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤4: CD021代碼無法更新至SO014.BTCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO016(BTCode,BTName)
  for crSO016 in ccSO016 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD021_Map where OldCode=crSO016.BTCode;      
      begin
        update SO016 set BTCode=v_NewCode, BTName=v_NewName where Rowid=crSO016.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤5: CD021代碼無法更新至SO016.BTCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: 更新使用到建築型態代碼檔(CD021)的相關檔案 ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


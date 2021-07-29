/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD020.SQL
  EXEC SP_UpdData_CD020
  更新使用到付費意願代碼檔(CD020)的相關檔案: 
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

  DBMS_OUTPUT.PUT_LINE('*** 執行: UpdData_CD020.SQL');
--刪除CD020_Map 檔不異動的資料
  DELETE CD020_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;


--更新 CD020(CodeNo,Description)
  for crCD020 in ccCD020 loop  --選出所有的代碼
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
          DBMS_OUTPUT.PUT_LINE('錯誤1: CD020代碼無法更新至CD020.CodeNo, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD020 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into CD020(CODENO,DESCRIPTION,SERVICETYPE,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('錯誤2: CD020代碼無法新增至CD020.CodeNo, 新代碼='||cr2.NewCode);
	    NULL;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('錯誤3: CD020代碼無法更新至CD020.CodeNo, 新代碼='||cr2.NewCode);
  	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	ROLLBACK;
	RETURN;
    end;
  end loop;


--更新 SO001(PWCode, PWName)
  for crSO001 in ccSO001 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD020_Map where OldCode=crSO001.PWCode;      
      begin
        update SO001 set PWCode=v_NewCode, PWName=v_NewName where Rowid=crSO001.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤4: CD020代碼無法更新至SO001.PWCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
  	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--更新 SO044(PWCode)
  for crSO044 in ccSO044 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD020_Map where OldCode=crSO044.PWCode;      
      begin
        update SO044 set PWCode=v_NewCode where Rowid=crSO044.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤5: CD020代碼無法更新至SO044.PWCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: 更新使用到付費意願代碼檔(CD020)的相關檔案 ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


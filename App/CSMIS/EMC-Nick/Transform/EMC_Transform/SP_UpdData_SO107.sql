/*

  @C:\Transform\EMC_Transform\SP_UpdData_CD019.SQL
  EXEC SP_UpdData_CD019;
  更新使用到收費項目代碼檔(CD019)的相關檔案: 
	SO107(CitemName)
  By: Jackal
  Date: 2002.12.20
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

  cursor ccSO107 is
    select Rowid,CitemName from SO107 where CitemName is not null;


begin

  DBMS_OUTPUT.PUT_LINE('*** 執行: UpdData_CD019.SQL');
--刪除CD019_Map 檔不異動的資料
  DELETE CD019_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;



--更新 SO107(CitemName)
  for crSO107 in ccSO107 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD019_Map where OldCode=crSO107.CitemName;      
      begin
        update SO107 set CitemName=v_NewName where Rowid=crSO107.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤1: CD019代碼無法更新至SO107.CitemName, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: 更新使用到收費項目別代碼檔(CD019)的相關檔案 ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


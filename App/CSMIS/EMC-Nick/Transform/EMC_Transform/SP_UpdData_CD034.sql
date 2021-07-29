/*
  @C:\Transform\EMC_Transform\SP_UpdData_CD034.SQL
  EXEC SP_UpdData_CD034
  更新使用到買賣方式代碼檔(CD034)的相關檔案: 
	CD022(DefBuyCode, DefBuyName)
	CD022CD019(BuyCode)
	SO004(BuyCode, BuyName)
	SO055(BuyCode, BuyName, BuyCodeB, BuyNameB)
	SO105D(BuyCode)
	CD034(CodeNo,Description) 


  By: Jackal
  Date: 2002.11.22
*/
set serveroutput on
CREATE OR REPLACE PROCEDURE SP_UpdData_CD034
  AS
--declare
  v_ChgCnt number := 0;
  v_ServiceType  char(1) := 'C';
  v_StopFlag number := 0; 


  v_OldCode number;
  v_OldName varchar2(20);
  v_NewCode number;
  v_NewName varchar2(20);
  v_Mark    number;

  cursor cc2 is
    select * from CD034_New order by NewCode;

  cursor ccCD034 is
    select Rowid,CodeNo from CD034 where CodeNo is not null;

  cursor ccCD022 is
    select Rowid,DefBuyCode from CD022 where DefBuyCode is not null;

  cursor ccCD022CD019 is
    select Rowid,BuyCode from CD022CD019 where BuyCode is not null;

  cursor ccSO004 is
    select Rowid,BuyCode from SO004 where BuyCode is not null;

  cursor ccSO055_1 is
    select Rowid,BuyCode from SO055 where BuyCode is not null;

  cursor ccSO055_2 is
    select Rowid,BuyCodeB from SO055 where BuyCodeB is not null;

  cursor ccSO105D is
    select Rowid,BuyCode from SO105D where BuyCode is not null;

begin

  DBMS_OUTPUT.PUT_LINE('*** 執行: UpdData_CD034.SQL');
--刪除CD034_Map 檔不異動的資料
  DELETE CD034_MAP WHERE OLDCODE=NEWCODE AND OLDNAME=NEWNAME;



--更新 CD034(CodeNo,Description)
  for crCD034 in ccCD034 loop  --選出所有的代碼
    begin
      select OldCode,OldName,NewCode,NewName,Mark 
             into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark 
             from CD034_Map where OldCode=crCD034.CodeNo;
      BEGIN
        if v_Mark = -1 then
          Delete CD034 WHERE Rowid=crCD034.Rowid;
        else
          update CD034 set CodeNo=v_NewCode, Description=v_NewName where Rowid=crCD034.Rowid;
        end if;
      EXCEPTION 
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤1: CD034代碼無法更新至CD034.CodeNo, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
      SELECT COUNT(CODENO) AS COUNTS INTO v_ChgCnt FROM CD034 WHERE CodeNo=cr2.NewCode;

      if v_ChgCnt = 0 then
        begin
          insert into CD034(CODENO,DESCRIPTION,SERVICETYPE,STOPFLAG) 
            values(cr2.NewCode,cr2.NewName,v_ServiceType,cr2.Mark);
        exception
          when others then
            --DBMS_OUTPUT.PUT_LINE('錯誤2: CD034代碼無法新增至CD034.CodeNo, 新代碼='||cr2.NewCode);
	    null;
        end;
      end if; 	
    exception
      when others then
        DBMS_OUTPUT.PUT_LINE('錯誤3: CD034代碼無法更新至CD034.CodeNo, 新代碼='||cr2.NewCode);
   	DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	ROLLBACK;
  	RETURN;
    end;
  end loop;


--更新 CD022(DefBuyCode, DefBuyName) 
  for crCD022 in ccCD022 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD034_Map where OldCode=crCD022.DefBuyCode;      
      begin
        update CD022 set DefBuyCode=v_NewCode, DefBuyName=v_NewName where Rowid=crCD022.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤4: CD034代碼無法更新至CD022.DefBuyCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--更新 CD022CD019(BuyCode)
  for crCD022CD019 in ccCD022CD019 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD034_Map where OldCode=crCD022CD019.BuyCode;      
      begin
        update CD022CD019 set BuyCode=v_NewCode where Rowid=crCD022CD019.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤5: CD034代碼無法更新至CD022CD019.BuyCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO004(BuyCode, BuyName)
  for crSO004 in ccSO004 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD034_Map where OldCode=crSO004.BuyCode;      
      begin
        update SO004 set BuyCode=v_NewCode, BuyName=v_NewName where Rowid=crSO004.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤6: CD034代碼無法更新至SO004.BuyCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;



--更新 SO055(BuyCode, BuyName)
  for crSO055_1 in ccSO055_1 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD034_Map where OldCode=crSO055_1.BuyCode;      
      begin
        update SO055 set BuyCode=v_NewCode, BuyName=v_NewName where Rowid=crSO055_1.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤7: CD034代碼無法更新至SO055.BuyCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;




--更新 SO055(BuyCodeB, BuyNameB)
  for crSO055_2 in ccSO055_2 loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD034_Map where OldCode=crSO055_2.BuyCodeB;      
      begin
        update SO055 set BuyCodeB=v_NewCode, BuyNameB=v_NewName where Rowid=crSO055_2.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤8: CD034代碼無法更新至SO055.BuyCodeB, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
   	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
  	  ROLLBACK;
  	  RETURN;
      end;
    exception
      when others then
        null;
    end;
  end loop;


--更新 SO105D(BuyCode)
  for crSO105D in ccSO105D loop
    begin      
      select OldCode,OldName,NewCode,NewName,Mark into v_OldCode,v_OldName,v_NewCode,v_NewName,v_Mark from CD034_Map where OldCode=crSO105D.BuyCode;      
      begin
        update SO105D set BuyCode=v_NewCode where Rowid=crSO105D.Rowid;
      exception
        when others then
          DBMS_OUTPUT.PUT_LINE('錯誤9: CD034代碼無法更新至SO105D.BuyCode, 新代碼='||v_NewCode||', 舊代碼='||v_OldCode);
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
  DBMS_OUTPUT.PUT_LINE('*** OK: 更新使用到買賣方式代碼檔(CD034)的相關檔案 ***');

exception
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    ROLLBACK;
end;
/


/*
--@c:\EMC_script\SP_ReCntAllMdu
exec SP_ReCntAllMdu;  

  批次作業: 重算所有大樓的各種戶數欄位(建檔戶,有效戶,拆機戶...) [呼叫: SF_MduCompute.sql]
	. 對象? 所有大樓
	. 何時執行? 每月18日23:00
  檔名: SP_ReCntAllMdu.sql

  流程: 
  . Loop所有大樓資料
  	.. 呼叫SF_MduCompute( )

  傳回碼代碼說明v_RetCode:
	0:'重新計算完成',
	-1: '取公司代碼時發生錯誤'
	-2: '重算此大樓時發生錯誤('錯誤訊息'), 大樓編號=xxx, 大樓名稱=xxxxxxxx'
	-3: '新增時發生錯誤',
	-4: '取公司代碼發生錯誤',
	-5: '取公司名稱發生錯誤',
	-99: '其他錯誤'

  Date: 2003.07.18
  By: Pierre
*/

CREATE OR REPLACE PROCEDURE SP_ReCntAllMdu  IS
  v_RcdCount	number:=0;
  v_CompCode	number;
  v_CompName	varchar2(20);
  v_JobID	number;
  v_StartExecTime   Date;
  v_ServiceType     char(1);
  v_RetMsg	varchar2(100);
  v_RetCode	number;
  v_FuncName	varchar2(100);
  v_PrgName	varchar2(100);

  cursor cc1 is 
    select MduId, Name from SO017;

BEGIN
  -- *************************************************************
  -- 0. 設定參數初值
  -- *************************************************************
  v_StartExecTime := sysdate;	
  v_JobID	:= 999989;
  v_StartExecTime := sysdate;
  v_FuncName	:= '重算所有大樓各種戶數欄位';
  v_PrgName	:= 'SP_RECNTALLMDU';

  -- 取公司代碼
  begin 
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501a where rownum<=1;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      v_RetMsg := '取公司代碼時發生錯誤'||SQLERRM;
      v_RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      commit;
      return;
  end;

  -- *************************************************************
  -- 1. Loop每個大樓來重算各種戶數
  -- *************************************************************
  for cr1 in cc1 loop
    v_RetCode := SF_MduCompute(cr1.MduId, v_CompCode, v_ServiceType, v_RetMsg);
    if v_RetCode >= 0 then		-- 正常執行成功
      v_RcdCount := v_RcdCount + 1;
    else				-- 執行有誤
      v_RetMsg := '重算此大樓時發生錯誤('||v_RetMsg||'), 大樓編號='||cr1.MduId||', 大樓名稱='||cr1.Name;
      v_RetCode := -2;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
      -- commit;			-- 最後再commit
    end if;
  end loop;

  -- *************************************************************
  -- 2. Log執行結果
  -- *************************************************************
  v_RetMsg :='執行完畢: 重算'||v_RcdCount||'筆大樓之各種戶數, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'秒';
  v_RetCode :=0;
  DBMS_OUTPUT.PUT_LINE(v_RetMsg); 
  insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg); 
  COMMIT;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code='||SQLCODE||' '|| 'Error Message='||SQLERRM);
    v_RetMsg := '其他錯誤 :'||SQLERRM;
    v_RetCode := -99;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
    return;
END;
/

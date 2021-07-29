/*
-- @c:\emc_script\SP_CntStbLog2
-- @C:\Gird\V400\Script\SP_CntStbLog2
set serveroutput on
exec SP_CntStbLog2;

  統計每日STB開機次數/成功率/失敗率，並存至TABLE(SO510)
  檔名: SP_CntStbLog2.sql

  注意:
	1. 假設Send_Nagra檔案是在各資料庫的Com資料區中, 若不是, 則程式中的'COM.'字串要拿掉
	2. 統計結果檔結構(SO510, SO510A): 請參考相關文件

  SO030A錯誤訊息:
	 0:  '統計完畢'
	-1:  '取該指令成功次數有錯誤: '||SQLERRM;
	-2:  '取該指令失敗次數有錯誤: '||SQLERRM;
	-3:  '記錄統計結果於SO510有錯誤: '||SQLERRM;
	-4:  'Night-run參數檔SO501A無資料'
	-5:  '記錄統計結果於SO510A有錯誤: '||SQLERRM;
	-99: '其他錯誤: '||SQLERRM;

  By: Pierre
  Date: 2003.08.28
	2003.11.03 改: EMC資訊長92.10.31需求，希望此程式能記錄STB開機指令執行錯誤的明細資料
*/

create or replace procedure SP_CntStbLog2
as
  c39 char(1) := chr(39);
  v_CmdId varchar2(5);
  v_CmdName varchar2(20);
  v_CompCode number;
  v_Tmp1 varchar2(20);
  v_Tmp2 varchar2(20);
  v_RptDate date;

  v_OkCnt number := 0;
  v_ErrCnt number := 0;
  v_OkRate number := 0;
  v_ErrRate number := 0;
  v_TotalCnt number := 0;

  v_Operator 	varchar2(10);
  v_JobId 	number;
  v_RetCode 	number := 0;
  v_RetMsg 	varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;
  v_PrgName	varchar2(100);
  v_FuncName	varchar2(100);
  x_CompCode number := 0;
  v_CompCnt	number := 0;

  v_SQL varchar2(4000);
  type CurTyp is ref cursor;  --自訂cursor型態
  v_DyCursor CurTyp;          --供dynamic SQL

  cursor cc1 is 
    select distinct CompCode from so501A;

begin
  -- ********************************************************
  -- (1) 設定night-run所需參數
  -- ********************************************************
  v_JobId		:= 999988;		-- 開博內定Job編輯
  v_StartExecTime 	:= sysdate;		-- 開始執行時間
  v_Operator		:= 'Night-run';		-- Night-run都用此名稱
  v_PrgName		:= 'SP_CntStbLog2';	-- 程式名稱
  v_FuncName		:= '統計每日STB開機次數/成功率/失敗率';	-- 功能名稱

  v_CmdId := 'A1';			-- 開機指令代號
  v_RptDate := trunc(sysdate)-1;	-- 資料日: 為目前時間之前一日
  --v_RptDate := to_date('20030329', 'YYYYMMDD'); -- 測試用
  v_Tmp1 := GiPackage.ED2CDString(v_RptDate);	-- 每次統計一整天的v_CmdId指令的成功/失敗...等資料
  v_Tmp2 := GiPackage.ED2CDString(v_RptDate+1);

  -- 取高階指令名稱			-- 此段不是必要的
  if v_CmdId = 'A1' then
    v_CmdName := '開機';
  elsif v_CmdId = 'A2' then
    v_CmdName := '關機';
  elsif v_CmdId = 'A3' then
    v_CmdName := '停機';
  elsif v_CmdId = 'A4' then
    v_CmdName := '復機';
  elsif v_CmdId = 'E1' then
    v_CmdName := '設定親子密碼';
  else
    DBMS_OUTPUT.PUT_LINE('無此高階指令代號: '||v_CmdId);
    return;
  end if;

  -- ********************************************************
  -- (2) Loop每一公司, 來統計該公司的該指令執行情形
  -- ********************************************************
  for cr1 in cc1 loop
    x_CompCode := cr1.CompCode;
    v_CompCnt := v_CompCnt + 1;		-- 累計公司別
    v_TotalCnt := 0;
    v_OkCnt := 0;
    v_ErrCnt := 0;

    -- 取該指令成功次數
    begin
      select count(*) into v_OkCnt from SOAC0202 where UpdTime>=v_Tmp1 and UpdTime<v_Tmp2
	 and ModeType=v_CmdId and CompCode=x_CompCode;
    exception
      when others then
	v_RetCode := -1;
	v_RetMsg := '取該指令成功次數有錯誤: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;

    -- 取該指令失敗次數
    -- 假設Send_Nagra檔案是在各資料庫的Com資料區中, 若不是, 則程式中的'COM.'字串要拿掉
    begin
      --select count(*) into v_ErrCnt from Send_Nagra where CompCode=x_CompCode and 
      select count(*) into v_ErrCnt from COM.Send_Nagra where CompCode=x_CompCode and 
	UpdTime>=v_RptDate and UpdTime<v_RptDate+1 and HIGH_LEVEL_CMD_ID=v_CmdId AND Cmd_Status='E';
    exception
      when others then
	v_RetCode := -2;
	v_RetMsg := '取該指令失敗次數有錯誤: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;

    -- 計算指令總數, 成功率, 失敗率
    v_TotalCnt := v_OkCnt + v_ErrCnt;			-- 指令總數
    if v_TotalCnt = 0 then
      v_OkRate := 0;					-- 成功率
      v_ErrRate := 0;					-- 失敗率
    else
      v_OkRate := round(v_OkCnt*100/v_TotalCnt, 2);	-- 至小數點兩位四捨五入
      v_ErrRate := 100 - v_OkRate;
    end if;

    -- 記錄統計結果於SO510
    delete SO510 where CompCode=x_CompCode and RptDate=v_RptDate; -- 先刪除同公司同一天的統計資料
    begin
      insert into SO510 (RptDate, CompCode, ModeType, OkCnt, ErrCnt, TotalCnt, OkRate, ErrRate) values
	(v_RptDate, x_CompCode, v_CmdId, v_OkCnt, v_ErrCnt, v_TotalCnt, v_OkRate, v_ErrRate);
    exception
      when others then
	v_RetCode := -3;
	v_RetMsg := '記錄統計結果於SO510有錯誤: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;


    -- 記錄指令執行錯誤明細資料於SO510A		2003.11.03
    -- 假設Send_Nagra檔案是在各資料庫的Com資料區中, 若不是, 則程式中的'COM.'字串要拿掉
    delete SO510A where CompCode=x_CompCode and RptDate=v_RptDate; -- 先刪除同公司同一天的統計資料
    begin
      insert into SO510A (RPTDATE, COMPCODE, MODETYPE, SEQNO, ICC_NO, STB_NO, OPERATOR, UPDTIME, ERR_CODE, ERR_MSG)
	(select trunc(UPDTIME), COMPCODE, HIGH_LEVEL_CMD_ID, SEQNO, ICC_NO, STB_NO, OPERATOR, UPDTIME, 
	ERR_CODE, ERR_MSG from COM.Send_Nagra where CompCode=x_CompCode and UpdTime>=v_RptDate and UpdTime<v_RptDate+1 and 
	HIGH_LEVEL_CMD_ID=v_CmdId AND Cmd_Status='E');
	--from Send_Nagra where CompCode=x_CompCode and UpdTime>=v_RptDate and UpdTime<v_RptDate+1 and 
    exception
      when others then
	v_RetCode := -5;
	v_RetMsg := '記錄統計結果於SO510A有錯誤: '||SQLERRM;
	DBMS_OUTPUT.PUT_LINE(v_RetMsg);
        insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	  TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	  values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	  trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
        commit;
        return;
    end;


    -- log一筆記錄至Night-run記錄檔SO030A
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
	trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, 0, '統計完畢');
    commit;

  end loop;		-- Loop每一公司


  if v_CompCnt = 0 then -- 環境不對
    v_RetCode := -4;
    v_RetMsg := 'Night-run參數檔SO501A無資料';
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
      TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
      values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
      trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
  else
    DBMS_OUTPUT.PUT_LINE('統計完畢');
  end if;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    v_RetCode := -99;
    v_RetMsg := '其他錯誤: '||SQLERRM;
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
      TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
      values (v_JobId, x_CompCode, null, v_StartExecTime, sysdate, 
      trunc(86400*(sysdate-v_StartExecTime)), v_PrgName, v_FuncName, v_RetCode, v_RetMsg);
    commit;
end;
/
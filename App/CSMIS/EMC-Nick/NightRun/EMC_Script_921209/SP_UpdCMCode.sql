/*
  -- compile與測試指令
  @c:\emc_script\SP_UpdCMCode
  set serveroutput on
  exec SP_UpdCMCode;

  EMC收費方式代碼檔(CD031)統一化之資料整批調整程式, 此程式應該只執行一次, 由StartUpdCmCode_920807.sql
  來掛上Oracle night-run, 執行時間預定為2003.08.10 23:30

  檔名: SP_UpdCMCode.sql

  說明: 根據EMC黃斗南提供之收費方式代碼新舊對照表, 將資料庫中有用到收費方式代碼的檔案內容
	調整成新代碼與新名稱, 最後再調整收費方式代碼檔(CD031)內容. 
	所調整的資料檔list, 請參考此段之[調整內容]說明.

  OUT	p_RetCode number: 結果代碼
	p_RetMsg VARCHAR2： 結果訊息, 呼叫端變數至少需300 bytes

        >0: 資料批號, 表示正常完畢
	-1: p_RetMsg='參數錯誤'
	-2: p_RetMsg='取公司代碼時發生錯誤'||SQLERRM
        -99：p_RetMsg='其他錯誤'

  *** 待調SO 
	1=觀昇, 2=屏南, 3=南天, 5=新頻道, 6=豐盟, 7=振道, 8=全聯, 9=陽明山, 
	10=新台北, 11=金頻道, 12=大安文山, 13=新唐城, 14=大新店, 16=北桃園

  *** 調整內容 ***
     1. 調整檔案list (以下都含欄位CMCode, CMName)
	SO002, SO003, SO033, SO033DEBT, SO034, SO035, SO036, SO050, SO051, SO061, SO106
     2. 不調整檔案list (除了SO044只含CMCode, 以下都含欄位CMCode, CMName)
	SO032, SO044, SO067, SO074, SO077
     3. CD031 收費方式代碼檔內容全部SO統一, 但新唐城與大新店保留CodeNo=9的該筆資料

  *** 注意: 
     1. 若此支程式於Oracle 9i平台的SQL*plus上執行, 因為並無RBS27此rollback segment,
	則會出現以下的錯誤訊息, 但不影響程式的正常執行, 仍會調整資料, 請忽略這些訊息:
	SQLCODE=-1534 SQLERRM=ORA-01534: 倒回區段 'RBS27' 不存在
	SQLCODE=-1453 SQLERRM=ORA-01453: SET TRANSACTION 必須是交易中的第一個敘述句
	SQLCODE=-1534 SQLERRM=ORA-01534: 倒回區段 'RBS27' 不存在

  By: Pierre
  Date: 2002.08.07
*/

create or replace procedure SP_UpdCMCode
as
  p_RetCode number;
  p_RetMsg varchar2(500);

  v_StartExecTime	date;
  v_EndExecTime		date;
  c39			char(1) := chr(39);
  v_CompCode		number;
  v_ServiceType 	char(1);
  v_TranCnt		number := 0;
  j number;
  v_UpdTime		varchar2(20);
  v_FuncName		varchar2(100);
  v_PrgName		varchar2(100);
  v_JobID	number;

  v_SQL			varchar2(1000);
  v_User		varchar2(100);
  v_TmpRBS		varchar2(100);

  TYPE NumberAry IS TABLE OF number(3) INDEX BY BINARY_INTEGER;
  OldCodeNo NumberAry;		-- 舊代碼 array
  NewCodeNo NumberAry;		-- 新代碼 array

  TYPE VC2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  NewCodeName VC2Ary;		-- 新代碼名稱 array

begin
  v_JobID	:= 0;
  v_FuncName	:= 'EMC收費方式代碼檔(CD031)統一化之資料整批調整程式';
  v_PrgName	:= 'SP_UPDCMCODE';
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_UpdTime	:= GiPackage.GetDTString2(v_StartExecTime);
  v_TmpRBS	:= 'RBS27';

  -- ************************************************
  -- 取公司代碼, 服務別
  -- ************************************************
  begin 
    select CompCode, ServiceType into v_CompCode, v_ServiceType from so501a where rownum<=1;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      p_RetMsg := '取公司代碼時發生錯誤'||SQLERRM;
      p_RetCode := -1;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(v_EndExecTime-v_StartExecTime)), v_PrgName, v_FuncName, 
	p_RetCode, p_RetMsg);
      commit;
      return;
  end;

  -- ************************************************
  -- 根據公司別代碼, 設定新舊代碼對照關係array
  -- ************************************************
  /* 1=觀昇, 2=屏南, 3=南天, 5=新頻道, 6=豐盟, 7=振道, 8=全聯, 9=陽明山, 
     10=新台北, 11=金頻道, 12=大安文山, 13=新唐城, 14=大新店, 16=北桃園 */
  if v_CompCode=1 or v_CompCode=2 then		-- 1=觀昇, 2=屏南
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '不寄劃撥';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 50;	NewCodeName(2) := '代收';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 20;	NewCodeName(3) := '信用卡';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 45;	NewCodeName(4) := '裝置收';
    OldCodeNo(5) := 9;	NewCodeNo(5) := 41;	NewCodeName(5) := '營業部收';
    OldCodeNo(6) := 10;	NewCodeNo(6) := 49;	NewCodeName(6) := '介紹優待';
    OldCodeNo(7) := 11;	NewCodeNo(7) := 44;	NewCodeName(7) := '工程裝置';
    OldCodeNo(8) := 12;	NewCodeNo(8) := 42;	NewCodeName(8) := '工程代收';
    OldCodeNo(9) := 13;	NewCodeNo(9) := 48;	NewCodeName(9) := '查緝私接';
    v_TranCnt := 9;
   
  elsif v_CompCode=3 then	-- 3=南天
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '代收';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '信用卡';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 200;	NewCodeName(3) := '金融';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 42;	NewCodeName(4) := '工程代收';
    v_TranCnt := 4;

  elsif v_CompCode=5 then	-- 5=新頻道
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '代收';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '信用卡';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 205;	NewCodeName(3) := '中華商銀';
    v_TranCnt := 3;

   elsif v_CompCode=6 then	-- 6=豐盟
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '代收';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '信用卡';
    v_TranCnt := 2;

  elsif v_CompCode=7 then	-- 7=振道
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '不寄劃撥';
    OldCodeNo(2) := 7;	NewCodeNo(2) := 32;	NewCodeName(2) := '媒體劃撥';
    OldCodeNo(3) := 8;	NewCodeNo(3) := 20;	NewCodeName(3) := '信用卡';
    v_TranCnt := 3;

  elsif v_CompCode=8 then	-- 8=全聯
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '不寄劃撥';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '信用卡';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 42;	NewCodeName(3) := '工程代收';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 46;	NewCodeName(4) := '手開單據';
    OldCodeNo(5) := 9;	NewCodeNo(5) := 47;	NewCodeName(5) := '專案';
    OldCodeNo(6) := 10;	NewCodeNo(6) := 43;	NewCodeName(6) := '移機順收';
    OldCodeNo(7) := 11;	NewCodeNo(7) := 101;	NewCodeName(7) := '7-11';
    OldCodeNo(8) := 12;	NewCodeNo(8) := 102;	NewCodeName(8) := '萊爾富';
    OldCodeNo(9) := 13;	NewCodeNo(9) := 103;	NewCodeName(9) := '全家';
    OldCodeNo(10) := 14; NewCodeNo(10) := 201;	NewCodeName(10) := '郵局';
    OldCodeNo(11) := 15; NewCodeNo(11) := 202;	NewCodeName(11) := '第一銀行';
    OldCodeNo(12) := 16; NewCodeNo(12) := 203;	NewCodeName(12) := '合作金庫';
    OldCodeNo(13) := 17; NewCodeNo(13) := 204;	NewCodeName(13) := '玉山銀行';
    OldCodeNo(14) := 18; NewCodeNo(14) := 104;	NewCodeName(14) := 'OK';
    OldCodeNo(15) := 19; NewCodeNo(15) := 105;	NewCodeName(15) := '福客多';
    v_TranCnt := 15;

  elsif v_CompCode=9 or v_CompCode=10 or v_CompCode=11 or v_CompCode=12 then	-- 北平台9~12
    OldCodeNo(1) := 5;	NewCodeNo(1) := 31;	NewCodeName(1) := '不寄劃撥';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 101;	NewCodeName(2) := '7-11';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 32;	NewCodeName(3) := '媒體劃撥';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 20;	NewCodeName(4) := '信用卡';
    v_TranCnt := 4;

  elsif v_CompCode=13 or v_CompCode=14 then	-- 13=新唐城, 14=大新店
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '代收';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '信用卡';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 100;	NewCodeName(3) := '便利商店代收';
    OldCodeNo(4) := 8;	NewCodeNo(4) := 200;	NewCodeName(4) := '銀行代收';
    OldCodeNo(5) := 10;	NewCodeNo(5) := 21;	NewCodeName(5) := '信卡授權';
    v_TranCnt := 5;

  elsif v_CompCode=16 then	-- 16=北桃園
    OldCodeNo(1) := 5;	NewCodeNo(1) := 50;	NewCodeName(1) := '代收';
    OldCodeNo(2) := 6;	NewCodeNo(2) := 20;	NewCodeName(2) := '信用卡';
    OldCodeNo(3) := 7;	NewCodeNo(3) := 32;	NewCodeName(3) := '媒體劃撥';
    v_TranCnt := 3;

  else
    p_RetMsg := '無此公司別: ' || v_CompCode;
    p_RetCode := -2;
    v_TranCnt := 0;
    return;
  end if;


  -- ************************************************************
  -- 給其alter rollback segment 權限
  begin
    select user into v_User from dual;
    --DBMS_OUTPUT.PUT_LINE('User = ' || v_User);
    v_SQL := 'GRANT ALTER ROLLBACK SEGMENT TO ' || v_User;
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;

  -- enable rollback segment RBS27
  begin
    v_SQL := 'alter rollback segment '||v_TmpRBS||' online';
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;

  -- 指定transaction至該rollback segment
  begin
    set transaction use rollback segment rbS6;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      p_RetMsg := '無法啟用RBS27: '||SQLERRM;
  end;
  -- ************************************************************


  -- ************************************************************
  -- 將資料庫中有用到收費方式代碼的檔案內容, 調整成新代碼與新名稱
  -- ************************************************************
  for j in 1..v_TranCnt loop
    update SO002 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO003 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO033 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO033Debt set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO034 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO035 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO036 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO050 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO051 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO051 set CMCodeB=NewCodeNo(j), CMNameB=NewCodeName(j) where CMCodeB=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO061 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;
  for j in 1..v_TranCnt loop
    update SO106 set CMCode=NewCodeNo(j), CMName=NewCodeName(j) where CMCode=OldCodeNo(j);
  end loop;


  -- ************************************************************
  -- 調整收費方式代碼檔(CD031)內容
  --	1. 刪除不用的代碼段落
  --	2. 新增新共用代碼
  -- ************************************************************
  --	1. 刪除不用的代碼段落
  if v_CompCode=13 or v_CompCode=14 then	-- 13=新唐城, 14=大新店
    delete from CD031 where CodeNo>=5 and CodeNo!=9;
  else
    delete from CD031 where CodeNo>=5;
  end if;

  --	2. 新增CD031新共用代碼
  begin
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(20, '信用卡', 4, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(21, '信卡授權', 4, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(30, '支票', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(31, '不寄劃撥', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(32, '媒體劃撥', null, v_UpdTime, '系統自動', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(41, '營業部收', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(42, '工程代收', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(43, '移機順收', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(44, '工程裝置', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(45, '裝置收', null, v_UpdTime, '系統自動', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(46, '手開單據', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(47, '專案', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(48, '查緝私接', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(49, '介紹優待', null, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(50, '代收', null, v_UpdTime, '系統自動', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(100, '便利商店代收', 3, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(101, '7-11', 3, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(102, '萊爾富', 3, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(103, '全家', 3, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(104, 'OK', 3, v_UpdTime, '系統自動', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(105, '福客多', 3, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(200, '銀行代收', 2, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(201, '郵局', 2, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(202, '第一銀行', 2, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(203, '合作金庫', 2, v_UpdTime, '系統自動', 'C', 0);

    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(204, '玉山銀行', 2, v_UpdTime, '系統自動', 'C', 0);
    insert into CD031 (CodeNo, Description, RefNo, UpdTime, UpdEn, ServiceType, StopFlag) values 
	(205, '中華商銀', 2, v_UpdTime, '系統自動', 'C', 0);
  
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('ErrCode='||SQLCODE||' '||'ErrMsg='||SQLERRM);
      p_RetMsg := '新增CD031新共用代碼錯誤: '||SQLERRM;
      p_RetCode := -2;
      insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, sysdate, 
	trunc(86400*(v_EndExecTime-v_StartExecTime)), v_PrgName, v_FuncName, 
	p_RetCode, p_RetMsg);
      commit;
      --return; 	-- 不要return
  end;


  -- ************************************************************
  v_EndExecTime := sysdate;		-- 執行結束時間
  p_RetMsg := '調整完畢, 共花費'||to_char(trunc(86400*(v_EndExecTime-v_StartExecTime)))||'秒';
  p_RetCode := 0;
  commit;


  -- ************************************************************
  -- disable該rollback segment
  begin
    v_SQL := 'alter rollback segment '||v_TmpRBS||' offline';
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;

  -- 取消其alter rollback segment 權限
  begin
    v_SQL := 'REVOKE ALTER ROLLBACK SEGMENT FROM ' || v_User;
    DBMS_UTILITY.EXEC_DDL_STATEMENT(v_SQL);
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
  end;
  -- ************************************************************


  -- ************************************************************
  -- Log一筆執行成功記錄至SO030A
  -- ************************************************************
  begin
    insert into SO030A (JobId, CompCode, ServiceType, BeginExecTime, EndExecTime, 
	TotalExecTime, PrgName, FuncName, ReturnCode, ReturnMsg)
	values (v_JobID, v_CompCode, v_ServiceType, v_StartExecTime, v_EndExecTime, 
	trunc(86400*(v_EndExecTime-v_StartExecTime)), v_PrgName, v_FuncName, 
	p_RetCode, p_RetMsg); 
    commit;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      rollback;
      p_RetMsg := 'Log至SO030A錯誤: '||SQLERRM;
      p_RetCode := -98;
  end;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    p_RetMsg := '其他錯誤: '||SQLERRM;
    p_RetCode := -99;
end;
/
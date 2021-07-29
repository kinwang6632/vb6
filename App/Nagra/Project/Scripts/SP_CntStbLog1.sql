/*
@C:\Gird\V300\Script\SP_CntStbLog1

set serveroutput on
exec SP_CntStbLog1(99, 'A1', '20030201', '20030218', 0);

--exec SP_CntStbLog1(99, 'A1', '20030201', '20030218', 1);

  統計某系統台於某段時間內, 開機或關機等某種高階指令的次數, 成功率, 失敗率
  檔名: SP_CntStbLog1.sql

  IN	p_CompCode number: 公司別
	p_CmdId varchar2: 高階指令代號, 
		'A1'=開機, 'A2'=關機, 'A3'=停機, 'A4'=復機, 'E1'=設定親子密碼
	p_D1 varchar2: 統計起始日期, 西曆'YYYYMMDD'
	p_D2 varchar2: 統計截止日期, 西曆'YYYYMMDD'
	p_Option number: 0=忽略起迄日期範圍, 統計所有log資料. 1=依日期範圍統計

  注意:
  1. 直接login至各家SO資料區中執行此程式即可, 但需注意公司別代碼要傳入
  2. 需啟動SQL*Plus的顯示開關: set serveroutput on
  3. 假設Send_Nagra檔案是在各資料庫的Com資料區中, 若不是, 則程式中的'COM.'字串要拿掉

  By: Pierre
  Date: 2003.02.19
	2003.02.20 加: 設定親子密碼(指令E1)的統計
	2003.03.13 加: 多一參數"公司別"

*/

create or replace procedure SP_CntStbLog1(p_CompCode number, p_CmdId varchar2, p_D1 varchar2, p_D2 varchar2,
	p_Option number)
as
  v_CmdName varchar2(20);

  v_CompCode number;
  v_CompName varchar2(20);

  v_OkCnt number := 0;
  v_ErrCnt number := 0;
  v_OkRate number := 0;
  v_ErrRate number := 0;
  v_TotalCnt number := 0;

  v_Tmp1 varchar2(20);
  v_Tmp2 varchar2(20);

begin
  if p_Option!=0 and p_Option!=1 then
    DBMS_OUTPUT.PUT_LINE('參數p_Option值需為0或1');
    return;
  end if;

  -- 取高階指令名稱
  if p_CmdId = 'A1' then
    v_CmdName := '開機';
  elsif p_CmdId = 'A2' then
    v_CmdName := '關機';
  elsif p_CmdId = 'A3' then
    v_CmdName := '停機';
  elsif p_CmdId = 'A4' then
    v_CmdName := '復機';
  elsif p_CmdId = 'E1' then
    v_CmdName := '設定親子密碼';
  else
    DBMS_OUTPUT.PUT_LINE('無此高階指令代號: '||p_CmdId);
    return;
  end if;

  -- 取公司別與名稱
  v_CompCode := p_CompCode;
  begin
    select Description into v_CompName from CD039 where CodeNo=v_CompCode;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('CD039無此公司別代碼: '||v_CompCode);
      return;
  end;

  -- 取該指令成功次數 (註: 日期將轉成中曆字串)
  begin
    if p_Option=0 then
      select count(*) into v_OkCnt from SOAC0202 where CompCode=v_CompCode and 
	ModeType=p_CmdId;
    else
      v_Tmp1 := GiPackage.ED2CDString(to_date(p_D1,'YYYYMMDD'));
      v_Tmp2 := GiPackage.ED2CDString(to_date(p_D2,'YYYYMMDD'))||' 23:59:59';
      select count(*) into v_OkCnt from SOAC0202 where CompCode=v_CompCode and 
	UpdTime>=v_Tmp1 and UpdTime<=v_Tmp2 and ModeType=p_CmdId;
    end if;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('取該指令成功次數有錯誤: ');
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      return;
  end;

  -- 取該指令失敗次數
  -- 假設Send_Nagra檔案是在各資料庫的Com資料區中, 若不是, 則程式中的'COM.'字串要拿掉
  begin
    if p_Option=0 then
      select count(*) into v_ErrCnt from COM.Send_Nagra where CompCode=v_CompCode and 
	HIGH_LEVEL_CMD_ID=p_CmdId AND Cmd_Status='E';
    else
      select count(*) into v_ErrCnt from COM.Send_Nagra where CompCode=v_CompCode and 
	UpdTime>=to_date(p_D1,'YYYYMMDD') and UpdTime<to_date(p_D2,'YYYYMMDD')+1 and 
	HIGH_LEVEL_CMD_ID=p_CmdId AND Cmd_Status='E';
    end if;
  exception
    when others then
      DBMS_OUTPUT.PUT_LINE('取該指令失敗次數有錯誤: ');
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      return;
  end;

  -- 計算指令總數, 成功率, 失敗率
  v_TotalCnt := v_OkCnt + v_ErrCnt;			-- 指令總數
  if v_TotalCnt = 0 then
    v_OkRate := 0;					-- 成功率
    v_ErrRate := 0;					-- 失敗率
  else
    v_OkRate := round(v_OkCnt*100/v_TotalCnt, 2);
    v_ErrRate := 100 - v_OkRate;
  end if;

  -- 顯示結果
  DBMS_OUTPUT.PUT_LINE('*** 高階指令統計結果 ***');
  if p_Option=0 then
    DBMS_OUTPUT.PUT_LINE('統計期間: 從開始log到現在 '||to_char(sysdate,'YYYY/MM/DD HH24:MI:SS'));
  else
    DBMS_OUTPUT.PUT_LINE('統計期間: '||p_D1||' 至 '||p_D2);
  end if;
  DBMS_OUTPUT.PUT_LINE('指令代號: '||p_CmdId||'  ('||v_CmdName||')');
  DBMS_OUTPUT.PUT_LINE('公司別  : '||v_CompCode||', 公司名稱: '||v_CompName);
  DBMS_OUTPUT.PUT_LINE('成功次數: '||v_OkCnt);
  DBMS_OUTPUT.PUT_LINE('失敗次數: '||v_ErrCnt);
  DBMS_OUTPUT.PUT_LINE('指令總數: '||v_TotalCnt);
  DBMS_OUTPUT.PUT_LINE('成功率  : '||v_OkRate);
  DBMS_OUTPUT.PUT_LINE('失敗率  : '||v_ErrRate);

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
end;
/